"""
Digitalização direta pela rede (protocolo eSCL / AirPrint Scan).
Um único trabalho puxa a pilha inteira do alimentador e a impressora entrega
as páginas uma a uma, sem perder folhas entre um pedido e outro.
"""
import time
import urllib.error
import urllib.request
from xml.sax.saxutils import escape

from monitor import fetch_status

ADF_MAX_DPI = 300
FEEDER_MAX = (2550, 4200)   # 1/300 de polegada
PLATEN_MAX = (2550, 3508)
PAGE_TIMEOUT = 180          # segundos esperando cada página ser escaneada


class EsclError(Exception):
    pass


class EsclPaperProblem(EsclError):
    """Trabalho interrompido por papel enroscado, tampa aberta etc."""


def _request(method, url, data=None, timeout=30, headers=None):
    req = urllib.request.Request(url, data=data, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers, e.read()


def settings_xml(feeder, duplex, color_mode, dpi, width, height):
    source = "Feeder" if feeder else "Platen"
    duplex_tag = f"<scan:Duplex>{'true' if duplex else 'false'}</scan:Duplex>" if feeder else ""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<scan:ScanSettings xmlns:scan="http://schemas.hp.com/imaging/escl/2011/05/03" xmlns:pwg="http://www.pwg.org/schemas/2010/12/sm">
  <pwg:Version>2.63</pwg:Version>
  <scan:Intent>Document</scan:Intent>
  <pwg:ScanRegions>
    <pwg:ScanRegion>
      <pwg:ContentRegionUnits>escl:ThreeHundredthsOfInches</pwg:ContentRegionUnits>
      <pwg:XOffset>0</pwg:XOffset>
      <pwg:YOffset>0</pwg:YOffset>
      <pwg:Width>{width}</pwg:Width>
      <pwg:Height>{height}</pwg:Height>
    </pwg:ScanRegion>
  </pwg:ScanRegions>
  <pwg:InputSource>{source}</pwg:InputSource>
  {duplex_tag}
  <scan:ColorMode>{escape(color_mode)}</scan:ColorMode>
  <scan:XResolution>{dpi}</scan:XResolution>
  <scan:YResolution>{dpi}</scan:YResolution>
  <pwg:DocumentFormat>image/jpeg</pwg:DocumentFormat>
  <scan:DocumentFormatExt>image/jpeg</scan:DocumentFormatExt>
</scan:ScanSettings>"""


def _job_result(ip, job_url):
    """(estado, motivos) do trabalho, conforme o histórico da impressora."""
    uuid = job_url.rstrip("/").rsplit("/", 1)[-1]
    try:
        import xml.etree.ElementTree as ET
        with urllib.request.urlopen(f"http://{ip}/eSCL/ScannerStatus", timeout=5) as r:
            root = ET.fromstring(r.read())
        for job in root.iter():
            if job.tag.endswith("JobInfo"):
                vals = {c.tag.rsplit("}", 1)[-1]: (c.text or "").strip() for c in job}
                if vals.get("JobUuid") == uuid or uuid in vals.get("JobUri", ""):
                    reasons = [r.text for r in job.iter() if r.tag.endswith("JobStateReason")]
                    return vals.get("JobState"), reasons
    except Exception:
        pass
    return None, []


def _problem_message(ip):
    st = fetch_status(ip)
    return st.problem if st.online else None


def _start_job(ip, xml, feeder, notify, stop_event):
    """Cria o trabalho de digitalização. Retorna a URL do trabalho, ou None se o alimentador estiver vazio."""
    waited = 0
    while not stop_event.is_set():
        try:
            status, headers, body = _request("POST", f"http://{ip}/eSCL/ScanJobs", xml.encode("utf-8"),
                                             headers={"Content-Type": "text/xml"})
        except OSError as e:
            raise EsclError(f"Sem comunicação com a impressora ({ip}). Verifique a rede.") from e
        if status == 201:
            loc = headers.get("Location", "")
            if loc.startswith("/"):
                loc = f"http://{ip}{loc}"
            return loc.rstrip("/")
        st = fetch_status(ip)
        if feeder and st.online and st.adf == "ScannerAdfEmpty":
            return None
        if status in (409, 503) and waited < 60:
            notify((st.problem if st.online and st.problem else "Impressora ocupada, aguardando...")
                   + "  O programa continua sozinho.")
            time.sleep(2)
            waited += 2 if not (st.online and st.problem) else 0
            continue
        detail = body.decode("utf-8", "ignore").strip()[:200]
        raise EsclError(f"A impressora recusou a digitalização (HTTP {status}). {detail}")
    return None


def scan(ip, feeder, duplex, color_mode, dpi, paper_inches, deliver, notify, stop_event):
    """Digitaliza tudo o que estiver no alimentador (ou uma folha do vidro).
    deliver(bytes_jpeg) é chamado para cada página.
    Retorna 'done', 'empty' ou 'stopped'."""
    max_w, max_h = FEEDER_MAX if feeder else PLATEN_MAX
    if feeder and dpi > ADF_MAX_DPI:
        notify(f"O alimentador aceita até {ADF_MAX_DPI} DPI; usando {ADF_MAX_DPI}.")
        dpi = ADF_MAX_DPI
    w_in, h_in = paper_inches or (8.27, 11.69)
    xml = settings_xml(feeder, duplex, color_mode, dpi, min(int(w_in * 300), max_w), min(int(h_in * 300), max_h))

    job = _start_job(ip, xml, feeder, notify, stop_event)
    if job is None:
        return "stopped" if stop_event.is_set() else "empty"

    count = 0
    finished = False
    try:
        while not stop_event.is_set():
            try:
                status, headers, body = _request("GET", job + "/NextDocument", timeout=PAGE_TIMEOUT)
            except OSError as e:
                raise EsclError("A comunicação com a impressora caiu durante a digitalização.") from e
            if status == 200:
                count += 1
                deliver(body)
            elif status == 503:          # página ainda sendo escaneada
                time.sleep(0.5)
            elif status == 404:          # não há mais páginas
                finished = True
                break
            else:
                raise EsclError(f"Erro ao receber a página {count + 1} (HTTP {status}).")
    finally:
        if not finished:
            try:
                _request("DELETE", job, timeout=10)
            except OSError:
                pass

    if stop_event.is_set():
        return "stopped"
    state, reasons = _job_result(ip, job)
    if state in ("Aborted", "Canceled") or (state is None and count == 0):
        problem = _problem_message(ip)
        if problem:
            raise EsclPaperProblem(problem)
        if state:
            raise EsclError(f"A impressora interrompeu a digitalização ({', '.join(reasons) or state}).")
    return "done" if count else "empty"
