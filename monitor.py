"""
Acompanhamento da impressora/scanner pela rede (protocolo eSCL, usado pelo AirPrint).
Informa se o scanner está livre ou ocupado e o estado do alimentador (com papel,
vazio, papel enroscado, tampa aberta...), para o programa só puxar a próxima
folha quando o equipamento estiver pronto.
"""
import concurrent.futures
import ssl
import threading
import time
import urllib.request
import xml.etree.ElementTree as ET

import pythoncom
import win32com.client

REQUEST_TIMEOUT = 3

ADF_PROBLEMS = {
    "ScannerAdfJam": "Papel enroscado no alimentador. Retire o papel na impressora.",
    "ScannerAdfHatchOpen": "A tampa do alimentador está aberta. Feche-a.",
    "ScannerAdfDoorOpen": "A tampa do alimentador está aberta. Feche-a.",
    "ScannerAdfMispick": "O alimentador não conseguiu puxar a folha. Ajeite o papel.",
    "ScannerAdfMultipickDetected": "O alimentador puxou mais de uma folha. Ajeite o papel.",
    "ScannerAdfInputTrayOverloaded": "Folhas demais no alimentador. Retire algumas.",
    "ScannerAdfInputTrayFailed": "Falha na bandeja do alimentador.",
    "ScannerAdfDuplexPageTooShort": "Folha curta demais para frente e verso.",
    "ScannerAdfDuplexPageTooLong": "Folha longa demais para frente e verso.",
}
STATE_PROBLEMS = {
    "Stopped": "A impressora está parada com erro. Verifique o painel.",
    "Down": "A impressora está com problema. Verifique o painel.",
}


class Status:
    def __init__(self, online, state=None, adf=None):
        self.online = online
        self.state = state      # Idle, Processing, Testing, Stopped, Down
        self.adf = adf          # ScannerAdfLoaded, ScannerAdfEmpty, ScannerAdfJam...
        self.time = time.monotonic()

    @property
    def problem(self):
        return ADF_PROBLEMS.get(self.adf) or STATE_PROBLEMS.get(self.state)

    @property
    def busy(self):
        return self.state not in ("Idle", None) or self.adf == "ScannerAdfProcessing"

    def describe(self):
        """(texto, cor) para mostrar na tela."""
        if not self.online:
            return "sem conexão", "#868e96"
        if self.problem:
            return self.problem, "#e03131"
        printer = "Ocupada" if self.busy else "Livre"
        paper = {"ScannerAdfLoaded": "com papel", "ScannerAdfEmpty": "vazio",
                 "ScannerAdfProcessing": "puxando folha"}.get(self.adf, self.adf or "?")
        color = "#f08c00" if self.busy else ("#2f9e44" if self.adf == "ScannerAdfLoaded" else "#1c7ed6")
        return f"{printer}  ·  Alimentador: {paper}", color


def _get(url):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(url, timeout=REQUEST_TIMEOUT, context=ctx) as r:
        return r.read()


def _find(root, name):
    for el in root.iter():
        if el.tag.rsplit("}", 1)[-1] == name:
            return (el.text or "").strip()
    return None


def fetch_status(ip):
    try:
        root = ET.fromstring(_get(f"http://{ip}/eSCL/ScannerStatus"))
        return Status(True, _find(root, "State"), _find(root, "AdfState"))
    except Exception:
        return Status(False)


def probe(ip):
    """Retorna (ip, modelo, uuid) se o endereço for um scanner eSCL, senão None."""
    try:
        root = ET.fromstring(_get(f"http://{ip}/eSCL/ScannerCapabilities"))
        return ip, _find(root, "MakeAndModel") or ip, (_find(root, "UUID") or "").lower()
    except Exception:
        return None


def printer_ips():
    """Endereços IP das impressoras de rede instaladas no Windows.
    Roda numa thread própria para não mexer no COM de quem chamou."""
    result = []

    def work():
        pythoncom.CoInitialize()
        try:
            wmi = win32com.client.GetObject(r"winmgmts:root\cimv2")
            ips = {p.HostAddress for p in wmi.ExecQuery("SELECT HostAddress FROM Win32_TCPIPPrinterPort")}
            result.extend(sorted(ip for ip in ips if ip))
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()

    t = threading.Thread(target=work, daemon=True)
    t.start()
    t.join(15)
    return result


def discover(wia_port=""):
    """Procura o scanner na rede. Se o dispositivo WIA for WSD (porta com urn:uuid),
    escolhe o equipamento com o mesmo UUID; senão, o primeiro que responder."""
    ips = printer_ips()
    if not ips:
        return None
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        found = [r for r in pool.map(probe, ips) if r]
    if not found:
        return None
    port = (wia_port or "").lower()
    for ip, model, uuid in found:
        if uuid and uuid in port:
            return ip, model
    return found[0][:2]


class PrinterMonitor:
    """Consulta o estado da impressora periodicamente em segundo plano."""

    def __init__(self):
        self.ip = None
        self.status = Status(False)
        self.fast = False
        self._stop = threading.Event()
        threading.Thread(target=self._loop, daemon=True).start()

    def set_ip(self, ip):
        self.ip = (ip or "").strip() or None
        self.status = Status(False)

    def _loop(self):
        while not self._stop.is_set():
            ip = self.ip
            if ip:
                st = fetch_status(ip)
                if ip == self.ip:
                    self.status = st
            self._stop.wait(0.7 if self.fast else 2.5)

    def close(self):
        self._stop.set()

    def wait_ready(self, stop_event, notify, check_adf=True, max_busy=60):
        """Espera o scanner ficar livre para puxar a próxima folha.
        Retorna 'ready', 'empty', 'stopped' ou None (acompanhamento indisponível).
        Enquanto houver papel enroscado/tampa aberta, espera indefinidamente
        (até o usuário resolver ou clicar em Parar)."""
        ip = self.ip
        if not ip:
            return None
        start = time.monotonic()
        idle_since = None
        last_msg = None
        failures = 0

        def say(msg):
            nonlocal last_msg
            if msg != last_msg:
                notify(msg)
                last_msg = msg

        while not stop_event.is_set():
            st = fetch_status(ip)
            self.status = st
            now = time.monotonic()
            if not st.online:
                failures += 1
                if failures >= 3:
                    return None
                time.sleep(0.5)
                continue
            failures = 0
            problem = st.problem if check_adf else STATE_PROBLEMS.get(st.state)
            if problem:
                say(problem + "  O programa continua sozinho quando for resolvido.")
                idle_since = None
                start = now
            elif st.state == "Idle" and not (check_adf and st.adf == "ScannerAdfProcessing"):
                if check_adf and st.adf == "ScannerAdfEmpty":
                    return "empty"
                idle_since = idle_since or now
                if now - idle_since >= 0.8:
                    return "ready"
            else:
                idle_since = None
                say("Aguardando a impressora terminar a folha anterior...")
                if now - start > max_busy:
                    return "ready"
            time.sleep(0.4)
        return "stopped"
