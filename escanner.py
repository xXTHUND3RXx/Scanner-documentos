"""
Escanner em Lote
Digitalização contínua de grandes volumes pelo alimentador automático (ADF),
sem limite de folhas. Usa o WIA do Windows e salva em PDF ou imagens.
"""
import json
import os
import queue
import shutil
import sys
import threading
import time
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from PIL import Image, ImageTk, ImageStat
import pythoncom
import pywintypes
import win32com.client
import img2pdf

from orientacao import detect_rotation, OCR_AVAILABLE
import monitor as M
import escl

APP_NAME = "Escanner em Lote"
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "EscannerLote")
SESSION_DIR = os.path.join(DATA_DIR, "sessao")
ORDER_FILE = os.path.join(SESSION_DIR, "ordem.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")

THUMB_SIZE = 130


def resource(name):
    """Caminho de um arquivo que acompanha o programa (funciona também dentro do .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

# ---------------------------------------------------------------- WIA
WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
WIA_DEVICE_TYPE_SCANNER = 1

DPS_DOCUMENT_HANDLING_SELECT = 3088
DPS_PAGES = 3096
IPS_CUR_INTENT = 6146
IPS_XRES = 6147
IPS_YRES = 6148
IPS_XPOS = 6149
IPS_YPOS = 6150
IPS_XEXTENT = 6151
IPS_YEXTENT = 6152
IPA_DATATYPE = 4103
IPA_DEPTH = 4104

FEEDER, FLATBED, DUPLEX = 1, 2, 4

WIA_ERROR_PAPER_EMPTY = 0x80210003
WIA_ERRORS = {
    0x80210001: "Erro geral no scanner.",
    0x80210002: "Papel enroscado no alimentador. Remova o papel e tente novamente.",
    0x80210003: "Não há papel no alimentador. Coloque as folhas e tente novamente.",
    0x80210004: "Problema com o papel no alimentador.",
    0x80210005: "O scanner está desligado ou desconectado.",
    0x80210006: "O scanner está ocupado. Aguarde e tente novamente.",
    0x80210007: "O scanner está aquecendo. Aguarde alguns segundos.",
    0x80210008: "O scanner precisa de atenção (verifique o painel do equipamento).",
    0x8021000A: "Falha de comunicação com o scanner (verifique rede/cabo).",
    0x8021000C: "Configuração não suportada por este scanner. Tente outra resolução ou origem.",
    0x8021000D: "O scanner está em uso por outro programa.",
    0x8021000E: "Erro no driver do scanner.",
    0x80210016: "A tampa do scanner está aberta.",
    0x80210017: "A lâmpada do scanner está desligada.",
}

SOURCES = {
    "Alimentador (ADF) - só frente": FEEDER,
    "Alimentador (ADF) - frente e verso": FEEDER | DUPLEX,
    "Vidro (mesa) - uma folha": FLATBED,
}
COLORS = {  # nome: (intent, datatype, depth)
    "Colorido": (1, 3, 24),
    "Tons de cinza": (2, 2, 8),
    "Preto e branco (texto)": (4, 0, 1),
}
ESCL_COLORS = {  # no P&B pedimos cinza e convertemos aqui, com limiar próprio para texto
    "Colorido": "RGB24",
    "Tons de cinza": "Grayscale8",
    "Preto e branco (texto)": "Grayscale8",
}
NETWORK_SUFFIX = " — pela rede (recomendado)"
RESOLUTIONS = ["150", "200", "300", "400", "600"]
PAPERS = {  # polegadas (largura, altura)
    "A4": (8.27, 11.69),
    "Carta": (8.5, 11.0),
    "Ofício": (8.5, 14.0),
    "Padrão do scanner": None,
}


# Erros temporários: o scanner ainda está terminando a folha anterior.
RETRY_ERRORS = {0x80210006, 0x80210007, 0x8021000D}  # ocupado, aquecendo, em uso
BUSY_RECONNECT_AFTER = 8  # segundos esperando antes de reconectar ao scanner
BUSY_TIMEOUT = 45         # segundos esperando antes de desistir

# Com o acompanhamento ligado, esperamos o usuário resolver estes problemas em vez de parar.
USER_FIXABLE_ERRORS = {0x80210002, 0x80210004, 0x80210008, 0x80210016}  # enrosco, papel, atenção, tampa

LOG_FILE = os.path.join(DATA_DIR, "log.txt")


def log(msg):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > 1_000_000:
            os.replace(LOG_FILE, LOG_FILE + ".old")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}\n")
    except OSError:
        pass


def hresult_of(err):
    code = getattr(err, "hresult", 0) or 0
    try:
        if err.excepinfo and err.excepinfo[5]:
            code = err.excepinfo[5]
    except (AttributeError, IndexError, TypeError):
        pass
    return code & 0xFFFFFFFF


def friendly_error(err):
    if isinstance(err, pywintypes.com_error):
        code = hresult_of(err)
        if code in WIA_ERRORS:
            return WIA_ERRORS[code]
        return f"Erro do scanner (código 0x{code:08X})."
    return str(err)


def find_prop(props, prop_id):
    for p in props:
        if p.PropertyID == prop_id:
            return p
    return None


def set_prop(props, prop_id, value, clamp_max=False):
    p = find_prop(props, prop_id)
    if p is None:
        return False
    try:
        if clamp_max:
            try:
                value = min(value, int(p.SubTypeMax))
            except Exception:
                pass
        p.Value = value
        return True
    except Exception:
        return False


def list_scanners():
    dm = win32com.client.Dispatch("WIA.DeviceManager")
    scanners = []
    for info in dm.DeviceInfos:
        if info.Type == WIA_DEVICE_TYPE_SCANNER:
            name = find_prop(info.Properties, 7)  # WIA_DIP_DEV_NAME
            port = find_prop(info.Properties, 6)  # WIA_DIP_PORT_NAME
            scanners.append((info.DeviceID, name.Value if name else info.DeviceID, port.Value if port else ""))
    return scanners


def is_blank(img):
    """Detecta página em branco ignorando as bordas."""
    gray = img.convert("L")
    w, h = gray.size
    gray = gray.crop((int(w * 0.06), int(h * 0.06), int(w * 0.94), int(h * 0.94)))
    gray.thumbnail((400, 400))
    hist = gray.histogram()
    dark = sum(hist[:170])
    return dark / max(1, sum(hist)) < 0.004


def make_thumb(path):
    with Image.open(path) as im:
        im = im.convert("RGB")
        im.thumbnail((THUMB_SIZE, THUMB_SIZE))
        return im.copy()


def rotate_file(path, angle):
    """Gira a imagem salva em disco (ângulo anti-horário), preservando o DPI."""
    with Image.open(path) as im:
        dpi = im.info.get("dpi")
        rotated = im.rotate(angle, expand=True)
    kwargs = {"dpi": dpi} if dpi else {}
    if path.lower().endswith(".jpg"):
        kwargs["quality"] = 90
    rotated.save(path, **kwargs)


def process_page(raw, settings, dpi, datatype):
    """Converte o arquivo bruto do scanner. Retorna o caminho final, ou None se a página estava em branco."""
    with Image.open(raw) as im:
        im.load()
    os.remove(raw)
    if settings.get("skip_blank") and is_blank(im):
        return None
    if settings.get("auto_rotate"):
        angle = detect_rotation(im)
        if angle:
            im = im.rotate(angle, expand=True)
    stamp = time.time_ns()
    if datatype == 0:
        out = os.path.join(SESSION_DIR, f"pag_{stamp}.png")
        if im.mode != "1":
            im = im.convert("L").point(lambda v: 255 if v > 155 else 0, mode="1")
        im.save(out, dpi=(dpi, dpi), optimize=True)
    else:
        out = os.path.join(SESSION_DIR, f"pag_{stamp}.jpg")
        im.convert("RGB" if datatype == 3 else "L").save(out, "JPEG", quality=85, dpi=(dpi, dpi))
    return out


def scan_worker(device_id, settings, events, stop_event, mon=None):
    """Executa em thread separada: puxa folhas do alimentador até acabar o papel.
    O tratamento de cada página (conversão, páginas em branco, rotação) roda em
    outra thread, para o alimentador não ficar esperando."""
    pythoncom.CoInitialize()
    intent, datatype, depth = COLORS[settings["color"]]
    dpi = int(settings["dpi"])
    source = SOURCES[settings["source"]]
    feeder = bool(source & FEEDER)
    if device_id.startswith("escl:") and feeder:
        dpi = min(dpi, escl.ADF_MAX_DPI)
    kept = [0]
    raw_queue = queue.Queue()

    def processor():
        while True:
            raw = raw_queue.get()
            if raw is None:
                return
            try:
                out = process_page(raw, settings, dpi, datatype)
            except Exception as e:
                events.put(("warning", f"Não foi possível processar uma página: {e}"))
                continue
            if out is None:
                events.put(("blank", None))
            else:
                kept[0] += 1
                events.put(("page", (out, make_thumb(out))))

    proc_thread = threading.Thread(target=processor, daemon=True)
    proc_thread.start()

    if device_id.startswith("escl:"):
        try:
            final_event = escl_session(device_id[5:], settings, source, dpi, raw_queue, events, stop_event)
        finally:
            raw_queue.put(None)
            proc_thread.join()
            pythoncom.CoUninitialize()
        emit_final(events, final_event, kept[0], feeder, stop_event)
        return

    transferred = 0
    final_event = None
    def connect():
        """Conecta ao scanner e aplica as configurações. Retorna o item de digitalização."""
        dm = win32com.client.Dispatch("WIA.DeviceManager")
        device = None
        for info in dm.DeviceInfos:
            if info.DeviceID == device_id:
                device = info.Connect()
                break
        if device is None:
            raise RuntimeError("Scanner não encontrado. Verifique se está ligado.")

        set_prop(device.Properties, DPS_DOCUMENT_HANDLING_SELECT, source)
        set_prop(device.Properties, DPS_PAGES, 1)

        item = device.Items.Item(1)
        set_prop(item.Properties, IPS_CUR_INTENT, intent)
        set_prop(item.Properties, IPA_DATATYPE, datatype)
        set_prop(item.Properties, IPA_DEPTH, depth)
        set_prop(item.Properties, IPS_XRES, dpi)
        set_prop(item.Properties, IPS_YRES, dpi)
        paper = PAPERS[settings["paper"]]
        if paper:
            set_prop(item.Properties, IPS_XPOS, 0)
            set_prop(item.Properties, IPS_YPOS, 0)
            set_prop(item.Properties, IPS_XEXTENT, int(paper[0] * dpi), clamp_max=True)
            set_prop(item.Properties, IPS_YEXTENT, int(paper[1] * dpi), clamp_max=True)
        return item

    try:
        log(f"Início: {settings}")
        item = connect()
        busy_since = None
        reconnected = False
        duplex = bool(source & DUPLEX)
        watching = mon is not None and mon.ip is not None
        while not stop_event.is_set():
            if watching:
                # Só puxa a próxima folha quando a impressora estiver livre.
                ready = mon.wait_ready(stop_event, lambda m: events.put(("warning", m)), check_adf=feeder)
                log(f"Acompanhamento: {ready} (estado={mon.status.state}, alimentador={mon.status.adf})")
                if ready == "stopped":
                    break
                if ready is None:
                    watching = False
                    events.put(("warning", "Sem resposta da impressora; continuando sem acompanhamento."))
                # No frente e verso, o verso da última folha ainda chega com o alimentador já vazio.
                elif ready == "empty" and (not duplex or transferred % 2 == 0):
                    if transferred == 0:
                        final_event = ("empty", None)
                    break
                if transferred:
                    events.put(("warning", f"Digitalizando... {transferred} página(s) recebidas."))
            try:
                wia_img = item.Transfer(WIA_FORMAT_BMP)
            except pywintypes.com_error as e:
                code = hresult_of(e)
                log(f"Transfer falhou após {transferred} página(s): 0x{code:08X}")
                if code == WIA_ERROR_PAPER_EMPTY:
                    if transferred == 0:
                        final_event = ("empty", None)
                    break
                if watching and code in USER_FIXABLE_ERRORS:
                    events.put(("warning", WIA_ERRORS.get(code, "Problema no scanner.")
                                + "  O programa continua sozinho quando for resolvido."))
                    time.sleep(2)
                    item = connect()
                    continue
                if code not in RETRY_ERRORS:
                    raise
                # Scanner ainda terminando a folha anterior: espera e tenta de novo.
                if busy_since is None:
                    busy_since = time.monotonic()
                    events.put(("warning", "Aguardando o scanner puxar a próxima folha..."))
                waited = time.monotonic() - busy_since
                if waited > BUSY_TIMEOUT:
                    raise
                if waited > BUSY_RECONNECT_AFTER and not reconnected:
                    log("Reconectando ao scanner")
                    item = None
                    time.sleep(1)
                    item = connect()
                    reconnected = True
                else:
                    time.sleep(1)
                continue
            busy_since = None
            reconnected = False
            raw = os.path.join(SESSION_DIR, f"raw_{time.time_ns()}.{wia_img.FileExtension}")
            wia_img.SaveFile(raw)
            del wia_img
            transferred += 1
            raw_queue.put(raw)
            if not feeder:
                break
        log(f"Fim: {transferred} página(s) recebidas")
    except Exception as e:
        log(f"Erro: {e!r}")
        final_event = ("error", friendly_error(e))
    finally:
        raw_queue.put(None)
        proc_thread.join()
        pythoncom.CoUninitialize()

    emit_final(events, final_event, kept[0], feeder, stop_event)


def emit_final(events, final_event, kept, feeder, stop_event):
    if final_event is None:
        events.put(("done", {"count": kept, "feeder": feeder, "stopped": stop_event.is_set()}))
    elif final_event[0] == "error":
        events.put(("error", (final_event[1], kept)))
    else:
        events.put(final_event)


def escl_session(ip, settings, source, dpi, raw_queue, events, stop_event):
    """Digitaliza direto pela rede. Se o papel enroscar, espera o usuário resolver e continua.
    Retorna o evento final (None = terminou normalmente)."""
    feeder = bool(source & FEEDER)
    duplex = bool(source & DUPLEX)
    received = [0]

    def notify(msg):
        events.put(("warning", msg))

    def deliver(data):
        raw = os.path.join(SESSION_DIR, f"raw_{time.time_ns()}.jpg")
        with open(raw, "wb") as f:
            f.write(data)
        received[0] += 1
        raw_queue.put(raw)

    log(f"Início (rede {ip}): {settings}")
    try:
        while True:
            try:
                result = escl.scan(ip, feeder, duplex, ESCL_COLORS[settings["color"]], dpi,
                                   PAPERS[settings["paper"]], deliver, notify, stop_event)
                log(f"Fim (rede): {result}, {received[0]} página(s) recebidas")
                if result == "empty" and received[0] == 0:
                    return ("empty", None)
                return None
            except escl.EsclPaperProblem as e:
                log(f"Problema no papel após {received[0]} página(s): {e}")
                notify(f"{e}  Depois recoloque as folhas que faltam: o programa continua sozinho "
                       f"(ou clique em Parar).")
                while not stop_event.is_set():
                    st = M.fetch_status(ip)
                    if st.online and not st.problem and st.state == "Idle" and st.adf == "ScannerAdfLoaded":
                        break
                    time.sleep(1)
                if stop_event.is_set():
                    return None
                time.sleep(1.5)
    except escl.EsclError as e:
        log(f"Erro (rede): {e}")
        return ("error", str(e))
    except Exception as e:
        log(f"Erro (rede): {e!r}")
        return ("error", friendly_error(e))


def autorotate_worker(paths, events):
    """Endireita páginas já digitalizadas."""
    fixed = 0
    for i, path in enumerate(paths, 1):
        try:
            with Image.open(path) as im:
                im.load()
            angle = detect_rotation(im)
            if angle:
                rotate_file(path, angle)
                fixed += 1
                events.put(("refresh", (path, make_thumb(path))))
        except Exception:
            pass
        events.put(("autorotate_progress", (i, len(paths))))
    events.put(("autorotate_done", (fixed, len(paths))))


# ---------------------------------------------------------------- Interface
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1150x760")
        self.minsize(900, 600)
        try:
            self.iconbitmap(default=resource("icone.ico"))
        except tk.TclError:
            pass
        os.makedirs(SESSION_DIR, exist_ok=True)

        self.config_data = self.load_config()
        self.pages = []          # lista de caminhos, na ordem do documento
        self.thumbs = {}         # caminho -> ImageTk.PhotoImage
        self.tiles = {}          # caminho -> frame
        self.selected = set()
        self.last_clicked = None
        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.scanning = False
        self.blank_count = 0
        self.batch_start = 0
        self.scanners = []

        self.monitor = M.PrinterMonitor()
        self.setup_style()
        self.build_ui()
        self.load_scanners()
        self.monitor.set_ip(self.ip_var.get())
        if not self.ip_var.get() or not self.config_data.get("printer_model"):
            self.detect_printer(silent=True)
        self.after(100, self.poll_events)
        self.after(500, self.poll_monitor)
        self.after(300, self.recover_session)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- configuração persistente
    def load_config(self):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def save_config(self):
        data = {
            "scanner": self.scanner_var.get(),
            "source": self.source_var.get(),
            "color": self.color_var.get(),
            "dpi": self.dpi_var.get(),
            "paper": self.paper_var.get(),
            "skip_blank": self.blank_var.get(),
            "auto_rotate": self.rotate_var.get(),
            "folder": self.folder_var.get(),
            "printer_ip": self.ip_var.get().strip(),
            "printer_model": self.config_data.get("printer_model", ""),
            "format": self.format_var.get(),
        }
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    # ---------- visual
    def setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")
        style.configure("Big.TButton", font=("Segoe UI", 13, "bold"), padding=(18, 10))
        style.configure("Save.TButton", font=("Segoe UI", 11, "bold"), padding=(14, 8))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Status.TLabel", font=("Segoe UI", 11))
        style.configure("Count.TLabel", font=("Segoe UI", 20, "bold"), foreground="#1a5fb4")
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def build_ui(self):
        cfg = self.config_data
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        # --- 1. Configurações
        top = ttk.LabelFrame(root, text=" 1. Configurações da digitalização ", padding=10)
        top.pack(fill="x")

        self.scanner_var = tk.StringVar()
        self.source_var = tk.StringVar(value=cfg.get("source", list(SOURCES)[0]))
        self.color_var = tk.StringVar(value=cfg.get("color", "Colorido"))
        self.dpi_var = tk.StringVar(value=cfg.get("dpi", "200"))
        self.paper_var = tk.StringVar(value=cfg.get("paper", "A4"))
        self.blank_var = tk.BooleanVar(value=cfg.get("skip_blank", False))
        self.rotate_var = tk.BooleanVar(value=cfg.get("auto_rotate", True) and OCR_AVAILABLE)

        ttk.Label(top, text="Scanner:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.scanner_combo = ttk.Combobox(top, textvariable=self.scanner_var, state="readonly", width=42)
        self.scanner_combo.grid(row=0, column=1, sticky="w")
        ttk.Button(top, text="⟳ Atualizar", command=self.load_scanners).grid(row=0, column=2, padx=6, sticky="w")

        ttk.Label(top, text="Origem:").grid(row=0, column=3, sticky="w", padx=(16, 6))
        ttk.Combobox(top, textvariable=self.source_var, values=list(SOURCES), state="readonly",
                     width=34).grid(row=0, column=4, sticky="w")

        ttk.Label(top, text="Impressora:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        mon_row = ttk.Frame(top)
        mon_row.grid(row=2, column=1, columnspan=4, sticky="w", pady=(8, 0))
        self.ip_var = tk.StringVar(value=cfg.get("printer_ip", ""))
        ttk.Entry(mon_row, textvariable=self.ip_var, width=16).pack(side="left")
        ttk.Button(mon_row, text="Detectar", command=self.detect_printer).pack(side="left", padx=6)
        self.mon_dot = tk.Label(mon_row, text="●", font=("Segoe UI", 14), fg="#868e96")
        self.mon_dot.pack(side="left", padx=(10, 4))
        self.mon_var = tk.StringVar(value="Acompanhamento desligado (informe o IP ou clique em Detectar)")
        ttk.Label(mon_row, textvariable=self.mon_var).pack(side="left")
        self.ip_var.trace_add("write", lambda *a: self.monitor.set_ip(self.ip_var.get()))

        ttk.Label(top, text="Cor:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        opts = ttk.Frame(top)
        opts.grid(row=1, column=1, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Combobox(opts, textvariable=self.color_var, values=list(COLORS), state="readonly",
                     width=22).pack(side="left")
        ttk.Label(opts, text="Resolução (DPI):").pack(side="left", padx=(16, 6))
        ttk.Combobox(opts, textvariable=self.dpi_var, values=RESOLUTIONS, state="readonly",
                     width=6).pack(side="left")
        ttk.Label(opts, text="Papel:").pack(side="left", padx=(16, 6))
        ttk.Combobox(opts, textvariable=self.paper_var, values=list(PAPERS), state="readonly",
                     width=17).pack(side="left")
        ttk.Checkbutton(opts, text="Descartar páginas em branco",
                        variable=self.blank_var).pack(side="left", padx=(16, 0))
        ttk.Checkbutton(opts, text="Girar automaticamente (texto em pé)", variable=self.rotate_var,
                        state="normal" if OCR_AVAILABLE else "disabled").pack(side="left", padx=(16, 0))

        # --- 2. Escanear
        mid = ttk.LabelFrame(root, text=" 2. Digitalizar ", padding=10)
        mid.pack(fill="x", pady=(10, 0))
        self.scan_btn = ttk.Button(mid, text="▶  ESCANEAR", style="Big.TButton", command=self.start_scan)
        self.scan_btn.pack(side="left")
        self.stop_btn = ttk.Button(mid, text="■  Parar", style="Big.TButton", command=self.stop_scan,
                                   state="disabled")
        self.stop_btn.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(mid, mode="indeterminate", length=160)
        self.progress.pack(side="left", padx=8)
        self.status_var = tk.StringVar(value="Coloque as folhas no alimentador e clique em ESCANEAR.")
        ttk.Label(mid, textvariable=self.status_var, style="Status.TLabel").pack(side="left", padx=8)
        count_box = ttk.Frame(mid)
        count_box.pack(side="right")
        self.count_var = tk.StringVar(value="0")
        ttk.Label(count_box, textvariable=self.count_var, style="Count.TLabel").pack(side="left")
        ttk.Label(count_box, text=" páginas").pack(side="left")

        # --- Páginas
        pages_box = ttk.LabelFrame(root, text=" Páginas digitalizadas  (clique para selecionar · Ctrl/Shift para várias · "
                                              "duplo clique para ampliar) ", padding=6)
        pages_box.pack(fill="both", expand=True, pady=(10, 0))

        tools = ttk.Frame(pages_box)
        tools.pack(fill="x", pady=(0, 6))
        ttk.Button(tools, text="↺ Girar esquerda", command=lambda: self.rotate_selected(90)).pack(side="left")
        ttk.Button(tools, text="↻ Girar direita", command=lambda: self.rotate_selected(-90)).pack(side="left", padx=4)
        ttk.Button(tools, text="⤒ Endireitar automático", command=self.autorotate_selected).pack(side="left",
                                                                                                 padx=(12, 0))
        ttk.Button(tools, text="◀ Mover p/ antes", command=lambda: self.move_selected(-1)).pack(side="left", padx=(12, 0))
        ttk.Button(tools, text="Mover p/ depois ▶", command=lambda: self.move_selected(1)).pack(side="left", padx=4)
        ttk.Button(tools, text="🗑 Excluir selecionadas", command=self.delete_selected).pack(side="left", padx=(12, 0))
        ttk.Button(tools, text="Selecionar todas", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(tools, text="Limpar tudo (novo lote)", command=self.clear_all).pack(side="right")

        area = ttk.Frame(pages_box)
        area.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(area, background="#e9ecef", highlightthickness=0)
        vsb = ttk.Scrollbar(area, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.grid_frame = tk.Frame(self.canvas, background="#e9ecef")
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.relayout())
        self.bind_all("<MouseWheel>", self.on_wheel)
        self.bind("<Delete>", lambda e: self.delete_selected())
        self.bind("<Control-a>", lambda e: self.select_all())
        self.empty_label = tk.Label(self.grid_frame, text="Nenhuma página ainda.\n\nAs páginas aparecem aqui "
                                    "conforme são digitalizadas.", bg="#e9ecef", fg="#6c757d",
                                    font=("Segoe UI", 12))
        self.empty_label.grid(row=0, column=0, padx=40, pady=60)

        # --- 3. Salvar
        bottom = ttk.LabelFrame(root, text=" 3. Salvar ", padding=10)
        bottom.pack(fill="x", pady=(10, 0))
        self.folder_var = tk.StringVar(value=cfg.get("folder", os.path.join(os.path.expanduser("~"), "Documents")))
        self.name_var = tk.StringVar(value=self.default_name())
        self.format_var = tk.StringVar(value=cfg.get("format", "PDF (um arquivo)"))
        ttk.Label(bottom, text="Pasta:").pack(side="left")
        ttk.Entry(bottom, textvariable=self.folder_var, width=40).pack(side="left", padx=4)
        ttk.Button(bottom, text="Procurar...", command=self.choose_folder).pack(side="left")
        ttk.Label(bottom, text="Nome:").pack(side="left", padx=(14, 4))
        ttk.Entry(bottom, textvariable=self.name_var, width=30).pack(side="left")
        ttk.Combobox(bottom, textvariable=self.format_var, state="readonly", width=18,
                     values=["PDF (um arquivo)", "Imagens (JPG/PNG)"]).pack(side="left", padx=8)
        self.save_btn = ttk.Button(bottom, text="💾  SALVAR", style="Save.TButton", command=self.save)
        self.save_btn.pack(side="right")

    def default_name(self):
        return "Digitalizacao_" + datetime.now().strftime("%Y-%m-%d_%H%M")

    # ---------- scanners
    def load_scanners(self):
        try:
            self.scanners = list_scanners()
        except Exception as e:
            self.scanners = []
            messagebox.showerror(APP_NAME, f"Não foi possível listar os scanners:\n{friendly_error(e)}")
        ip = self.ip_var.get().strip()
        if ip:
            model = self.config_data.get("printer_model") or "Impressora"
            self.scanners.insert(0, ("escl:", model + NETWORK_SUFFIX, ""))
        names = []
        for i, (_, name, _port) in enumerate(self.scanners):
            names.append(name if name not in names else f"{name} ({i + 1})")
        self.scanner_names = names
        self.scanner_combo["values"] = names
        if not names:
            self.scanner_var.set("")
            self.status_var.set("Nenhum scanner encontrado. Verifique se está ligado e clique em Atualizar.")
            return
        saved = self.config_data.get("scanner")
        self.scanner_var.set(saved if saved in names else names[0])

    def current_device_id(self):
        name = self.scanner_var.get()
        if name in self.scanner_names:
            device_id = self.scanners[self.scanner_names.index(name)][0]
            if device_id == "escl:":
                ip = self.ip_var.get().strip()
                return "escl:" + ip if ip else None
            return device_id
        return None

    # ---------- acompanhamento da impressora
    def detect_printer(self, silent=False):
        port = ""
        name = self.scanner_var.get()
        if name in self.scanner_names:
            port = self.scanners[self.scanner_names.index(name)][2]
        self.mon_var.set("Procurando a impressora na rede...")

        def work():
            found = M.discover(port)
            self.events.put(("printer_found", (found, silent)))

        threading.Thread(target=work, daemon=True).start()

    def poll_monitor(self):
        if self.monitor.ip:
            text, color = self.monitor.status.describe()
            self.mon_var.set(f"{self.monitor.ip}:  {text}")
            self.mon_dot.configure(fg=color)
        elif not self.mon_var.get().startswith("Procurando"):
            self.mon_var.set("Acompanhamento desligado (informe o IP ou clique em Detectar)")
            self.mon_dot.configure(fg="#868e96")
        self.after(500, self.poll_monitor)

    # ---------- digitalização
    def start_scan(self):
        if self.scanning:
            return
        device_id = self.current_device_id()
        if not device_id:
            messagebox.showwarning(APP_NAME, "Selecione um scanner.")
            return
        settings = {
            "source": self.source_var.get(),
            "color": self.color_var.get(),
            "dpi": self.dpi_var.get(),
            "paper": self.paper_var.get(),
            "skip_blank": self.blank_var.get(),
            "auto_rotate": self.rotate_var.get(),
        }
        self.save_config()
        self.scanning = True
        self.monitor.fast = True
        self.stop_event.clear()
        self.batch_start = len(self.pages)
        self.scan_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.save_btn.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set("Digitalizando... aguarde.")
        threading.Thread(target=scan_worker, daemon=True,
                         args=(device_id, settings, self.events, self.stop_event, self.monitor)).start()

    def stop_scan(self):
        self.stop_event.set()
        self.status_var.set("Parando após a folha atual...")

    def end_scan_ui(self):
        self.scanning = False
        self.monitor.fast = False
        self.progress.stop()
        self.scan_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.save_btn.configure(state="normal")

    def poll_events(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == "page":
                    path, thumb = data
                    self.add_page(path, thumb)
                    self.status_var.set(f"Digitalizando... {len(self.pages) - self.batch_start} página(s) neste lote.")
                elif kind == "blank":
                    self.blank_count += 1
                    self.status_var.set(f"Página em branco descartada ({self.blank_count}).")
                elif kind == "printer_found":
                    found, silent = data
                    if found:
                        self.ip_var.set(found[0])
                        self.config_data["printer_model"] = found[1]
                        self.save_config()
                        self.load_scanners()
                        self.scanner_var.set(self.scanner_names[0])  # opção pela rede (recomendada)
                        self.save_config()
                        if not silent:
                            messagebox.showinfo(APP_NAME, f"Impressora encontrada:\n{found[1]} ({found[0]})")
                    else:
                        self.mon_var.set("Impressora não encontrada na rede")
                        if not silent:
                            messagebox.showwarning(APP_NAME, "Não encontrei a impressora na rede.\n\n"
                                                             "Você pode digitar o endereço IP manualmente "
                                                             "(aparece no painel da impressora).")
                elif kind == "warning":
                    self.status_var.set(data)
                elif kind == "refresh":
                    path, thumb = data
                    if path in self.tiles:
                        self.thumbs[path] = ImageTk.PhotoImage(thumb)
                        self.tiles[path].img_lbl.configure(image=self.thumbs[path])
                elif kind == "autorotate_progress":
                    self.status_var.set(f"Endireitando páginas... {data[0]} de {data[1]}")
                elif kind == "autorotate_done":
                    self.end_scan_ui()
                    self.status_var.set(f"Pronto: {data[0]} de {data[1]} página(s) foram giradas.")
                elif kind == "empty":
                    self.end_scan_ui()
                    self.status_var.set("Alimentador vazio.")
                    messagebox.showwarning(APP_NAME, "Não há papel no alimentador.\n\n"
                                                     "Coloque as folhas (face para cima) e clique em ESCANEAR.")
                elif kind == "done":
                    self.end_scan_ui()
                    self.on_batch_done(data)
                elif kind == "error":
                    msg, count = data
                    self.end_scan_ui()
                    self.status_var.set("Erro na digitalização.")
                    extra = f"\n\n{count} página(s) deste lote foram mantidas." if count else ""
                    messagebox.showerror(APP_NAME, msg + extra)
        except queue.Empty:
            pass
        self.after(100, self.poll_events)

    def on_batch_done(self, info):
        total = len(self.pages)
        if info["stopped"]:
            self.status_var.set(f"Digitalização interrompida. Total: {total} página(s).")
            return
        if not info["feeder"]:
            self.status_var.set(f"Página digitalizada. Total: {total} página(s).")
            return
        self.status_var.set(f"Lote concluído: {info['count']} página(s). Total: {total}.")
        self.bell()
        again = messagebox.askyesno(
            APP_NAME,
            f"O alimentador esvaziou.\n\n"
            f"Neste lote: {info['count']} página(s)\nTotal até agora: {total} página(s)\n\n"
            f"Deseja CONTINUAR escaneando no mesmo documento?\n"
            f"Coloque mais folhas no alimentador e clique em SIM.\n\n"
            f"Clique em NÃO para terminar e salvar.")
        if again:
            self.start_scan()
        else:
            self.status_var.set(f"Pronto! {total} página(s). Confira e clique em SALVAR.")

    # ---------- miniaturas
    def add_page(self, path, thumb_img=None, index=None):
        if thumb_img is None:
            thumb_img = make_thumb(path)
        self.thumbs[path] = ImageTk.PhotoImage(thumb_img)
        tile = tk.Frame(self.grid_frame, bg="#e9ecef", padx=3, pady=3)
        img_lbl = tk.Label(tile, image=self.thumbs[path], bg="white", bd=1, relief="solid")
        img_lbl.pack()
        num_lbl = tk.Label(tile, bg="#e9ecef", font=("Segoe UI", 9))
        num_lbl.pack()
        tile.num_lbl = num_lbl
        tile.img_lbl = img_lbl
        for w in (tile, img_lbl, num_lbl):
            w.bind("<Button-1>", lambda e, p=path: self.on_click(e, p))
            w.bind("<Double-Button-1>", lambda e, p=path: self.preview(p))
        self.tiles[path] = tile
        if index is None:
            self.pages.append(path)
        else:
            self.pages.insert(index, path)
        self.relayout()
        self.save_order()
        if index is None:
            self.canvas.update_idletasks()
            self.canvas.yview_moveto(1.0)

    def relayout(self):
        if not self.pages:
            self.empty_label.grid(row=0, column=0, padx=40, pady=60)
        else:
            self.empty_label.grid_forget()
        width = max(self.canvas.winfo_width(), 200)
        cols = max(1, width // (THUMB_SIZE + 16))
        for i, path in enumerate(self.pages):
            tile = self.tiles[path]
            tile.grid(row=i // cols, column=i % cols, padx=4, pady=4)
            tile.num_lbl.configure(text=f"Página {i + 1}")
            sel = path in self.selected
            color = "#1a5fb4" if sel else "#e9ecef"
            tile.configure(bg=color)
            tile.num_lbl.configure(bg=color, fg="white" if sel else "black")
        self.count_var.set(str(len(self.pages)))

    def on_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def on_click(self, event, path):
        ctrl = event.state & 0x0004
        shift = event.state & 0x0001
        if shift and self.last_clicked in self.pages:
            a, b = sorted((self.pages.index(self.last_clicked), self.pages.index(path)))
            self.selected = set(self.pages[a:b + 1])
        elif ctrl:
            self.selected ^= {path}
            self.last_clicked = path
        else:
            self.selected = {path}
            self.last_clicked = path
        self.relayout()

    def select_all(self):
        self.selected = set(self.pages)
        self.relayout()

    def selected_in_order(self):
        return [p for p in self.pages if p in self.selected]

    def need_selection(self):
        if not self.selected:
            messagebox.showinfo(APP_NAME, "Selecione uma ou mais páginas clicando nelas.")
            return False
        return True

    def rotate_selected(self, angle):
        if self.scanning or not self.need_selection():
            return
        for path in self.selected_in_order():
            rotate_file(path, angle)
            self.thumbs[path] = ImageTk.PhotoImage(make_thumb(path))
            self.tiles[path].img_lbl.configure(image=self.thumbs[path])

    def autorotate_selected(self):
        """Endireita as páginas selecionadas (ou todas, se nenhuma estiver selecionada)."""
        if self.scanning:
            return
        if not OCR_AVAILABLE:
            messagebox.showwarning(APP_NAME, "O reconhecimento de texto do Windows não está disponível.")
            return
        paths = self.selected_in_order() or list(self.pages)
        if not paths:
            return
        self.scanning = True
        self.scan_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set(f"Endireitando {len(paths)} página(s)...")
        threading.Thread(target=autorotate_worker, args=(paths, self.events), daemon=True).start()

    def move_selected(self, direction):
        if not self.need_selection():
            return
        order = self.pages
        idxs = sorted(order.index(p) for p in self.selected)
        if direction < 0:
            if idxs[0] == 0:
                return
            for i in idxs:
                order[i - 1], order[i] = order[i], order[i - 1]
        else:
            if idxs[-1] == len(order) - 1:
                return
            for i in reversed(idxs):
                order[i + 1], order[i] = order[i], order[i + 1]
        self.relayout()
        self.save_order()

    def remove_pages(self, paths):
        for path in paths:
            self.pages.remove(path)
            self.tiles.pop(path).destroy()
            self.thumbs.pop(path, None)
            self.selected.discard(path)
            try:
                os.remove(path)
            except OSError:
                pass
        self.relayout()
        self.save_order()

    def delete_selected(self):
        if self.scanning or not self.need_selection():
            return
        n = len(self.selected)
        if messagebox.askyesno(APP_NAME, f"Excluir {n} página(s) selecionada(s)?"):
            self.remove_pages(self.selected_in_order())

    def clear_all(self, ask=True):
        if self.scanning or not self.pages:
            return
        if ask and not messagebox.askyesno(APP_NAME, f"Apagar todas as {len(self.pages)} páginas e começar "
                                                     f"um novo lote?\n\n(Páginas não salvas serão perdidas.)"):
            return
        self.remove_pages(list(self.pages))
        self.blank_count = 0
        self.name_var.set(self.default_name())
        self.status_var.set("Novo lote. Coloque as folhas no alimentador e clique em ESCANEAR.")

    def preview(self, path):
        win = tk.Toplevel(self)
        win.title(f"Página {self.pages.index(path) + 1} de {len(self.pages)}")
        with Image.open(path) as im:
            im = im.convert("RGB")
            h = int(self.winfo_screenheight() * 0.85)
            im.thumbnail((int(h * 0.9), h))
            photo = ImageTk.PhotoImage(im)
        lbl = tk.Label(win, image=photo)
        lbl.image = photo
        lbl.pack()
        win.bind("<Escape>", lambda e: win.destroy())
        win.focus_set()

    # ---------- sessão (recuperação em caso de fechamento inesperado)
    def save_order(self):
        try:
            with open(ORDER_FILE, "w", encoding="utf-8") as f:
                json.dump([os.path.basename(p) for p in self.pages], f)
        except OSError:
            pass

    def recover_session(self):
        files = [f for f in os.listdir(SESSION_DIR) if f.startswith("pag_")]
        for f in os.listdir(SESSION_DIR):
            if f.startswith("raw_"):
                try:
                    os.remove(os.path.join(SESSION_DIR, f))
                except OSError:
                    pass
        if not files:
            return
        if not messagebox.askyesno(APP_NAME, f"Foram encontradas {len(files)} página(s) de uma digitalização "
                                             f"anterior que não foi concluída.\n\nDeseja recuperá-las?"):
            for f in files:
                try:
                    os.remove(os.path.join(SESSION_DIR, f))
                except OSError:
                    pass
            self.save_order()
            return
        try:
            with open(ORDER_FILE, encoding="utf-8") as fh:
                order = [f for f in json.load(fh) if f in files]
        except (OSError, ValueError):
            order = []
        order += sorted(set(files) - set(order))
        for f in order:
            try:
                self.add_page(os.path.join(SESSION_DIR, f))
            except Exception:
                pass
        self.status_var.set(f"{len(self.pages)} página(s) recuperada(s).")

    # ---------- salvar
    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get() or None)
        if folder:
            self.folder_var.set(os.path.normpath(folder))

    def save(self):
        if not self.pages:
            messagebox.showinfo(APP_NAME, "Não há páginas para salvar.")
            return
        folder = self.folder_var.get().strip()
        name = self.name_var.get().strip()
        for ch in '<>:"/\\|?*':
            name = name.replace(ch, "_")
        if not name:
            messagebox.showwarning(APP_NAME, "Digite um nome para o arquivo.")
            return
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            messagebox.showerror(APP_NAME, "Pasta inválida. Escolha outra pasta.")
            return
        self.save_config()
        pages = list(self.pages)
        as_pdf = self.format_var.get().startswith("PDF")
        if as_pdf:
            target = os.path.join(folder, name + ".pdf")
            if os.path.exists(target) and not messagebox.askyesno(APP_NAME, f"O arquivo já existe:\n{target}\n\n"
                                                                          f"Deseja substituí-lo?"):
                return
        else:
            target = os.path.join(folder, name)
            if os.path.exists(target) and os.listdir(target):
                messagebox.showwarning(APP_NAME, f"A pasta já existe e não está vazia:\n{target}\n\n"
                                                 f"Escolha outro nome.")
                return

        self.save_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")
        self.progress.start(12)
        self.status_var.set(f"Salvando {len(pages)} página(s)...")
        result = {}

        def work():
            try:
                if as_pdf:
                    tmp = target + ".tmp"
                    with open(tmp, "wb") as f:
                        img2pdf.convert(pages, outputstream=f)
                    os.replace(tmp, target)
                else:
                    os.makedirs(target, exist_ok=True)
                    for i, p in enumerate(pages, 1):
                        shutil.copy2(p, os.path.join(target, f"{name}_{i:04d}{os.path.splitext(p)[1]}"))
            except Exception as e:
                result["error"] = str(e)

        t = threading.Thread(target=work, daemon=True)
        t.start()

        def check():
            if t.is_alive():
                self.after(150, check)
                return
            self.progress.stop()
            self.save_btn.configure(state="normal")
            self.scan_btn.configure(state="normal")
            if "error" in result:
                self.status_var.set("Erro ao salvar.")
                messagebox.showerror(APP_NAME, f"Erro ao salvar:\n{result['error']}")
                return
            self.status_var.set(f"Salvo: {target}")
            answer = messagebox.askyesnocancel(
                APP_NAME, f"{len(pages)} página(s) salvas com sucesso em:\n{target}\n\n"
                          f"SIM = abrir a pasta e começar novo lote\n"
                          f"NÃO = começar novo lote\n"
                          f"CANCELAR = manter as páginas na tela")
            if answer is not None:
                if answer:
                    os.startfile(folder)
                self.clear_all(ask=False)

        check()

    def on_close(self):
        if self.scanning and not messagebox.askyesno(APP_NAME, "Uma digitalização está em andamento. Sair mesmo assim?"):
            return
        self.save_config()
        self.monitor.close()
        self.destroy()


def self_test():
    """Verifica os componentes do programa (usado para testar o .exe): escanner --autoteste"""
    from PIL import ImageDraw, ImageFont
    results = []

    def check(name, fn):
        try:
            results.append(f"OK    {name}: {fn()}")
        except Exception as e:
            results.append(f"FALHA {name}: {e!r}")

    check("scanners", lambda: [n for _, n, _ in list_scanners()])
    check("ocr disponivel", lambda: OCR_AVAILABLE)

    def ocr():
        page = Image.new("RGB", (1654, 2339), "white")
        d = ImageDraw.Draw(page)
        font = ImageFont.truetype(os.path.join(os.environ["WINDIR"], "Fonts", "arial.ttf"), 34)
        for i in range(30):
            d.text((150, 150 + i * 65), "Pelo presente instrumento particular, de um lado a empresa", font=font, fill="black")
        return {a: detect_rotation(page.rotate(a, expand=True)) for a in (0, 90, 180, 270)}

    check("rotacao", ocr)

    def pdf():
        import io
        buf = io.BytesIO()
        Image.new("L", (100, 100), 255).save(buf, "JPEG")
        return len(img2pdf.convert([buf.getvalue()])) > 0

    check("pdf", pdf)
    ip = App.load_config(None).get("printer_ip") or "192.168.2.248"
    check("impressora", lambda: M.fetch_status(ip).describe()[0])
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(os.path.join(DATA_DIR, "autoteste.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(results) + "\n")


if __name__ == "__main__":
    if "--autoteste" in sys.argv:
        self_test()
    else:
        App().mainloop()
