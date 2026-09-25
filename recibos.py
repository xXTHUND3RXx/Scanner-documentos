"""
Separação em um PDF por página e nome automático dos arquivos a partir de um modelo,
por exemplo "RECIBO DE INSUMO - RT {RT}", onde {RT} é lido da página pelo OCR.
"""
import os
import re
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox

from PIL import Image, ImageTk

from orientacao import read_text

DEFAULT_TEMPLATE = "RECIBO DE INSUMO - RT {RT}"
MISSING = "SEM NUMERO"

# "RT: 01", "RT 12345", "RT nº 12345", "R.T. 12.345", "RT-12345/2026"...
# Aceita a letra O no lugar do zero ("RT: O1"), confusão comum do OCR; o número precisa terminar em dígito.
RT_PATTERN = re.compile(
    r"\bR\s?\.?\s?T\b\.?\s*(?:N\s?[º°o.]*\s*)?[:#\-–]?\s*([\dO][\dO./\-]*\d|\d)",
    re.IGNORECASE)

INVALID_CHARS = '<>:"/\\|?*'


# ---------------------------------------------------------------- texto das páginas
def text_cache_path(page_path):
    folder, name = os.path.split(page_path)
    return os.path.join(folder, f"ocr_{name}.txt")


def cache_text(page_path, img=None):
    """Lê o texto da página pelo OCR e guarda ao lado da imagem."""
    if img is None:
        with Image.open(page_path) as im:
            im.load()
            text = read_text(im)
    else:
        text = read_text(img)
    with open(text_cache_path(page_path), "w", encoding="utf-8") as f:
        f.write(text)
    return text


def page_text(page_path):
    try:
        with open(text_cache_path(page_path), encoding="utf-8") as f:
            return f.read()
    except OSError:
        return cache_text(page_path)


def has_cached_text(page_path):
    return os.path.exists(text_cache_path(page_path))


def forget_text(page_path):
    try:
        os.remove(text_cache_path(page_path))
    except OSError:
        pass


def find_rt(text):
    m = RT_PATTERN.search(text or "")
    return re.sub(r"[Oo]", "0", m.group(1)).strip(".-/") if m else None


# ---------------------------------------------------------------- nomes
def uses_rt(template):
    return "{RT}" in template.upper()


def fill_template(template, rt, number):
    name = re.sub(r"\{RT\}", rt or MISSING, template, flags=re.IGNORECASE)
    name = re.sub(r"\{N\}", str(number), name, flags=re.IGNORECASE)
    name = re.sub(r"\{DATA\}", datetime.now().strftime("%d-%m-%Y"), name, flags=re.IGNORECASE)
    for ch in INVALID_CHARS:
        name = name.replace(ch, "-")
    return name.strip().rstrip(".") or f"Pagina {number}"


def final_names(template, rts, folder):
    """Nomes de arquivo (sem extensão) para cada página. Nomes repetidos, ou que já
    existem na pasta, ganham " (2)", " (3)"... para nada ser sobrescrito."""
    taken = set()
    try:
        taken = {os.path.splitext(f)[0].lower() for f in os.listdir(folder) if f.lower().endswith(".pdf")}
    except OSError:
        pass
    names = []
    for i, rt in enumerate(rts, 1):
        base = fill_template(template, rt, i)
        name, n = base, 2
        while name.lower() in taken:
            name = f"{base} ({n})"
            n += 1
        taken.add(name.lower())
        names.append(name)
    return names


# ---------------------------------------------------------------- tela de conferência
class ReviewDialog(tk.Toplevel):
    """Mostra cada página com o número lido e o nome final, para conferir e corrigir antes de salvar."""

    PREVIEW_HEIGHT = 560

    def __init__(self, master, pages, rts, template, folder, on_confirm):
        super().__init__(master)
        self.title("Conferir nomes dos arquivos")
        self.transient(master)
        self.grab_set()
        self.pages = pages
        self.rts = list(rts)
        self.template = template
        self.folder = folder
        self.on_confirm = on_confirm
        self.need_rt = uses_rt(template)
        self.current = None
        self.photo = None

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)

        self.info_var = tk.StringVar()
        ttk.Label(body, textvariable=self.info_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        main = ttk.Frame(body)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)
        self.tree = ttk.Treeview(left, columns=("pag", "rt", "nome"), show="headings", height=20)
        self.tree.heading("pag", text="Pág.")
        self.tree.heading("rt", text="Número da RT")
        self.tree.heading("nome", text="Nome do arquivo")
        self.tree.column("pag", width=50, anchor="center", stretch=False)
        self.tree.column("rt", width=130, anchor="center", stretch=False)
        self.tree.column("nome", width=380)
        self.tree.tag_configure("missing", background="#ffe3e3")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        right = ttk.Frame(main, padding=(10, 0, 0, 0))
        right.pack(side="left", fill="y")
        self.preview = tk.Label(right, bg="#e9ecef", width=56, height=30)
        self.preview.pack()
        edit = ttk.Frame(right)
        edit.pack(fill="x", pady=(8, 0))
        ttk.Label(edit, text="Número da RT:").pack(side="left")
        self.rt_var = tk.StringVar()
        self.rt_entry = ttk.Entry(edit, textvariable=self.rt_var, width=18, font=("Segoe UI", 11))
        self.rt_entry.pack(side="left", padx=6)
        self.rt_entry.bind("<Return>", self.on_enter)
        ttk.Label(right, text="Digite o número e tecle Enter para ir ao próximo.",
                  foreground="#6c757d").pack(anchor="w", pady=(4, 0))
        if not self.need_rt:
            self.rt_entry.configure(state="disabled")

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(10, 0))
        self.save_btn = ttk.Button(buttons, text="💾  Salvar", style="Save.TButton", command=self.confirm)
        self.save_btn.pack(side="right")
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="right", padx=8)

        self.refresh()
        first = next((i for i, rt in enumerate(self.rts) if self.need_rt and not rt), 0)
        self.select(first)
        self.bind("<Escape>", lambda e: self.destroy())
        self.rt_entry.focus_set()

    def refresh(self):
        names = final_names(self.template, self.rts, self.folder)
        self.names = names
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i, (rt, name) in enumerate(zip(self.rts, names)):
            missing = self.need_rt and not rt
            self.tree.insert("", "end", iid=str(i), values=(i + 1, rt or "— digite —", name + ".pdf"),
                             tags=("missing",) if missing else ())
        if selected:
            self.tree.selection_set(selected)
        found = sum(1 for rt in self.rts if rt)
        total = len(self.rts)
        if self.need_rt:
            self.info_var.set(f"{total} PDF(s) serão criados. RT encontrada em {found} de {total}. "
                              + ("Os destacados em vermelho precisam do número." if found < total else "Tudo certo!"))
        else:
            self.info_var.set(f"{total} PDF(s) serão criados.")
        self.save_btn.configure(text=f"💾  Salvar {total} PDF(s)")

    def select(self, index):
        iid = str(index)
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.tree.see(iid)

    def on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        self.current = int(sel[0])
        self.rt_var.set(self.rts[self.current] or "")
        with Image.open(self.pages[self.current]) as im:
            im = im.convert("RGB")
            im.thumbnail((int(self.PREVIEW_HEIGHT * 0.75), self.PREVIEW_HEIGHT))
            self.photo = ImageTk.PhotoImage(im)
        self.preview.configure(image=self.photo, width=self.photo.width(), height=self.photo.height())
        self.rt_entry.select_range(0, "end")

    def on_enter(self, _event=None):
        if self.current is None:
            return
        value = self.rt_var.get().strip()
        for ch in INVALID_CHARS:
            value = value.replace(ch, "-")
        self.rts[self.current] = value or None
        nxt = next((i for i in range(self.current + 1, len(self.rts)) if not self.rts[i]), None)
        if nxt is None:
            nxt = min(self.current + 1, len(self.rts) - 1)
        self.refresh()
        self.select(nxt)

    def confirm(self):
        if self.current is not None and self.need_rt:
            typed = self.rt_var.get().strip()
            if typed and typed != (self.rts[self.current] or ""):
                self.on_enter()
        missing = sum(1 for rt in self.rts if self.need_rt and not rt)
        if missing and not messagebox.askyesno(
                "Conferir nomes", f"{missing} página(s) estão sem número da RT e serão salvas como "
                                  f"\"{MISSING}\".\n\nSalvar mesmo assim?", parent=self):
            return
        names = final_names(self.template, self.rts, self.folder)
        self.destroy()
        self.on_confirm(names)
