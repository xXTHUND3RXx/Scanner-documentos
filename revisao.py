"""
Tela de conferência antes de salvar os PDFs separados: mostra cada página com o tipo
identificado, o trecho do campo ampliado e o nome final do arquivo, para conferir e
completar o que o OCR não conseguiu ler (ex.: nomes escritos à mão).
"""
import tkinter as tk
from tkinter import ttk, messagebox

from PIL import Image, ImageTk

import recibos as R

NAVIGATION_KEYS = {"BackSpace", "Delete", "Left", "Right", "Up", "Down", "Home", "End", "Return", "Tab",
                   "Escape", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"}


class ReviewDialog(tk.Toplevel):
    PREVIEW_SIZE = (300, 420)
    ZOOM_SIZE = (620, 240)

    def __init__(self, master, rows, models, folder, on_confirm):
        """rows: lista de dicts {path, page (OCR), model (Model ou None), values {campo: valor}}"""
        super().__init__(master)
        self.title("Conferir nomes dos arquivos")
        self.transient(master)
        self.grab_set()
        self.rows = rows
        self.models = models
        self.by_name = {m.name: m for m in models}
        self.folder = folder
        self.on_confirm = on_confirm
        self.known = R.load_known()
        self.current = None
        self.field_vars = {}
        self.field_entries = []
        self.photos = {}

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        self.info_var = tk.StringVar()
        ttk.Label(body, textvariable=self.info_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        main = ttk.Frame(body)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)
        self.tree = ttk.Treeview(left, columns=("pag", "tipo", "nome"), show="headings", height=22)
        self.tree.heading("pag", text="Pág.")
        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("nome", text="Nome do arquivo")
        self.tree.column("pag", width=45, anchor="center", stretch=False)
        self.tree.column("tipo", width=140, stretch=False)
        self.tree.column("nome", width=380)
        self.tree.tag_configure("missing", background="#ffe3e3")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        right = ttk.Frame(main, padding=(10, 0, 0, 0))
        right.pack(side="left", fill="both")
        top = ttk.Frame(right)
        top.pack(fill="x")
        self.preview = tk.Label(top, bg="#e9ecef", cursor="hand2")
        self.preview.pack(side="left", anchor="n")
        self.preview.bind("<Button-1>", lambda e: self.open_full())
        form = ttk.Frame(top, padding=(10, 0, 0, 0))
        form.pack(side="left", fill="both", expand=True, anchor="n")
        ttk.Label(form, text="Tipo de documento:").pack(anchor="w")
        self.type_var = tk.StringVar()
        self.type_combo = ttk.Combobox(form, textvariable=self.type_var, state="readonly", width=28,
                                       values=[m.name for m in models])
        self.type_combo.pack(anchor="w", pady=(2, 10))
        self.type_combo.bind("<<ComboboxSelected>>", self.on_type_change)
        self.fields_frame = ttk.Frame(form)
        self.fields_frame.pack(fill="x")
        self.suggest_var = tk.StringVar()
        ttk.Label(form, textvariable=self.suggest_var, foreground="#1a5fb4", wraplength=260,
                  justify="left").pack(anchor="w", pady=(6, 0))
        ttk.Label(form, text="Enter = próximo campo / próxima página.\nClique na página para ampliar.",
                  foreground="#6c757d", justify="left").pack(anchor="w", pady=(10, 0))

        ttk.Label(right, text="Trecho do campo:").pack(anchor="w", pady=(10, 2))
        self.zoom = tk.Label(right, bg="#e9ecef", text="", width=80, height=10)
        self.zoom.pack(anchor="w")

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(10, 0))
        self.save_btn = ttk.Button(buttons, text="💾  Salvar", style="Save.TButton", command=self.confirm)
        self.save_btn.pack(side="right")
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="right", padx=8)
        self.bind("<Escape>", lambda e: self.destroy())

        self.refresh()
        self.select(self.next_pending(-1) or 0)

    # ---------- estado
    def is_pending(self, row):
        return row["model"] is None or any(not row["values"].get(f) for f in row["model"].fields)

    def next_pending(self, after):
        for i in range(after + 1, len(self.rows)):
            if self.is_pending(self.rows[i]):
                return i
        return None

    def base_names(self):
        bases = []
        for i, row in enumerate(self.rows, 1):
            if row["model"] is None:
                bases.append(f"NAO IDENTIFICADO - PAGINA {i}")
            else:
                bases.append(R.fill_name(row["model"], row["values"], i))
        return bases

    def refresh(self):
        names = R.unique_names(self.base_names(), self.folder)
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i, (row, name) in enumerate(zip(self.rows, names)):
            tipo = row["model"].name if row["model"] else "— escolha —"
            self.tree.insert("", "end", iid=str(i), values=(i + 1, tipo, name + ".pdf"),
                             tags=("missing",) if self.is_pending(row) else ())
        if selected:
            self.tree.selection_set(selected)
        pending = sum(1 for r in self.rows if self.is_pending(r))
        total = len(self.rows)
        self.info_var.set(f"{total} PDF(s) serão criados. "
                          + (f"{pending} página(s) em vermelho precisam de informação." if pending
                             else "Tudo preenchido!"))
        self.save_btn.configure(text=f"💾  Salvar {total} PDF(s)")

    # ---------- seleção e formulário
    def select(self, index):
        iid = str(index)
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.tree.see(iid)

    def on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        index = int(sel[0])
        if index == self.current:
            return
        self.commit_fields()
        self.current = index
        row = self.rows[index]
        self.type_var.set(row["model"].name if row["model"] else "")
        self.show_preview(row)
        self.build_fields(row)

    def build_fields(self, row):
        for child in self.fields_frame.winfo_children():
            child.destroy()
        self.field_vars = {}
        self.field_entries = []
        self.suggest_var.set("")
        if row["model"] is None:
            ttk.Label(self.fields_frame, text="Escolha o tipo de documento acima.",
                      foreground="#c92a2a").pack(anchor="w")
            self.show_zoom(row, None)
            self.type_combo.focus_set()
            return
        for i, field in enumerate(row["model"].fields):
            cfg = row["model"].cfg(field)
            ttk.Label(self.fields_frame, text=f"{field}:").pack(anchor="w")
            var = tk.StringVar(value=row["values"].get(field) or "")
            entry = ttk.Entry(self.fields_frame, textvariable=var, width=30, font=("Segoe UI", 12))
            entry.pack(anchor="w", pady=(0, 8))
            entry.bind("<Return>", lambda e, i=i: self.next_field(i))
            entry.bind("<FocusIn>", lambda e, f=field: self.show_zoom(self.rows[self.current], f))
            if cfg["tipo"] == R.TEXT:
                entry.bind("<KeyRelease>", lambda e, f=field, ent=entry: self.autocomplete(e, f, ent))
            self.field_vars[field] = var
            self.field_entries.append(entry)
        if self.field_entries:
            first_empty = next((e for f, e in zip(row["model"].fields, self.field_entries)
                                if not row["values"].get(f)), self.field_entries[0])
            first_empty.focus_set()
            first_empty.select_range(0, "end")

    def commit_fields(self):
        if self.current is None or self.rows[self.current]["model"] is None:
            return
        row = self.rows[self.current]
        for field, var in self.field_vars.items():
            cfg = row["model"].cfg(field)
            row["values"][field] = R.clean_value(var.get(), cfg["tipo"], cfg["remover"])

    def next_field(self, i):
        if i + 1 < len(self.field_entries):
            self.field_entries[i + 1].focus_set()
            self.field_entries[i + 1].select_range(0, "end")
            return
        self.commit_fields()
        self.refresh()
        nxt = self.next_pending(self.current)
        if nxt is None:
            nxt = min(self.current + 1, len(self.rows) - 1)
            if not any(self.is_pending(r) for r in self.rows):
                self.save_btn.focus_set()
        self.select(nxt)

    def on_type_change(self, _event=None):
        if self.current is None:
            return
        row = self.rows[self.current]
        model = self.by_name.get(self.type_var.get())
        if model is row["model"]:
            return
        row["model"] = model
        row["values"] = R.extract(model, row["page"]) if model else {}
        self.build_fields(row)
        self.refresh()

    # ---------- preenchimento automático
    def autocomplete(self, event, field, entry):
        typed = entry.get()[:entry.index("insert")]
        options = R.complete(field, typed, self.known) if typed else []
        self.suggest_var.set(("Sugestões: " + ", ".join(options[:5])) if options else "")
        if event.keysym in NAVIGATION_KEYS or not typed:
            return
        starts = [o for o in options if R.norm(o).startswith(R.norm(typed))]
        if starts and len(starts[0]) > len(typed):
            entry.delete(0, "end")
            entry.insert(0, starts[0])
            entry.icursor(len(typed))
            entry.select_range(len(typed), "end")

    # ---------- imagens
    def show_preview(self, row):
        with Image.open(row["path"]) as im:
            im = im.convert("RGB")
            im.thumbnail(self.PREVIEW_SIZE)
            self.photos["preview"] = ImageTk.PhotoImage(im)
        self.preview.configure(image=self.photos["preview"])

    def show_zoom(self, row, field):
        box = R.field_box(row["model"], field, row["page"]) if row["model"] and field else None
        if box is None:
            self.zoom.configure(image="", text="(campo não localizado na página)", width=80, height=10)
            return
        with Image.open(row["path"]) as im:
            crop = im.convert("RGB").crop(box)
        crop.thumbnail(self.ZOOM_SIZE)
        self.photos["zoom"] = ImageTk.PhotoImage(crop)
        self.zoom.configure(image=self.photos["zoom"], text="", width=crop.width, height=crop.height)

    def open_full(self):
        if self.current is None:
            return
        win = tk.Toplevel(self)
        win.title(f"Página {self.current + 1}")
        with Image.open(self.rows[self.current]["path"]) as im:
            im = im.convert("RGB")
            h = int(self.winfo_screenheight() * 0.85)
            im.thumbnail((int(h * 0.9), h))
            photo = ImageTk.PhotoImage(im)
        lbl = tk.Label(win, image=photo)
        lbl.image = photo
        lbl.pack()
        win.bind("<Escape>", lambda e: win.destroy())
        win.focus_set()

    # ---------- salvar
    def confirm(self):
        self.commit_fields()
        self.refresh()
        pending = sum(1 for r in self.rows if self.is_pending(r))
        if pending and not messagebox.askyesno(
                "Conferir nomes", f"{pending} página(s) ainda estão sem alguma informação e serão salvas "
                                  f"com \"SEM ...\" ou \"NAO IDENTIFICADO\" no nome.\n\nSalvar mesmo assim?",
                parent=self):
            return
        names = R.unique_names(self.base_names(), self.folder)
        R.remember((f, v) for r in self.rows if r["model"]
                   for f, v in r["values"].items() if r["model"].cfg(f)["tipo"] == R.TEXT)
        self.destroy()
        self.on_confirm(names)
