"""
Tela para criar e editar os modelos de documento (tipo de recibo, como reconhecê-lo,
nome do arquivo e de onde tirar cada campo), sem precisar mexer no código.
"""
import tkinter as tk
from tkinter import ttk, messagebox

import recibos as R

KIND_LABELS = {R.NUMBER: "Número", R.TEXT: "Texto"}
KIND_VALUES = {v: k for k, v in KIND_LABELS.items()}

HELP = ("Como reconhecer: frases que precisam aparecer no texto impresso, separadas por ponto e vírgula.\n"
        "Use * para qualquer final de palavra (livro* = livro, livros) e | para alternativas (RT|R.T.).\n"
        "Nome do arquivo: use {CAMPO} para cada informação lida da página, além de {DATA} e {N} (sequência).")


def split_phrases(text):
    return [p.strip() for p in text.split(";") if p.strip()]


class ModelsDialog(tk.Toplevel):
    def __init__(self, master, on_saved=None):
        super().__init__(master)
        self.title("Modelos de documento")
        self.transient(master)
        self.grab_set()
        self.on_saved = on_saved
        self.models = [R.Model(m.to_dict()) for m in R.load_models()]
        self.known = R.load_known()
        self.current = None
        self.field_widgets = {}

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        left.pack(side="left", fill="y")
        ttk.Label(left, text="Modelos:").pack(anchor="w")
        self.listbox = tk.Listbox(left, width=26, height=18, exportselection=False, font=("Segoe UI", 10))
        self.listbox.pack(fill="y", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_select)
        lb_buttons = ttk.Frame(left)
        lb_buttons.pack(fill="x", pady=(6, 0))
        ttk.Button(lb_buttons, text="+ Novo", command=self.add).pack(side="left")
        ttk.Button(lb_buttons, text="Duplicar", command=self.duplicate).pack(side="left", padx=4)
        ttk.Button(lb_buttons, text="Excluir", command=self.delete).pack(side="left")
        ttk.Button(left, text="Restaurar modelos padrão", command=self.restore).pack(fill="x", pady=(6, 0))

        right = ttk.Frame(body, padding=(14, 0, 0, 0))
        right.pack(side="left", fill="both", expand=True)
        self.name_var = tk.StringVar()
        self.include_var = tk.StringVar()
        self.exclude_var = tk.StringVar()
        self.template_var = tk.StringVar()
        for label, var in (("Nome do modelo:", self.name_var),
                           ("Como reconhecer (o texto contém):", self.include_var),
                           ("E NÃO contém (opcional):", self.exclude_var),
                           ("Nome do arquivo:", self.template_var)):
            ttk.Label(right, text=label).pack(anchor="w")
            entry = ttk.Entry(right, textvariable=var, width=62, font=("Segoe UI", 10))
            entry.pack(anchor="w", pady=(0, 8))
        self.template_var.trace_add("write", lambda *a: self.after_idle(self.rebuild_fields))
        ttk.Label(right, text=HELP, foreground="#6c757d", justify="left", wraplength=520).pack(anchor="w")

        self.fields_box = ttk.LabelFrame(right, text=" Campos (de onde tirar cada informação) ", padding=8)
        self.fields_box.pack(fill="x", pady=(10, 0))

        buttons = ttk.Frame(self, padding=(10, 0, 10, 10))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="💾  Salvar modelos", style="Save.TButton", command=self.save).pack(side="right")
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="right", padx=8)

        self.fill_list()
        if self.models:
            self.listbox.selection_set(0)
            self.on_select()

    # ---------- lista
    def fill_list(self):
        self.listbox.delete(0, "end")
        for m in self.models:
            self.listbox.insert("end", m.name)

    def on_select(self, _event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        self.store_current()
        self.current = sel[0]
        m = self.models[self.current]
        for child in self.fields_box.winfo_children():
            child.destroy()
        self.field_widgets = {}
        self.loading = True
        self.name_var.set(m.name)
        self.include_var.set("; ".join(m.include))
        self.exclude_var.set("; ".join(m.exclude))
        self.template_var.set(m.template)
        self.loading = False
        self.rebuild_fields()

    def add(self):
        self.store_current()
        self.models.append(R.Model({"nome": "Novo modelo", "contem": [], "arquivo": "NOVO DOCUMENTO - {CAMPO}"}))
        self.select_index(len(self.models) - 1)

    def duplicate(self):
        if self.current is None:
            return
        self.store_current()
        data = self.models[self.current].to_dict()
        data["nome"] += " (cópia)"
        self.models.append(R.Model(data))
        self.select_index(len(self.models) - 1)

    def delete(self):
        if self.current is None:
            return
        m = self.models[self.current]
        if not messagebox.askyesno("Modelos", f"Excluir o modelo \"{m.name}\"?", parent=self):
            return
        del self.models[self.current]
        self.current = None
        self.fill_list()
        if self.models:
            self.select_index(0)

    def restore(self):
        if not messagebox.askyesno("Modelos", "Voltar para os modelos padrão?\n\n"
                                              "Os modelos criados ou alterados serão perdidos.", parent=self):
            return
        self.models = R.default_models()
        self.current = None
        self.fill_list()
        self.select_index(0)

    def select_index(self, index):
        self.fill_list()
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(index)
        self.listbox.see(index)
        self.current = None
        self.on_select()

    # ---------- campos
    def rebuild_fields(self):
        if getattr(self, "loading", False) or self.current is None:
            return
        model = self.models[self.current]
        fields = R.template_fields(self.template_var.get())
        previous = self.read_field_widgets()
        if list(self.field_widgets) == fields:
            return
        for child in self.fields_box.winfo_children():
            child.destroy()
        self.field_widgets = {}
        if not fields:
            ttk.Label(self.fields_box, text="O nome do arquivo não usa nenhum {CAMPO}.",
                      foreground="#6c757d").pack(anchor="w")
            return
        header = ttk.Frame(self.fields_box)
        header.pack(fill="x")
        for text, width in (("Campo", 11), ("Vem depois de", 15), ("Termina antes de", 15), ("Remover", 12),
                            ("Tipo", 9), ("À mão", 6), ("", 16)):
            ttk.Label(header, text=text, width=width, font=("Segoe UI", 9, "bold")).pack(side="left", padx=2)
        for field in fields:
            cfg = previous.get(field) or model.cfg(field)
            row = ttk.Frame(self.fields_box)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text="{" + field + "}", width=11).pack(side="left", padx=2)
            after = tk.StringVar(value=cfg["depois"])
            before = tk.StringVar(value=cfg["antes"])
            remove = tk.StringVar(value="; ".join(cfg.get("remover", [])))
            kind = tk.StringVar(value=KIND_LABELS.get(cfg["tipo"], "Texto"))
            hand = tk.BooleanVar(value=cfg["manuscrito"])
            ttk.Entry(row, textvariable=after, width=15).pack(side="left", padx=2)
            ttk.Entry(row, textvariable=before, width=15).pack(side="left", padx=2)
            ttk.Entry(row, textvariable=remove, width=12).pack(side="left", padx=2)
            ttk.Combobox(row, textvariable=kind, values=list(KIND_VALUES), state="readonly",
                         width=8).pack(side="left", padx=2)
            ttk.Checkbutton(row, variable=hand, width=4).pack(side="left", padx=2)
            ttk.Button(row, text="Valores conhecidos...",
                       command=lambda f=field: self.edit_known(f)).pack(side="left", padx=2)
            self.field_widgets[field] = (after, before, remove, kind, hand)
        ttk.Label(self.fields_box, foreground="#6c757d", justify="left", wraplength=560,
                  text="\"Remover\": textos a tirar do valor, separados por ponto e vírgula (ex.: RJ).\n"
                       "\"À mão\": o campo é escrito à mão; o programa não confia na leitura e pede para "
                       "digitar, completando com os valores conhecidos.").pack(anchor="w", pady=(6, 0))

    def read_field_widgets(self):
        return {f: {"depois": a.get(), "antes": b.get(), "remover": split_phrases(r.get()),
                    "tipo": KIND_VALUES.get(k.get(), R.TEXT), "manuscrito": h.get()}
                for f, (a, b, r, k, h) in self.field_widgets.items()}

    def store_current(self):
        if self.current is None or self.current >= len(self.models):
            return
        self.models[self.current] = R.Model({
            "nome": self.name_var.get(),
            "contem": split_phrases(self.include_var.get()),
            "nao_contem": split_phrases(self.exclude_var.get()),
            "arquivo": self.template_var.get(),
            "campos": self.read_field_widgets(),
        })
        self.listbox.delete(self.current)
        self.listbox.insert(self.current, self.models[self.current].name)
        self.listbox.selection_set(self.current)

    def edit_known(self, field):
        win = tk.Toplevel(self)
        win.title(f"Valores conhecidos de {{{field}}}")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, padding=8, justify="left",
                  text=f"Um valor por linha. Eles são sugeridos ao digitar o campo {{{field}}}.\n"
                       f"Os valores digitados na conferência são adicionados aqui automaticamente.").pack(anchor="w")
        text = tk.Text(win, width=50, height=18, font=("Segoe UI", 10))
        text.pack(padx=8, fill="both", expand=True)
        text.insert("1.0", "\n".join(sorted(self.known.get(field, []))))

        def save():
            values = [v.strip().upper() for v in text.get("1.0", "end").splitlines() if v.strip()]
            self.known[field] = sorted(set(values))
            win.destroy()

        buttons = ttk.Frame(win, padding=8)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="OK", command=save).pack(side="right")
        ttk.Button(buttons, text="Cancelar", command=win.destroy).pack(side="right", padx=6)

    # ---------- salvar
    def save(self):
        self.store_current()
        names = [m.name for m in self.models]
        problems = []
        if len(set(n.lower() for n in names)) != len(names):
            problems.append("Há dois modelos com o mesmo nome.")
        for m in self.models:
            if not m.include:
                problems.append(f"\"{m.name}\": informe como reconhecer o documento.")
            if not m.template:
                problems.append(f"\"{m.name}\": informe o nome do arquivo.")
        if problems:
            messagebox.showwarning("Modelos", "\n".join(problems), parent=self)
            return
        R.save_models(self.models)
        R.save_known(self.known)
        self.destroy()
        if self.on_saved:
            self.on_saved()
