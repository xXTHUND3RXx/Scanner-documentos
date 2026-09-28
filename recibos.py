"""
Modelos de documento: identificam o tipo de cada página pelo texto impresso e montam
o nome do arquivo a partir dos campos lidos, por exemplo:

    "RECIBO DE INSUMO - RT {RT}"   ->  RECIBO DE INSUMO - RT 01.pdf
    "LIVRO ATA - {CAPS}"           ->  LIVRO ATA - NOME DO CAPS.pdf

Os modelos ficam num arquivo JSON e podem ser criados/alterados pela tela "Modelos",
sem mexer no código. Os valores já digitados (ex.: nomes de CAPS) são guardados para
completar automaticamente as próximas digitações.
"""
import copy
import difflib
import json
import os
import re
import unicodedata
from datetime import datetime

from PIL import Image

from orientacao import read_lines, ordered_text

# Mesmo local de dados do programa principal (perfil do usuário no Windows).
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "EscannerLote")
MODELS_FILE = os.path.join(DATA_DIR, "modelos.json")
KNOWN_FILE = os.path.join(DATA_DIR, "valores_conhecidos.json")

AUTO = "Automático (pelo conteúdo)"
NUMBER, TEXT = "numero", "texto"
BUILTIN_PLACEHOLDERS = {"DATA", "N"}
INVALID_CHARS = '<>:"/\\|?*'

# Regras de reconhecimento:  "*" = qualquer final de palavra (livro* = livro, livros)
#                            "|" = alternativas (RT|R.T.)
DEFAULT_MODELS = [
    {
        "nome": "Recibo de insumo",
        "contem": ["insumo*"],
        "nao_contem": [],
        "arquivo": "RECIBO DE INSUMO - RT {RT}",
        "campos": {"RT": {"depois": "RT|R.T.", "antes": "", "tipo": NUMBER, "manuscrito": False}},
    },
    {
        "nome": "Livro ATA",
        "contem": ["livro* ata"],
        "nao_contem": [],
        "arquivo": "LIVRO ATA - {CAPS}",
        "campos": {"CAPS": {"depois": "CAPS", "antes": "para serem", "tipo": TEXT, "manuscrito": True}},
    },
    {
        "nome": "Livro CAIXA",
        "contem": ["livro* caixa"],
        "nao_contem": [],
        "arquivo": "LIVRO CAIXA - {CAPS}",
        "campos": {"CAPS": {"depois": "CAPS", "antes": "para serem", "tipo": TEXT, "manuscrito": True}},
    },
    {
        "nome": "Caixa BOX",
        "contem": ["caixa* box"],
        "nao_contem": [],
        "arquivo": "CAIXA BOX - {CAPS}",
        "campos": {"CAPS": {"depois": "CAPS", "antes": "para serem", "tipo": TEXT, "manuscrito": True}},
    },
    {
        "nome": "Recebimento de chaves",
        "contem": ["chave*"],
        "nao_contem": [],
        "arquivo": "RECEBIMENTO DE CHAVES - {ENDEREÇO}",
        "campos": {"ENDEREÇO": {"depois": "endere*", "antes": "vinculad*", "tipo": TEXT, "manuscrito": False,
                                "remover": ["RJ"]}},
    },
]


# ---------------------------------------------------------------- modelos
class Model:
    def __init__(self, data):
        self.name = data.get("nome", "").strip() or "Sem nome"
        self.include = [p for p in data.get("contem", []) if p.strip()]
        self.exclude = [p for p in data.get("nao_contem", []) if p.strip()]
        self.template = data.get("arquivo", "").strip()
        self.field_cfg = data.get("campos", {})

    @property
    def fields(self):
        """Campos usados no nome do arquivo, na ordem em que aparecem no modelo."""
        return template_fields(self.template)

    def cfg(self, field):
        base = {"depois": field, "antes": "", "tipo": TEXT, "manuscrito": False, "remover": []}
        base.update(self.field_cfg.get(field, {}))
        return base

    def to_dict(self):
        return {"nome": self.name, "contem": self.include, "nao_contem": self.exclude,
                "arquivo": self.template,
                "campos": {f: self.cfg(f) for f in self.fields}}


def template_fields(template):
    seen = []
    for name in re.findall(r"\{([^{}]+)\}", template):
        name = name.strip().upper()
        if name and name not in BUILTIN_PLACEHOLDERS and name not in seen:
            seen.append(name)
    return seen


def load_models():
    """Modelos salvos pelo usuário. Modelos padrão adicionados em versões novas do programa
    entram automaticamente (os que o usuário excluiu de propósito não voltam)."""
    try:
        with open(MODELS_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return default_models()
    if isinstance(data, list):  # formato antigo
        data = {"modelos": data, "padroes_vistos": []}
    models = [Model(d) for d in data.get("modelos", [])]
    seen = set(data.get("padroes_vistos", [])) | {m.name for m in models}
    models += [Model(d) for d in copy.deepcopy(DEFAULT_MODELS) if d["nome"] not in seen]
    return models or default_models()


def save_models(models):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = MODELS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"modelos": [m.to_dict() for m in models],
                   "padroes_vistos": [d["nome"] for d in DEFAULT_MODELS]}, f, ensure_ascii=False, indent=2)
    os.replace(tmp, MODELS_FILE)


def default_models():
    return [Model(d) for d in copy.deepcopy(DEFAULT_MODELS)]


# ---------------------------------------------------------------- texto
def norm(text):
    """Minúsculas e sem acentos, mantendo o mesmo tamanho (posições batem com o texto original)."""
    return "".join(unicodedata.normalize("NFD", c)[0] for c in text).lower()


def phrase_regex(phrase):
    alternatives = []
    for alt in phrase.split("|"):
        alt = norm(alt.strip())
        if not alt:
            continue
        parts = []
        for ch in alt:
            if ch == "*":
                parts.append(r"[a-z0-9]*")
            elif ch.isspace():
                parts.append(r"\s+")
            else:
                parts.append(re.escape(ch))
        alternatives.append("".join(parts))
    if not alternatives:
        return None
    return r"(?<![a-z0-9])(?:" + "|".join(alternatives) + r")(?![a-z0-9])"


def contains(text_norm, phrase):
    rx = phrase_regex(phrase)
    return bool(rx and re.search(rx, text_norm))


def identify(text, models):
    """Modelo cujo reconhecimento bate com o texto da página; None se nenhum ou se ficar ambíguo."""
    t = norm(text)
    matches = [m for m in models
               if m.include and all(contains(t, p) for p in m.include)
               and not any(contains(t, p) for p in m.exclude)]
    if not matches:
        return None
    matches.sort(key=lambda m: len(m.include), reverse=True)
    if len(matches) > 1 and len(matches[0].include) == len(matches[1].include):
        return None  # ambíguo: o usuário escolhe na conferência
    return matches[0]


def clean_value(value, kind, remove=()):
    """Padroniza o valor para o nome do arquivo:
    "Rua das Flores, 10 APT: 201- Centro - RJ," -> "RUA DAS FLORES 10 APT 201 - CENTRO"
    (com "RJ" na lista de remover)."""
    if value is None:
        return None
    if kind == NUMBER:
        value = re.sub(r"[Oo]", "0", value).strip(".-/ ")
    else:
        value = re.sub(r"[._…·]{2,}", " ", value)            # linhas pontilhadas do formulário
        value = re.sub(r"[,;:]", " ", value)                  # vírgulas e dois-pontos
        value = re.sub(r"\s*[-–—]\s*", " - ", value)          # hífens sempre com espaço
        for phrase in remove:
            rx = phrase_regex(phrase)
            if rx:
                spans = [m.span() for m in re.finditer(rx, norm(value))]
                for a, b in reversed(spans):
                    value = value[:a] + value[b:]
        value = re.sub(r"\s+", " ", value)
        value = re.sub(r"(?:\s*-\s*)+$", "", value)          # hífen sobrando no fim
        value = re.sub(r"^(?:\s*-\s*)+", "", value)
        value = re.sub(r"(?:\s-\s)+", " - ", value).strip(" .-_")
    for ch in INVALID_CHARS:
        value = value.replace(ch, "-")
    return value.upper() or None


def is_plausible(value):
    """Descarta leituras que viraram lixo ("(210WU SEE", "CJ.S'P"): exige letras e poucos símbolos."""
    if not value or sum(c.isalpha() for c in value) < 2:
        return False
    odd = sum(1 for c in value if not (c.isalnum() or c in " -'/()º°"))
    return odd <= len(value) * 0.1


def extract_from_text(model, text):
    """{campo: valor ou None} lendo o texto da página."""
    t = norm(text)
    values = {}
    for field in model.fields:
        cfg = model.cfg(field)
        after = phrase_regex(cfg["depois"])
        value = None
        if after:
            if cfg["tipo"] == NUMBER:
                m = re.search(after + r"\s*(?:n\s?[º°o.]*\s*)?[:#\-–.]*\s*([\do][\do./\-]*\d|\d)", t)
            else:
                before = phrase_regex(cfg["antes"]) if cfg["antes"].strip() else None
                tail = r"(.{1,160}?)\s*" + before if before else r"([^\n]{1,120})"
                m = re.search(after + r"\s*[:\-–]?\s*" + tail, t, re.DOTALL)
            if m:
                value = clean_value(text[m.start(1):m.end(1)], cfg["tipo"], cfg["remover"])
        if value and cfg["tipo"] == TEXT:
            if cfg["manuscrito"]:
                # letra de mão: só aceita se for bem parecido com um valor já conhecido
                value = best_known(field, value)
            elif not is_plausible(value):
                value = None  # leitura não confiável: fica para digitar
        values[field] = value
    return values


def extract(model, page):
    return extract_from_text(model, ordered_text(page))


def field_box(model, field, page):
    """Região da página (x0, y0, x1, y1) onde está o campo, para mostrar ampliada."""
    cfg = model.cfg(field)
    lines = page.get("lines", [])
    size = page.get("size", [0, 0])
    after = phrase_regex(cfg["depois"])
    before = phrase_regex(cfg["antes"]) if cfg["antes"].strip() else None
    if not after or not lines:
        return None
    ordered = sorted(lines, key=lambda l: (l["box"][1], l["box"][0]))
    start = next((l for l in ordered if re.search(after, norm(l["text"]))), None)
    if start is None:
        return None
    end = start
    if before:
        end = next((l for l in ordered if l["box"][1] >= start["box"][1] and re.search(before, norm(l["text"]))),
                   start)
    y0 = start["box"][1]
    y1 = end["box"][1] + end["box"][3]
    margin = max(start["box"][3], 20)
    xs = [l["box"][0] for l in lines]
    xe = [l["box"][0] + l["box"][2] for l in lines]
    return (max(0, min(xs) - margin), max(0, y0 - margin * 2),
            min(size[0], max(xe) + margin), min(size[1], y1 + margin))


# ---------------------------------------------------------------- valores conhecidos
def load_known():
    try:
        with open(KNOWN_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return {k.upper(): list(v) for k, v in data.items()}
    except (OSError, ValueError):
        return {}


def save_known(known):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(KNOWN_FILE, "w", encoding="utf-8") as f:
        json.dump({k: sorted(set(v)) for k, v in known.items()}, f, ensure_ascii=False, indent=2)


def remember(field_values):
    """Guarda os valores de texto usados, para sugerir nas próximas vezes."""
    known = load_known()
    changed = False
    for field, value in field_values:
        if value and not value.isdigit():
            items = known.setdefault(field.upper(), [])
            if value not in items:
                items.append(value)
                changed = True
    if changed:
        save_known(known)


def best_known(field, value, cutoff=0.85):
    match = difflib.get_close_matches(value, load_known().get(field.upper(), []), n=1, cutoff=cutoff)
    return match[0] if match else None


def complete(field, typed, known=None):
    """Valores conhecidos que começam com o que foi digitado (depois os que contêm)."""
    items = (known if known is not None else load_known()).get(field.upper(), [])
    t = norm(typed)
    starts = [v for v in items if norm(v).startswith(t)]
    inside = [v for v in items if t in norm(v) and v not in starts]
    return sorted(starts) + sorted(inside)


# ---------------------------------------------------------------- nomes de arquivo
def fill_name(model, values, number):
    name = model.template
    for field in model.fields:
        name = re.sub(r"\{\s*" + re.escape(field) + r"\s*\}", values.get(field) or f"SEM {field}",
                      name, flags=re.IGNORECASE)
    name = re.sub(r"\{\s*N\s*\}", str(number), name, flags=re.IGNORECASE)
    name = re.sub(r"\{\s*DATA\s*\}", datetime.now().strftime("%d-%m-%Y"), name, flags=re.IGNORECASE)
    for ch in INVALID_CHARS:
        name = name.replace(ch, "-")
    return name.strip().rstrip(".") or f"Pagina {number}"


def unique_names(bases, folder):
    """Nomes repetidos, ou que já existem na pasta, ganham " (2)", " (3)"..."""
    taken = set()
    try:
        taken = {os.path.splitext(f)[0].lower() for f in os.listdir(folder) if f.lower().endswith(".pdf")}
    except OSError:
        pass
    names = []
    for base in bases:
        name, n = base, 2
        while name.lower() in taken:
            name = f"{base} ({n})"
            n += 1
        taken.add(name.lower())
        names.append(name)
    return names


# ---------------------------------------------------------------- OCR guardado de cada página
def ocr_cache_path(page_path):
    folder, name = os.path.split(page_path)
    return os.path.join(folder, f"ocr_{name}.json")


def cache_ocr(page_path, img=None):
    if img is None:
        with Image.open(page_path) as im:
            im.load()
            page = read_lines(im)
    else:
        page = read_lines(img)
    with open(ocr_cache_path(page_path), "w", encoding="utf-8") as f:
        json.dump(page, f, ensure_ascii=False)
    return page


def page_ocr(page_path):
    try:
        with open(ocr_cache_path(page_path), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return cache_ocr(page_path)


def forget_ocr(page_path):
    try:
        os.remove(ocr_cache_path(page_path))
    except OSError:
        pass
