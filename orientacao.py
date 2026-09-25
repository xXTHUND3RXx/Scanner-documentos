"""
Detecção automática de orientação da página usando o OCR nativo do Windows.
Lê o texto da página em cada rotação e escolhe aquela em que o texto faz sentido.
"""
import asyncio
import re

try:
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat
    from winrt.windows.storage.streams import DataWriter
    from winrt.windows.globalization import Language
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False

WORK_SIZE = 1400  # lado maior usado na análise (mais rápido que a imagem cheia)

COMMON_WORDS = set("""
de a o que e do da em um para é com não nao uma os no se na por mais as dos como mas foi ao ele das
tem à seu sua ou ser quando muito há nos já está esta eu também só pelo pela até isso ela entre era
depois sem mesmo aos ter seus quem nas me esse eles estão você tinha foram essa num nem suas meu às
minha numa pelos elas havia seja qual será nós tenho lhe deles essas esses pelas este fosse dele
valor data nome rua cpf cnpj total nº n° art lei sr sra dia mês ano rio janeiro brasil são the and of
""".split())

_engine = None


def _get_engine():
    global _engine
    if _engine is None:
        _engine = OcrEngine.try_create_from_language(Language("pt-BR")) \
            or OcrEngine.try_create_from_user_profile_languages()
    return _engine


async def _recognize(img):
    rgba = img.convert("RGBA")
    writer = DataWriter()
    writer.write_bytes(rgba.tobytes("raw", "BGRA"))
    bitmap = SoftwareBitmap.create_copy_from_buffer(
        writer.detach_buffer(), BitmapPixelFormat.BGRA8, rgba.width, rgba.height)
    result = await _get_engine().recognize_async(bitmap)
    return [w.text for line in result.lines for w in line.words]


def _score(words):
    score = 0
    for w in words:
        w = w.strip(".,;:!?()[]\"'").lower()
        if w in COMMON_WORDS:
            score += 3
        elif len(w) >= 3 and re.fullmatch(r"[a-záàâãéêíóôõúç]+", w) and re.search(r"[aeiouáéíóúâêôãõ]", w):
            score += 1
    return score


def _score_rotation(img, angle):
    rotated = img.rotate(angle, expand=True) if angle else img
    return _score(asyncio.run(_recognize(rotated)))


def detect_rotation(img):
    """Retorna o ângulo (0, 90, 180 ou 270, sentido anti-horário, como em PIL.rotate)
    que deixa a página em pé. Retorna 0 se não houver texto suficiente para decidir."""
    if not OCR_AVAILABLE or _get_engine() is None:
        return 0
    small = img.convert("L")
    small.thumbnail((WORK_SIZE, WORK_SIZE))
    layout = _line_direction(small)

    if layout == "sideways":
        # Deitada com certeza: só falta saber para qual lado girar.
        scores = {angle: _score_rotation(small, angle) for angle in (90, 270)}
        best = max(scores, key=scores.get)
        return best if scores[best] >= 3 and scores[best] > min(scores.values()) else 0

    angles = (0, 180) if layout == "upright" else (0, 90, 180, 270)
    scores = {angle: _score_rotation(small, angle) for angle in angles}
    best = max(scores, key=scores.get)
    if best != 0 and scores[best] >= 8 and scores[best] > scores[0] * 1.3 + 5:
        return best
    return 0


def _line_direction(gray):
    """Diz se as linhas de texto estão na horizontal ('upright'), na vertical ('sideways')
    ou se não dá para saber (None), pelas faixas brancas entre as linhas."""
    g = gray.copy()
    w, h = g.size
    g = g.crop((int(w * 0.04), int(h * 0.04), int(w * 0.96), int(h * 0.96)))  # ignora sombras da borda
    g.thumbnail((600, 600))
    ink = g.point(lambda v: 1 if v < 150 else 0)
    box = ink.getbbox()
    if not box:
        return None
    ink = ink.crop(box)
    w, h = ink.size
    if w < 20 or h < 20:
        return None
    px = ink.load()
    empty_rows = sum(1 for y in range(h) if sum(px[x, y] for x in range(w)) <= w * 0.01) / h
    empty_cols = sum(1 for x in range(w) if sum(px[x, y] for y in range(h)) <= h * 0.01) / w
    if empty_rows - empty_cols >= 0.15:
        return "upright"
    if empty_cols - empty_rows >= 0.15:
        return "sideways"
    return None
