# -*- coding: utf-8 -*-
"""Склеює однакові медіафайли в .pptx.

pptxgenjs записує окрему копію зображення на кожен слайд, тому один логотип
перетворюється на 22 однакові частини. Лишаємо по одному екземпляру
й переспрямовуємо посилання у *.rels.
"""
import hashlib
import re
import shutil
import zipfile
from pathlib import Path

DECK = Path(__file__).resolve().parent.parent / "01_OBRIO_creative_testing_framework.pptx"
TMP = DECK.with_suffix(".dedup.pptx")

with zipfile.ZipFile(DECK) as z:
    names = z.namelist()
    media = {n: z.read(n) for n in names if n.startswith("ppt/media/") and not n.endswith("/")}

    canonical = {}          # md5 -> ім'я, яке лишаємо
    remap = {}              # старе ім'я -> канонічне
    for name in sorted(media):
        digest = hashlib.md5(media[name]).hexdigest()
        if digest not in canonical:
            canonical[digest] = name
        remap[name] = canonical[digest]

    dropped = {n for n, keep in remap.items() if n != keep}
    saved = sum(len(media[n]) for n in dropped)

    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as out:
        for item in zipfile.ZipFile(DECK).infolist():
            if item.filename in dropped:
                continue
            data = zipfile.ZipFile(DECK).read(item.filename)
            if item.filename.endswith(".rels"):
                text = data.decode("utf-8")
                for old, keep in remap.items():
                    if old == keep:
                        continue
                    text = text.replace(
                        f'Target="../media/{Path(old).name}"',
                        f'Target="../media/{Path(keep).name}"',
                    )
                data = text.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                text = data.decode("utf-8")
                for old in dropped:
                    text = re.sub(
                        r'<Override[^>]*PartName="/' + re.escape(old) + r'"[^>]*/>', "", text
                    )
                data = text.encode("utf-8")
            out.writestr(item, data)

shutil.move(str(TMP), str(DECK))
print(f"Унікальних зображень: {len(canonical)}; прибрано копій: {len(dropped)}; "
      f"звільнено {saved / 1024:,.0f} KB")

# перевірка: жоден .rels не посилається на видалену частину
with zipfile.ZipFile(DECK) as z:
    present = {Path(n).name for n in z.namelist() if n.startswith("ppt/media/")}
    broken = []
    for name in z.namelist():
        if not name.endswith(".rels"):
            continue
        for target in re.findall(r'Target="\.\./media/([^"]+)"', z.read(name).decode("utf-8")):
            if target not in present:
                broken.append((name, target))
    assert not broken, f"биті посилання: {broken}"
print(f"Перевірка пройдена: {len(present)} медіафайлів, биті посилання відсутні. "
      f"Розмір деку: {DECK.stat().st_size / 1024 / 1024:.2f} MB")
