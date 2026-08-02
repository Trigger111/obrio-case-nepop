# -*- coding: utf-8 -*-
"""Прибирає з блоків графіків третій <c:axId>, який pptxgenjs пише, але не оголошує.

За схемою OOXML у 2-D barChart/lineChart рівно два axId. Зайве посилання
формально робить файл невалідним, тому видаляємо його після генерації.
"""
import re
import shutil
import zipfile
from pathlib import Path

DECK = Path(__file__).resolve().parent.parent / "01_OBRIO_creative_testing_framework.pptx"
TMP = DECK.with_suffix(".fixed.pptx")

axid = re.compile(r'<c:axId val="(\d+)"/>')
plot = re.compile(r"<c:(barChart|lineChart|areaChart|pieChart|doughnutChart)>(.*?)</c:\1>", re.S)
declared_axis = re.compile(r"<c:(catAx|valAx|dateAx|serAx)>\s*<c:axId val=\"(\d+)\"/>")

removed = 0
with zipfile.ZipFile(DECK) as src, zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as dst:
    for item in src.infolist():
        data = src.read(item.filename)
        if re.match(r"ppt/charts/chart\d+\.xml$", item.filename):
            text = data.decode("utf-8")
            declared = {m.group(2) for m in declared_axis.finditer(text)}

            def clean(match):
                global removed
                body = match.group(2)
                ids = axid.findall(body)
                extra = [i for i in ids if i not in declared]
                for value in set(extra):
                    body = body.replace(f'<c:axId val="{value}"/>', "")
                    removed += 1
                return f"<c:{match.group(1)}>{body}</c:{match.group(1)}>"

            text = plot.sub(clean, text)
            data = text.encode("utf-8")
        dst.writestr(item, data)

shutil.move(str(TMP), str(DECK))
print(f"Прибрано неоголошених посилань на осі: {removed}")

# контрольна перевірка
with zipfile.ZipFile(DECK) as z:
    for name in sorted(n for n in z.namelist() if re.match(r"ppt/charts/chart\d+\.xml$", n)):
        text = z.read(name).decode("utf-8")
        declared = {m.group(2) for m in declared_axis.finditer(text)}
        for match in plot.finditer(text):
            ids = axid.findall(match.group(2))
            assert len(ids) == 2, f"{name}: {match.group(1)} має {len(ids)} axId"
            assert set(ids) <= declared, f"{name}: посилання на неоголошену вісь"
print("Перевірка пройдена: у кожному блоці графіка рівно дві оголошені осі.")
