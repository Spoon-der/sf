import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
import mtk2
doc = docx.Document("D:/variant/variant06.docx")
bukw_1, bukw_2, bukw_3, bukw_4, bukw_5 = "", "", "", "", ""

for paragraph in doc.paragraphs:
    for run in paragraph.runs:

        
        if run.font.size.pt == 11.5:
            bukw_1 += "1" * len(run.text)
        else:
            bukw_1 += "0" * len(run.text)

        
        if run.font.color.rgb != RGBColor(0,0,0):
            bukw_2 += "1" * len(run.text)
        else:
            bukw_2 += "0" * len(run.text)

        
        if run.font.highlight_color != WD_COLOR_INDEX.WHITE:
            bukw_3 += "1" * len(run.text)
        else:
            bukw_3 += "0" * len(run.text)

        
        if run._r.get_or_add_rPr().xpath("./w:spacing"):
            bukw_4 += "1" * len(run.text)
        else:
            bukw_4 += "0" * len(run.text)

        
        if run._r.get_or_add_rPr().xpath("./w:w"):
            bukw_5 += "1" * len(run.text)
        else:
            bukw_5 += "0" * len(run.text)

def dlina(tex):
    while len(tex) % 8 != 0:
        tex += "0"
    return tex
bukw_1 = dlina(bukw_1)
bukw_2 = dlina(bukw_2)
bukw_3 = dlina(bukw_3)
bukw_4 = dlina(bukw_4)
bukw_5 = dlina(bukw_5)
print("РАЗМЕР ШРИФТА")
print(bukw_1)
print("ЦВЕТ СИМВОЛОВ")
print(bukw_2)
print("ЦВЕТ ФОНА")
print(bukw_3)
print("МЕЖСИМВОЛЬНЫЙ ИНТЕРВАЛ")
print(bukw_4)
print("МАСШТАБ ШРИФТА")
print(bukw_5)

def encode(code):
    print("cp1251 - ", bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp1251"))
    print("koi8_r - ", bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="koi8_r"))
    print("cp866 - ", bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp866"))
    print("mtk2 - ", mtk2.MTK2_decode(code))

if bukw_1 != "0" * len(bukw_1):
    print("РАЗМЕР ШРИФТА")
    encode(bukw_1)

if bukw_2 != "0" * len(bukw_2):
    print("ЦВЕТ СИМВОЛОВ")
    encode(bukw_2)

if bukw_3 != "0" * len(bukw_3):
    print("ЦВЕТ ФОНА")
    encode(bukw_3)

if bukw_4 != "0" * len(bukw_4):
    print("МЕЖСИМВОЛЬНЫЙ ИНТЕРВАЛ")
    encode(bukw_4)

if bukw_5 != "0" * len(bukw_5):
    print("МАСШТАБ ШРИФТА")
    encode(bukw_5)
