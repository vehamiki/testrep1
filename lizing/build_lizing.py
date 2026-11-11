"""Kiegészítő tájékoztató a lízingkérelemhez – Mac 180 tésztagyártó gépsor.

Futtatás:  python3 build_lizing.py   ->  .docx, majd LibreOffice-szal PDF.
"""
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT = Path(__file__).resolve().parent
NAME = "Naturtrade_lizing_kiegeszito_tajekoztato_Mac180_2026-11-11"

TEAL = RGBColor(0x2F, 0x7F, 0x72)
INK = RGBColor(0x1F, 0x2D, 0x2B)
MUTED = RGBColor(0x5A, 0x6B, 0x68)
FILL_KPI = "E6F4F0"
FILL_ALT = "F5F9F8"
FILL_SUM = "D3E7E2"
BORDER = "C9DED8"
FONT = "Carlito"

# ---- Finanszírozás (a két korábbi lízingajánlattal azonos díjszorzó) ----
PRICE = 40_000.00
DOWN = PRICE * 0.10
FIN = PRICE - DOWN
TERM = 36  # hónap
# Effektív éves kamat a két korábbi 60 hónapos ajánlatból visszaszámolva
# (havi díj / finanszírozott összeg = 0,019540  ->  kb. 6,444% annuitás).
EFF_RATE = 0.06444
_r = EFF_RATE / 12
MONTHLY = round(FIN * _r / (1 - (1 + _r) ** -TERM), 2)
INTEREST = round(MONTHLY * TERM - FIN)
FX = 390


def huf(x, dec=0):
    s = f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")
    return s


doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(1.9)
sec.top_margin = Cm(1.6)
sec.bottom_margin = Cm(1.6)

st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10.5)
st.font.color.rgb = INK
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.12


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def borders(table, color=BORDER, inside=True):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
    for e in edges:
        el = OxmlElement(f"w:{e}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)


def cell_margins(table, top=40, bottom=40, left=80, right=80):
    tblPr = table._tbl.tblPr
    m = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        m.append(el)
    tblPr.append(m)


def fix_widths(t, widths):
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    grid = t._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w.twips)))
    for col, w in zip(t.columns, widths):
        col.width = w


def run(p, text, bold=False, italic=False, size=None, color=None):
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.name = FONT
    if size:
        r.font.size = Pt(size)
    if color is not None:
        r.font.color.rgb = color
    return r


def para(text="", bold=False, size=None, color=None, after=4, before=0, align=None, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    if align:
        p.alignment = align
    if text:
        run(p, text, bold=bold, size=size, color=color, italic=italic)
    return p


def rich(parts, after=4, before=0):
    """parts: list of (text, bold)"""
    p = para(after=after, before=before)
    for t, b in parts:
        run(p, t, bold=b)
    return p


def h2(text):
    p = para(text, bold=True, size=15, color=TEAL, before=12, after=5)
    p.paragraph_format.keep_with_next = True
    return p


def h3(text):
    p = para(text, bold=True, size=11.5, color=INK, before=6, after=3)
    p.paragraph_format.keep_with_next = True
    return p


def bullet(parts):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    if isinstance(parts, str):
        parts = [(parts, False)]
    for t, b in parts:
        run(p, t, bold=b)
    return p


def table(rows, widths, header=True, num_cols=(), bold_rows=(), size=9, valign_top=False):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t)
    cell_margins(t)
    for i, row in enumerate(rows):
        tr = t.rows[i]._tr
        trPr = tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.width = widths[j]
            c.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.TOP if valign_top else WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            is_head = header and i == 0
            is_bold = is_head or i in bold_rows
            if j in num_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if is_head and j > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j not in num_cols else WD_ALIGN_PARAGRAPH.RIGHT
            run(
                p,
                val,
                bold=is_bold,
                size=size,
                color=RGBColor(0xFF, 0xFF, 0xFF) if is_head else INK,
            )
            if is_head:
                shade(c, "2F7F72")
            elif i in bold_rows:
                shade(c, FILL_SUM)
            elif i % 2 == 0:
                shade(c, FILL_ALT)
    fix_widths(t, widths)
    if header:
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        trPr.append(th)
    return t


def footer():
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(p, "Naturtrade Hungary Kft. · Kiegészítő tájékoztató lízingkérelemhez · ", size=8, color=MUTED)
    r = p.add_run()
    r.font.size = Pt(8)
    r.font.color.rgb = MUTED
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), tag)
            r._r.append(f)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            r._r.append(it)


footer()

# ---------------------------------------------------------------- Fejléc
para("NATURTRADE HUNGARY KFT.", bold=True, size=9, color=TEAL, after=2)
para("Kiegészítő tájékoztató a lízingkérelemhez", bold=True, size=20, color=INK, after=1)
para(
    "Tésztagyártó gépsor (Mac 180 présgép, használt, 2018) · Szeged, 2026. november 11.",
    size=10,
    color=MUTED,
    after=8,
)
para("Tisztelt Partnerünk!", after=4)
para(
    "Köszönjük megkeresésüket. Az alábbiakban összefoglaljuk a kért dokumentumokat és válaszainkat. "
    "A számszerű adatok a 2025. évi lezárt és a 2026. szeptember 30-i előzetes főkönyvi kivonaton, "
    "valamint a vállalatirányítási rendszerünk október végi adatain alapulnak.",
    after=8,
)

# KPI sor
kpis = [
    ("716 M Ft", "árbevétel 2026 I–III. n.év,\na 2025. évi teljes árbevétel 98%-a"),
    ("~1,09 Mrd Ft", "várható 2026. évi árbevétel\n(+49% a 2025. évhez képest)"),
    ("45 M Ft", "adózás előtti eredmény 9 hónap alatt\n(2025 teljes év: 39 M Ft)"),
    ("53%", "exportarány 2026-ban\n(2025: 34%), EUR-alapú bevétel"),
]
kt = doc.add_table(rows=1, cols=4)
kt.alignment = WD_TABLE_ALIGNMENT.CENTER
cell_margins(kt, 80, 80, 110, 110)
tblPr = kt._tbl.tblPr
b = OxmlElement("w:tblBorders")
for e in ("top", "left", "bottom", "right", "insideV"):
    el = OxmlElement(f"w:{e}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), "18")
    el.set(qn("w:color"), "FFFFFF")
    b.append(el)
tblPr.append(b)
fix_widths(kt, [Cm(4.3)] * 4)
for j, (big, small) in enumerate(kpis):
    c = kt.cell(0, j)
    c.width = Cm(4.3)
    shade(c, FILL_KPI)
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(1)
    run(p, big, bold=True, size=16, color=TEAL)
    p2 = c.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.0
    run(p2, small, size=8, color=INK)

# ---------------------------------------------------------------- Társaság
h2("A társaságról röviden")
para(
    "A Naturtrade Hungary Kft. szegedi élelmiszer-gyártó vállalat. Négy generációs malomipari háttérrel "
    "saját maga állítja elő legfontosabb alapanyagát, a prémium szemcseméretű pszeudogabona- és "
    "hüvelyeslisztet, és ebből gyárt gluténmentes, vegán termékeket. 27 munkatársával 100-nál több "
    "terméket kínál öt kategóriában:"
)
bullet([("Lisztek és lisztkeverékek: ", True),
        ("15 féle egykomponensű liszt és 15 féle konyhakész, sütésre kész lisztkeverék.", False)])
bullet([("Tészták, instant ételek és növényi premixek: ", True),
        ("30 féle gluténmentes tészta, instant levesek és tészták, valamint növényi alapú sajt- és hús-premixek.", False)])
bullet([("Értékesítés: ", True),
        ("több mint 250 hazai viszonteladó (lánc, nagykereskedés, drogéria, bio- és reformbolt), saját "
         "webáruház, export az Egyesült Királyságba, Nyugat-Európába és Szaúd-Arábiába, valamint bérgyártás "
         "más márkák számára.", False)])
bullet([("Innováció: ", True),
        ("folyamatos, támogatásokkal is elismert K+F tevékenység; 2026-ban 36 M Ft fejlesztés aktiválva.", False)])

h3("Negyedéves nettó árbevétel, 2025 vs. 2026 (M Ft)")
W4 = [Cm(7.6), Cm(3.2), Cm(3.2), Cm(3.2)]
table(
    [
        ["Időszak", "2025", "2026", "Változás"],
        ["I. negyedév", "136,9", "202,9", "+48%"],
        ["II. negyedév", "157,4", "231,7", "+47%"],
        ["III. negyedév", "159,1", "277,2", "+74%"],
        ["I–III. negyedév összesen", "453,4", "711,8", "+57%"],
        ["IV. negyedév (2026: várható)", "266,4", "381,8", "+43%"],
        ["Teljes év (2026: várható)", "719,8", "1 093,6", "+52%"],
    ],
    W4, num_cols=(1, 2, 3), bold_rows=(4, 6),
)
para(
    "Forrás: vállalatirányítási rendszer, számlázott nettó árbevétel (a főkönyv szerinti 2025. évi árbevétel 732 M Ft).",
    size=8, color=MUTED, italic=True, before=3, after=4,
)

# ---------------------------------------------------------------- 1.
h2("1. Aktuális főkönyvi kivonat")
para(
    "Mellékeljük a 2026. szeptember 30-i fordulónapú előzetes főkönyvi kivonatot (készült: 2026. október 6.), "
    "amely 3 hónapnál nem régebbi. Összehasonlításként mellékeljük a 2025. évi lezárt kivonatot is. "
    "A legfontosabb állományadatok:"
)
table(
    [
        ["Mérlegtétel (M Ft)", "2025.12.31", "2026.09.30", "Változás"],
        ["Befektetett eszközök", "792,9", "796,8", "+0%"],
        ["Készletek", "95,6", "107,8", "+13%"],
        ["Vevőkövetelések (nettó)", "44,5", "252,8", "export felfutás"],
        ["Saját tőke (tárgyévi eredménnyel)", "409,0", "454,0", "+11%"],
        ["Hosszú lejáratú hitelek és lízingek", "203,6", "192,2", "–6%"],
        ["Vevőktől kapott előlegek", "0,0", "51,8", "új"],
    ],
    W4, num_cols=(1, 2, 3),
)

# ---------------------------------------------------------------- 2.
h2("2. A beruházás indoklása")
para(
    "A finanszírozandó eszköz egy komplett, használt ipari tésztagyártó gépsor, amely a gluténmentes "
    "tésztagyártásunk kapacitását bővíti. A gépsor a saját gyártású pszeudogabona- és hüvelyeslisztünket "
    "magasabb hozzáadott értékű késztermékké, tésztává alakítja, így a lisztüzem kibocsátásának egyre nagyobb "
    "része a legjobb árrésű termékkategóriánkban értékesíthető. A 2018-as gyártású gépsort a szállító telepíti "
    "és üzembe helyezi, így új berendezés áránál lényegesen kedvezőbben és rövid átfutással bővül a kapacitás."
)
W2 = [Cm(4.2), Cm(13.0)]
table(
    [
        ["", "Tésztagyártó gépsor (Mac 180 présgép, 2018)"],
        ["Eszköz összetétele",
         "1 db Mac 180 présgép, 1 db zárt rendszerű hűtő, 1 db szállítószalag, 1 db előszárító tálca "
         "adagolóval és tálca-továbbító sínnel, 10 db matrica, vágókések. Gyártási év: 2018."],
        ["Szállító / ajánlat",
         "Brio Kft. (Békésszentandrás), ajánlat száma 28-26, kelte 2026.10.28., érvényes 30 napig. "
         "Paritás: vevő telephelye; a csomagolást, szállítást, telepítést és beüzemelést az ár tartalmazza."],
        ["Nettó ár / finanszírozás", f"{huf(PRICE)} EUR / {huf(FIN)} EUR, {TERM} hó"],
        ["Havi díj", f"{huf(MONTHLY, 2)} EUR"],
        ["Felhasználás",
         "Gluténmentes tészták folyamatos, gépi gyártása (préselés, formázás a cserélhető matricákkal, vágás, "
         "előszárítás). A 10 db matrica több tésztaforma gyártását teszi lehetővé átállással, így a hazai "
         "kiskereskedelmi, export- és bérgyártási megrendelések egyaránt kiszolgálhatók. A zárt rendszerű hűtő "
         "stabil, ellenőrzött gyártási körülményeket biztosít."],
        ["Hatás a jövedelemre",
         "Nagyobb tésztagyártó kapacitás, a saját liszt magasabb hozzáadott értékű feldolgozása, kevesebb kézi "
         "munka a formázásban és az előszárításba adagolásban, valamint a növekvő export- és bérgyártási "
         "tésztarendelések kiszolgálása arányos létszámbővítés nélkül."],
        ["Újrahasznosíthatóság",
         "Elterjedt típusú ipari tésztaprés, aktív hazai és európai használtgép-piaccal. A gépsor egységei "
         "(prés, hűtő, szalag, előszárító) külön is értékesíthetők, bonthatók és áttelepíthetők. "
         "Gluténmentes és hagyományos tészta gyártására egyaránt alkalmas, így a vevőkör széles."],
    ],
    W2, header=True, size=9, valign_top=True,
)

h3("A finanszírozás főbb adatai (EUR-alapú, fix kamatozású zártvégű pénzügyi lízing):")
table(
    [
        ["EUR", "Tésztagyártó gépsor"],
        ["Nettó beszerzési ár", huf(PRICE)],
        ["Önerő (10%, nettó)", huf(DOWN)],
        ["Finanszírozott összeg", huf(FIN)],
        ["Futamidő / kamat", f"{TERM} hó / {huf(EFF_RATE * 100, 2)}% fix"],
        ["Havi díj", huf(MONTHLY, 2)],
        ["Összes kamat a futamidő alatt", huf(INTEREST)],
    ],
    [Cm(9.0), Cm(5.0)], num_cols=(1,), bold_rows=(5,),
)
para(after=2)

yearly_huf = MONTHLY * 12 * FX / 1e6
rich(
    [
        ("Megtérülés. ", True),
        (f"A gépsor lízingdíja havi {huf(MONTHLY, 2)} EUR, ami évente kb. {huf(yearly_huf, 1)} M Ft "
         f"({FX} Ft/EUR árfolyamon). Óvatosan, a liszt átlagos 200 Ft/kg-os nettó árrésével számolva már heti "
         "1 tonna többlet gluténmentes tészta évi kb. 10 M Ft fedezetet termel; a tészta árrése ennél magasabb, "
         "mert a saját liszt feldolgozott formában kerül értékesítésre. A beruházás így várhatóan kb. 1,5–2 év "
         "alatt megtérül, a többletfedezet pedig már a futamidő alatt is kb. kétszeresen fedezi a lízingdíjat. A lízingdíjat EUR-ban "
         "fizetjük, és EUR-alapú exportbevételünk ennek sokszorosát fedezi, így árfolyamkockázatunk nincs. "
         f"A 10%-os önerőt ({huf(DOWN)} EUR + ÁFA) saját forrásból biztosítjuk.", False),
    ],
    after=6,
)

# ---------------------------------------------------------------- 3.
h2("3. Eredményterv 2026–2027").paragraph_format.page_break_before = True
table(
    [
        ["M Ft", "2025 tény", "2026 I–III. tény", "2026 eredeti terv", "2026 várható", "2027 terv"],
        ["Belföldi árbevétel", "480,7", "336,5", "539,7", "494,7", "563,4"],
        ["Export árbevétel", "251,3", "379,6", "333,6", "598,9", "849,3"],
        ["Értékesítés nettó árbevétele", "732,0", "716,1", "873,3", "1 093,6", "1 412,7"],
        ["Egyéb bevételek (főként támogatás)", "56,1", "21,9", "96,3", "29,2", "25,0"],
        ["Anyagjellegű ráfordítások", "–442,5", "–461,4", "–529,9", "–704,6", "–904,1"],
        ["Személyi jellegű ráfordítások", "–125,3", "–145,5", "–197,2", "–199,8", "–223,8"],
        ["Egyéb ráfordítások", "–38,9", "–7,0", "–41,3", "–9,4", "–12,0"],
        ["EBITDA", "181,3", "124,1", "201,2", "209,0", "297,8"],
        ["Működési EBITDA (egyéb bevételek nélkül)", "125,2", "102,2", "104,9", "179,9", "272,8"],
        ["Értékcsökkenés", "–150,8", "–77,0", "–159,4", "–139,6", "–195,3"],
        ["Üzemi eredmény", "30,6", "47,2", "41,8", "69,4", "102,5"],
        ["Pénzügyi műveletek eredménye", "8,2", "–2,2", "–3,7", "–3,0", "–14,2"],
        ["Adózás előtti eredmény", "38,7", "45,0", "38,1", "66,5", "88,3"],
        ["Adózott eredmény", "37,9", "45,0", "34,7", "60,5", "80,4"],
    ],
    [Cm(6.2), Cm(2.2), Cm(2.2), Cm(2.2), Cm(2.2), Cm(2.2)],
    num_cols=(1, 2, 3, 4, 5), bold_rows=(3, 8, 13), size=8.5,
)
para(after=0)
para("Feltételezések:", bold=True, after=2)
bullet([("2026 várható: ", True),
        ("9 havi tény, plusz az október havi tényleges számlázás (94,6 M Ft), plusz a november–decemberi "
         "szezonális csúcs. A IV. negyedévi árbevétel kb. 380 M Ft (2025 IV. n.év: 266 M Ft). A költségarányok "
         "a 9 havi tényből számoltak, a bérköltség a csúcsszezon többletmunkaidejével. Az értékcsökkenés az év "
         "végi elszámolással és az aktivált K+F-fel együtt kb. 140 M Ft (2025: 151 M Ft).", False)])
bullet([("2027 terv: ", True),
        ("1,41 Mrd Ft árbevétel (+29%), 60%-os exportaránnyal, 64%-os anyagjellegű ráfordítás-aránnyal és "
         "12%-os bérnövekedéssel. Az értékcsökkenés kb. 195 M Ft a kapacitásbővítés új eszközeivel, a pénzügyi "
         "ráfordítás a többletfinanszírozás miatt nő. A tervezett zöldmezős kapacitásbővítés tőkebevonásból "
         "valósul meg, hatását ez a terv nem tartalmazza.", False)])
bullet("Társasági adó óvatosan 9%-kal számolva, adókedvezmények nélkül.")

# ---------------------------------------------------------------- 4.
h2("4. Válaszok a kérdéseikre")
h3("4.1 Várható árbevétel és eredmény, rendelésállomány, terv–tény eltérés")
para(
    "2026-ban kb. 1,09 Mrd Ft árbevételre (+49%) és 66 M Ft adózás előtti eredményre számítunk (2025: 732 M Ft, "
    "illetve 39 M Ft). Szeptember végéig már a teljes 2025. évi árbevétel 98%-át realizáltuk, októberben pedig "
    "további 94,6 M Ft-ot számláztunk."
)
rich([("Rendelésállomány. ", True),
      ("Belföldi B2B megrendeléseink 3–10 napon belül teljesülnek, ezért a nyitott belföldi rendelésállomány "
       "jellemzően rövid távú. Bevételünk stabilitását a fő exportpartnerrel kötött keretmegállapodás adja. "
       "Ennek alapján 2026 februárja óta 2–3 hetente érkeznek ismétlődő megrendelések, és 2026-ban eddig közel "
       "40 szállítást számláztunk a partnernek. Az október–novemberi szállítások értéke már meghaladja a 70 M "
       "Ft-ot, az exportmegrendelések folyamatosan bővülnek.", False)])
rich([("Terv–tény. ", True),
      ("A 9 havi árbevétel +14%-kal haladja meg az időarányos tervet (629 M Ft), az export +63%-kal. Az adózás "
       "előtti eredmény 45 M Ft, szemben az időarányos 16 M Ft-os tervvel. A tervet két ponton korrigáltuk: "
       "(1) az export arányát a tervezett 38%-ról 53%-ra emeltük, és ehhez igazítottuk a logisztikai költségeket; "
       "(2) a zöldmezős beruházás ütemezése 2027-re tolódott, ezért az értékcsökkenés a tervezettnél alacsonyabb. "
       "A várható éves árbevétel (1,09 Mrd Ft) +25%-kal, az EBITDA +4%-kal, az adózás előtti eredmény +75%-kal "
       "haladja meg az eredeti éves tervet.", False)])

h3("4.2 Jövedelmezőség 2026-ban")
para(
    "A jövedelmezőség javul: az adózás előtti eredmény +72%-kal, a működési EBITDA +44%-kal nő 2025-höz képest, "
    "mert a működés egyre kevésbé függ támogatásoktól. Az anyagköltség az árbevétel 41%-a. Ez a saját "
    "lisztgyártásnak köszönhető, mivel a legfontosabb alapanyagot önköltségen állítjuk elő. A gépesítés (köztük "
    "a most finanszírozandó tésztagyártó gépsor) a termelékenységet tovább javítja, és a saját lisztet magasabb "
    "hozzáadott értékű termékként értékesíthetjük. A saját tőke 454 M Ft, a mérlegfőösszeg kb. 36%-a."
)

h3("4.3 Piaci körülmények")
bullet([("Kereslet: ", True),
        ("a gluténmentes élelmiszerpiac Európában évi 9–10%-kal, a Közel-Keleten kb. 10%-kal nő. Saját gyártású, "
         "prémium pszeudogabona-lisztjeink, lisztkeverékeink és tésztáink exportban versenyképesek, az export "
         "aránya 2025-ben 34%, 2026-ban 53%.", False)])
bullet([("Fizetési morál: ", True),
        ("fő exportpartnerünk számlái azonnali esedékességűek, és a partner 2026-ban valamennyi szeptember 30. "
         "előtti számláját kiegyenlítette. Szeptember végén ezen felül 52 M Ft vevői előleg állt rendelkezésünkre. "
         "Hazai lánc- és nagykereskedelmi vevőink 30–60 napos határidővel, fegyelmezetten fizetnek. 2026-ban "
         "követelés miatt értékvesztést nem kellett elszámolni.", False)])
bullet([("TOP vevők: ", True),
        ("a legnagyobb vevő részaránya a szaúdi ipari gluténmentes pékséggel kötött keretmegállapodás felfutásával "
         "kb. 18%-ról (2025) kb. 50%-ra (2026) nőtt. A partner mögött állami megrendelő áll, és előleges fizetéssel "
         "dolgozik. A kockázatot több mint 250 hazai viszonteladó, a hazai kiskereskedelmi láncok stabil forgalma, "
         "a bérgyártás és a további exportpiacok (Egyesült Királyság, Nyugat-Európa, Öböl-térség) bővítése "
         "ellensúlyozza.", False)])
bullet([("Kockázatkezelés: ", True),
        ("az exportbevétel EUR-alapú, ez természetes fedezetet ad az EUR-alapú lízingdíjra. A saját "
         "alapanyaggyártás mérsékli a beszerzési árak ingadozásának hatását.", False)])

h3("4.4 Likviditás")
para(
    "Likviditási helyzetünk stabil. A szeptember végi vevőállomány az exportforgalom felfutását tükrözi, és a fő "
    "exportpartner számlái rendre, rövid határidőn belül befolynak. A forgóeszköz-igényt a meglévő rulírozó és "
    "folyószámlahitel-keretek, a vevői előlegek és a IV. negyedév szezonálisan erős bevétele fedezik. A hosszú "
    "lejáratú hitel- és lízingállomány 2026-ban csökkent (204 M Ft-ról 192 M Ft-ra). A várható EBITDA a teljes "
    "éves adósságszolgálat (tőke és kamat, az új lízinggel együtt) kb. 2,6-szeresét fedezi."
)
para(
    "Új folyószámlahitel felvételét vagy a meglévő keret bővítését nem tervezzük. A gépsor önerejét és "
    "lízingdíját a folyó működésből finanszírozzuk."
)

# ---------------------------------------------------------------- 5.
h2("5. Összegzés: miért biztonságos a finanszírozás")
table(
    [
        ["Szempont", "Adat"],
        ["Növekedés", "árbevétel +49% 2026-ban (1,09 Mrd Ft), 2027-re 1,41 Mrd Ft a terv; I–III. negyedév +57%"],
        ["Eredményesség", "adózás előtti eredmény 9 hónap alatt 45 M Ft, ami már meghaladja a teljes 2025. évit"],
        ["Adósságszolgálati fedezet", "a várható EBITDA a teljes éves adósságszolgálat kb. 2,6-szerese (az új lízinggel együtt)"],
        ["Tőkeerő", "saját tőke 454 M Ft, a mérlegfőösszeg kb. 36%-a"],
        ["Devizafedezet", "53%-os, EUR-alapú exportbevétel az EUR-alapú lízingdíjjal szemben"],
        ["Biztosítéki érték",
         f"komplett ipari tésztagyártó gépsor, {huf(PRICE)} EUR nettó értékben; elterjedt típus aktív "
         "használtgép-piaccal, egységei külön is értékesíthetők; 3 éves futamidő"],
    ],
    [Cm(4.6), Cm(12.6)], size=9,
)

para(
    "Bízunk benne, hogy a fenti információk alapján kedvezően bírálják el kérelmünket. Kérdés esetén szívesen "
    "állunk rendelkezésükre.",
    before=6,
)
para("Tisztelettel:", before=4, after=10)
para("Véha Miklós", bold=True, after=0)
para("ügyvezető, Naturtrade Hungary Kft.", after=4)
para(
    "Mellékletek: 2026.09.30-i előzetes főkönyvi kivonat · 2025. évi főkönyvi kivonat · a tésztagyártó gépsor "
    "árajánlata (Brio Kft., 28-26) és finanszírozási ajánlata",
    size=8.5, color=MUTED,
)

docx_path = OUT / f"{NAME}.docx"
doc.save(docx_path)
subprocess.run(
    ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(OUT), str(docx_path)],
    check=True, capture_output=True,
)
print(OUT / f"{NAME}.pdf", MONTHLY, INTEREST, round(yearly_huf, 2))
