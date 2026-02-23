#!/usr/bin/env python3
"""
Generuje prezentacje Trening Wechowy w formacie PDF - styl Aromagic.
Kazdy slajd to osobna strona A4 landscape.
"""

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
import os

# === COLORS ===
PURPLE = HexColor("#7E57C2")
PURPLE_SOFT = HexColor("#F3EEFA")
PURPLE_DARK = HexColor("#5E35B1")
GREEN = HexColor("#4CAF50")
GREEN_SOFT = HexColor("#E8F5E9")
TEXT = HexColor("#111827")
TEXT_SEC = HexColor("#6B7280")
TEXT_MUTED = HexColor("#9CA3AF")
BG = HexColor("#FFFFFF")
BG_SOFT = HexColor("#F9FAFB")
BORDER = HexColor("#E5E7EB")

# === PAGE ===
W, H = landscape(A4)
MARGIN = 40
CONTENT_W = W - 2 * MARGIN

# === OUTPUT ===
OUTDIR = os.path.dirname(os.path.abspath(__file__))
OUTFILE = os.path.join(OUTDIR, "Trening W\u0119chowy \u2014 Prezentacja \u2014 Aromagic.pdf")
LOGO_PATH = os.path.join(OUTDIR, "aromagic_logo.png")

# Polish quotes
LQ = "\u201E"  # opening lower
RQ = "\u201D"  # closing upper


def register_fonts():
    font_dir = os.path.expanduser("~/Library/Fonts")
    inter_files = {
        "Inter": "Inter_24pt-Regular.ttf",
        "Inter-Bold": "Inter_28pt-Bold.ttf",
        "Inter-SemiBold": "Inter_24pt-SemiBold.ttf",
        "Inter-Medium": "Inter_24pt-Medium.ttf",
    }
    found = False
    for name, filename in inter_files.items():
        path = os.path.join(font_dir, filename)
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont(name, path))
            found = True
    if not found:
        for base in ["/System/Library/Fonts", "/Library/Fonts"]:
            for name, filename in inter_files.items():
                path = os.path.join(base, filename)
                if os.path.exists(path):
                    pdfmetrics.registerFont(TTFont(name, path))
                    found = True
    return found


HAS_INTER = register_fonts()
FONT = "Inter" if HAS_INTER else "Helvetica"
FONT_BOLD = "Inter-Bold" if HAS_INTER else "Helvetica-Bold"
FONT_SEMI = "Inter-SemiBold" if HAS_INTER else "Helvetica-Bold"
FONT_MED = "Inter-Medium" if HAS_INTER else "Helvetica"

# Bottom margin for content
BOTTOM = MARGIN


class SlideBuilder:
    def __init__(self):
        self.c = canvas.Canvas(OUTFILE, pagesize=landscape(A4))
        self.c.setTitle("Trening W\u0119chowy \u2014 Aromagic")
        self.slide_num = 0

    def new_slide(self, bg_color=BG):
        if self.slide_num > 0:
            self.c.showPage()
        self.slide_num += 1
        self.c.setFillColor(bg_color)
        self.c.rect(0, 0, W, H, fill=1, stroke=0)

    def draw_logo(self, y=None):
        if y is None:
            y = H - MARGIN
        if os.path.exists(LOGO_PATH):
            self.c.drawImage(LOGO_PATH, MARGIN, y - 18, width=72, height=18,
                           preserveAspectRatio=True, mask="auto")
            return y - 26
        return y - 10

    def draw_pill(self, text, x, y):
        self.c.setFont(FONT_SEMI, 9)
        tw = self.c.stringWidth(text, FONT_SEMI, 9)
        pill_w = tw + 22
        pill_h = 20
        self.c.setFillColor(PURPLE_SOFT)
        self.c.roundRect(x, y - pill_h + 4, pill_w, pill_h, 10, fill=1, stroke=0)
        self.c.setFillColor(PURPLE)
        self.c.drawString(x + 11, y - 9, text)
        return y - pill_h - 8

    def draw_title(self, text, y, size=28, color=TEXT):
        self.c.setFont(FONT_BOLD, size)
        self.c.setFillColor(color)
        self.c.drawString(MARGIN, y, text)
        return y - size - 10

    def draw_sub(self, text, y, size=14, color=TEXT_SEC, max_width=None):
        if max_width is None:
            max_width = CONTENT_W
        style = ParagraphStyle("sub", fontName=FONT, fontSize=size, textColor=color, leading=size * 1.5)
        p = Paragraph(text, style)
        pw, ph = p.wrap(max_width, 200)
        p.drawOn(self.c, MARGIN, y - ph)
        return y - ph - 6

    def draw_body(self, text, y, x=None, size=13, color=TEXT_SEC, max_width=None):
        if x is None:
            x = MARGIN
        if max_width is None:
            max_width = CONTENT_W
        style = ParagraphStyle("body", fontName=FONT, fontSize=size, textColor=color, leading=size * 1.5)
        p = Paragraph(text, style)
        pw, ph = p.wrap(max_width, 300)
        p.drawOn(self.c, x, y - ph)
        return y - ph - 4

    def draw_card(self, x, y, w, h, title=None, body=None, accent_color=None, body_size=13):
        self.c.setFillColor(BG)
        self.c.setStrokeColor(BORDER)
        self.c.setLineWidth(0.5)
        self.c.roundRect(x, y - h, w, h, 8, fill=1, stroke=1)
        if accent_color:
            self.c.setFillColor(accent_color)
            self.c.rect(x, y - h + 8, 3, h - 16, fill=1, stroke=0)
        inner_x = x + 16 + (6 if accent_color else 0)
        inner_w = w - 32 - (6 if accent_color else 0)
        cy = y - 18
        if title:
            title_color = accent_color if accent_color else PURPLE
            self.c.setFont(FONT_SEMI, 14)
            self.c.setFillColor(title_color)
            self.c.drawString(inner_x, cy, title)
            cy -= 24
        if body:
            style = ParagraphStyle("cb", fontName=FONT, fontSize=body_size, textColor=TEXT_SEC, leading=body_size * 1.5)
            p = Paragraph(body, style)
            pw, ph = p.wrap(inner_w, 200)
            p.drawOn(self.c, inner_x, cy - ph)
        return y - h - 10

    def draw_accent_box(self, text, y, bg=PURPLE_SOFT, text_color=TEXT, max_width=None, font_size=12):
        if max_width is None:
            max_width = CONTENT_W
        style = ParagraphStyle("ab", fontName=FONT, fontSize=font_size, textColor=text_color, leading=font_size * 1.5)
        p = Paragraph(text, style)
        pw, ph = p.wrap(max_width - 36, 200)
        box_h = ph + 24
        self.c.setFillColor(bg)
        self.c.roundRect(MARGIN, y - box_h, max_width, box_h, 8, fill=1, stroke=0)
        p.drawOn(self.c, MARGIN + 18, y - box_h + 12)
        return y - box_h - 10

    def draw_blockquote(self, text, cite, y, font_size=13):
        style = ParagraphStyle("bq", fontName=FONT, fontSize=font_size, textColor=TEXT, leading=font_size * 1.6)
        p = Paragraph("<i>" + text + "</i>", style)
        pw, ph = p.wrap(CONTENT_W - 50, 200)
        box_h = ph + 36
        self.c.setFillColor(PURPLE_SOFT)
        self.c.roundRect(MARGIN, y - box_h, CONTENT_W, box_h, 8, fill=1, stroke=0)
        self.c.setFillColor(PURPLE)
        self.c.rect(MARGIN, y - box_h + 4, 4, box_h - 8, fill=1, stroke=0)
        p.drawOn(self.c, MARGIN + 24, y - box_h + 22)
        self.c.setFont(FONT, 9)
        self.c.setFillColor(TEXT_MUTED)
        self.c.drawRightString(MARGIN + CONTENT_W - 18, y - box_h + 8, cite)
        return y - box_h - 10

    def draw_table(self, headers, rows, y, col_widths=None, font_size=11):
        if col_widths is None:
            col_widths = [CONTENT_W / len(headers)] * len(headers)
        data = [headers] + rows
        style_cmds = [
            ("FONTNAME", (0, 0), (-1, 0), FONT_SEMI),
            ("FONTSIZE", (0, 0), (-1, 0), font_size),
            ("FONTNAME", (0, 1), (-1, -1), FONT),
            ("FONTSIZE", (0, 1), (-1, -1), font_size),
            ("TEXTCOLOR", (0, 0), (-1, 0), TEXT),
            ("TEXTCOLOR", (0, 1), (-1, -1), TEXT_SEC),
            ("BACKGROUND", (0, 0), (-1, 0), BG_SOFT),
            ("BACKGROUND", (0, 1), (-1, -1), BG),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ("LINEBELOW", (0, 0), (-1, -2), 0.5, BORDER),
            ("ROUNDEDCORNERS", [8, 8, 8, 8]),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER),
        ]
        t = Table(data, colWidths=col_widths)
        t.setStyle(TableStyle(style_cmds))
        tw, th = t.wrap(CONTENT_W, 300)
        t.drawOn(self.c, MARGIN, y - th)
        return y - th - 10

    def draw_checklist(self, items, y, font_size=14):
        x = MARGIN
        for item in items:
            self.c.setFillColor(PURPLE_SOFT)
            self.c.roundRect(x, y - 12, 16, 16, 3, fill=1, stroke=0)
            self.c.setFont(FONT_BOLD, 11)
            self.c.setFillColor(PURPLE)
            self.c.drawString(x + 3, y - 8, "\u2713")
            style = ParagraphStyle("cl", fontName=FONT, fontSize=font_size, textColor=TEXT_SEC, leading=font_size * 1.5)
            p = Paragraph(item, style)
            pw, ph = p.wrap(CONTENT_W - 34, 60)
            p.drawOn(self.c, x + 26, y - ph + 2)
            y -= max(ph, 20) + 8
        return y

    def draw_ordered_list(self, items, y, font_size=14):
        x = MARGIN
        for i, item in enumerate(items, 1):
            self.c.setFont(FONT_BOLD, 13)
            self.c.setFillColor(PURPLE)
            self.c.drawString(x, y - 2, str(i) + ".")
            style = ParagraphStyle("ol", fontName=FONT, fontSize=font_size, textColor=TEXT_SEC, leading=font_size * 1.5)
            p = Paragraph(item, style)
            pw, ph = p.wrap(CONTENT_W - 34, 60)
            p.drawOn(self.c, x + 26, y - ph + 2)
            y -= max(ph, 20) + 8
        return y

    def draw_stat_card(self, x, y, w, h, number, label, num_size=36):
        self.c.setFillColor(BG)
        self.c.setStrokeColor(BORDER)
        self.c.setLineWidth(0.5)
        self.c.roundRect(x, y - h, w, h, 8, fill=1, stroke=1)
        self.c.setFont(FONT_BOLD, num_size)
        self.c.setFillColor(PURPLE)
        self.c.drawCentredString(x + w / 2, y - num_size - 14, number)
        self.c.setFont(FONT, 13)
        self.c.setFillColor(TEXT_SEC)
        self.c.drawCentredString(x + w / 2, y - num_size - 34, label)
        return y - h - 10

    def save(self):
        self.c.save()


def build():
    s = SlideBuilder()

    # SLIDE 1: Tytul
    s.new_slide()
    y = s.draw_logo(H - MARGIN)
    y = s.draw_pill("Aromapsychologia", MARGIN, y)
    y -= 20
    s.c.setFont(FONT_BOLD, 42)
    s.c.setFillColor(PURPLE)
    s.c.drawString(MARGIN, y, "Trening w\u0119chowy")
    y -= 50
    s.c.setFont(FONT_BOLD, 42)
    s.c.setFillColor(TEXT)
    s.c.drawString(MARGIN, y, "w warunkach domowych")
    y -= 50
    s.c.setFont(FONT, 14)
    s.c.setFillColor(TEXT_SEC)
    s.c.drawString(MARGIN, y, "Opracowanie: Emilia Chodorowska")
    y -= 22
    s.c.setFont(FONT, 12)
    s.c.setFillColor(TEXT_MUTED)
    s.c.drawString(MARGIN, y, "na podstawie kursu Aromapsychologia Anny Bober")

    # SLIDE 2: Dlaczego nos zamilkl
    s.new_slide(BG_SOFT)
    y = s.draw_logo()
    y = s.draw_pill("1 \u00b7 Wst\u0119p", MARGIN, y)
    y = s.draw_title("Dlaczego Tw\u00f3j nos " + LQ + "zamilk\u0142" + RQ + "?", y, size=28)
    y = s.draw_sub("Utrata w\u0119chu w COVID-19 to zjawisko inne ni\u017c zatkany nos przy grypie.", y)
    y -= 6
    card_w = (CONTENT_W - 14) / 2
    avail = y - BOTTOM
    card_top_h = int(avail * 0.48)
    card_bot_h = int(avail * 0.42)
    s.draw_card(MARGIN, y, card_w, card_top_h, "Grypa",
                "Obrz\u0119k tkanek fizycznie blokuje dost\u0119p aromat\u00f3w do nab\u0142onka w\u0119chowego. Nos jest zatkany \u2014 powietrze nie przechodzi.")
    s.draw_card(MARGIN + card_w + 14, y, card_w, card_top_h, "COVID-19",
                "Dro\u017cne przewody nosowe, ale wirus atakuje <b>kom\u00f3rki podporowe</b> i <b>gruczo\u0142y Bowmana</b>. Zapach nie dociera mimo wolnych dr\u00f3g oddechowych.")
    y -= card_top_h + 12
    s.draw_card(MARGIN, y, CONTENT_W, card_bot_h, None,
                "Neurony trac\u0105 " + LQ + "system podtrzymywania \u017cycia" + RQ + " \u2014 jak sprawne odbiorniki, kt\u00f3rym odci\u0119to zasilanie. Brak stymulacji prowadzi do <b>atrofii opuszki w\u0119chowej</b> i zmian w hipokampie, co wp\u0142ywa na pami\u0119\u0107 i emocje.")

    # SLIDE 3: Dlaczego to minie
    s.new_slide()
    y = s.draw_logo()
    y = s.draw_pill("1 \u00b7 Wst\u0119p", MARGIN, y)
    y = s.draw_title("Dlaczego to minie?", y, size=28)
    y = s.draw_body("<b><font color='#7E57C2'>Neurony w\u0119chowe maj\u0105 unikaln\u0105 zdolno\u015b\u0107 do regeneracji \u2014 jako jedyne w organizmie odnawiaj\u0105 si\u0119 przez ca\u0142e \u017cycie.</font></b>", y, size=14)
    y -= 10
    y = s.draw_blockquote(
        LQ + "Systematyczny trening w\u0119chowy wykazuje skuteczno\u015b\u0107 por\u00f3wnywaln\u0105 z terapi\u0105 sterydow\u0105. Regularna stymulacja zwi\u0119ksza obj\u0119to\u015b\u0107 istoty szarej w m\u00f3zgu. Tw\u00f3j m\u00f3zg jest plastyczny \u2014 trening to proces jego fizycznej odbudowy." + RQ,
        "\u2014 metaanalizy prof. Thomasa Hummela",
        y, font_size=14
    )
    y -= 6
    s.draw_accent_box(
        "\U0001f4a1 <b>Kluczowy wniosek:</b> Trening w\u0119chowy to nie " + LQ + "alternatywna medycyna" + RQ + " \u2014 to metoda poparta setkami bada\u0144 naukowych, w tym badaniami obrazowania m\u00f3zgu (fMRI/MRI).",
        y, font_size=12
    )

    # SLIDE 4: Co przygotowac
    s.new_slide(BG_SOFT)
    y = s.draw_logo()
    y = s.draw_pill("2 \u00b7 Warsztat zapachowy", MARGIN, y)
    y = s.draw_title("Co przygotowa\u0107?", y, size=28)
    y = s.draw_sub("Potrzebujemy stworzy\u0107 <b>headspace</b> \u2014 nasycon\u0105 cz\u0105steczkami przestrze\u0144 nad \u017ar\u00f3d\u0142em zapachu.", y)
    y -= 6
    cw = (CONTENT_W - 28) / 3
    avail = y - BOTTOM
    card_h = int(avail * 0.92)
    items = [
        ("S\u0142oiczki z ciemnego szk\u0142a", "15\u201330 ml pojemno\u015bci. Chroni\u0105 olejki przed \u015bwiat\u0142em i koncentruj\u0105 opary wewn\u0105trz. Najlepiej z zakr\u0119tkami \u2014 szczelno\u015b\u0107 jest kluczowa."),
        ("Papier akwarelowy", "Porowato\u015b\u0107 idealnie trzyma aromat wewn\u0105trz s\u0142oiczka. Wycinamy pasek dopasowany do wielko\u015bci s\u0142oiczka."),
        ("Olejki eteryczne", "Wy\u0142\u0105cznie naturalne koncentraty wysokiej jako\u015bci. Syntetyczne odpowiedniki nie aktywuj\u0105 w\u0142a\u015bciwych receptor\u00f3w."),
    ]
    for i, (title, body) in enumerate(items):
        s.draw_card(MARGIN + i * (cw + 14), y, cw, card_h, title, body)

    # SLIDE 5: Przygotowanie sloiczka
    s.new_slide()
    y = s.draw_logo()
    y = s.draw_pill("2 \u00b7 Warsztat zapachowy", MARGIN, y)
    y = s.draw_title("Jak przygotowa\u0107 s\u0142oiczek?", y, size=28)
    y -= 6
    y = s.draw_ordered_list([
        "W\u0142\u00f3\u017c do s\u0142oiczka pasek papieru akwarelowego",
        "Nas\u0105cz go <b>4\u20138 kroplami</b> wybranego olejku eterycznego",
        "Szczelnie zakr\u0119\u0107 i odczekaj <b>minimum godzin\u0119</b> na nasycenie",
        "<b>Co tydzie\u0144</b> wymieniaj papier i dolewaj \u015bwie\u017cego olejku",
        "Popro\u015b kogo\u015b ze sprawnym w\u0119chem o <b>weryfikacj\u0119 intensywno\u015bci</b>",
    ], y)
    y -= 8
    s.draw_accent_box(
        "\U0001f4a1 <b>Cytrusy szybko oksyduj\u0105</b> \u2014 myj s\u0142oiczki po nich dok\u0142adnie myd\u0142em, poniewa\u017c utlenione olejki trac\u0105 w\u0142a\u015bciwo\u015bci terapeutyczne i mog\u0105 podra\u017cnia\u0107 sk\u00f3r\u0119.",
        y, font_size=12
    )

    # SLIDE 6: Jakie zapachy wybrac
    s.new_slide(BG_SOFT)
    y = s.draw_logo()
    y = s.draw_pill("3 \u00b7 Wyb\u00f3r zapach\u00f3w", MARGIN, y)
    y = s.draw_title("Jakie zapachy wybra\u0107?", y, size=28)
    y -= 6
    y = s.draw_table(
        ["Grupa zapachowa", "Zamienniki", "Dlaczego?"],
        [
            ["Kwiatowa (R\u00f3\u017ca)", "Geranium, ylang-ylang", "Pobudza subtelne receptory"],
            ["Owocowa (Cytryna)", "Pomara\u0144cza, grejpfrut", "Wysoka intensywno\u015b\u0107"],
            ["Korzenna (Go\u017adziki)", "Cynamon, wanilia", "Zakotwiczenie w pami\u0119ci"],
            ["\u017bywicza (Eukaliptus)", "Mi\u0119ta, rozmaryn", "Nerw tr\u00f3jdzielny (ch\u0142\u00f3d)"],
        ],
        y,
        col_widths=[CONTENT_W * 0.35, CONTENT_W * 0.35, CONTENT_W * 0.30],
        font_size=12,
    )
    y -= 4
    s.draw_accent_box(
        "\U0001f9e0 <b>Pami\u0119\u0107 w\u0119chowa:</b> Wybieraj aromaty budz\u0105ce silne wspomnienia \u2014 emocjonalny \u015blad u\u0142atwia regeneracj\u0119 po\u0142\u0105cze\u0144 synaptycznych. Im silniejsze skojarzenie, tym lepszy efekt terapeutyczny.",
        y, font_size=12
    )

    # SLIDE 7: Technika malych wdechow
    s.new_slide()
    y = s.draw_logo()
    y = s.draw_pill("4 \u00b7 Technika oddechowa", MARGIN, y)
    y = s.draw_title("Technika " + LQ + "ma\u0142ych wdech\u00f3w" + RQ, y, size=28)
    y = s.draw_sub("<b>G\u0142\u0119boki wdech omija nab\u0142onek w\u0119chowy</b> \u2014 kieruje powietrze prosto do p\u0142uc, zamiast do pola w\u0119chowego.", y)
    y -= 8
    card_w = (CONTENT_W - 14) / 2
    avail = y - BOTTOM
    card_h = int(avail * 0.92)
    s.draw_card(MARGIN, y, card_w, card_h, "Prawid\u0142owa technika",
                "Kr\u00f3tkie, ma\u0142e wdechy \u2014 jak pies na spacerze. Tworzysz <b>zawirowania powietrza</b>, kt\u00f3re kieruj\u0105 headspace bezpo\u015brednio na pole w\u0119chowe w g\u00f3rnej cz\u0119\u015bci jamy nosowej.",
                accent_color=PURPLE)
    s.draw_card(MARGIN + card_w + 14, y, card_w, card_h, "B\u0142\u0105d do unikania",
                "G\u0142\u0119boki, d\u0142ugi wdech nosem \u2014 powietrze omija nab\u0142onek w\u0119chowy i trafia wprost do p\u0142uc. <b>Nie stymuluje receptor\u00f3w</b> i nie przynosi efektu terapeutycznego.")

    # SLIDE 8: Sesja treningowa
    s.new_slide(BG_SOFT)
    y = s.draw_logo()
    y = s.draw_pill("4 \u00b7 Sesja treningowa", MARGIN, y)
    y = s.draw_title("Jak wygl\u0105da sesja treningowa?", y, size=28)
    y -= 6
    y = s.draw_checklist([
        "Wybierz spokojne miejsce, wycisz telefon",
        "Otw\u00f3rz s\u0142oiczek i zbli\u017c go do nosa (ok. 2\u20133 cm)",
        "<b>20 sekund</b> w\u0105chania technik\u0105 ma\u0142ych wdech\u00f3w",
        "Zamknij s\u0142oiczek \u2014 <b>10\u201315 sekund przerwy</b> mi\u0119dzy zapachami",
        "Przejd\u017a do kolejnego zapachu (4 zapachy = 1 sesja)",
        "Powtarzaj <b>2\u00d7 dziennie: rano i wieczorem</b>",
    ], y, font_size=14)

    # SLIDE 9: Wachaj wyobraznia
    s.new_slide()
    y = s.draw_logo()
    y = s.draw_pill("5 \u00b7 Praca mentalna", MARGIN, y)
    y = s.draw_title("W\u0105chaj wyobra\u017ani\u0105", y, size=28)
    y = s.draw_sub("Trening w\u0119chowy to w po\u0142owie praca umys\u0142u. Kora w\u0119chowa wykazuje aktywno\u015b\u0107 nawet przy braku fizycznego bod\u017aca.", y)
    y -= 8
    card_w = (CONTENT_W - 14) / 2
    avail = y - BOTTOM
    card_h = int(avail * 0.92)
    s.draw_card(MARGIN, y, card_w, card_h, "Wizualizacja",
                "<b>Zamknij oczy</b> i przywo\u0142aj obraz obiektu \u2014 kolor cytryny, porowato\u015b\u0107 sk\u00f3rki, kwa\u015bny smak na j\u0119zyku, ch\u0142\u00f3d z lod\u00f3wki. Anga\u017cuj <b>wszystkie zmys\u0142y</b> naraz.")
    s.draw_card(MARGIN + card_w + 14, y, card_w, card_h, "Wsparcie wizualne",
                "Patrz na <b>zdj\u0119cia</b> w\u0105chanych obiekt\u00f3w podczas sesji. Medytacja sensoryczna zapobiega degradacji neuron\u00f3w i <b>wzmacnia \u015bcie\u017cki pami\u0119ciowe</b>.")

    # SLIDE 10: Jak sledzic postepy
    s.new_slide(BG_SOFT)
    y = s.draw_logo()
    y = s.draw_pill("6 \u00b7 Dzienniczek post\u0119p\u00f3w", MARGIN, y)
    y = s.draw_title("Jak \u015bledzi\u0107 post\u0119py?", y, size=28)
    y -= 6
    stat_w = 240
    stat_h = 90
    sx = MARGIN + (CONTENT_W - 2 * stat_w - 20) / 2
    s.draw_stat_card(sx, y, stat_w, stat_h, "4 mies.", "Pierwsze efekty", num_size=40)
    s.draw_stat_card(sx + stat_w + 20, y, stat_w, stat_h, "14\u201324", "Miesi\u0105ce rehabilitacji", num_size=40)
    y -= stat_h + 14
    y = s.draw_accent_box(
        "\u2705 <b>Parosmia = dobry znak!</b> Zniekszta\u0142cone zapachy (np. zapach gumy zamiast kawy) to dow\u00f3d, \u017ce neurony nawi\u0105zuj\u0105 nowe po\u0142\u0105czenia synaptyczne.",
        y, bg=GREEN_SOFT, font_size=12
    )
    y -= 4
    y = s.draw_table(
        ["Pole", "Wpis"],
        [
            ["Data", ".................."],
            ["Zapach", ".................."],
            ["Odczucia", "nic / ch\u0142\u00f3d / zniekszta\u0142cony / czysty"],
            ["Intensywno\u015b\u0107", "0 \u2013 1 \u2013 2 \u2013 3 \u2013 4 \u2013 5"],
        ],
        y,
        col_widths=[CONTENT_W * 0.3, CONTENT_W * 0.7],
        font_size=11,
    )

    # SLIDE 11: Nie tylko po wirusie
    s.new_slide()
    y = s.draw_logo()
    y = s.draw_pill("7 \u00b7 Szersze korzy\u015bci", MARGIN, y)
    y = s.draw_title("Nie tylko po wirusie", y, size=28)
    y = s.draw_sub("Trening w\u0119chowy przynosi szersze korzy\u015bci dla m\u00f3zgu i zdrowia psychicznego:", y)
    y -= 6
    card_w = (CONTENT_W - 14) / 2
    avail = y - BOTTOM
    card_h = int((avail - 14) / 2)
    cards = [
        ("Funkcje poznawcze", "Udowodniona poprawa pami\u0119ci i koncentracji, szczeg\u00f3lnie u os\u00f3b starszych i po urazach."),
        ("P\u0142ynno\u015b\u0107 werbalna", "Badania potwierdzaj\u0105 popraw\u0119 p\u0142ynno\u015bci semantycznej i zdolno\u015bci nazywania."),
        ("Istota szara", "Zwi\u0119kszenie obj\u0119to\u015bci istoty szarej \u2014 odwraca skutki anosmii potwierdzone w MRI."),
        ("Nastr\u00f3j i emocje", "Poprawa nastroju i redukcja objaw\u00f3w depresji potwierdzona klinicznie."),
    ]
    for i, (title, body) in enumerate(cards):
        col = i % 2
        row = i // 2
        cx = MARGIN + col * (card_w + 14)
        cy = y - row * (card_h + 14)
        s.draw_card(cx, cy, card_w, card_h, title, body)

    # SLIDE 12: Neuroplastycznosc
    s.new_slide(BG_SOFT)
    y = s.draw_logo()
    y = s.draw_pill("7 \u00b7 Neuroplastyczno\u015b\u0107", MARGIN, y)
    y = s.draw_title("Jak m\u00f3zg si\u0119 odbudowuje?", y, size=28)
    y -= 4
    avail = y - BOTTOM
    card_h = int((avail - 24) / 3)
    y = s.draw_card(MARGIN, y, CONTENT_W, card_h, "Istota szara",
                    "Anosmia powoduje utrat\u0119 istoty szarej w obszarach odpowiedzialnych za w\u0119ch. Systematyczny trening <b>fizycznie zwi\u0119ksza jej obj\u0119to\u015b\u0107</b>, odwracaj\u0105c negatywne skutki utraty powonienia. Potwierdzone w badaniach MRI.",
                    accent_color=PURPLE)
    y -= 2
    y = s.draw_card(MARGIN, y, CONTENT_W, card_h, "\u0141\u0105czno\u015b\u0107 strukturalna",
                    "D\u0142ugoterminowa ekspozycja na bod\u017ace w\u0119chowe <b>przebudowuje szlaki nerwowe</b>. Nawet nocna ekspozycja (2h/noc przez 6 miesi\u0119cy) poprawia fizyczn\u0105 \u0142\u0105czno\u015b\u0107 mi\u0119dzy uk\u0142adem limbicznym a kor\u0105 m\u00f3zgow\u0105.",
                    accent_color=PURPLE)
    y -= 2
    s.draw_accent_box(
        "\U0001f9d3 Trening w\u0119chowy to tak\u017ce skuteczny <b>trening umys\u0142u i pami\u0119ci dla senior\u00f3w</b> \u2014 niezale\u017cnie od tego, czy dosz\u0142o do utraty w\u0119chu.",
        y, font_size=13
    )

    # SLIDE 13: Zapamietaj te zasady
    s.new_slide()
    y = s.draw_logo()
    y = s.draw_pill("8 \u00b7 Podsumowanie", MARGIN, y)
    y = s.draw_title("Zapami\u0119taj te zasady", y, size=28)
    y -= 6
    card_w = (CONTENT_W - 14) / 2
    avail = y - BOTTOM
    top_card_h = int((avail - 24) * 0.38)
    bot_card_h = int((avail - 24) * 0.28)
    zasady = [
        ("1. Systematyczno\u015b\u0107", "2\u00d7 dziennie, codziennie \u2014 rano i wieczorem. To Twoje lekarstwo."),
        ("2. Technika oddechu", "Kr\u00f3tkie, " + LQ + "w\u0119sz\u0105ce" + RQ + " wdechy jak pies. Nie omijaj receptor\u00f3w."),
        ("3. Wyobra\u017ania", "M\u00f3zg reaguje na wspomnienie zapachu tak samo intensywnie jak na prawdziwy."),
        ("4. Stymulacja tr\u00f3jdzielna", "Zawsze mi\u0119ta lub eukaliptus w zestawie \u2014 aktywuj\u0105 nerw tr\u00f3jdzielny."),
    ]
    for i, (title, body) in enumerate(zasady):
        col = i % 2
        row = i // 2
        cx = MARGIN + col * (card_w + 14)
        cy = y - row * (top_card_h + 14)
        s.draw_card(cx, cy, card_w, top_card_h, title, body, accent_color=PURPLE)
    y -= 2 * (top_card_h + 14)
    s.draw_card(MARGIN, y, CONTENT_W, bot_card_h, "5. Czas i cierpliwo\u015b\u0107",
                "Daj sobie minimum 4 miesi\u0105ce na pierwszy sygna\u0142 powrotu. Pe\u0142na rehabilitacja to 14\u201324 miesi\u0105ce \u2014 ale ka\u017cdy dzie\u0144 treningu przybli\u017ca Ci\u0119 do celu.",
                accent_color=GREEN)

    s.save()
    print("PDF zapisany: " + OUTFILE)
    print("   " + str(s.slide_num) + " slajdow")


if __name__ == "__main__":
    build()
