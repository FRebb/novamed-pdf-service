"""
nova:med PDF-Generator Service
Render.com Web Service — empfängt JSON von n8n, gibt PDF zurück
"""
from flask import Flask, request, send_file, jsonify
import io, os

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                  TableStyle, PageBreak, HRFlowable)

app = Flask(__name__)

# ==============================================================
# SICHERHEIT: Einfacher API-Key-Schutz
# Setze in Render.com unter "Environment Variables": PDF_API_KEY=<dein-key>
# ==============================================================
API_KEY = os.environ.get('PDF_API_KEY', 'novamed-pdf-2026')

# ==============================================================
# CORPORATE DESIGN — nova:med (nicht verändern)
# ==============================================================
GREEN      = colors.Color(106/255, 172/255, 70/255)
LIGHTGREEN = colors.Color(207/255, 225/255, 195/255)
SUBKPI_BG  = colors.Color(239/255, 246/255, 238/255)
FOOTER_BG  = colors.Color(207/255, 225/255, 195/255)
BLACK      = colors.HexColor('#1a1a1a')
GREY       = colors.HexColor('#888888')
RED        = colors.HexColor('#cc0000')

W, H = A4
ML = 20*mm; MR = 20*mm; MT = 18*mm; MB = 20*mm


# ==============================================================
# HILFSFUNKTIONEN
# ==============================================================
def fmt(n):
    s = f"{abs(n):,.0f}".replace(",","X").replace(".","," ).replace("X",".")
    return ("-" if n < 0 else "") + s + " €"

def fmtM(n):
    return f"{n/1e6:.2f}".replace(".",",") + " Mio, €"

def pct(a, b):
    if not b: return "—"
    return f"{(a-b)/b*100:.1f}".replace(".",",") + " %"

def sign(v):
    if v == "—": return "—"
    return v if v.startswith('-') else '+' + v

def ppFn(a, b):
    dv = a - b
    s = "+" if dv >= 0 else ""
    return f"{s}{dv:.1f}".replace(".",",") + " PP"

def fmtA(n):
    return f"{n:.1f}".replace(".",",") + " %"


# ==============================================================
# SEITEN-TEMPLATE (Kopf-/Fußzeile)
# ==============================================================
def make_on_page(monat, quartal):
    page_num = [0]
    titel = f'Umsatzbericht {monat} I {quartal}'

    def on_page(canvas, doc):
        page_num[0] += 1
        canvas.saveState()
        by = H - MT + 2*mm
        bh = 7*mm
        canvas.setFillColor(GREEN)
        canvas.rect(ML, by, W-ML-MR, bh, fill=1, stroke=0)
        canvas.setFont('Helvetica-Bold', 8)
        canvas.setFillColor(colors.white)
        canvas.drawString(ML+3*mm, by+2.2*mm, 'nova:')
        canvas.setFont('Helvetica', 8)
        x = ML+3*mm + canvas.stringWidth('nova:', 'Helvetica-Bold', 8)
        canvas.drawString(x, by+2.2*mm, 'med')
        canvas.drawRightString(W-MR-1*mm, by+2.2*mm, titel)
        fh = 6.5*mm
        fy = MB - fh - 1*mm
        canvas.setFillColor(FOOTER_BG)
        canvas.rect(ML, fy, W-ML-MR, fh, fill=1, stroke=0)
        canvas.setFont('Helvetica', 7.5)
        canvas.setFillColor(BLACK)
        canvas.drawString(ML+3*mm, fy+2*mm,
                          'nova:med GmbH & Co. KG · Höchstadt a. d. Aisch · www.novamed.de')
        canvas.drawRightString(W-MR-1*mm, fy+2*mm,
                               f'Seite {page_num[0]} · Vertraulich')
        canvas.restoreState()

    return on_page


# ==============================================================
# FORMATVORLAGEN
# ==============================================================
def make_styles():
    def S(name, **kw):
        base = dict(fontName='Helvetica', fontSize=9, leading=12, textColor=BLACK, spaceAfter=2)
        base.update(kw)
        return ParagraphStyle(name, **base)

    return dict(
        sNormal   = S('Normal'),
        sBold     = S('Bold',     fontName='Helvetica-Bold', fontSize=10, leading=14),
        sSmall    = S('Small',    fontSize=7.5, textColor=GREY),
        sItalic   = S('Italic',   fontName='Helvetica-Oblique', fontSize=8.5,
                       textColor=colors.HexColor('#333333')),
        sH1       = S('H1',       fontName='Helvetica-Bold', fontSize=22, leading=26, spaceAfter=2),
        sH1sub    = S('H1sub',    fontName='Helvetica', fontSize=13, leading=16,
                       spaceAfter=2, textColor=GREEN),
        sMeta     = S('Meta',     fontSize=9, textColor=colors.HexColor('#555555'), leading=13),
        sH2       = S('H2',       fontName='Helvetica-Bold', fontSize=12, leading=16,
                       spaceAfter=5, textColor=GREEN),
        sKernHead = S('KernHead', fontName='Helvetica-Bold', fontSize=10, leading=14,
                       textColor=GREEN),
        sFootnote = S('Footnote', fontSize=7.5, textColor=GREY),
    )


TH_STYLE = [
    ('BACKGROUND',    (0,0),  (-1,0),  GREEN),
    ('TEXTCOLOR',     (0,0),  (-1,0),  colors.white),
    ('FONTNAME',      (0,0),  (-1,0),  'Helvetica-Bold'),
    ('FONTSIZE',      (0,0),  (-1,0),  8),
    ('TOPPADDING',    (0,0),  (-1,0),  5),
    ('BOTTOMPADDING', (0,0),  (-1,0),  5),
    ('LEFTPADDING',   (0,0),  (-1,-1), 6),
    ('RIGHTPADDING',  (0,0),  (-1,-1), 6),
    ('FONTNAME',      (0,1),  (-1,-1), 'Helvetica'),
    ('FONTSIZE',      (0,1),  (-1,-1), 8.5),
    ('TOPPADDING',    (0,1),  (-1,-1), 3.5),
    ('BOTTOMPADDING', (0,1),  (-1,-1), 3.5),
    ('ROWBACKGROUNDS',(0,1),  (-1,-1), [colors.white, LIGHTGREEN]),
    ('LINEBELOW',     (0,0),  (-1,-1), 0.3, colors.HexColor('#d0d0d0')),
    ('VALIGN',        (0,0),  (-1,-1), 'MIDDLE'),
]

def make_table(data, col_widths, extra=None):
    ts = list(TH_STYLE)
    if extra:
        ts.extend(extra)
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle(ts))
    return t


# ==============================================================
# PDF AUFBAUEN
# ==============================================================
def build_pdf(d):
    st = make_styles()
    sNormal=st['sNormal']; sBold=st['sBold']; sSmall=st['sSmall']
    sItalic=st['sItalic']; sH1=st['sH1']; sH1sub=st['sH1sub']
    sMeta=st['sMeta']; sH2=st['sH2']
    sKernHead=st['sKernHead']; sFootnote=st['sFootnote']

    MONAT   = d['monat']
    DATUM   = d['datum']
    QUARTAL = d['quartal']
    u=d['u']; v=d['v']; p=d['p']; th=d['th']; epfp=d['epfp']
    hist=d['hist']; ad=d['ad']
    q26=d['q26']; q25=d['q25']; pq=d['pq']
    jsum2025=d['jsum2025total']; cagr=d['cagr']
    prog_basis=d['prog_basis']; prog_kons=d['prog_kons']; prog_opt=d['prog_opt']
    kernaussagen = d.get('kernaussagen', [['Kernaussagen:','Automatisch generiert.']])
    k1=d.get('k1',''); k2=d.get('k2',''); k3=d.get('k3','')
    k4=d.get('k4',''); ka=d.get('ka','')
    he = d.get('he', [])

    story = []
    cw = W - ML - MR

    # --- SEITE 1 ---
    story.append(Spacer(1, 8*mm))
    story.append(Paragraph('Umsatzbericht', sH1))
    story.append(Paragraph(f'{MONAT} & {QUARTAL}', sH1sub))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph('Vertraulicher Bericht für Gesellschafter / Beirat', sMeta))
    story.append(Paragraph('nova:med GmbH & Co. KG · Höchstadt a. d. Aisch', sMeta))
    story.append(Paragraph(f"Erstellt: {d.get('erstellt', 'April 2026')}", sMeta))
    story.append(Spacer(1, 5*mm))

    monat_kurz = d.get('monat_kurz', MONAT[:3])
    vorjahr    = d.get('vorjahr', 2025)
    aktuell_jahr = d.get('aktuell_jahr', 2026)
    kpi = [
        ['Kennzahl', f'{MONAT}',
         f'Vorjahr {monat_kurz} {vorjahr}',
         f'Plan {monat_kurz} {aktuell_jahr}',
         'Abw. VJ / Plan'],
        ['Monatsumsatz', fmtM(u['m']), fmtM(v['m']), fmtM(p['m']),
         f"{sign(pct(u['m'],v['m']))} / {sign(pct(u['m'],p['m']))}"],
        [f'YTD-Umsatz {QUARTAL}', fmtM(q26), fmtM(q25), fmtM(pq),
         f"{sign(pct(q26,q25))} / {sign(pct(q26,pq))}"],
        [f'{QUARTAL} gesamt', fmtM(q26), fmtM(q25), '—',
         f"{sign(pct(q26,q25))} / —"],
    ]
    story.append(make_table(kpi, [cw*0.22, cw*0.16, cw*0.16, cw*0.16, cw*0.30], [
        ('FONTNAME',  (1,1),(1,-1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (4,1),(4,-1), GREEN),
        ('FONTNAME',  (4,1),(4,-1), 'Helvetica-Bold'),
    ]))
    story.append(Spacer(1, 5*mm))

    story.append(Paragraph('<b>Kernaussagen</b>', sKernHead))
    story.append(Spacer(1, 2*mm))
    for topic, text in kernaussagen:
        story.append(Paragraph(f'· <b>{topic}</b> {text}', sNormal))
        story.append(HRFlowable(width=cw, thickness=0.3,
                                color=colors.HexColor('#dddddd'), spaceAfter=2))
    story.append(PageBreak())

    # --- SEITE 2 ---
    story.append(Paragraph(f'1. Monatsumsätze {QUARTAL} im Detail', sH2))
    story.append(HRFlowable(width=cw, thickness=1, color=GREEN, spaceAfter=6))

    def sv(name, **kw):
        base = dict(fontName='Helvetica', fontSize=9, leading=12, textColor=BLACK, spaceAfter=0)
        base.update(kw)
        return ParagraphStyle(name, **base)

    sub = [
        [Paragraph(f'Umsatz {MONAT}', sSmall),
         Paragraph('Abw. zum Plan', sSmall),
         Paragraph('Abw. zum Vorjahr', sSmall),
         Paragraph(f'YTD-Umsatz {QUARTAL}', sSmall)],
        [Paragraph(f"<b>{fmtM(u['m'])}</b>",
                   sv('v1', fontName='Helvetica-Bold', fontSize=13, textColor=GREEN)),
         Paragraph(f"<b>{sign(pct(u['m'],p['m']))}</b>",
                   sv('v2', fontName='Helvetica-Bold', fontSize=13, textColor=GREEN)),
         Paragraph(f"<b>{sign(pct(u['m'],v['m']))}</b>",
                   sv('v3', fontName='Helvetica-Bold', fontSize=13, textColor=GREEN)),
         Paragraph(f"<b>{fmtM(q26)}</b>",
                   sv('v4', fontName='Helvetica-Bold', fontSize=13, textColor=GREEN))],
        [Paragraph(f"Abw. Plan: {sign(pct(u['m'],p['m']))}", sSmall),
         Paragraph('Planzielerreichung', sSmall),
         Paragraph(f"{d.get('monat_kurz', MONAT[:3])} {d.get('vorjahr', 2025)}: {fmtM(v['m'])}", sSmall),
         Paragraph(f"Plan: {fmtM(pq)}", sSmall)],
        [Paragraph(f"Plan: {fmtM(p['m'])}", sSmall),
         Paragraph('', sSmall), Paragraph('', sSmall), Paragraph('', sSmall)],
    ]
    sub_t = Table(sub, colWidths=[cw/4]*4)
    sub_t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), SUBKPI_BG),
        ('BOX',           (0,0),(-1,-1), 0.5, colors.HexColor('#cccccc')),
        ('LINEAFTER',     (0,0),(2,-1),  0.5, colors.HexColor('#cccccc')),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 3),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
    ]))
    story.append(sub_t)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph(f'<b>Monatliche Umsatzentwicklung {QUARTAL}</b>', sBold))
    story.append(Spacer(1, 2*mm))
    # Monatsnamen und Quartalslabel dynamisch aus d['monate_labels']
    monate_labels = d.get('monate_labels', [
        ['Jan', 'Feb', 'Mrz'], f'{QUARTAL} Gesamt'])
    mon_zeilen = monate_labels[0]  # ['Jan','Feb','Mrz'] oder ['Apr','Mai','Jun'] etc.
    quartal_label = monate_labels[1]
    vorjahr = d.get('vorjahr', 2025)
    aktuell_jahr = d.get('aktuell_jahr', 2026)
    mon = [[f'Monat',f'Umsatz {aktuell_jahr}',f'Umsatz {vorjahr}',f'Plan {aktuell_jahr}','Abw. VJ','Abw. Plan']]
    for m2,a,b,c in zip(mon_zeilen,
                        [u['j'],u['f'],u['m']],
                        [v['j'],v['f'],v['m']],
                        [p['j'],p['f'],p['m']]):
        mon.append([m2, fmt(a), fmt(b), fmt(c), sign(pct(a,b)), sign(pct(a,c))])
    mon.append([quartal_label, fmt(q26), fmt(q25), fmt(pq),
                sign(pct(q26,q25)), sign(pct(q26,pq))])
    story.append(make_table(mon, [cw*0.1, cw*0.18, cw*0.18, cw*0.18, cw*0.18, cw*0.18], [
        ('TEXTCOLOR', (4,1),(5,-1), GREEN),
        ('FONTNAME',  (4,1),(5,-1), 'Helvetica-Bold'),
    ]))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph('<b>Kommentar</b>', sBold))
    story.append(Paragraph(k1, sItalic))
    story.append(PageBreak())

    # --- SEITE 3 ---
    story.append(Paragraph('2. Quartalsvergleich — Langfristentwicklung', sH2))
    story.append(HRFlowable(width=cw, thickness=1, color=GREEN, spaceAfter=6))
    story.append(Paragraph('<b>Umsatz Netto nach Quartal 2021–2026</b>', sBold))
    story.append(Spacer(1, 2*mm))

    hd = [['Jahr','Q1','Q2','Q3','Q4','Gesamt','∆ Q1 VJ']]
    for i, h in enumerate(hist):
        ges = (h['q1']or 0)+(h['q2']or 0)+(h['q3']or 0)+(h['q4']or 0)
        gs = fmt(h['q1'])+' *' if h['j']==2026 else fmt(ges)
        q1vj = pct(h['q1'], hist[i-1]['q1']) if i>0 else '—'
        sgn = '+' if q1vj != '—' and not q1vj.startswith('-') else ''
        hd.append([str(h['j']), fmt(h['q1']),
                   fmt(h['q2']) if h['q2'] else '—',
                   fmt(h['q3']) if h['q3'] else '—',
                   fmt(h['q4']) if h['q4'] else '—',
                   gs, f"{sgn}{q1vj}" if q1vj != '—' else '—'])
    hs = []
    for r in range(1, len(hd)):
        val = hd[r][6]
        c = GREEN if val != '—' and not val.startswith('-') else \
            (RED if val.startswith('-') else BLACK)
        hs += [('TEXTCOLOR',(6,r),(6,r),c), ('FONTNAME',(6,r),(6,r),'Helvetica-Bold')]
    story.append(make_table(hd, [cw*0.08,cw*0.15,cw*0.15,cw*0.15,cw*0.15,cw*0.17,cw*0.15], hs))
    story.append(Paragraph(f'* 2026: nur Q1 abgeschlossen (Stand {DATUM})', sFootnote))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph('<b>Q1-Wachstum Year-over-Year</b>', sBold))
    story.append(Spacer(1, 2*mm))
    yd = [['Jahr','Q1 Vorjahr','Q1 aktuell','Abs. Differenz','Wachstum']]
    for i in range(1, len(hist)):
        h = hist[i]; prev = hist[i-1]
        w = pct(h['q1'], prev['q1'])
        sgn = '+' if not w.startswith('-') else ''
        yd.append([str(h['j']), fmt(prev['q1']), fmt(h['q1']),
                   fmt(h['q1']-prev['q1']), f"{sgn}{w}"])
    ys = []
    for r in range(1, len(yd)):
        c = GREEN if not yd[r][4].startswith('-') else RED
        ys += [('TEXTCOLOR',(4,r),(4,r),c), ('FONTNAME',(4,r),(4,r),'Helvetica-Bold')]
    story.append(make_table(yd, [cw*0.1,cw*0.22,cw*0.22,cw*0.22,cw*0.24], ys))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph('<b>Kommentar</b>', sBold))
    story.append(Paragraph(k2, sItalic))
    story.append(PageBreak())

    # --- SEITE 4 ---
    story.append(Paragraph('3. Umsatz nach Therapiegebiet', sH2))
    story.append(HRFlowable(width=cw, thickness=1, color=GREEN, spaceAfter=6))
    td = [['Therapiegebiet','Q1 2024','Q1 2025','Q1 2026','∆ 25→26','Anteil']]
    for x in th:
        dv = pct(x['a'], x['b'])
        sgn = '+' if not dv.startswith('-') else ''
        td.append([x['n'], fmt(x['c']), fmt(x['b']), fmt(x['a']),
                   f"{sgn}{dv}", fmtA(x['ant26'])])
    td.append(['Gesamt', fmt(5352428), fmt(5997361), fmt(8085004), '+34,8 %', '100,0 %'])
    ts2 = []
    for r in range(1, len(td)):
        dv = td[r][4]
        c = GREEN if not dv.startswith('-') else RED
        ts2 += [('TEXTCOLOR',(4,r),(4,r),c), ('FONTNAME',(4,r),(4,r),'Helvetica-Bold')]
    ts2.append(('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'))
    story.append(make_table(td, [cw*0.18,cw*0.16,cw*0.16,cw*0.16,cw*0.16,cw*0.18], ts2))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph('<b>Kommentar</b>', sBold))
    story.append(Paragraph(k3, sItalic))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph('<b>Umsatzanteile nach Therapiegebiet — Q1 2024 bis Q1 2026</b>', sBold))
    story.append(Spacer(1, 2*mm))
    ad2 = [['Therapiegebiet','Anteil Q1 2024','Anteil Q1 2025','Anteil Q1 2026','∆ Anteil (PP)']]
    for x in th:
        dv = ppFn(x['ant26'], x['ant25'])
        ad2.append([x['n'], fmtA(x['ant24']), fmtA(x['ant25']), fmtA(x['ant26']), dv])
    as2 = []
    for r in range(1, len(ad2)):
        v2 = ad2[r][4]
        c = GREEN if not v2.startswith('-') else RED
        as2 += [('TEXTCOLOR',(4,r),(4,r),c), ('FONTNAME',(4,r),(4,r),'Helvetica-Bold')]
    story.append(make_table(ad2, [cw*0.22,cw*0.19,cw*0.19,cw*0.19,cw*0.21], as2))
    story.append(Paragraph('PP = Prozentpunkte  ∆ gegenüber Q1 2025', sFootnote))
    story.append(PageBreak())

    # --- SEITE 5 ---
    story.append(Paragraph('4. Abrechnungsstruktur nach Therapiegebiet — EP / FP / Kauf/WE', sH2))
    story.append(HRFlowable(width=cw, thickness=1, color=GREEN, spaceAfter=6))
    ep = [['Segment / Abrechnungsart','Q1 2024','Q1 2025','Q1 2026','∆ 25→26','Anteil']]
    ep_seg = []; ep_sub = []
    for x in epfp:
        ep.append([x['seg'], fmt(x['ges24']), fmt(x['ges25']), fmt(x['ges26']),
                   sign(pct(x['ges26'],x['ges25'])), fmtA(x['ant'])])
        ep_seg.append(len(ep)-1)
        for lbl, k24, k25, k26 in [('EP','ep24','ep25','ep26'),
                                    ('FP','fp24','fp25','fp26'),
                                    ('Kauf/WE','kw24','kw25','kw26')]:
            ep.append(['    '+lbl, fmt(x[k24]), fmt(x[k25]), fmt(x[k26]),
                       sign(pct(x[k26],x[k25])), fmtA(x[k26]/x['ges26']*100)])
            ep_sub.append(len(ep)-1)
    ep.append(['Gesamt (aufgeschl. Segmente)', fmt(5137227), fmt(5835890),
               fmt(7880178), '+35,0 %', '100,0 %'])
    eps = [('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold')]
    for r in ep_seg:
        eps += [('TEXTCOLOR',(0,r),(-1,r),GREEN), ('FONTNAME',(0,r),(-1,r),'Helvetica-Bold')]
    for r in ep_sub:
        eps += [
            ('TEXTCOLOR',(0,r),(0,r),colors.HexColor('#333333')),
            ('FONTNAME', (0,r),(0,r),'Helvetica'),
            ('FONTSIZE', (0,r),(0,r),8),
            ('FONTNAME', (1,r),(3,r),'Helvetica'),
            ('FONTNAME', (5,r),(5,r),'Helvetica'),
        ]
    for r in range(1, len(ep)):
        dv = ep[r][4]
        c = GREEN if not dv.startswith('-') else RED
        eps += [('TEXTCOLOR',(4,r),(4,r),c), ('FONTNAME',(4,r),(4,r),'Helvetica-Bold')]
    story.append(make_table(ep, [cw*0.28,cw*0.13,cw*0.13,cw*0.13,cw*0.15,cw*0.18], eps))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph('<b>Kommentar</b>', sBold))
    story.append(Paragraph(k4, sItalic))
    story.append(PageBreak())

    # --- SEITE 6 ---
    story.append(Paragraph('5. Ausblick und strategische Einschätzung', sH2))
    story.append(HRFlowable(width=cw, thickness=1, color=GREEN, spaceAfter=6))
    story.append(Paragraph('<b>Jahresprognose 2026 (indikativ)</b>', sBold))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"Basierend auf dem Q1-Anteil am Jahresumsatz 2025 (22,0 %) ergibt sich eine indikative "
        f"Jahresprognose 2026 von ca. {fmtM(prog_basis)}. "
        f"Wachstum ggü. Jahresumsatz 2025 (27,25 Mio, €): {sign(pct(prog_basis,jsum2025))}.",
        sNormal))
    story.append(Spacer(1, 2*mm))
    pr = [['Szenario','Prognose Jahresumsatz 2026','Wachstum ggü. 2025']]
    for lbl, val in [('Konservativ (Q1 × 4)',prog_kons),
                     ('Basis (Q1-Anteil 2025)',prog_basis),
                     ('Optimistisch (Basis +5 %)',prog_opt)]:
        pr.append([lbl, fmtM(val), sign(pct(val,jsum2025))])
    story.append(make_table(pr, [cw*0.40,cw*0.33,cw*0.27], [
        ('TEXTCOLOR',(2,1),(2,-1),GREEN), ('FONTNAME',(2,1),(2,-1),'Helvetica-Bold'),
        ('FONTNAME', (1,1),(1,-1),'Helvetica-Bold'),
    ]))
    story.append(Spacer(1, 5*mm))

    story.append(Paragraph('<b>Außendienst YTD</b>', sBold))
    story.append(Spacer(1, 2*mm))
    adr = [['Mitarbeiter','YTD 2026','YTD 2025','vs. VJ']]
    for x in ad:
        adr.append([x['n'], fmtM(x['a']), fmtM(x['b']), sign(pct(x['a'],x['b']))])
    ads = []
    for r in range(1, len(adr)):
        c = GREEN if not adr[r][3].startswith('-') else RED
        ads += [('TEXTCOLOR',(3,r),(3,r),c), ('FONTNAME',(3,r),(3,r),'Helvetica-Bold')]
    story.append(make_table(adr, [cw*0.35,cw*0.21,cw*0.21,cw*0.23], ads))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph('<b>Kommentar</b>', sBold))
    story.append(Paragraph(ka, sItalic))
    story.append(Spacer(1, 5*mm))

    story.append(Paragraph('<b>Strategische Handlungshinweise</b>', sBold))
    story.append(Spacer(1, 2*mm))
    for topic, text in he:
        story.append(Paragraph(f'· <b>{topic}</b> {text}', sNormal))
        story.append(HRFlowable(width=cw, thickness=0.3,
                                color=colors.HexColor('#dddddd'), spaceAfter=2))
    story.append(Spacer(1, 8*mm))
    story.append(Paragraph(
        f'Datenquelle: Qlik Sense Dashboard „Umsatz FR", nova:med GmbH & Co. KG '
        f'(Stand: {DATUM}). Alle Angaben in Euro netto. '
        'Vertraulich — ausschließlich für Gesellschafter und Beirat.',
        sFootnote))

    # PDF in Memory erzeugen
    buf = io.BytesIO()
    on_page = make_on_page(MONAT, QUARTAL)
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=ML, rightMargin=MR,
        topMargin=MT+8*mm, bottomMargin=MB,
        title=f'nova:med Umsatzbericht {MONAT} | {QUARTAL}'
    )
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    buf.seek(0)
    return buf


# ==============================================================
# API ENDPOINTS
# ==============================================================
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'nova:med PDF Generator'})


@app.route('/generate', methods=['POST'])
def generate():
    # API-Key prüfen
    key = request.headers.get('X-API-Key', '')
    if key != API_KEY:
        return jsonify({'error': 'Unauthorized'}), 401

    data = request.get_json(force=True)
    if not data:
        return jsonify({'error': 'No JSON body'}), 400

    try:
        buf = build_pdf(data)
        fn = data.get('fn', 'novamed_umsatzbericht.pdf')
        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=fn
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
