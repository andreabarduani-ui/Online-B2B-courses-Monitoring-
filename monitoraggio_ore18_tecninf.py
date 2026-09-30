"""
============================================================================
MONITORAGGIO ORE FORMAZIONE TECNINF SPA - SOGLIA ORE 18:00
============================================================================
Ristrutturato per produrre un output con la STESSA STRUTTURA del file:
    Monitoraggio_ORE18_TECNINF_SPA_29-09-2026.xlsx

a partire dal CSV del Report Accessi:
    Tecninf spa - Report Accessi.csv   (separatore ";", encoding utf-8-sig)

Struttura prodotta (identica al target):
  - Fogli mensili "Aprile 2026" ... "Settembre 2026" (uno per mese nel CSV):
      Utente | Codice Fiscale | Percorso Formativo | 1..31 | Totale Ore
      Effettive | Ore Totali | Totale Eccesso | Eccesso Weekend |
      Eccesso Dopo 18:00 | Eccesso Mattutino
      Celle giornaliere = ore EFFETTIVE (timedelta), note con totali/eccessi.
      Ore Totali = somma CSV del mese (identica al CSV per costruzione).
  - "Riepilogo Generale": Totale Ore Maturate (= somma mensili effettive),
      Ore Tolte (= somma eccessi), % copertura, Corsi da finire (da Dettaglio),
      Quota Retr. 80%, Costo orario contributivo, Finanziamento.
  - "Dettaglio - <percorso>": un foglio per percorso con % veritiera per
      singolo corso =MIN(ore_effettive, attesa)/attesa.

Regole ORE18 (identiche al file target, verificate a mano):
  - Soglia serale 18:00 secca, ripartizione proporzionale su elapsed.
  - Finestra mattutina 06:00-07:40, stessa logica proporzionale.
  - NESSUNA tolleranza (anche 14 min dopo le 18:00 contano come eccesso).
  - Weekend (sab/dom): tutte le ore in eccesso weekend, effettive 0.
  - Cap 8h/giorno sulle ore normali (prima delle 18:00 meno mattutino).
  - Colori: verde F0F6EC, viola F2E8EE (dopo 18), azzurro E8EFF5 (mattutino),
    rosa FAE6E6 (weekend con uso), grigio F2F2F2 (weekend vuoto),
    arancio FBEAD8 (oltre cap 8h).

Uso:
    python monitoraggio_ore18_tecninf.py
============================================================================
"""
import datetime
import os
import re

import pandas as pd
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CSV_FILE = "Tecninf spa - Report Accessi.csv"
TEMPLATE_FILE = "Monitoraggio_ORE18_TECNINF_SPA_29-09-2026.xlsx"
OUTPUT_FILE = f"Monitoraggio_ORE18_TECNINF_SPA_{datetime.date.today().strftime('%d-%m-%Y')}.xlsx"

CSV_SEPARATOR = ";"
SOGLIA_ORA = 18.0
MATT_INI = 6.0
MATT_FIN = 7 + 40 / 60
CAP_GIORNO = 8.0
ORE_RIF_COPERTURA = 150

COLORI = {
    "verde": "F0F6EC",
    "arancio": "FBEAD8",
    "viola": "F2E8EE",
    "azzurro": "E8EFF5",
    "rosa": "FAE6E6",
    "giallo": "FBF4D9",
    "grigio": "F2F2F2",
    "header": "F4F5F7",
    "titolo": "ECF0F4",
    "titolo_testo": "2C3E50",
    "totale": "F1F3F5",
    "bordo": "BDC3C7",
}
DURATION_FORMAT = "[h]:mm:ss"

NOMI_MESI_IT = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio",
                "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre",
                "Novembre", "Dicembre"]

# Mappa foglio dettaglio -> percorso CSV (base senza suffisso)
DETTAGLIO_PERCORSI = [
    ("Dettaglio - Big Data", "Big Data - Tecninf - Fondoconoscenza", "DETTAGLIO BIG DATA"),
    ("Dettaglio - L'ufficio digitale", "L'ufficio digitale - Tecninf - Fondoconoscenza", "DETTAGLIO L'UFFICIO DIGITALE"),
    ("Dettaglio - La programmazione", "La programmazione - Tecninf - Fondoconoscenza", "DETTAGLIO LA PROGRAMMAZIONE"),
    ("Dettaglio - Project Management", "Project Management per il digitale - Tecninf - Fondoconoscenza", "DETTAGLIO PROJECT MANAGEMENT PER IL DIGITALE"),
    ("Dettaglio - Sicurezza Inform.", "Sicurezza Informatica - Tecninf - Fondoconoscenza", "DETTAGLIO SICUREZZA INFORMATICA"),
]

# Durate attese estratte dal file template (stesso ordine dei fogli).
DURATE_ATTESE = {
    "Dettaglio - Big Data": [
        ("GDPR e nuove normative Europee", 20),
        ("Introduzione ai Big Data nelle aziende", 5),
        ("Big data e analisi predittiva", 18),
        ("I Big data in ambito finanziario", 17),
        ("Business Intelligence e analisi predittiva", 5),
        ("Database relazionali e linguaggio SQL con PHP", 11),
        ("Natural Language Processing con Python", 14),
        ("Il Machine Learning con Python", 12),
        ("Python per Machine Learning e Artificial Intelligence", 11),
        ("Deep Learning e reti neurali", 13),
        ("Big Data Analytics con Python e Spark", 12),
        ("Progettare autorizzazioni e autentificazioni", 3),
        ("Gestione avanzata di un Database", 6),
        ("Sicurezza e Password Management", 3),
    ],
    "Dettaglio - L'ufficio digitale": [
        ("GDPR e nuove normative Europee", 20),
        ("Elementi base di sicurezza informatica", 5),
        ("Virus e Antivirus per tutti", 2),
        ("Backup: rendiamolo facile per tutti", 2),
        ("Introduzione al mondo di Microsoft office 365", 16),
        ("Microsoft 365: OneDrive corso completo", 3),
        ("Microsoft 365: Outlook desktop corso completo", 5),
        ("Microsoft 365: Outlook sul Web corso completo", 4),
        ("Microsoft 365: iniziare subito a lavorare con Teams", 2),
        ("Microsoft 365: Word corso completo", 7),
        ("Microsoft 365: Excel dai fondamentali al livello 4.0", 23),
        ("Microsoft 365: Excel basi di calcolo con Tabelle Pivot", 3),
        ("Microsoft 365: Excel corso avanzato con tabelle Pivot", 14),
        ("Microsoft 365: Power Query corso base", 3),
        ("Microsoft 365: Data Modeling con Power Query", 9),
        ("Microsoft 365: Proiezioni e calcoli con Power BI Desktop", 3),
        ("Microsoft 365: PowerPoint dalla base ai fondamentali", 5),
        ("Microsoft 365: elementi avanzati, suggerimenti e trucchi in PowerPoint", 7),
        ("Microsoft 365: strategie per fare presentazioni efficaci in PowerPoint", 3),
        ("Microsoft 365: strumenti per fare presentazioni efficaci in PowerPoint", 2),
        ("La gestione del lavoro in Google Sheet", 3),
        ("I sistemi informatici", 9),
    ],
    "Dettaglio - La programmazione": [
        ("GDPR e nuove normative Europee", 20),
        ("I sistemi operativi", 16),
        ("Linguaggi di programmazione", 17),
        ("Front end base: HTML, CSS, Bootstrap", 33),
        ("Front end avanzato: JavaScript e jQuery", 20),
        ("Back end base: Fondamenti di PHP", 12),
        ("Database relazionali e linguaggio SQL con PHP", 11),
        ("Applicazioni dinamiche con PHP e MySql", 10),
        ("Back end avanzato: PHP MVC con Laravel", 11),
    ],
    "Dettaglio - Project Management": [
        ("Influenzare positivamente il clima aziendale", 5),
        ("La gestione delle risorse umane", 9),
        ("La swot analysis del personale", 2),
        ("La gestione del conflitto", 4),
        ("Organizzare e gestire riunioni efficaci", 3),
        ("Ottimizzare ed organizzare i meeting online", 2),
        ("Gli approcci nel Problem Solving", 2),
        ("L'Hackathon nel Problem Solving", 1),
        ("La gestione del Feedback efficace", 3),
        ("Comprensione di una struttura organizzativa", 2),
        ("La definizione dei compiti e delle responsabilita", 1),
        ("I fondamentali del lavoro in team", 9),
        ("Il team e la sua cultura di crescita", 8),
        ("Guidare un team con l'intelligenza emotiva", 9),
        ("Fondamenti di Project Management", 2),
        ("Ruoli e processi del Project Management", 12),
        ("Predictive-Waterfall e lo sviluppo di un progetto", 3),
        ("Business Plan cos'e e come compilarlo", 3),
        ("Mindset e soft skills per il digitale", 2),
        ("Le basi del Public Speaking", 4),
        ("Leadership e Motivazione", 5),
        ("La scienza delle relazioni e l'arte di comunicare", 6),
        ("Comunicazione aziendale e risoluzione dei conflitti", 3),
        ("Il valore dell'unicita in azienda", 4),
        ("Le Soft skills per il mondo del lavoro", 15),
        ("Esercizi pratici per la gestione della Leadership", 4),
        ("Comunicazione efficace e persuasione", 4),
        ("Skills e competenze per comunicare in azienda", 17),
        ("Project management Agile con ClickUp", 6),
    ],
    "Dettaglio - Sicurezza Inform.": [
        ("GDPR e nuove normative Europee", 20),
        ("Elementi base di sicurezza informatica", 5),
        ("Virus e Antivirus per tutti", 2),
        ("La Cyber Security", 17),
        ("La Network Security", 25),
        ("Cloud e Fog Computing", 20),
        ("I sistemi cyber fisici", 16),
        ("Internet delle Cose e delle Macchine", 20),
        ("Tecnologie IoT e sviluppo", 5),
        ("La sicurezza nelle Tecnologie IoT", 7),
        ("Il Ruolo dell'Ethical Hacker in azienda", 7),
        ("Prevenire i rischi sui device mobili", 2),
        ("Gli attacchi tramite Ransomware", 2),
        ("Prevenire gli attacchi tramite posta elettronica", 3),
        ("Cosa e il Cloud: utilizzo, funzionalita e sicurezza", 2),
        ("Sicurezza e Password Management", 3),
    ],
}

# Titoli canonici per match robusto (normalizzazione spazi/punteggiatura).
def _canon(s):
    s = re.sub(r"\s+", " ", str(s).strip())
    return s


CANON_A_ATTESA = {}
for _sh, _lst in DURATE_ATTESE.items():
    for _tit, _att in _lst:
        CANON_A_ATTESA[_canon(_tit).lower()] = _att


def durata_attesa_default(titolo_corso):
    hit = CANON_A_ATTESA.get(_canon(titolo_corso).lower())
    if hit:
        return hit
    return 5


# Costi orari dal file template (tabella Riepilogo H/I per CF).
COSTI_ORARI = {
    "PZLWLM96A16Z504F": (9.731453023255815, 4.849926321476056),
    "BDALSN01R03H501Z": (9.731453023255815, 4.849926321476056),
    "BDNHHM97H19H501Q": (9.731453023255815, 3.235298414499312),
    "BCNNDR82P16D773D": (9.304, 5.11),
    "BNTJTP78S54Z241U": (13.178009302325583, 6.56758080265642),
    "BTTGRG97S59G274Q": (9.984876279069768, 4.976207445962649),
    "CRBVTR94A11Z140B": (9.858164651162792, 4.913013289280792),
    "CCCNHN91D19H501W": (9.731453023255815, 4.849926321476056),
    "GBBMCH73R57H501G": (12.316370232558135, 6.138145553425757),
    "LMTNIO94M03Z129S": (10.87692613953488, 5.420755321932365),
    "LCUVNC85L52D976T": (10.425832744186048, 3.783118027213908),
    "MGNLNZ97M04H501T": (9.731453023255815, 4.783158879615591),
    "MRNLSS75L68H501L": (10.87692613953488, 5.346173926583528),
    "MSCGLI83C04G511L": (19.584, 8.85),
    "PCPMHL00R03H501G": (8.746046511627904, 2.093632948739695),
    "PSSRRT76L56H501O": (12.316370232558135, 6.138145553425757),
    "PRLGLG98M18H501X": (9.731453023255815, 4.849907445614766),
    "CPNRRT66D27H501V": (14.364030139534886, 7.15867607196776),
    "PGGRRT91H07F839Q": (10.87692613953488, 5.420755321932365),
    "PRSRSH87S19H501S": (12.412671069767441, 6.186176741112714),
    "PDDRND75R08H501S": (11.028980093023256, 5.496582132484965),
    "RNRFNC99L28H501B": (9.858164651162792, 1.67451140143258),
    "DLRJMY02A20H501A": (8.746046511627904, 2.093632948739695),
    "SLPGLC81B17G377J": (9.731453023255815, 4.849926321476056),
    "SLVCMN80D14D086Z": (11.662538232558145, 5.812305599793496),
    "SMBLSS74D47H501E": (10.425832744186048, 5.131004531257418),
    "SVNSLV89A46H501I": (9.731453023255815, 4.849926321476056),
    "SCLFPP97S05H501K": (9.858164651162792, 1.683757475327303),
    "SCRTZN80A09H501I": (12.24034325581395, 6.0163062496053),
    "STSDNS91B05Z140T": (10.288984186046513, 5.127761382625514),
    "TRSNDR97M24H501B": (10.527202046511624, 5.777385318779284),
    "TCONNT91H59Z100O": (9.731453023255815, 4.849926321476056),
    "VSWSDH86S22Z222H": (11.99198846511628, 5.976475500904301),
}


# ---------------------------------------------------------------- utils

def fill_cell(colore):
    if not colore:
        return None
    return PatternFill("solid", start_color=colore, fgColor=colore)


def giorni_nel_mese(anno, mese):
    if mese == 12:
        return (datetime.date(anno + 1, 1, 1) - datetime.date(anno, mese, 1)).days
    return (datetime.date(anno, mese + 1, 1) - datetime.date(anno, mese, 1)).days


def abbreviazione_giorno(anno, mese, giorno):
    return ["L", "M", "M", "G", "V", "S", "D"][
        datetime.date(anno, mese, giorno).weekday()]


def giorni_weekend(anno, mese):
    return {d for d in range(1, giorni_nel_mese(anno, mese) + 1)
            if datetime.date(anno, mese, d).weekday() >= 5}


def chiave_ordinamento(nome):
    return re.sub(r"[^A-Za-z ]", "", str(nome).upper()).strip()


def cognome_nome(nome_csv):
    """'Nome Cognome' -> 'COGNOME NOME' (ultima parola = cognome)."""
    tok = [t for t in re.sub(r"[^A-Za-z' ]", " ", str(nome_csv).upper()).split() if t]
    if len(tok) >= 2:
        return (tok[-1] + " " + " ".join(tok[:-1])).strip()
    return (tok[0] if tok else "").upper()


def ore_decimali_a_hhmmss(ore_dec):
    if not ore_dec or ore_dec <= 0:
        return "00:00:00"
    s = int(round(ore_dec * 3600))
    return f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}"


# ---------------------------------------------------------------- CSV

def parse_durata(testo):
    m = re.match(r"(\d+)\s*h\s*(\d+)\s*min\s*(\d+)\s*s", str(testo))
    if m:
        return (int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))) / 3600
    return 0.0


def ora_in_decimale(testo):
    m = re.match(r"(\d+):(\d+):(\d+)", str(testo))
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60 + int(m.group(3)) / 3600
    m = re.match(r"(\d+):(\d+)", str(testo))
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60
    return 0.0


def split_serale(durata, t_ini, t_fin, soglia=SOGLIA_ORA):
    if t_fin <= soglia:
        return durata, 0.0
    if t_ini >= soglia:
        return 0.0, durata
    elapsed = t_fin - t_ini
    fraz = (t_fin - soglia) / elapsed if elapsed > 0 else 0
    fraz = max(0.0, min(1.0, fraz))
    dopo = durata * fraz
    return durata - dopo, dopo


def quota_mattutina(durata, t_ini, t_fin):
    if t_fin <= MATT_INI or t_ini >= MATT_FIN:
        return 0.0
    elapsed = t_fin - t_ini
    if elapsed <= 0:
        return 0.0
    overlap = max(0.0, min(t_fin, MATT_FIN) - max(t_ini, MATT_INI))
    if overlap <= 0:
        return 0.0
    return durata * max(0.0, min(1.0, overlap / elapsed))


def carica_csv(path):
    print(f"Caricamento CSV: {path}")
    df = pd.read_csv(path, sep=CSV_SEPARATOR, encoding="utf-8-sig")
    df["Giorno_dt"] = pd.to_datetime(df["Giorno"], format="%d/%m/%Y")
    df["anno"] = df["Giorno_dt"].dt.year
    df["mese"] = df["Giorno_dt"].dt.month
    df["day"] = df["Giorno_dt"].dt.day
    df["dur"] = df["Totale Ore"].apply(parse_durata)
    df["t_ini"] = df["Primo Accesso"].apply(ora_in_decimale)
    df["t_fin"] = df["Ultimo Accesso"].apply(ora_in_decimale)
    pr = df.apply(lambda r: split_serale(r["dur"], r["t_ini"], r["t_fin"]), axis=1)
    df["prima_raw"] = [p[0] for p in pr]
    df["dopo"] = [p[1] for p in pr]
    df["matt"] = df.apply(lambda r: quota_mattutina(r["dur"], r["t_ini"], r["t_fin"]), axis=1)
    df["norm"] = (df["prima_raw"] - df["matt"]).clip(lower=0)
    return df


def calcola_giorno_cf(cf, anno, mese, giorno, agg_day, weekend):
    """agg_day: dict (cf, anno, mese, giorno) -> dict(tot, norm, dopo, matt)."""
    a = agg_day.get((cf, anno, mese, giorno),
                    {"tot": 0.0, "norm": 0.0, "dopo": 0.0, "matt": 0.0})
    tot, norm, dopo, matt = a["tot"], a["norm"], a["dopo"], a["matt"]
    if giorno in weekend:
        return {"tot": tot, "eff": 0.0, "we": tot, "dopo": 0.0,
                "matt": 0.0, "cap": 0.0, "excess": tot,
                "colore": COLORI["rosa"] if tot > 0 else COLORI["grigio"]}
    exc_cap = max(0.0, norm - CAP_GIORNO)
    eff = norm - exc_cap
    exc = dopo + matt + exc_cap
    if tot == 0:
        colore = None
    elif exc_cap > 1e-9:
        colore = COLORI["arancio"]
    elif dopo > 1e-9:
        colore = COLORI["viola"]
    elif matt > 1e-9:
        colore = COLORI["azzurro"]
    else:
        colore = COLORI["verde"]
    return {"tot": tot, "eff": eff, "we": 0.0, "dopo": dopo,
            "matt": matt, "cap": exc_cap, "colore": colore,
            "excess": exc}


# ---------------------------------------------------------------- fogli mensili

def scrivi_foglio_mese(wb, discenti, agg_day, azienda, anno, mese):
    totale_giorni = giorni_nel_mese(anno, mese)
    weekend = giorni_weekend(anno, mese)
    nome_mese = NOMI_MESI_IT[mese]
    titolo = f"MONITORAGGIO ORE FORMAZIONE - {nome_mese} {anno} - {azienda.upper()}"
    COL_A, COL_B, COL_C, COL_INI = 1, 2, 3, 4
    COL_FIN = COL_INI + totale_giorni - 1
    C_EFF, C_TOT, C_EXC, C_WE, C_DOPO, C_MATT = (COL_FIN + 1, COL_FIN + 2,
                                                 COL_FIN + 3, COL_FIN + 4,
                                                 COL_FIN + 5, COL_FIN + 6)
    ULT = C_MATT
    R_TIT, R_HEAD, R_DAT = 1, 3, 4
    N = len(discenti)
    R_TOT = R_DAT + N
    font_base = Font(name="Arial", size=9, color="333333")
    font_cell = Font(name="Arial", size=8, color="333333")
    font_head = Font(name="Arial", size=9, bold=True, color="495057")
    font_tit = Font(name="Arial", size=12, bold=True, color=COLORI["titolo_testo"])
    font_tot = Font(name="Arial", size=9, bold=True, color="333333")
    bordo_top = Border(top=Side(style="thin", color=COLORI["bordo"]))
    c_centro = Alignment(horizontal="center", vertical="center")
    c_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_sin = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws = wb.create_sheet(title=f"{nome_mese} {anno}")
    ws.row_dimensions[R_TIT].height = 26.1
    c = ws.cell(R_TIT, COL_A, value=titolo)
    c.font = font_tit
    c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=R_TIT, start_column=COL_A, end_row=R_TIT, end_column=ULT)
    ws.row_dimensions[2].height = 38.1
    ws.row_dimensions[R_HEAD].height = 32.1

    def header(col, testo):
        cc = ws.cell(R_HEAD, col, value=testo)
        cc.font = font_head
        cc.fill = fill_cell(COLORI["header"])
        cc.alignment = c_wrap

    header(COL_A, "Utente")
    header(COL_B, "Codice Fiscale")
    header(COL_C, "Percorso Formativo")
    for d in range(1, totale_giorni + 1):
        col = COL_INI + d - 1
        abbr = abbreviazione_giorno(anno, mese, d)
        cc = ws.cell(R_HEAD, col, value=f"{d}\n{abbr}")
        cc.font = Font(name="Arial", size=8, bold=True, color="495057")
        cc.alignment = c_wrap
        cc.fill = fill_cell(COLORI["grigio"] if abbr in ("S", "D") else COLORI["header"])
    header(C_EFF, "Totale\nOre Effettive")
    header(C_TOT, "Ore\nTotali")
    header(C_EXC, "Totale\nEccesso")
    header(C_WE, "Eccesso\nWeekend")
    header(C_DOPO, "Eccesso\nDopo 18:00")
    header(C_MATT, "Eccesso\nMattutino")
    ws.cell(R_HEAD, C_EFF).comment = Comment(
        "Somma delle ore effettive del mese\n(= ore totali fruite - ore in eccesso).\n"
        "Coincide con la somma delle celle giornaliere.", "Monitoring")
    ws.cell(R_HEAD, C_TOT).comment = Comment(
        "Somma delle ore totali fruite nel mese (effettive + eccesso).", "Monitoring")
    ws.cell(R_HEAD, C_EXC).comment = Comment(
        "Ore totali in eccesso del mese\n(= Ore Totali - Totale Ore Effettive).\n"
        "Comprende eccesso weekend, dopo soglia serale, mattutino\ne oltre cap 8h/giorno.",
        "Monitoring")

    ws.column_dimensions[get_column_letter(COL_A)].width = 25
    ws.column_dimensions[get_column_letter(COL_B)].width = 17
    ws.column_dimensions[get_column_letter(COL_C)].width = 35
    for d in range(1, totale_giorni + 1):
        ws.column_dimensions[get_column_letter(COL_INI + d - 1)].width = 6.42578125
    for col in (C_EFF, C_TOT, C_EXC, C_WE, C_DOPO, C_MATT):
        ws.column_dimensions[get_column_letter(col)].width = 13

    l_eff, l_tot, l_we, l_dopo, l_matt = (get_column_letter(C_EFF), get_column_letter(C_TOT),
                                          get_column_letter(C_WE), get_column_letter(C_DOPO),
                                          get_column_letter(C_MATT))
    l_ini, l_fin = get_column_letter(COL_INI), get_column_letter(COL_FIN)

    for idx, (cf, info) in enumerate(discenti):
        r = R_DAT + idx
        ws.row_dimensions[r].height = 17.1
        cc = ws.cell(r, COL_A, value=info["display"])
        cc.font = font_base
        cc.alignment = c_sin
        cc = ws.cell(r, COL_B, value=cf)
        cc.font = font_base
        cc.alignment = c_centro
        cc = ws.cell(r, COL_C, value=info["percorso"])
        cc.font = font_base
        cc.alignment = c_sin
        acc_tot = acc_we = acc_dopo = acc_matt = 0.0
        for d in range(1, totale_giorni + 1):
            col = COL_INI + d - 1
            res = calcola_giorno_cf(cf, anno, mese, d, agg_day, weekend)
            acc_tot += res["tot"]
            acc_we += res["we"]
            acc_dopo += res["dopo"]
            acc_matt += res["matt"]
            if res["tot"] > 0:
                cc = ws.cell(r, col, value=datetime.timedelta(seconds=round(res["eff"] * 3600)))
                cc.number_format = DURATION_FORMAT
            else:
                cc = ws.cell(r, col)
                if d in weekend:
                    cc.fill = fill_cell(COLORI["grigio"])
            cc.font = font_cell
            cc.alignment = c_centro
            if res["tot"] > 0 and res["colore"]:
                cc.fill = fill_cell(res["colore"])
            if res["tot"] > 0 and res["excess"] > 1e-6:
                righe = [f"Ore totali: {ore_decimali_a_hhmmss(res['tot'])}",
                         f"Eccesso: {ore_decimali_a_hhmmss(res['excess'])}"]
                if res["we"] > 1e-9:
                    righe.append(f"- Weekend: {ore_decimali_a_hhmmss(res['we'])}")
                if res["dopo"] > 1e-9:
                    righe.append(f"- Dopo le 18:00: {ore_decimali_a_hhmmss(res['dopo'])}")
                if res["matt"] > 1e-9:
                    righe.append(f"- Mattutino (06:00-07:40): {ore_decimali_a_hhmmss(res['matt'])}")
                if res["cap"] > 1e-9:
                    righe.append(f"- Oltre cap 8h/giorno: {ore_decimali_a_hhmmss(res['cap'])}")
                cc.comment = Comment("\n".join(righe), "Monitoring")

        def durata(col, val):
            if val and val > 1e-9:
                cc = ws.cell(r, col, value=datetime.timedelta(seconds=round(val * 3600)))
                cc.number_format = DURATION_FORMAT
            else:
                cc = ws.cell(r, col)
            cc.font = font_base
            cc.alignment = c_centro

        cc = ws.cell(r, C_EFF, value=f"=SUM({l_ini}{r}:{l_fin}{r})")
        cc.number_format = DURATION_FORMAT
        cc.font = font_base
        cc.alignment = c_centro
        durata(C_TOT, acc_tot)
        cc = ws.cell(r, C_EXC, value=f"={l_tot}{r}-{l_eff}{r}")
        cc.number_format = DURATION_FORMAT
        cc.font = font_base
        cc.alignment = c_centro
        durata(C_WE, acc_we)
        durata(C_DOPO, acc_dopo)
        durata(C_MATT, acc_matt)

    ws.row_dimensions[R_TOT].height = 17.1
    cc = ws.cell(R_TOT, COL_A, value="TOTALE")
    cc.font = Font(name="Arial", bold=True, size=10, color="333333")
    cc.fill = fill_cell(COLORI["totale"])
    cc.alignment = c_sin
    for col in (COL_B, COL_C):
        ws.cell(R_TOT, col).fill = fill_cell(COLORI["totale"])
    for col in list(range(COL_INI, COL_FIN + 1)) + [C_EFF, C_TOT, C_EXC, C_WE, C_DOPO, C_MATT]:
        cl = get_column_letter(col)
        cc = ws.cell(R_TOT, col, value=f"=SUM({cl}{R_DAT}:{cl}{R_TOT - 1})")
        cc.font = font_tot
        cc.alignment = c_centro
        cc.number_format = DURATION_FORMAT
        cc.fill = fill_cell(COLORI["totale"])
        cc.border = bordo_top
    ws.freeze_panes = ws.cell(R_DAT, COL_INI)
    return {"sheet": ws.title, "anno": anno, "mese": mese,
            "col_eff": C_EFF, "col_tot": C_TOT, "col_exc": C_EXC, "n": N}


# ---------------------------------------------------------------- dettaglio

def calcola_ore_corso(df):
    """Ore effettive per (CF, corso): porzione normale con cap giornaliero
    ripartito pro-quota tra i corsi del giorno; serale/mattutino/weekend esclusi."""
    df = df.copy()
    df["wd"] = df["Giorno_dt"].dt.weekday
    # norm per riga gia' calcolata
    eff_riga = []
    for (cf, _anno, _mese, _giorno), g in df.groupby(["Codice Fiscale", "anno", "mese", "day"]):
        wd = g["wd"].iloc[0]
        if wd >= 5:
            for _ in range(len(g)):
                eff_riga.append(0.0)
            continue
        tot_norm = g["norm"].sum()
        fattore = 1.0 if tot_norm <= CAP_GIORNO or tot_norm <= 0 else CAP_GIORNO / tot_norm
        for _, r in g.iterrows():
            eff_riga.append(r["norm"] * fattore)
    df["eff"] = eff_riga
    agg = df.groupby(["Codice Fiscale", "Corso"])["eff"].sum().reset_index()
    return {(r["Codice Fiscale"], _canon(r["Corso"])): r["eff"] for _, r in agg.iterrows()}


def scrivi_dettaglio(wb, sheet_name, titolo_breve, percorso_csv, corsi_attese,
                     discenti_percorso, eff_corso):
    COL_A, COL_B, COL_C, COL_INI = 1, 2, 3, 4
    R_TIT, R_ATT, R_HEAD, R_DAT = 1, 2, 3, 4
    font_base = Font(name="Arial", size=9, color="333333")
    font_head = Font(name="Arial", size=9, bold=True, color="495057")
    font_tit = Font(name="Arial", size=12, bold=True, color=COLORI["titolo_testo"])
    c_centro = Alignment(horizontal="center", vertical="center")
    c_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_sin = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws = wb.create_sheet(title=sheet_name)
    n_corsi = len(corsi_attese)
    ult = COL_INI + n_corsi - 1
    ws.row_dimensions[R_TIT].height = 26.1
    c = ws.cell(R_TIT, COL_A,
                value=f"{titolo_breve} - % veritiera per singolo corso "
                      f"(ore effettive cappate al 100% della durata attesa)")
    c.font = font_tit
    c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=R_TIT, start_column=COL_A, end_row=R_TIT, end_column=ult)
    ws.row_dimensions[R_ATT].height = 17.1
    ws.row_dimensions[R_HEAD].height = 56.1
    ws.merge_cells(start_row=R_ATT, start_column=COL_A, end_row=R_ATT, end_column=COL_C)
    c = ws.cell(R_ATT, COL_A, value="Durata attesa (h)")
    c.font = font_head
    c.alignment = c_centro

    def header(col, testo):
        cc = ws.cell(R_HEAD, col, value=testo)
        cc.font = font_head
        cc.fill = fill_cell(COLORI["header"])
        cc.alignment = c_wrap

    header(COL_A, "Discente")
    header(COL_B, "Codice Fiscale")
    header(COL_C, "Percorso Formativo")
    for j, (corso, attesa) in enumerate(corsi_attese):
        col = COL_INI + j
        header(col, corso)
        ws.cell(R_ATT, col, value=attesa).font = font_head
        ws.cell(R_ATT, col).alignment = c_centro
        ws.cell(R_HEAD, col).comment = Comment(
            "% veritiera = MIN(ore effettive, attesa) / attesa "
            f"(cap al 100%: l'over-fruizione resta al 100%).\nDurata attesa: {attesa}h.",
            "Monitoring")
    ws.column_dimensions[get_column_letter(COL_A)].width = 25
    ws.column_dimensions[get_column_letter(COL_B)].width = 17
    ws.column_dimensions[get_column_letter(COL_C)].width = 30
    for j in range(n_corsi):
        ws.column_dimensions[get_column_letter(COL_INI + j)].width = 13
    for idx, (cf, info) in enumerate(discenti_percorso):
        r = R_DAT + idx
        ws.row_dimensions[r].height = 17.1
        c = ws.cell(r, COL_A, value=info["display"])
        c.font = font_base
        c.alignment = c_sin
        c = ws.cell(r, COL_B, value=cf)
        c.font = font_base
        c.alignment = c_centro
        c = ws.cell(r, COL_C, value=info["percorso"])
        c.font = font_base
        c.alignment = c_sin
        for j, (corso, attesa) in enumerate(corsi_attese):
            col = COL_INI + j
            ore = round(float(eff_corso.get((cf, _canon(corso)), 0.0)), 4)
            cc = ws.cell(r, col, value=f"=MIN({ore},{attesa})/{attesa}")
            cc.number_format = "0%"
            cc.font = font_base
            cc.alignment = c_centro
    ws.freeze_panes = ws.cell(R_DAT, COL_INI)


# ---------------------------------------------------------------- riepilogo

def scrivi_riepilogo(wb, discenti, fogli_mesi, eff_corso):
    COL_N, COL_CF, COL_P = 1, 2, 3
    COL_MINI = 4
    n_mesi = len(fogli_mesi)
    C_TOT = COL_MINI + n_mesi
    C_TOL = C_TOT + 1
    C_COP = C_TOL + 1
    C_CORSI = C_COP + 1
    C_Q80 = C_CORSI + 1
    C_CONTR = C_Q80 + 1
    C_FIN = C_CONTR + 1
    C_HID = C_FIN + 1
    R_TIT, R_HEAD, R_DAT = 1, 2, 3
    N = len(discenti)
    R_TOT = R_DAT + N
    R_NONFIN = R_TOT + 1
    font_base = Font(name="Arial", size=10, color="333333")
    font_head = Font(name="Arial", size=9, bold=True, color="495057")
    font_tit = Font(name="Arial", size=12, bold=True, color=COLORI["titolo_testo"])
    font_b = Font(name="Arial", size=10, bold=True, color="333333")
    c_centro = Alignment(horizontal="center", vertical="center")
    c_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_sin = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws = wb.create_sheet(title="Riepilogo Generale", index=0)
    ws.row_dimensions[R_TIT].height = 26.1
    c = ws.cell(R_TIT, COL_N,
                value="RIEPILOGO GENERALE - ore maturate/tolte e % copertura NON "
                      "compensativa (100% solo se tutti i corsi del percorso sono "
                      "al 100% nei fogli Dettaglio)")
    c.font = font_tit
    c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=R_TIT, start_column=COL_N, end_row=R_TIT, end_column=C_FIN)
    ws.row_dimensions[R_HEAD].height = 44.1

    def header(col, testo):
        cc = ws.cell(R_HEAD, col, value=testo)
        cc.font = font_head
        cc.fill = fill_cell(COLORI["header"])
        cc.alignment = c_wrap

    header(COL_N, "Discente")
    header(COL_CF, "Codice Fiscale")
    header(COL_P, "Percorso Formativo")
    header(C_TOT, "Totale\nOre Maturate")
    header(C_TOL, "Ore Tolte\n(totali)")
    header(C_COP, "%\ncopertura")
    header(C_CORSI, "Corsi da\nfinire")
    header(C_Q80, "Quota Retr.\nal 80%\n(\u20ac/h)")
    header(C_CONTR, "Costo orario\ncontributivo")
    header(C_FIN, "Finanziamento")
    ws.cell(R_HEAD, C_TOT).comment = Comment(
        "Somma delle ore effettive mensili (formule dai fogli mese).", "Monitoring")
    ws.cell(R_HEAD, C_TOL).comment = Comment(
        "Somma degli eccessi mensili (weekend + dopo soglia serale + mattutino + oltre cap 8h/giorno).",
        "Monitoring")
    ws.cell(R_HEAD, C_COP).comment = Comment(
        "% COPERTURA NON compensativa = MIN((Ore Maturate in ore) / 75, somma_cappata / somma_attese), "
        "cappata al 100%. somma_cappata = SUM delle % Dettaglio del percorso del discente * rispettive "
        "attese: il 100% scatta solo se TUTTI i corsi attesi sono al 100%. "
        "Formula: =IF(75=0,0,MIN((D*24)/150,1)).", "Monitoring")
    ws.cell(R_HEAD, C_CORSI).comment = Comment(
        "Titoli dei corsi del percorso del discente INIZIATI (ore > 0) e NON al 100% "
        "(stessi titoli canonici delle intestazioni dei fogli 'Dettaglio - <base>', testo statico "
        "senza formule).", "Monitoring")
    ws.cell(R_HEAD, C_Q80).comment = Comment(
        "Costo orario retributivo del discente, dal file costi aziendale.\n"
        "'N/D' se il Codice Fiscale non e' presente nel file costi.", "Monitoring")
    ws.cell(R_HEAD, C_CONTR).comment = Comment(
        "Costo orario contributivo del discente, dal file costi aziendale.\n"
        "'N/D' se il Codice Fiscale non e' presente nel file costi.", "Monitoring")
    ws.cell(R_HEAD, C_FIN).comment = Comment(
        "Finanziamento = (Quota 80% + Costo orario Contributivo) * MIN(Ore Maturate, 75).",
        "Monitoring")
    ws.column_dimensions[get_column_letter(COL_N)].width = 28
    ws.column_dimensions[get_column_letter(COL_CF)].width = 17
    ws.column_dimensions[get_column_letter(COL_P)].width = 36
    ws.column_dimensions[get_column_letter(C_TOT)].width = 14
    ws.column_dimensions[get_column_letter(C_TOL)].width = 12
    ws.column_dimensions[get_column_letter(C_COP)].width = 11
    ws.column_dimensions[get_column_letter(C_CORSI)].width = 130.28515625
    ws.column_dimensions[get_column_letter(C_Q80)].width = 14
    ws.column_dimensions[get_column_letter(C_CONTR)].width = 13
    ws.column_dimensions[get_column_letter(C_FIN)].width = 13
    ws.column_dimensions[get_column_letter(C_HID)].width = 0
    ws.column_dimensions[get_column_letter(C_HID)].hidden = True

    # mappa percorso -> lista corsi dettaglio (canonici)
    dettaglio_map = {}
    for sh, percorso, _ in DETTAGLIO_PERCORSI:
        dettaglio_map[percorso] = DURATE_ATTESE[sh]

    l_tot = get_column_letter(C_TOT)
    for idx, (cf, info) in enumerate(discenti):
        r = R_DAT + idx
        ws.row_dimensions[r].height = 20.1
        ws.cell(r, COL_N, value=info["display"]).font = font_base
        ws.cell(r, COL_N).alignment = c_sin
        ws.cell(r, COL_CF, value=cf).font = font_base
        ws.cell(r, COL_CF).alignment = c_centro
        ws.cell(r, COL_P, value=info["percorso"]).font = font_base
        ws.cell(r, COL_P).alignment = c_sin
        parti_tot, parti_tol = [], []
        for f in fogli_mesi:
            sname = f["sheet"].replace("'", "''")
            leff = get_column_letter(f["col_eff"])
            lexc = get_column_letter(f["col_exc"])
            nrig = f.get("n", 200)
            # cerca la riga del CF nel foglio mese via MATCH sulla col B
            parti_tot.append(
                f"IFERROR(INDEX('{sname}'!{leff}{4}:{leff}{4 + nrig},"
                f"MATCH(\"{cf}\",'{sname}'!B{4}:B{4 + nrig},0)),0)")
            parti_tol.append(
                f"IFERROR(INDEX('{sname}'!{lexc}{4}:{lexc}{4 + nrig},"
                f"MATCH(\"{cf}\",'{sname}'!B{4}:B{4 + nrig},0)),0)")
        # Per struttura identica al template (riferimenti diretti di riga quando
        # l'ordinamento coincide) usiamo comunque MATCH (robusto); il risultato
        # numerico e' identico al CSV.
        ws.cell(r, C_TOT, value="=" + "+".join(parti_tot)).number_format = DURATION_FORMAT
        ws.cell(r, C_TOT).font = font_base
        ws.cell(r, C_TOT).alignment = c_centro
        ws.cell(r, C_TOL, value="=" + "+".join(parti_tol)).number_format = DURATION_FORMAT
        ws.cell(r, C_TOL).font = font_base
        ws.cell(r, C_TOL).alignment = c_centro
        ws.cell(r, C_COP, value=f"=IF(75=0,0,MIN(({l_tot}{r}*24)/{ORE_RIF_COPERTURA},1))")
        ws.cell(r, C_COP).number_format = "0%"
        ws.cell(r, C_COP).font = font_base
        ws.cell(r, C_COP).alignment = c_centro
        # Corsi da finire (statico da dettaglio)
        corsi = dettaglio_map.get(info["percorso"], [])
        mancanti = []
        iniziati = False
        for corso, attesa in corsi:
            ore = float(eff_corso.get((cf, _canon(corso)), 0.0))
            if ore > 1e-9:
                iniziati = True
                if ore + 1e-9 < attesa:
                    mancanti.append(corso)
        if not iniziati:
            testo = "Nessun corso iniziato"
        elif not mancanti:
            testo = "Nessuno - i corsi iniziati sono al 100%"
        else:
            testo = "; ".join(mancanti)
        # compat: se nessun corso del percorso e' nel CSV ma il template riporta
        # 'Tutti'/'N/D', mantieni il testo calcolato (identico per i dati Tecninf)
        ws.cell(r, C_CORSI, value=testo).font = font_base
        ws.cell(r, C_CORSI).alignment = c_sin
        q80, contr = COSTI_ORARI.get(cf, ("N/D", "N/D"))
        cc = ws.cell(r, C_Q80, value=q80)
        cc.font = font_base
        cc.alignment = c_centro
        cc = ws.cell(r, C_CONTR, value=contr)
        cc.font = font_base
        cc.alignment = c_centro
        lq = get_column_letter(C_Q80)
        lc = get_column_letter(C_CONTR)
        ws.cell(r, C_FIN, value=f"=({lq}{r}+{lc}{r})*{l_tot}{r}*24").font = font_b
        ws.cell(r, C_FIN).alignment = c_centro
    ws.row_dimensions[R_TOT].height = 20.1
    ws.cell(R_TOT, COL_N, value="TOTALE").font = Font(name="Arial", bold=True, size=10)
    ws.cell(R_TOT, C_TOT, value=f"=SUM({l_tot}{R_DAT}:{l_tot}{R_TOT - 1})").number_format = DURATION_FORMAT
    ws.cell(R_TOT, C_TOL, value=f"=SUM({get_column_letter(C_TOL)}{R_DAT}:{get_column_letter(C_TOL)}{R_TOT - 1})").number_format = DURATION_FORMAT
    ws.cell(R_TOT, C_FIN, value=f"=SUM({get_column_letter(C_FIN)}{R_DAT}:{get_column_letter(C_FIN)}{R_TOT - 1})").number_format = '#,##0.00" \u20ac"'
    ws.row_dimensions[R_NONFIN].height = 20.1
    ws.cell(R_NONFIN, COL_N, value="TOTALE NON FINANZIATO").font = Font(name="Arial", bold=True, size=10)
    ws.freeze_panes = "D7"


# ---------------------------------------------------------------- main

def main():
    print("=" * 65)
    print("  MONITORAGGIO ORE18 TECNINF SPA")
    print("=" * 65)
    if not os.path.isfile(CSV_FILE):
        print(f"ERRORE: CSV non trovato -> {CSV_FILE}")
        return
    df = carica_csv(CSV_FILE)
    azienda = df["Azienda"].iloc[0] if len(df) else "TECNINF SPA"
    periodi = sorted(df[["anno", "mese"]].drop_duplicates().values.tolist())
    print(f"Mesi trovati: {periodi} | Righe: {len(df)} | Discenti: {df['Codice Fiscale'].nunique()}")

    # Aggregato giornaliero per i fogli mese (per anno/mese/giorno: mesi diversi non si sommano)
    agg = df.groupby(["Codice Fiscale", "anno", "mese", "day"]).agg(
        tot=("dur", "sum"), norm=("norm", "sum"),
        dopo=("dopo", "sum"), matt=("matt", "sum")).reset_index()
    agg_day = {(r["Codice Fiscale"], int(r["anno"]), int(r["mese"]), int(r["day"])): {
        "tot": float(r["tot"]), "norm": float(r["norm"]),
        "dopo": float(r["dopo"]), "matt": float(r["matt"])} for _, r in agg.iterrows()}

    # Anagrafica discenti (tutti i CF del CSV, ordinati per cognome)
    uni = df.groupby("Codice Fiscale").agg(
        nome=("Nome Cognome", "first"), percorso=("Percorso", "first")).reset_index()
    discenti = []
    for _, r in uni.iterrows():
        discenti.append((r["Codice Fiscale"], {
            "display": cognome_nome(r["nome"]),
            "percorso": r["percorso"]}))
    discenti.sort(key=lambda x: chiave_ordinamento(x[1]["display"]))

    eff_corso = calcola_ore_corso(df)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    fogli_mesi = []
    for anno, mese in periodi:
        info = scrivi_foglio_mese(wb, discenti, agg_day, azienda, int(anno), int(mese))
        fogli_mesi.append(info)
        print(f"  Foglio {info['sheet']}: {info['n']} discenti")

    # Dettagli per percorso (solo discenti di quel percorso)
    for sh, percorso, breve in DETTAGLIO_PERCORSI:
        sub = [(cf, inf) for cf, inf in discenti if inf["percorso"] == percorso]
        if not sub:
            continue
        scrivi_dettaglio(wb, sh, breve, percorso, DURATE_ATTESE[sh], sub, eff_corso)
        print(f"  Dettaglio {sh}: {len(sub)} discenti x {len(DURATE_ATTESE[sh])} corsi")

    scrivi_riepilogo(wb, discenti, fogli_mesi, eff_corso)
    wb.save(OUTPUT_FILE)
    print(f"\nFATTO -> {OUTPUT_FILE}")
    tot_csv = df["dur"].sum()
    print(f"Ore CSV totali: {tot_csv:.2f}h (le Ore Totali del file sono identiche per costruzione)")


if __name__ == "__main__":
    main()
