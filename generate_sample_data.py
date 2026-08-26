# -*- coding: utf-8 -*-
"""Generates SYNTHETIC demo data for training_hours_monitor.py:

  - sample_data/Report_Accessi.csv  : export fittizio di accessi alla piattaforma
  - sample_data/Presenze.xlsx       : registro presenze (LUL) fittizio, giugno 2026

All names, tax IDs, emails and data are invented and deterministically
generated (fixed seed). No real data is contained in this repository.

The data deliberately covers every case the engine can handle:
evening usage (after 18:00), weekends, early-morning windows (06:00-07:40),
days over 8 hours, learners with multiple programmes, and employees
(present in the attendance register) vs external learners (not in it).

Usage:
    python generate_sample_data.py
"""

import csv
import datetime
import os
import random
import string

import openpyxl

random.seed(42)

BASE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE, "sample_data")
CSV_PATH = os.path.join(SAMPLE_DIR, "Report_Accessi.csv")
LUL_PATH = os.path.join(SAMPLE_DIR, "Presenze.xlsx")

AZIENDA_CSV = "Azienda Demo S.r.l."
MESE_LUL = (2026, 6)          # il registro presenze riguarda giugno 2026
MESI_CSV = [(2026, 6), (2026, 7)]

CORSI = [
    "Sicurezza sul Lavoro (D.Lgs. 81/08)",
    "Excel Avanzato per il Data Analysis",
    "Competenze Digitali di Base",
    "Privacy e GDPR in Azienda",
    "Productivity e Gestione del Tempo",
]
PERCORSI = ["Percorso A - Fondamenti", "Percorso B - Specializzazione"]

# (nome, cognome, dipendente, profilo)
# profilo: normal | evening | weekend | early | heavy
DISCENTI = [
    ("Maria",   "Rossi",    True,  "normal"),
    ("Luca",    "Bianchi",  True,  "normal"),
    ("Giulia",  "Ferrari",  True,  "evening"),
    ("Paolo",   "Romano",   True,  "evening"),
    ("Sara",    "Gallo",    True,  "weekend"),
    ("Marco",   "Conti",    True,  "weekend"),
    ("Elena",   "Ricci",    True,  "early"),
    ("Davide",  "Moretti",  True,  "heavy"),
    ("Chiara",  "Fontana",  False, "normal"),
    ("Simone",  "Villa",    False, "evening"),
    ("Martina", "Serra",    False, "weekend"),
    ("Alessandro", "Longo", False, "early"),
]


def codice_fiscale_fittizio():
    """Genera un codice fiscale PLAUSIBILE ma completamente casuale (16 caratteri)."""
    lettere = string.ascii_uppercase
    return "".join(random.choice(lettere) for _ in range(6)) \
        + f"{random.randint(0, 99):02d}" + random.choice(lettere) \
        + f"{random.randint(0, 99):02d}" + random.choice(lettere) \
        + f"{random.randint(0, 999):03d}" + random.choice(lettere)


def fmt_hms(ore_dec):
    sec = int(round(ore_dec * 3600))
    return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}:{sec % 60:02d}"


def fmt_durata(ore_dec):
    sec = int(round(ore_dec * 3600))
    return f"{sec // 3600}h {(sec % 3600) // 60}min {sec % 60}s"


def sessione(profilo, giorno):
    """Restituisce (primo_accesso_dec, durata_ore) secondo il profilo del discente."""
    if profilo == "evening":
        inizio = random.uniform(20.0, 21.5)
        durata = random.uniform(0.8, 2.0)
    elif profilo == "early":
        inizio = random.uniform(6.10, 6.85)          # dentro la finestra 06:00-07:40
        durata = random.uniform(0.5, 1.5)
    elif profilo == "heavy" and random.random() < 0.55:
        inizio = random.uniform(8.5, 9.5)
        durata = random.uniform(2.0, 3.5)            # piu' sessioni nello stesso giorno
    else:
        inizio = random.uniform(9.0, 11.5)
        durata = random.uniform(0.7, 2.8)
    return inizio, durata


def genera_csv():
    righe = []
    for nome, cognome, _, profilo in DISCENTI:
        cf = codice_fiscale_fittizio()
        email = f"{nome.lower()}.{cognome.lower()}@example.com"
        corsi_discente = random.sample(CORSI, random.choice([1, 1, 2]))
        percorso = random.choice(PERCORSI)
        for anno, mese in MESI_CSV:
            giorni_mese = (datetime.date(anno + (mese == 12), mese % 12 + 1, 1)
                           - datetime.date(anno, mese, 1)).days
            for giorno in range(1, giorni_mese + 1):
                d = datetime.date(anno, mese, giorno)
                # frequenza di fruizione
                if profilo == "weekend":
                    frequenza = 0.70 if d.weekday() >= 5 else 0.10
                else:
                    frequenza = 0.10 if d.weekday() >= 5 else 0.65
                if random.random() > frequenza:
                    continue
                n_sessioni = random.choice([1, 1, 1, 2, 2, 3]) if profilo == "heavy" else 1
                for _ in range(n_sessioni):
                    inizio, durata = sessione(profilo, d)
                    corso = random.choice(corsi_discente)
                    righe.append({
                        "Azienda": AZIENDA_CSV,
                        "Email": email,
                        "Nome Cognome": f"{nome} {cognome}",
                        "Codice Fiscale": cf,
                        "Percorso": percorso,
                        "Giorno": d.strftime("%d/%m/%Y"),
                        "Primo Accesso": fmt_hms(inizio),
                        "Ultimo Accesso": fmt_hms(inizio + durata),
                        "Totale Ore": fmt_durata(durata),
                        "Corso": corso,
                    })
    os.makedirs(SAMPLE_DIR, exist_ok=True)
    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=[
            "Azienda", "Email", "Nome Cognome", "Codice Fiscale", "Percorso",
            "Giorno", "Primo Accesso", "Ultimo Accesso", "Totale Ore", "Corso"],
            delimiter=";", quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        w.writerows(righe)
    print(f"[OK] {len(righe)} righe -> {CSV_PATH}")


def genera_lul():
    """Crea un registro presenze (formato LUL) sintetico per MESE_LUL.

    Layout atteso dal parser (vedi carica_lul in training_hours_monitor.py):
      - una riga con 'ORE' in colonna A e i giorni ('1 L', '2 M', ...) a partire
        dalla colonna C;
      - un blocco per dipendente: riga 'Dipendente: <matr> - COGNOME NOME',
        riga 'LAV' con le ore lavorate allineate alle colonne dei giorni,
        opzionale riga 'A S S E N Z E' con i codici di assenza.
    """
    anno, mese = MESE_LUL
    giorni_mese = (datetime.date(anno + (mese == 12), mese % 12 + 1, 1)
                   - datetime.date(anno, mese, 1)).days
    lettere = ["L", "M", "M", "G", "V", "S", "D"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registro Presenze"

    ws.append(["REGISTRO PRESENZE - DATI DIMOSTRATIVI SINTETICI"])
    ws.append(["Azienda Demo S.r.l."])
    ws.append([])
    ws.append([])

    # Riga 'ORE' con la mappa dei giorni (colonna C in poi)
    riga_ore = ["ORE", None]
    for g in range(1, giorni_mese + 1):
        riga_ore.append(f"{g} {lettere[datetime.date(anno, mese, g).weekday()]}")
    ws.append(riga_ore)
    col_giorno = {c: g + 1 for c, g in enumerate(range(giorni_mese), start=2)}

    matricola = 100
    for nome, cognome, dipendente, _ in DISCENTI:
        if not dipendente:
            continue                    # gli esterni non compaiono nel registro
        matricola += 1
        ws.append([])
        ws.append([f"Dipendente: {matricola} - {cognome.upper()} {nome.upper()}"])
        riga_lav = ["LAV", None] + [None] * giorni_mese
        for c, g in col_giorno.items():
            if datetime.date(anno, mese, g).weekday() < 5:
                riga_lav[c] = 8         # 8 ore nei giorni feriali
        ws.append(riga_lav)
        # Un dipendente con un giorno di ferie codificato (mezza giornata)
        if cognome == "Conti":
            riga_ass = ["A S S E N Z E", None] + [None] * giorni_mese
            riga_ass[col_giorno[9]] = "F"
            ws.append(riga_ass)
            riga_ass_ore = ["", None] + [None] * giorni_mese
            riga_ass_ore[col_giorno[9]] = 8
            ws.append(riga_ass_ore)

    wb.save(LUL_PATH)
    n_dip = sum(1 for _, _, d, _ in DISCENTI if d)
    print(f"[OK] registro presenze ({n_dip} dipendenti, {giorni_mese} giorni) -> {LUL_PATH}")


if __name__ == "__main__":
    genera_csv()
    genera_lul()
    print("\nDati dimostrativi generati. Ora esegui:  python training_hours_monitor.py")
