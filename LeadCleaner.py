import re
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# 1) HARTE NO-GO WORTLISTE (fest im Script)
#    -> nur Teilwörter, klein, möglichst "hart"
# =========================
HARD_NO_GO = [
    # Staat / Verwaltung
    "polizei", "bundespolizei", "zoll", "finanzamt", "steueramt", "gericht", "staatsan",
    "ministerium", "regierung", "behörde", "bürgeramt", "einwohn", "ordnungsamt",
    "sozialamt", "jobcenter", "arbeitsagentur", "agentur für arbeit",

    # Gesundheit
    "kranken", "klinik", "klinikum", "arzt", "ärztin", "zahnarzt", "apothek",
    "gesundheitsamt", "rehazentrum",

    # Bildung / Forschung
    "schule", "grundschule", "realschule", "hauptschule", "gymnasium", "berufsschule",
    "hochschule", "universität", "fachhochschule", "akademie", "bildungszentrum",
    "vhs", "forsch", "institut",

    # Rettung / Katastrophe
    "feuerwehr", "rettung", "notarzt", "katastroph", "zivilschutz",

    # Religion
    "kirche", "pfarr", "gemeinde", "bistum", "diözese", "mosche", "islam", "synagog",

    # Non-Profit / Vereine
    "verein", "stiftung", "gemeinn", "e.v.", "caritas", "diakonie", "drk",
    "rotes kreuz", "malteser", "johanniter",

    # Recht
    "rechtsanwalt", "anwalt", "kanzlei", "notar", "notariat",

    # Sonstige klare No-Gos
    "friedhof", "denkmal", "archiv", "museum", "theater", "bibli",
]

WHITELIST = [
    "gmbh", "ug", "kg", "gbr", "ohg", "ag", "e.k", "ek",
    "gmbh & co", "gmbh&co", "se", "kgaa"
]

# Welche Spalten geprüft werden sollen (Header müssen exakt passen)
CHECK_COLS_CANDIDATES = [
    ["Vorname", "Name", "Zusatz"],
    ["NAME", "NACHNAME", "Zusatz"],
    ["NAME", "NACHNAME", "ZUSATZ"],
]

def compile_pattern(words, min_len):
    cleaned = []
    for w in words:
        w = (w or "").strip().lower()
        if len(w) >= min_len:
            cleaned.append(re.escape(w))
    if not cleaned:
        return re.compile(r"(?!x)x")  # matcht nie
    return re.compile("(" + "|".join(cleaned) + ")", re.IGNORECASE)

def find_check_cols(df):
    cols = list(df.columns)
    for cand in CHECK_COLS_CANDIDATES:
        if all(c in cols for c in cand):
            return cand
    raise KeyError(
        "Spalten nicht gefunden.\n"
        f"Vorhandene Spalten: {cols}\n"
        "Passe CHECK_COLS_CANDIDATES im Script an."
    )

def first_match(pattern, s):
    m = pattern.search(s)
    return m.group(0).lower() if m else ""

def clean_file(path: Path, no_go_pat, wl_pat):
    df = pd.read_excel(path, sheet_name=0)
    check_cols = find_check_cols(df)

    text = (
        df[check_cols]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )

    df["_WHITELIST_HIT"] = text.apply(lambda s: first_match(wl_pat, s))
    df["_MATCH_WORD"] = text.apply(lambda s: first_match(no_go_pat, s))

    mask_delete = (df["_MATCH_WORD"] != "") & (df["_WHITELIST_HIT"] == "")
    cleaned = df.loc[~mask_delete].drop(columns=["_MATCH_WORD", "_WHITELIST_HIT"], errors="ignore").copy()
    deleted = df.loc[mask_delete].copy()

    out_path = path.with_name(path.stem + "_CLEANED.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        cleaned.to_excel(writer, index=False, sheet_name="CLEANED")
        deleted.to_excel(writer, index=False, sheet_name="DELETED")

    return out_path, int(mask_delete.sum()), check_cols

# ---------------- GUI ----------------
selected_files = []

def log(msg: str):
    txt.insert("end", msg + "\n")
    txt.see("end")
    root.update_idletasks()

def pick_files():
    global selected_files
    files = filedialog.askopenfilenames(
        title="Excel-Dateien auswählen",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
    )
    selected_files = [Path(f) for f in files] if files else []
    lbl_files.config(text=f"{len(selected_files)} Datei(en) ausgewählt")
    if selected_files:
        log(f"Ausgewählt: {len(selected_files)} Datei(en)")

def run_clean():
    if not selected_files:
        messagebox.showwarning("Hinweis", "Bitte zuerst Excel-Dateien auswählen.")
        return

    no_go_pat = compile_pattern(HARD_NO_GO, min_len=3)
    wl_pat = compile_pattern(WHITELIST, min_len=2)

    ok = 0
    for f in selected_files:
        try:
            out, deleted_count, cols = clean_file(f, no_go_pat, wl_pat)
            log(f"✓ {f.name} -> {out.name} | gelöscht: {deleted_count} | geprüft: {cols}")
            ok += 1
        except Exception as e:
            log(f"✗ FEHLER {f.name}: {e}")

    messagebox.showinfo("Fertig", f"Fertig. Erfolgreich: {ok}/{len(selected_files)}\n\nOutput liegt im gleichen Ordner wie die Datei(en).")

root = tk.Tk()
root.title("Lead Cleaner")
root.geometry("640x360")

frame = tk.Frame(root)
frame.pack(padx=12, pady=10, fill="x")

btn_pick = tk.Button(frame, text="Excel-Dateien auswählen", command=pick_files, width=28)
btn_pick.pack(side="left", padx=(0, 10))

btn_run = tk.Button(frame, text="Bereinigen", command=run_clean, width=18)
btn_run.pack(side="left")

lbl_files = tk.Label(frame, text="0 Datei(en) ausgewählt")
lbl_files.pack(side="left", padx=12)

txt = tk.Text(root, height=16)
txt.pack(padx=12, pady=10, fill="both", expand=True)
log("Output: *_CLEANED.xlsx (Sheets: CLEANED, DELETED)")

root.mainloop()