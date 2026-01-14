import re
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# HARTE NO-GO WORTLISTE (fest im Script)
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

# Whitelist: wenn Treffer -> NIE löschen
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

# Debug-Spalten nur im DELETED-Sheet behalten?
DEBUG_IN_DELETED_ONLY = True


def compile_pattern(words, min_len: int) -> re.Pattern:
    cleaned = []
    for w in words:
        w = (w or "").strip().lower()
        if len(w) >= min_len:
            cleaned.append(re.escape(w))
    if not cleaned:
        # Pattern, das nie matcht
        return re.compile(r"(?!x)x")
    return re.compile("(" + "|".join(cleaned) + ")", re.IGNORECASE)


def find_check_cols(df: pd.DataFrame) -> list[str]:
    cols = list(df.columns)
    for cand in CHECK_COLS_CANDIDATES:
        if all(c in cols for c in cand):
            return cand
    raise KeyError(
        "Spalten nicht gefunden.\n"
        f"Vorhandene Spalten: {cols}\n"
        "Passe CHECK_COLS_CANDIDATES im Script an."
    )


def first_match(pattern: re.Pattern, s: str) -> str:
    m = pattern.search(s)
    return m.group(0).lower() if m else ""


def clean_file(path: Path, no_go_pat: re.Pattern, wl_pat: re.Pattern):
    # Lies ALLE Spalten, damit Output identisch bleibt
    df = pd.read_excel(path, sheet_name=0)

    check_cols = find_check_cols(df)

    # Kombinierter Text nur aus den Check-Spalten
    text = (
        df[check_cols]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )

    # Schnell: vektorisierte Matches (viel schneller als apply)
    mask_no_go = text.str.contains(no_go_pat, na=False)
    mask_wl = text.str.contains(wl_pat, na=False)

    # DELETE: No-Go getroffen UND NICHT Whitelist
    mask_delete = mask_no_go & (~mask_wl)

    # Debug-Spalten (optional – kosten extra, aber hilfreich)
    df["_MATCH_WORD"] = ""
    df["_WHITELIST_HIT"] = ""

    # Nur dort Match-Wörter berechnen, wo es relevant ist (spart Zeit)
    idx_no_go = df.index[mask_no_go]
    idx_wl = df.index[mask_wl]

    df.loc[idx_no_go, "_MATCH_WORD"] = text.loc[idx_no_go].apply(lambda s: first_match(no_go_pat, s))
    df.loc[idx_wl, "_WHITELIST_HIT"] = text.loc[idx_wl].apply(lambda s: first_match(wl_pat, s))

    cleaned = df.loc[~mask_delete].copy()
    deleted = df.loc[mask_delete].copy()

    # Wenn du Debug nur im Deleted willst: aus CLEANED entfernen
    if DEBUG_IN_DELETED_ONLY:
        cleaned = cleaned.drop(columns=["_MATCH_WORD", "_WHITELIST_HIT"], errors="ignore")

    out_path = path.with_name(path.stem + "_CLEANED.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        cleaned.to_excel(writer, index=False, sheet_name="CLEANED")
        deleted.to_excel(writer, index=False, sheet_name="DELETED")

    return out_path, int(mask_delete.sum()), check_cols


# ---------------- GUI ----------------
selected_files: list[Path] = []

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

    # No-Go min 3 Zeichen, Whitelist min 2 (UG/KG)
    no_go_pat = compile_pattern(HARD_NO_GO, min_len=3)
    wl_pat = compile_pattern(WHITELIST, min_len=2)

    ok = 0
    for f in selected_files:
        try:
            log(f"Starte: {f.name}")
            out, deleted_count, cols = clean_file(f, no_go_pat, wl_pat)
            log(f"✓ Fertig: {f.name} -> {out.name} | gelöscht: {deleted_count} | geprüft: {cols}")
            ok += 1
        except Exception as e:
            log(f"✗ FEHLER {f.name}: {e}")

    messagebox.showinfo(
        "Fertig",
        f"Fertig. Erfolgreich: {ok}/{len(selected_files)}\n\n"
        "Output liegt im gleichen Ordner wie die Datei(en)."
    )

root = tk.Tk()
root.title("Lead Cleaner")
root.geometry("720x420")

frame = tk.Frame(root)
frame.pack(padx=12, pady=10, fill="x")

btn_pick = tk.Button(frame, text="Excel-Dateien auswählen", command=pick_files, width=28)
btn_pick.pack(side="left", padx=(0, 10))

btn_run = tk.Button(frame, text="Bereinigen", command=run_clean, width=18)
btn_run.pack(side="left")

lbl_files = tk.Label(frame, text="0 Datei(en) ausgewählt")
lbl_files.pack(side="left", padx=12)

txt = tk.Text(root, height=18)
txt.pack(padx=12, pady=10, fill="both", expand=True)
log("Output: *_CLEANED.xlsx (Sheets: CLEANED, DELETED)")

root.mainloop()
