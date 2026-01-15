import re
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# NO-GO KEYWORDS (Teilwörter)
# =========================
HARD_NO_GO = [
    # Staat / Verwaltung
    "polizei", "bundespolizei", "polizeipräsidium", "zoll", "finanzamt", "steueramt", "gericht", "staatsan",
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

    # Sonstige
    "friedhof", "denkmal", "archiv", "museum", "theater", "bibli",
]

# =========================
# WHITELIST (wenn Treffer -> NIE löschen)
# =========================
def compile_no_go_pattern(words: list[str], min_len: int = 3) -> re.Pattern:
    cleaned = [re.escape(w.strip().lower()) for w in words if w and len(w.strip()) >= min_len]
    if not cleaned:
        return re.compile(r"(?!x)x")
    return re.compile("(" + "|".join(cleaned) + ")", re.IGNORECASE)

def compile_whitelist_pattern() -> re.Pattern:
    # Matcht Rechtsformen nur als eigenständige Tokens (nicht mitten im Wort!)
    # Beispiele: "GmbH", "UG", "KG", "GbR", "OHG", "AG", "SE", "KGaA", "e.K."
    return re.compile(
        r"(?<!\w)("
        r"gmbh(\s*&\s*co(\.\s*kg)?)?|"
        r"ug|kg|gbr|ohg|ag|se|kgaa|"
        r"e\.?\s*k\.?"
        r")(?!\w)",
        re.IGNORECASE
    )
def build_whitelist_text(df: pd.DataFrame) -> pd.Series:
    # Nur Felder, wo Rechtsformen realistisch sind
    prefer = []
    for c in df.columns:
        lc = str(c).lower()
        if any(k in lc for k in ["zusatz", "firma", "unternehmen", "company", "name", "nachname"]):
            prefer.append(c)

    if not prefer:
        # Fallback: wenn nichts passt, nimm trotzdem alle Textspalten
        prefer = df.select_dtypes(include=["object"]).columns.tolist()

    return (
        df[prefer]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )


# Debug-Spalten in CLEANED entfernen (sinnvoll für Endnutzer)
DEBUG_IN_DELETED_ONLY = True


def compile_pattern(words: list[str], min_len: int) -> re.Pattern:
    cleaned = []
    for w in words:
        w = (w or "").strip().lower()
        if len(w) >= min_len:
            cleaned.append(re.escape(w))
    if not cleaned:
        return re.compile(r"(?!x)x")  # matcht nie
    return re.compile("(" + "|".join(cleaned) + ")", re.IGNORECASE)


def first_match(pattern: re.Pattern, s: str) -> str:
    m = pattern.search(s)
    return m.group(0).lower() if m else ""


def build_search_text_all_text_columns(df: pd.DataFrame) -> pd.Series:
    # Prüft ALLE Text-Spalten -> robust gegen "polizei" in irgendeiner Spalte
    text_cols = df.select_dtypes(include=["object"]).columns.tolist()
    if not text_cols:
        # Fallback: alles zu string, falls excel komische Typen hat
        text_cols = list(df.columns)

    return (
        df[text_cols]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )


def clean_file(path: Path, no_go_pat: re.Pattern, wl_pat: re.Pattern):
    df = pd.read_excel(path, sheet_name=0)

    text = build_search_text_all_text_columns(df)

    # Schnell: vektorisierte Contains
text_all = build_search_text_all_text_columns(df)   # No-Go sucht überall
text_wl  = build_whitelist_text(df)                 # Whitelist nur in "Firma/Zusatz/Name"

mask_no_go = text_all.str.contains(no_go_pat, na=False)
mask_wl    = text_wl.str.contains(wl_pat, na=False)

mask_delete = mask_no_go & (~mask_wl)


    # Debug-Spalten
df.loc[idx_no_go, "_MATCH_WORD"] = text_all.loc[idx_no_go].apply(lambda s: first_match(no_go_pat, s))
df.loc[idx_wl, "_WHITELIST_HIT"] = text_wl.loc[idx_wl].apply(lambda s: first_match(wl_pat, s))


    # Treffer-Wörter nur wo nötig berechnen
    idx_no_go = df.index[mask_no_go]
    idx_wl = df.index[mask_wl]

    df.loc[idx_no_go, "_MATCH_WORD"] = text.loc[idx_no_go].apply(lambda s: first_match(no_go_pat, s))
    df.loc[idx_wl, "_WHITELIST_HIT"] = text.loc[idx_wl].apply(lambda s: first_match(wl_pat, s))

    cleaned = df.loc[~mask_delete].copy()
    deleted = df.loc[mask_delete].copy()

    if DEBUG_IN_DELETED_ONLY:
        cleaned = cleaned.drop(columns=["_MATCH_WORD", "_WHITELIST_HIT"], errors="ignore")

    out_path = path.with_name(path.stem + "_CLEANED.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        cleaned.to_excel(writer, index=False, sheet_name="CLEANED")
        deleted.to_excel(writer, index=False, sheet_name="DELETED")

    return out_path, int(mask_delete.sum())


# ---------------- GUI ----------------
selected_files: list[Path] = []


def log(msg: str):
    txt.insert("end", msg + "\n")
    txt.see("end")
    root.update_idletasks()


def pick_files():
    global selected_files
    files = filedialog.askopenfilenames(
        title="Excel-Dateien auswählen (mehrere möglich)",
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

no_go_pat = compile_no_go_pattern(HARD_NO_GO, min_len=3)
wl_pat = compile_whitelist_pattern()


    ok = 0
    for f in selected_files:
        try:
            log(f"Starte: {f.name}")
            out, deleted_count = clean_file(f, no_go_pat, wl_pat)
            log(f"✓ Fertig: {f.name} -> {out.name} | gelöscht: {deleted_count}")
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
