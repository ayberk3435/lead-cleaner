import re
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# NO-GO (Teilwörter)
# =========================
HARD_NO_GO = [
    "polizei", "bundespolizei", "zoll", "finanzamt", "steueramt", "gericht", "staatsan",
    "ministerium", "regierung", "behörde", "bürgeramt", "einwohn", "ordnungsamt",
    "sozialamt", "jobcenter", "arbeitsagentur", "agentur für arbeit",

    "kranken", "klinik", "klinikum", "arzt", "ärztin", "zahnarzt", "apothek",
    "gesundheitsamt", "rehazentrum",

    "schule", "grundschule", "realschule", "hauptschule", "gymnasium", "berufsschule",
    "hochschule", "universität", "fachhochschule", "akademie", "bildungszentrum",
    "vhs", "forsch", "institut",

    "feuerwehr", "rettung", "notarzt", "katastroph", "zivilschutz",

    "kirche", "pfarr", "gemeinde", "bistum", "diözese", "mosche", "islam", "synagog",

    "verein", "stiftung", "gemeinn", "e.v.", "caritas", "diakonie", "drk",
    "rotes kreuz", "malteser", "johanniter",

    "rechtsanwalt", "anwalt", "kanzlei", "notar", "notariat",

    "friedhof", "denkmal", "archiv", "museum", "theater", "bibli",
]

DEBUG_IN_DELETED_ONLY = True


def compile_no_go_pattern(words: list[str], min_len: int = 3) -> re.Pattern:
    cleaned = []
    for w in words:
        w = (w or "").strip().lower()
        if len(w) >= min_len:
            cleaned.append(re.escape(w))
    if not cleaned:
        return re.compile(r"(?!x)x")
    return re.compile("(" + "|".join(cleaned) + ")", re.IGNORECASE)


def compile_whitelist_pattern() -> re.Pattern:
    # Rechtsformen nur als eigenständige Tokens, nicht mitten im Wort
    return re.compile(
        r"(?<!\w)("
        r"gmbh(\s*&\s*co(\.\s*kg)?)?|"
        r"ug|kg|gbr|ohg|ag|se|kgaa|"
        r"e\.?\s*k\.?"
        r")(?!\w)",
        re.IGNORECASE
    )


def first_match(pattern: re.Pattern, s: str) -> str:
    m = pattern.search(s)
    return m.group(0).lower() if m else ""


def build_no_go_text(df: pd.DataFrame) -> pd.Series:
    # Spalten, in denen No-Go wirklich Sinn macht (Firma/Einrichtung/Zusatz/Bezeichnung)
    prefer = []
    for c in df.columns:
        lc = str(c).lower()
        if any(k in lc for k in [
            "zusatz", "firma", "unternehmen", "company", "betrieb", "bezeichnung",
            "einrichtung", "organisation", "org", "branche", "art", "notiz",
            "name"  # <- nur wenn du "NAME" als Firmenname-Spalte nutzt
        ]):
            prefer.append(c)

    # Bewusst ausschließen: Ort/Straße/PLZ/Tel/Mail usw.
    blocked = []
    for c in prefer:
        lc = str(c).lower()
        if any(k in lc for k in ["ort", "straße", "strasse", "plz", "telefon", "tel", "mail", "e-mail", "email"]):
            blocked.append(c)
    prefer = [c for c in prefer if c not in blocked]

    # Fallback, falls nichts gefunden
    if not prefer:
        # nimm nur typische Firmenfelder, die fast immer existieren:
        for c in df.columns:
            if str(c).lower() in ["zusatz", "firma", "unternehmen", "name"]:
                prefer.append(c)
        if not prefer:
            # notfalls gar nichts -> dann wird nichts gelöscht
            return pd.Series([""] * len(df), index=df.index)

    return (
        df[prefer]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )



def build_whitelist_text(df: pd.DataFrame) -> pd.Series:
    prefer = []
    for c in df.columns:
        lc = str(c).lower()
        if any(k in lc for k in ["zusatz", "firma", "unternehmen", "company", "name", "nachname"]):
            prefer.append(c)

    if not prefer:
        prefer = df.select_dtypes(include=["object"]).columns.tolist()
        if not prefer:
            prefer = list(df.columns)

    return (
        df[prefer]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .str.lower()
    )


def clean_file(path: Path, no_go_pat: re.Pattern, wl_pat: re.Pattern):
    df = pd.read_excel(path, sheet_name=0)

    text_no_go = build_no_go_text(df)
    text_wl = build_whitelist_text(df)

    mask_no_go = text_no_go.str.contains(no_go_pat, na=False)
    mask_wl = text_wl.str.contains(wl_pat, na=False)

    mask_delete = mask_no_go & (~mask_wl)

    # Debug-Spalten
    df["_MATCH_WORD"] = ""
    df["_WHITELIST_HIT"] = ""

    idx_no_go = df.index[mask_no_go]
    idx_wl = df.index[mask_wl]

    if len(idx_no_go) > 0:
        df.loc[idx_no_go, "_MATCH_WORD"] = (
            text_no_go.loc[idx_no_go]
            .apply(lambda s: first_match(no_go_pat, s))
    )


    if len(idx_wl) > 0:
        df.loc[idx_wl, "_WHITELIST_HIT"] = (
            text_wl.loc[idx_wl]
            .apply(lambda s: first_match(wl_pat, s))
    )


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

