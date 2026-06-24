import re
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================
# NO-GO (Teilwörter)
# =========================
HARD_NO_GO = [
     # ==================================================
    # POLIZEI / SICHERHEIT / RETTUNG
    # ==================================================
    "polizei", "bundespolizei", "landespolizei", "kripo", "kriminalpolizei",
    "schutzpolizei", "verkehrspolizei", "wasserschutzpolizei",
    "zoll", "hauptzollamt", "zollamt", "finanzkontrolle schwarzarbeit",
    "feuerwehr", "berufsfeuerwehr", "freiwillige feuerwehr",
    "rettungsdienst", "notarzt", "leitstelle", "rettungsleitstelle",
    "katastrophenschutz", "zivilschutz", "technisches hilfswerk", "thw",
    "ordnungsdienst", "kommunaler ordnungsdienst", "stadtwache",

    # ==================================================
    # REGIERUNG / STAAT / MINISTERIEN
    # ==================================================
    "bundesregierung", "landesregierung", "regierung",
    "ministerium", "bundesministerium", "landesministerium",
    "innenministerium", "finanzministerium", "justizministerium",
    "gesundheitsministerium", "arbeitsministerium", "wirtschaftsministerium",
    "verkehrsministerium", "bildungsministerium", "kultusministerium",
    "umweltministerium", "familienministerium", "sozialministerium",
    "staatskanzlei", "kanzleramt", "bundeskanzleramt",
    "senatsverwaltung", "bezirksregierung", "regierungspraesidium", "regierungspräsidium",

    # ==================================================
    # BEHÖRDEN ALLGEMEIN
    # ==================================================
    "behörde", "behoerde", "bundesbehörde", "bundesbehoerde",
    "landesbehörde", "landesbehoerde", "verwaltungsbehörde", "verwaltungsbehoerde",
    "amt für", "amt fuer", "bundesamt", "landesamt", "kreisamt",
    "dienststelle", "verwaltung", "verwaltungsstelle",

    # ==================================================
    # STADT / KOMMUNE / RATHAUS
    # ==================================================
    "rathaus", "stadtverwaltung", "gemeindeverwaltung", "kreisverwaltung",
    "landratsamt", "bezirksamt", "bürgeramt", "buergeramt",
    "bürgerbüro", "buergerbuero", "servicecenter stadt",
    "einwohnermeldeamt", "meldeamt", "meldestelle",
    "standesamt", "ordnungsamt", "gewerbeamt",
    "fundbüro", "fundbuero", "wahlamt", "statistikamt",
    "bürgerdienste", "buergerdienste", "kommunalverwaltung",
    "stadt", "gemeinde", "landkreis", "kreis",

    # ==================================================
    # FINANZ / STEUERN / ZOLL
    # ==================================================
    "finanzamt", "steueramt", "bundeszentralamt für steuern",
    "bundeszentralamt fuer steuern", "steuerverwaltung",
    "oberfinanzdirektion", "finanzverwaltung", "hauptzollamt", "zollamt",

    # ==================================================
    # ARBEIT / SOZIALES / FAMILIE
    # ==================================================
    "jobcenter", "arbeitsagentur", "agentur für arbeit", "agentur fuer arbeit",
    "sozialamt", "jugendamt", "familienkasse", "elterngeldstelle",
    "wohngeldstelle", "versorgungsamt", "integrationsamt",
    "ausländerbehörde", "auslaenderbehoerde", "migrationsamt",
    "flüchtlingshilfe", "fluechtlingshilfe", "asyl", "sozialdienst",
    "seniorenbüro", "seniorenbuero", "pflegestützpunkt", "pflegestuetzpunkt",

    # ==================================================
    # GESUNDHEIT / ÖFFENTLICHE MEDIZIN
    # ==================================================
    "gesundheitsamt", "veterinäramt", "veterinaeramt",
    "lebensmittelüberwachung", "lebensmittelueberwachung",
    "krankenhaus", "klinikum", "klinik", "universitätsklinikum", "universitaetsklinikum",
    "arzt", "ärztin", "aerztin", "zahnarzt", "zahnärztin", "zahnaerztin",
    "apotheke", "notfallpraxis", "reha", "rehazentrum",
    "sanatorium",
    "krankenkasse", "aok", "barmer", "tk", "techniker krankenkasse",
    "dak", "ikk", "knappschaft", "bkk",

    # ==================================================
    # SCHULEN / HOCHSCHULEN / BILDUNG
    # ==================================================
    "schule", "grundschule", "hauptschule", "realschule",
    "gesamtschule", "gymnasium", "berufsschule", "förderschule", "foerderschule",
    "sekundarschule", "privatschule", "internat",
    "hochschule", "universität", "universitaet", "fachhochschule",
    "uni ", "fh ", "campus", "studentenwerk", "studierendenwerk",
    "akademie", "bildungszentrum", "weiterbildungszentrum",
    "volkshochschule", "vhs",
    "familienzentrum", "schulamt", "bildungswerk",

    # ==================================================
    # JUSTIZ / GERICHTE / RECHT
    # ==================================================
    "gericht", "amtsgericht", "landgericht", "oberlandesgericht",
    "verwaltungsgericht", "arbeitsgericht", "sozialgericht",
    "finanzgericht", "verfassungsgericht", "bundesgerichtshof",
    "bundesverfassungsgericht", "staatsanwaltschaft", "generalstaatsanwaltschaft",
    "justizvollzugsanstalt", "jva", "vollzugsanstalt",
    "rechtsanwalt", "rechtsanwältin", "rechtsanwaeltin",
    "anwalt", "anwältin", "anwaeltin", "kanzlei",
    "notar", "notarin", "notariat", "gerichtsvollzieher",

    # ==================================================
    # BAU / VERKEHR / UMWELT / ÄMTER
    # ==================================================
    "bauamt", "bauaufsichtsamt", "bauordnungsamt",
    "stadtplanungsamt", "planungsamt", "vermessungsamt",
    "katasteramt", "liegenschaftsamt", "immobilienamt",
    "umweltamt", "naturschutzbehörde", "naturschutzbehoerde",
    "straßenverkehrsamt", "strassenverkehrsamt",
    "verkehrsamt", "kfz-zulassungsstelle", "zulassungsstelle",
    "führerscheinstelle", "fuehrerscheinstelle",
    "tiefbauamt", "hochbauamt", "grünflächenamt", "gruenflaechenamt",
    "gartenamt", "forstamt", "wasserwirtschaftsamt",

    # ==================================================
    # ÖFFENTLICHE BETRIEBE / EINRICHTUNGEN
    # ==================================================
    "stadtwerke", "gemeindewerke", "wasserwerk", "wasserwerke",
    "abwasserbetrieb", "abwasserverband", "entsorgungsbetrieb",
    "abfallwirtschaft", "müllabfuhr", "muellabfuhr", "recyclinghof",
    "wertstoffhof", "friedhof", "friedhofsamt", "bestattung",
    "stadtarchiv", "archiv", "museum", "bibliothek", "bücherei", "buecherei",
    "theater", "oper", "konzerthaus", "philharmonie",
    "volksbank-stadion", "stadthalle", "bürgerhaus", "buergerhaus",

    # ==================================================
    # KIRCHEN / RELIGION
    # ==================================================
    "kirche", "pfarr", "pfarramt", "kirchengemeinde",
    "bistum", "diözese", "dioezese", "evangelisch", "katholisch",
    "moschee", "islam", "islamisch", "synagoge", "jüdisch", "juedisch",
    "tempel", "religionsgemeinschaft",

    # ==================================================
    # VEREINE / GEMEINNÜTZIG / HILFSORGANISATIONEN
    # ==================================================
    "verein", "e.v.", "ev.", "gemeinnützig", "gemeinnuetzig",
    "stiftung", "förderverein", "foerderverein",
    "caritas", "diakonie", "drk", "deutsches rotes kreuz",
    "rotes kreuz", "malteser", "johanniter", "awo", "arbeiterwohlfahrt",
    "lebenshilfe", "tafel", "kinderschutzbund", "sozialverband",
    "vdk", "paritätischer", "paritaetischer", "nabu", "bUND", "greenpeace",
    "tierschutzverein", "sportverein", "turnverein", "schützenverein", "schuetzenverein",

    # ==================================================
    # MILITÄR / STAATLICHE SICHERHEIT
    # ==================================================
    "bundeswehr", "wehrverwaltung", "verteidigungsministerium",
    "kaserne", "marine", "heer", "luftwaffe",
    "bundesnachrichtendienst", "bnd", "verfassungsschutz",

    # ==================================================
    # KAMMERN / VERBÄNDE / INSTITUTIONEN
    # ==================================================
    "ihk", "industrie- und handelskammer", "handelskammer",
    "handwerkskammer", "hwk", "ärztekammer", "aerztekammer",
    "apothekerkammer", "rechtsanwaltskammer", "notarkammer",
    "steuerberaterkammer", "architektenkammer",
    "verband", "bundesverband", "landesverband",
    "innung", "kreishandwerkerschaft",

    # ==================================================
    # VERSICHERUNG / RENTE / BG
    # ==================================================
    "rentenversicherung", "deutsche rentenversicherung",
    "unfallkasse", "berufsgenossenschaft", "bg ",
    "versicherung", "versicherungsbüro", "versicherungsbuero",
    "allianz", "axa", "ergo", "huk", "huk-coburg",
    "devk", "signal iduna", "generali", "zurich", "nürnberger", "nuernberger",
    "debeka", "continentale", "provinzial", "lvm",

    # ==================================================
    # BANKEN / FINANZEN
    # ==================================================
    "bank", "sparkasse", "volksbank", "raiffeisenbank",
    "sparda-bank", "postbank", "deutsche bank", "commerzbank",
    "targobank", "santander", "ing-diba", "ing ", "dkb",
    "n26", "finanzberatung", "vermögensberatung", "vermoegensberatung",
    "dvag", "mlp", "tecis", "ovb",

    # ==================================================
    # GROSSE FAST-FOOD-KETTEN
    # ==================================================
    "mcdonalds", "mc donalds", "mcdonald's", "mc donald's",
    "burger king", "kfc", "kentucky fried chicken",
    "subway", "dominos", "domino's", "pizza hut",
    "nordsee", "vapiano", "dean & david", "dean and david",
    "backwerk", "le crobag", "ditsch", "starbucks",
    "tchibo", "yormas", "immergrün", "immergruen",

    # ==================================================
    # SUPERMÄRKTE / DISCOUNTER / DROGERIEN
    # ==================================================
    "aldi", "aldi nord", "aldi süd", "aldi sued",
    "lidl", "rewe", "edeka", "kaufland", "netto", "netto marken-discount",
    "penny", "tegut", "globus", "real", "metro",
    "dm-drogerie", "dm drogerie", "dm markt", "rossmann",
    "müller drogerie", "mueller drogerie", "budni", "budnikowsky",
    "denns", "denn's", "alnatura",

    # ==================================================
    # GROSSE EINZELHANDELSKETTEN
    # ==================================================
    "ikea", "poco", "mömax", "moemax", "xxxlutz", "roller",
    "hornbach", "obi", "bauhaus", "toom", "hagebaumarkt",
    "mediamarkt", "media markt", "saturn", "expert",
    "conrad", "cyberport", "gravis",
    "decathlon", "intersport", "sportcheck",
    "h&m", "hm ", "zara", "primark", "c&a", "ca ",
    "peek & cloppenburg", "p&c", "reserved", "new yorker",
    "tk maxx", "deichmann", "snipes", "foot locker",
    "douglas", "flaconi",

    # ==================================================
    # TANKSTELLEN / AUTO-KETTEN
    # ==================================================
    "aral", "shell", "esso", "jet tankstelle", "totalenergies",
    "avia", "hem tankstelle", "star tankstelle", "bft tankstelle",
    "autohaus", "autohändler", "autohaendler",
    "bmw", "mercedes", "mercedes-benz", "audi", "volkswagen", "vw ",
    "opel", "ford", "toyota", "hyundai", "kia", "skoda", "seat",
    "nissan", "renault", "peugeot", "citroen", "fiat",
    "atu", "pitstop", "vergoelst", "euromaster",
    "tüv", "tuev", "dekra", "gtü", "gtue", "küs", "kues",

    # ==================================================
    # LOGISTIK / POST / PAKETDIENSTE
    # ==================================================
    "deutsche post", "dhl", "dpd", "gls", "hermes", "ups",
    "fedex", "trans-o-flex", "postfiliale", "briefzentrum",



    # ==================================================
    # CALLCENTER / AGENTUREN / B2B, falls keine Zielgruppe
    # ==================================================
    "callcenter", "call center", "telemarketing",

    # ==================================================
    # ONLINE / TECH-GROSSKONZERNE / PLATTFORMEN
    # ==================================================
    "amazon", "amazon logistics", "google", "meta", "facebook",
    "instagram", "apple", "microsoft", "sap", "telekom",
    "vodafone", "o2 shop", "telefonica", "1&1", "1und1",

    # ==================================================
    # ÖPNV / BAHN / FLUGHAFEN
    # ==================================================
    "deutsche bahn", "db reisezentrum", "bahnhof",
    "verkehrsbetriebe", "stadtbahn", "straßenbahn", "strassenbahn",
    "busbetrieb", "busbahnhof", "flughafen", "airport",
    "hafen", "taxi", "taxizentrale",

    # ==================================================
    # ZEITARBEIT / PERSONALDIENSTLEISTER
    # ==================================================
    "zeitarbeit", "personalvermittlung", "personaldienstleistung",
    "randstad", "adecco", "manpower", "persona service",
    "tempton", "runtime", "timepartner", "iperdi",

    # ==================================================
    # WETTEN / CASINO / SPIELHALLE
    # ==================================================
    "spielhalle", "casino", "wettbüro", "wettbuero",
    "tipico", "bet365", "merkur casino", "löwen play", "loewen play",

    # ==================================================
    # SONSTIGE NO-GO BEGRIFFE
    # ==================================================
    "beamt", "beamter", "beamtin", "öffentlich", "oeffentlich",
    "öffentliche einrichtung", "oeffentliche einrichtung",
    "kommunal", "staatlich", "bundesweit",
    "gmbh & co kg", "ag ", "konzern", "filiale", "zentrale",


    # ==================================================
# AMTLICHE / ÖFFENTLICHE PERSONEN / BEAMTE
# ==================================================

"beamter", "beamtin", "beamte", "beamtinnen",
"beamt", "verbeamtet", "staatsbediensteter", "staatsbedienstete",
"angestellter im öffentlichen dienst", "angestellte im öffentlichen dienst",
"angestellter im oeffentlichen dienst", "angestellte im oeffentlichen dienst",
"öffentlicher dienst", "oeffentlicher dienst",
"public service", "civil servant",

# Sachbearbeitung / Verwaltung
"sachbearbeiter", "sachbearbeiterin", "sachbearbeitung",
"verwaltungsmitarbeiter", "verwaltungsmitarbeiterin",
"verwaltungsangestellter", "verwaltungsangestellte",
"verwaltungsfachangestellter", "verwaltungsfachangestellte",
"verwaltungsbeamter", "verwaltungsbeamtin",
"amtsmitarbeiter", "amtsmitarbeiterin",
"behördenmitarbeiter", "behoerdenmitarbeiter",
"behördenmitarbeiterin", "behoerdenmitarbeiterin",

# Leitung im Amt / Behörde
"amtsleiter", "amtsleiterin",
"fachbereichsleiter", "fachbereichsleiterin",
"bereichsleiter amt", "bereichsleiterin amt",
"dezernent", "dezernentin",
"dezernatsleiter", "dezernatsleiterin",
"referatsleiter", "referatsleiterin",
"abteilungsleiter behörde", "abteilungsleiter behoerde",
"abteilungsleiterin behörde", "abteilungsleiterin behoerde",
"dienststellenleiter", "dienststellenleiterin",

# Stadt / Kommune / Politiknah
"bürgermeister", "buergermeister",
"bürgermeisterin", "buergermeisterin",
"oberbürgermeister", "oberbuergermeister",
"oberbürgermeisterin", "oberbuergermeisterin",
"landrat", "landrätin", "landraetin",
"stadtrat", "stadträtin", "stadtraetin",
"gemeinderat", "gemeinderätin", "gemeinderaetin",
"kreisrat", "kreisrätin", "kreisraetin",
"bezirksbürgermeister", "bezirksbuergermeister",
"bezirksbürgermeisterin", "bezirksbuergermeisterin",
"ratsherr", "ratsfrau",
"ratsmitglied", "fraktionsvorsitzender", "fraktionsvorsitzende",

# Meldeamt / Bürgeramt / Ordnungsamt Rollen
"standesbeamter", "standesbeamtin",
"meldebeamter", "meldebeamtin",
"ordnungsbeamter", "ordnungsbeamtin",
"vollzugsbeamter", "vollzugsbeamtin",
"kommunaler vollzugsdienst",
"ordnungsdienstmitarbeiter", "ordnungsdienstmitarbeiterin",
"stadtwache mitarbeiter", "stadtwache mitarbeiterin",

# Polizei / Zoll / Sicherheit Personen
"polizist", "polizistin",
"polizeibeamter", "polizeibeamtin",
"kriminalbeamter", "kriminalbeamtin",
"zollbeamter", "zollbeamtin",
"zollinspektor", "zollinspektorin",
"hauptkommissar", "hauptkommissarin",
"kommissar", "kommissarin",
"wachtmeister", "wachtmeisterin",

# Feuerwehr / Rettung / Katastrophenschutz Personen
"feuerwehrmann", "feuerwehrfrau",
"brandmeister", "brandmeisterin",
"brandinspektor", "brandinspektorin",
"rettungsbeamter", "rettungsbeamtin",
"notfallsanitäter", "notfallsanitaeter",
"notfallsanitäterin", "notfallsanitaeterin",
"leitstellendisponent", "leitstellendisponentin",

# Justiz Personen
"richter", "richterin",
"staatsanwalt", "staatsanwältin", "staatsanwaeltin",
"oberstaatsanwalt", "oberstaatsanwältin", "oberstaatsanwaeltin",
"rechtspfleger", "rechtspflegerin",
"justizbeamter", "justizbeamtin",
"justizvollzugsbeamter", "justizvollzugsbeamtin",
"gerichtsvollzieher", "gerichtsvollzieherin",

# Schule / Hochschule öffentlich beschäftigte Personen
"schulleiter", "schulleiterin",
"rektor", "rektorin",
"konrektor", "konrektorin",
"lehrer", "lehrerin",
"studienrat", "studienrätin", "studienraetin",
"oberstudienrat", "oberstudienrätin", "oberstudienraetin",
"dozent", "dozentin",
"wissenschaftlicher mitarbeiter", "wissenschaftliche mitarbeiterin",

# Gesundheitsamt / öffentliche Medizin Rollen
"amtsarzt", "amtsärztin", "amtsaerztin",
"gesundheitsbeamter", "gesundheitsbeamtin",
"hygienekontrolleur", "hygienekontrolleurin",
"lebensmittelkontrolleur", "lebensmittelkontrolleurin",
"veterinärbeamter", "veterinaerbeamter",
"veterinärbeamtin", "veterinaerbeamtin",

# Arbeitsagentur / Jobcenter / Sozialamt Rollen
"fallmanager", "fallmanagerin",
"arbeitsvermittler", "arbeitsvermittlerin",
"leistungsberater", "leistungsberaterin",
"sozialarbeiter amt", "sozialarbeiterin amt",
"sozialpädagoge amt", "sozialpaedagoge amt",
"sozialpädagogin amt", "sozialpaedagogin amt",
"jugendamtsmitarbeiter", "jugendamtsmitarbeiterin",

# Ausländerbehörde / Migration
"sachbearbeiter ausländerbehörde", "sachbearbeiter auslaenderbehoerde",
"sachbearbeiterin ausländerbehörde", "sachbearbeiterin auslaenderbehoerde",
"integrationsbeauftragter", "integrationsbeauftragte",
"migrationsbeauftragter", "migrationsbeauftragte",

# Finanzamt / Steuer / Prüfer
"finanzbeamter", "finanzbeamtin",
"steuerbeamter", "steuerbeamtin",
"betriebsprüfer", "betriebspruefer",
"betriebsprüferin", "betriebsprueferin",
"steuerprüfer", "steuerpruefer",
"steuerprüferin", "steuerprueferin",

# Bauamt / Umweltamt / Verkehr
"bauprüfer", "baupruefer",
"bauprüferin", "bauprueferin",
"baukontrolleur", "baukontrolleurin",
"verkehrsüberwacher", "verkehrsueberwacher",
"verkehrsüberwacherin", "verkehrsueberwacherin",
"umweltbeauftragter", "umweltbeauftragte",
"naturschutzbeauftragter", "naturschutzbeauftragte"
    
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

