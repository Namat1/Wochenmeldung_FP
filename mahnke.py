import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

st.info("Die rot gefärbten Disponent-Zeilen müssen manuell eingetragen werden.")

RELEVANT_WORDS = [
    "ausgleich", "krank", "sonderurlaub", "urlaub",
    "berufsschule", "fahrschule", "homeoffice",
    "schulung", "dienstreise", "seminar", "fortbildung",
    "elternzeit", "kur", "reha", "kur und reha", "reha und kur",
]
EXCLUDED_WORDS = ["hoffahrer", "waschteam", "aushilfsfahrer"]

WEEKDAYS = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"]
DAY_COL_PAIRS = [(4, 5), (6, 7), (8, 9), (10, 11), (12, 13), (14, 15), (16, 17)]

# Gewünschte Reihenfolge der Gruppen
GROUP_ORDER = ["Fahrer", "Disponent", "Auszubildende", "Aushilfsfahrer"]
GROUP_LABELS = set(GROUP_ORDER)

# Disponenten (feste, rote Zeilen - manuell auszufüllen)
DISPO_PEOPLE = [
    ("Carstensen", "Martin"), ("Lau", "Eike"), ("Pham Manh", "Chris"),
    ("Ohlenroth", "Nadja"), ("Schulz", "Julian"), ("Aniol", "Przemyslaw"),
    ("Packmohr", "Gina"),
]

# Farbschema je Gruppe: (Balkenfarbe, Zeilenfarbe)
GROUP_COLORS = {
    "Fahrer":         ("2E75B6", "DDEBF7"),  # blau
    "Disponent":      ("C0392B", "F1948A"),  # rot
    "Auszubildende":  ("D68910", "FDEBD0"),  # orange
    "Aushilfsfahrer": ("1E8449", "ABEBC0"),  # grün
}


def _clean(value) -> str:
    text = str(value).strip()
    return "" if text.lower() in {"none", "nan", "nat"} else text


def extract_grouped_data(df: pd.DataFrame) -> pd.DataFrame:
    """Liest ALLE Fahrer (ohne Filter) und ordnet sie einer Kategorie zu."""
    col_b = df.iloc[:, 1].astype(str).str.strip().str.lower()

    if "adler" not in col_b.values or "steckel" not in col_b.values:
        st.error("'adler' oder 'steckel' wurden in Spalte B nicht gefunden.")
        st.stop()

    start_index = col_b[col_b == "adler"].index[0]
    end_index = col_b[col_b == "steckel"].index[0]

    rows = []
    current_category = "Fahrer"

    for i in range(start_index, end_index + 1):
        row_text = " ".join(_clean(v) for v in df.iloc[i].tolist()).lower()

        # Abschnittsmarker erkennen
        if "hoffahrer" in row_text and "waschteam" in row_text:
            current_category = "Fahrer"          # Hoffahrer/Waschteam unter Fahrer
            continue
        if "ausbildung" in row_text and "berufskraftfahrer" in row_text:
            current_category = "Auszubildende"
            continue
        if "aushilfsfahrer" in row_text:
            current_category = "Aushilfsfahrer"
            continue

        lastname = _clean(df.iloc[i, 1])
        firstname = _clean(df.iloc[i, 2])
        if not lastname or not firstname:
            continue  # Tour-Zeilen / Platzhalter ohne Namen

        activities_row = i + 1
        row = {"Kategorie": current_category,
               "Nachname": lastname.title(), "Vorname": firstname.title()}
        for day, (c1, c2) in zip(WEEKDAYS, DAY_COL_PAIRS):
            a1 = _clean(df.iloc[activities_row, c1])
            a2 = _clean(df.iloc[activities_row, c2])
            activity = " ".join(x for x in (a1, a2) if x and x != "0").strip().lower()
            if any(w in activity for w in RELEVANT_WORDS) and not any(e in activity for e in EXCLUDED_WORDS):
                row[day] = activity.title()
            else:
                row[day] = ""
        rows.append(row)

    return pd.DataFrame(rows)


def build_output(df: pd.DataFrame) -> pd.DataFrame:
    """Gruppiert in der gewünschten Reihenfolge und fügt Balkenzeilen ein."""
    def empty_row(kat, last="", first=""):
        r = {"Nachname": last, "Vorname": first}
        for d in WEEKDAYS:
            r[d] = ""
        return r

    groups = {
        "Fahrer": df[df["Kategorie"] == "Fahrer"].drop(columns="Kategorie").to_dict("records"),
        "Disponent": [empty_row("Disponent", l, f) for l, f in DISPO_PEOPLE],
        "Auszubildende": df[df["Kategorie"] == "Auszubildende"].drop(columns="Kategorie").to_dict("records"),
        "Aushilfsfahrer": df[df["Kategorie"] == "Aushilfsfahrer"].drop(columns="Kategorie").to_dict("records"),
    }

    out_rows = []
    for group in GROUP_ORDER:
        out_rows.append(empty_row(group, last=group))  # Balkenzeile
        out_rows.extend(groups[group])

    return pd.DataFrame(out_rows, columns=["Nachname", "Vorname"] + WEEKDAYS)


def create_header_with_dates(df):
    return [pd.to_datetime(df.iloc[1, c]).strftime("%d.%m.%Y") for c in (4, 6, 8, 10, 12, 14, 16)]


def style_excel(ws, calendar_week):
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="4472C4")
    thin = Border(*[Side(style="thin", color="CCCCCC")] * 4)
    medium = Border(*[Side(style="medium", color="1F4E78")] * 4)
    last_col = ws.max_column

    ws["A1"].value = f"Kalenderwoche: {calendar_week + 1}"
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = title_fill
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.row_dimensions[1].height = 30

    ws["A2"].value = "Abteilung: Fuhrpark - NFC"
    ws["A2"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].fill = title_fill
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws.row_dimensions[2].height = 26

    for cell in ws[3]:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = Font(bold=True, size=11, color="FFFFFF")
        cell.fill = header_fill
        cell.border = medium
    ws.row_dimensions[3].height = 24

    current_row_fill = None
    band_index = 0
    for r in range(4, ws.max_row + 1):
        a_val = str(ws.cell(r, 1).value or "").strip()
        b_val = str(ws.cell(r, 2).value or "").strip()

        # Balkenzeile einer Gruppe?
        if a_val in GROUP_LABELS and b_val == "":
            bar, row_color = GROUP_COLORS[a_val]
            current_row_fill = row_color
            band_index = 0
            # zuerst Inhalt der Spalten B..letzte leeren, dann mergen
            for c in range(2, last_col + 1):
                ws.cell(r, c).value = None
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell = ws.cell(r, 1)
            cell.value = a_val
            cell.fill = PatternFill("solid", fgColor=bar)
            cell.font = Font(bold=True, size=12, color="FFFFFF")
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            cell.border = medium
            ws.row_dimensions[r].height = 22
            continue

        # Datenzeile in aktueller Gruppenfarbe
        fill = PatternFill("solid", fgColor=current_row_fill) if current_row_fill else None
        for cell in ws[r]:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font = Font(size=10, color="2C3E50", bold=(cell.column <= 2))
            cell.border = thin
            if fill:
                cell.fill = fill
        ws.row_dimensions[r].height = 20
        band_index += 1

    widths = {1: 20, 2: 16}
    for idx, col in enumerate(ws.columns, start=1):
        letter = get_column_letter(col[0].column)
        longest = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[letter].width = min(max(longest + 4, widths.get(idx, 16)), 60)

    ws.freeze_panes = "A4"
    # bewusst KEIN ws.auto_filter -> keine Filter in der Tabelle


# ------------------------------------------------------------ Streamlit App
st.title("Wochenarbeitsbericht Fuhrpark")
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    data = pd.DataFrame(wb["Druck Fahrer"].values)

    grouped = extract_grouped_data(data)
    output_data = build_output(grouped)

    dates = create_header_with_dates(data)
    calendar_week = pd.to_datetime(dates[0], format="%d.%m.%Y").isocalendar()[1]
    output_data.columns = ["Nachname", "Vorname"] + [f"{w} ({d})" for w, d in zip(WEEKDAYS, dates)]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        output_data.to_excel(writer, index=False, sheet_name="Wochenübersicht", startrow=2)
        style_excel(writer.sheets["Wochenübersicht"], calendar_week)

    st.success("Verarbeitung abgeschlossen.")
    st.download_button(
        "Download als Excel",
        data=output.getvalue(),
        file_name=f"Fuhrpark_Meldung_KW_{calendar_week + 1}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
