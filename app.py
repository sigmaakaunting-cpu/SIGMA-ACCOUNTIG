import streamlit as st
import pandas as pd
import re
import ast
from io import BytesIO
from PIL import Image


from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="AOP Finansiski Izvestai", layout="wide")
USERS = {
    "sigma": "12345",
    "client1": "test123"
}

def login():
    st.sidebar.title("🔐 Login")

    username = st.sidebar.text_input("Корисник")
    password = st.sidebar.text_input("Лозинка", type="password")

    if st.sidebar.button("Најави се"):

        if username in USERS and USERS[username] == password:
            st.session_state["logged_in"] = True
            st.session_state["username"] = username
            st.rerun()
        else:
            st.sidebar.error("Погрешно корисничко име или лозинка")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login()
    st.stop()

logo = Image.open("assets/sigma_logo.png")

st.image(logo,width=420)
st.markdown("""
            ### Добредојдовте во SIGMA Accounting Intelligence!
            Контактирајте не за повеќе информации и прилагодени решенија за вашата компанија:
            - 📧 Email:sigmaakaunting@gmail.com
            - 📞 Телефон: 070/229-057   
            - Адреса: ул. 121 3-1 Тетово
            🌏 SIGMA Accounting 
<style>
[data-testid="metric-container"] {
    background-color: #f8f9fa;
    border: 1px solid #e6e6e6;
    padding: 20px;
    border-radius: 14px;
}
[data-testid="metric-container"] label {
    font-size: 22px !important;
    font-weight: 700 !important;
}
[data-testid="metric-container"] p {
    font-size: 42px !important;
    font-weight: 800 !important;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 SIGMA Accounting Intelligence")
st.caption("Автоматизирана финансиска анализа од заклучен лист")

zaklucen_file = st.file_uploader("📤 Прикачи заклучен лист", type=["xls", "xlsx"])
pravila_file = "data/pravila bilansi.xlsx"


def clean_number(x):
    try:
        if pd.isna(x):
            return 0.0
        x = str(x).strip().replace(".", "").replace(",", ".").replace(" ", "")
        return float(x)
    except Exception:
        return 0.0


def format_aop(x):
    try:
        if pd.isna(x):
            return ""
        s = str(x).strip()
        if s.endswith(".0"):
            s = s[:-2]
        if s.isdigit():
            return s.zfill(3)
        return s
    except Exception:
        return str(x).strip()


def normalize_key(x):
    if pd.isna(x):
        return ""
    return str(x).strip().replace(" ", "").replace(".0", "").upper()


def read_zaklucen(file):
    engine = "xlrd" if file.name.lower().endswith(".xls") else "openpyxl"
    raw = pd.read_excel(file, engine=engine, header=None).dropna(how="all")

    start_row = None
    for i in range(len(raw)):
        val = str(raw.iloc[i, 0]).strip().lower()
        if val in ["konto", "конто"] or val.isdigit():
            start_row = i
            break

    if start_row is None:
        st.error("Ne moze da se najde pocetok na zaklucniot list.")
        st.stop()

    first_val = str(raw.iloc[start_row, 0]).strip().lower()
    if first_val in ["konto", "конто"]:
        start_row += 1

    df = raw.iloc[start_row:].copy()

    if df.shape[1] >= 10:
        df = df.iloc[:, :10]
        df.columns = [
            "konto", "naziv",
            "prethodna_dolzi", "prethodna_pobaruva",
            "tekovna_dolzi", "tekovna_pobaruva",
            "vkupno_dolzi", "vkupno_pobaruva",
            "krajno_dolzi", "krajno_pobaruva"
        ]
    else:
        df = df.iloc[:, :9]
        df.columns = [
            "konto", "naziv", "saldo",
            "prethodna_dolzi", "prethodna_pobaruva",
            "tekovna_dolzi", "tekovna_pobaruva",
            "vkupno_dolzi", "vkupno_pobaruva"
        ]
        df["saldo"] = df["saldo"].apply(clean_number)
        df["krajno_dolzi"] = df["saldo"].apply(lambda x: x if x > 0 else 0)
        df["krajno_pobaruva"] = df["saldo"].apply(lambda x: abs(x) if x < 0 else 0)

    df["konto"] = df["konto"].astype(str).str.replace(".0", "", regex=False).str.strip()
    df["naziv"] = df["naziv"].astype(str)

    for col in df.columns:
        if col not in ["konto", "naziv"]:
            df[col] = df[col].apply(clean_number)

    return df[(df["konto"] != "") & (df["konto"] != "nan")]


def find_header_row(file, sheet_name):
    temp = pd.read_excel(file, sheet_name=sheet_name, header=None)
    for i in range(len(temp)):
        row = [str(x).strip().lower() for x in temp.iloc[i].tolist()]
        if "aop" in row:
            return i
    raise ValueError(f"Ne ja najdov kolonata AOP vo sheet: {sheet_name}")


def read_rules(file, sheet_name):
    header_row = find_header_row(file, sheet_name)
    rules = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    rules.columns = [str(c).strip() for c in rules.columns]

    rename = {}
    for col in rules.columns:
        c = str(col).strip().lower()
        if c in ["aop", "ознака на аоп"]:
            rename[col] = "AOP"
        elif c in ["pozicija", "позиција"]:
            rename[col] = "Pozicija"
        elif c == "naziv":
            rename[col] = "Naziv"
        elif c == "opis":
            rename[col] = "Opis"
        elif c == "tip":
            rename[col] = "Tip"
        elif c == "izvor":
            rename[col] = "Izvor"
        elif c == "key":
            rename[col] = "Key"
        elif c == "kategorija":
            rename[col] = "Kategorija"
        elif c == "konto":
            rename[col] = "konto"
        elif c == "konto.1":
            rename[col] = "konto_prethodna"
        elif c == "kolona":
            rename[col] = "kolona"
        elif c == "kolona.1":
            rename[col] = "kolona_prethodna"
        elif c == "operacija":
            rename[col] = "Operacija"
        elif c == "operacija.1":
            rename[col] = "Operacija_prethodna"
        elif c == "formula":
            rename[col] = "Formula"
        elif c == "logika":
            rename[col] = "Logika"

    rules = rules.rename(columns=rename)
    if "AOP" not in rules.columns:
        raise ValueError(f"Vo {sheet_name} ne postoi kolona AOP.")

    rules = rules[rules["AOP"].notna()].copy()
    rules["AOP"] = rules["AOP"].apply(format_aop)
    return rules


def konto_list(konto_text):
    if pd.isna(konto_text):
        return []
    text = str(konto_text).replace("(", "").replace(")", "").replace(" ", "").replace(";", ",")
    result = []
    for part in text.split(","):
        if not part:
            continue
        if "-" in part:
            try:
                a, b = part.split("-")
                width = max(len(a), len(b))
                for n in range(int(a), int(b) + 1):
                    result.append(str(n).zfill(width))
            except Exception:
                pass
        else:
            result.append(part)
    return result


def sum_rule(df, konto_text, kolona, operacija):
    if pd.isna(konto_text) or pd.isna(kolona):
        return 0.0
    kolona = str(kolona).strip()
    if kolona not in df.columns:
        return 0.0
    konta = konto_list(konto_text)
    if not konta:
        return 0.0
    mask = df["konto"].astype(str).str.startswith(tuple(konta), na=False)
    value = df.loc[mask, kolona].sum()
    if str(operacija).strip() == "-":
        return -value
    return value


def safe_eval_expression(expr):
    allowed_nodes = (
        ast.Expression, ast.BinOp, ast.UnaryOp,
        ast.Add, ast.Sub, ast.Mult, ast.Div,
        ast.USub, ast.UAdd,
        ast.Constant, ast.Load, ast.Expr
    )
    tree = ast.parse(expr, mode="eval")
    for node in ast.walk(tree):
        if not isinstance(node, allowed_nodes):
            raise ValueError(f"Invalid formula node: {type(node).__name__}")
    return eval(compile(tree, "<string>", "eval"))


def apply_formulas_and_logic(rules, values):
    values = values.copy()
    for _ in range(30):
        for _, row in rules.iterrows():
            aop = format_aop(row["AOP"])
            formula = row.get("Formula", "")
            if pd.isna(formula) or str(formula).strip() == "":
                continue
            formula = str(formula).strip().replace("=", "").replace(" ", "")

            def repl(match):
                ref = match.group(0).zfill(3)
                return str(values.get(ref, 0.0))

            expression = re.sub(r"\d{3}", repl, formula)
            try:
                values[aop] = float(safe_eval_expression(expression))
            except Exception:
                values[aop] = 0.0

    for _, row in rules.iterrows():
        aop = format_aop(row["AOP"])
        logic = str(row.get("Logika", "")).strip().lower()
        value = values.get(aop, 0.0)
        if logic == "positive":
            values[aop] = value if value > 0 else 0.0
        elif logic == "negative":
            values[aop] = abs(value) if value < 0 else 0.0
    return values


def calculate_bilans_uspeh(df, rules_file):
    rules = read_rules(rules_file, "Pravila Bilans Uspeh")
    values, positions, order = {}, {}, []
    for _, row in rules.iterrows():
        aop = format_aop(row["AOP"])
        if aop not in order:
            order.append(aop)
        positions[aop] = row.get("Pozicija", "")
        val = sum_rule(df, row.get("konto", None), row.get("kolona", None), row.get("Operacija", "+"))
        values[aop] = values.get(aop, 0.0) + val
    values = apply_formulas_and_logic(rules, values)
    return pd.DataFrame([{"AOP": aop, "Pozicija": positions.get(aop, ""), "Iznos": values.get(aop, 0.0)} for aop in order])


def calculate_bilans_sostojba(df, rules_file):
    rules = read_rules(rules_file, "Pravila bilans Sostojba")
    current_values, previous_values, positions, order = {}, {}, {}, []
    for _, row in rules.iterrows():
        aop = format_aop(row["AOP"])
        if aop not in order:
            order.append(aop)
        positions[aop] = row.get("Pozicija", "")
        current_val = sum_rule(df, row.get("konto", None), row.get("kolona", None), row.get("Operacija", "+"))
        previous_val = sum_rule(df, row.get("konto_prethodna", row.get("konto", None)), row.get("kolona_prethodna", None), row.get("Operacija_prethodna", row.get("Operacija", "+")))
        current_values[aop] = current_values.get(aop, 0.0) + current_val
        previous_values[aop] = previous_values.get(aop, 0.0) + previous_val
    current_values = apply_formulas_and_logic(rules, current_values)
    previous_values = apply_formulas_and_logic(rules, previous_values)
    return pd.DataFrame([{"AOP": aop, "Pozicija": positions.get(aop, ""), "Tekovna godina": current_values.get(aop, 0.0), "Prethodna godina": previous_values.get(aop, 0.0)} for aop in order])


def build_data_map(bu, bs):
    data_map = {}
    for _, row in bu.iterrows():
        aop = str(row["AOP"]).strip().replace(".0", "").zfill(3)
        data_map[f"BU{aop}"] = row["Iznos"]
    for _, row in bs.iterrows():
        aop = str(row["AOP"]).strip().replace(".0", "").zfill(3)
        tekovna = row["Tekovna godina"]
        prethodna = row["Prethodna godina"]
        data_map[f"BS{aop}_T"] = tekovna
        data_map[f"BS{aop}_P"] = prethodna
        data_map[f"BS{aop}_R"] = tekovna - prethodna
    return data_map


def calculate_cash_flow(df, rules_file, bu, bs):
    rules = read_rules(rules_file, "Pravila CF")
    data_map = build_data_map(bu, bs)
    values, positions, order, formula_rows = {}, {}, [], []
    for _, row in rules.iterrows():
        aop = normalize_key(row.get("AOP", ""))
        if aop not in order:
            order.append(aop)
        positions[aop] = row.get("Pozicija", "")
        key = normalize_key(row.get("Key", ""))
        operacija = str(row.get("Operacija", "+")).strip()
        formula = str(row.get("Formula", "")).strip()
        value = 0.0
        if key != "" and key.lower() != "nan":
            value = data_map.get(key, 0.0)
        elif formula != "" and formula.lower() != "nan":
            formula_rows.append(row)
        if operacija == "-":
            value = -value
        values[aop] = values.get(aop, 0.0) + value

    for _ in range(20):
        for row in formula_rows:
            aop = normalize_key(row.get("AOP", ""))
            formula = str(row.get("Formula", "")).strip().replace(" ", "").replace("=", "").upper()
            def repl(match):
                ref = normalize_key(match.group(0))
                return str(values.get(ref, 0.0))
            expression = re.sub(r"CF\d+", repl, formula)
            try:
                values[aop] = float(safe_eval_expression(expression))
            except Exception:
                values[aop] = 0.0
    return pd.DataFrame([{"AOP": aop, "Pozicija": positions.get(aop, ""), "Iznos": values.get(aop, 0.0)} for aop in order])


def calculate_kpi(rules_file, bu, bs):
    rules = read_rules(rules_file, "Pravila KPI")
    data_map = build_data_map(bu, bs)

    results = []

    for _, row in rules.iterrows():
        kpi = str(row.get("AOP", "")).strip()
        naziv = str(row.get("Naziv", "")).strip()
        opis = str(row.get("Opis", "")).strip()
        tip = str(row.get("Tip", "")).strip().lower()
        formula = str(row.get("Formula", "")).strip().replace(" ", "")

        expression = formula

        for key in sorted(data_map.keys(), key=len, reverse=True):
            expression = re.sub(
                rf"\b{re.escape(key)}\b",
                str(data_map[key]),
                expression
            )

        try:
            value = float(eval(expression))
        except Exception:
            value = 0.0

        if tip == "procent":
            prikaz = f"{value * 100:.2f}%"
        elif tip == "iznos":
            prikaz = f"{value:,.0f}"
        else:
            prikaz = f"{value:,.2f}"

        dobro_od = row.get("Dobro_od", None)
        rizik_pod = row.get("Rizik_pod", None)

        status = "ℹ️"

        try:
            dobro_od = float(dobro_od)
            rizik_pod = float(rizik_pod)

            if value >= dobro_od:
                status = "🟢 Добро"
            elif value <= rizik_pod:
                status = "🔴 Ризик"
            else:
                status = "🟡 Средно"

        except:
            pass

        results.append({
            "AOP": kpi,
            "Naziv": naziv,
            "Kategorija": row.get("Kategorija", ""),
            "Opis": opis,
            "Status": status,
            "Vrednost": prikaz
        })

    return pd.DataFrame(results)

def export_excel(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            if data is not None:
                data.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()


def style_formula_rows(row):
    formula_rows = ["201", "207", "213", "223", "224", "234", "235", "246", "247", "250", "251", "001", "002", "009", "010", "020", "021", "031", "036", "037", "045", "052", "053", "059", "063", "065", "071", "081", "082", "085", "095", "111"]
    if str(row["AOP"]).zfill(3) in formula_rows:
        return ["font-weight: bold; font-size: 18px;" for _ in row]
    return ["" for _ in row]


def style_cashflow_rows(row):
    total_rows = ["CF08", "CF12", "CF16", "CF17", "CF20"]
    if row["AOP"] in total_rows:
        if row["AOP"] == "CF17":
            return ["background-color: #2F65D9; color: white; font-weight: 700;" for _ in row]
        return ["background-color: #D9EAD3; font-weight: 700;" for _ in row]
    return ["" for _ in row]


def export_single_pdf(df, title, file_title):
    output = BytesIO()
    pdfmetrics.registerFont(TTFont("Arial", "C:/Windows/Fonts/arial.ttf"))
    pdfmetrics.registerFont(TTFont("Arial-Bold", "C:/Windows/Fonts/arialbd.ttf"))
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    styles["Title"].fontName = "Arial-Bold"
    elements = [Paragraph(title, styles["Title"]), Spacer(1, 12)]
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(table)
    doc.build(elements)
    output.seek(0)
    return output.getvalue()


def export_cash_flow_pdf(df):
    output = BytesIO()
    pdfmetrics.registerFont(TTFont("Arial", "C:/Windows/Fonts/arial.ttf"))
    pdfmetrics.registerFont(TTFont("Arial-Bold", "C:/Windows/Fonts/arialbd.ttf"))
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    styles["Title"].fontName = "Arial-Bold"
    elements = [Paragraph("Извештај за паричните текови", styles["Title"]), Spacer(1, 14)]
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1, colWidths=[50, 420, 120])
    total_rows = ["CF08", "CF12", "CF16", "CF17", "CF20"]
    style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B1F4D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("ALIGN", (2, 1), (2, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
    ])
    for i, row in enumerate(df.itertuples(index=False), start=1):
        aop = str(row.AOP)
        if aop in total_rows:
            style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#D9EAD3"))
            style.add("FONTNAME", (0, i), (-1, i), "Arial-Bold")
            style.add("FONTSIZE", (0, i), (-1, i), 10)
        if aop == "CF17":
            style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#2F65D9"))
            style.add("TEXTCOLOR", (0, i), (-1, i), colors.white)
            style.add("FONTNAME", (0, i), (-1, i), "Arial-Bold")
            style.add("FONTSIZE", (0, i), (-1, i), 11)
    table.setStyle(style)
    elements.append(table)
    doc.build(elements)
    output.seek(0)
    return output.getvalue()


if zaklucen_file :
    try:
        df = read_zaklucen(zaklucen_file)
        st.subheader("📑 Zaklucen list")
        sum_row = {"konto": "", "naziv": "VKUPNO"}
        for col in df.columns:
            if col not in ["konto", "naziv"]:
                sum_row[col] = int(df[col].sum())
        df_total = pd.concat([df, pd.DataFrame([sum_row])], ignore_index=True)
        for col in df_total.columns:
            if col not in ["konto", "naziv"]:
                df_total[col] = df_total[col].round(0).astype(int)

        def style_total_row(row):
            if row["naziv"] == "VKUPNO":
                return ["font-weight: bold; font-size: 18px; background-color: #d9d9d9;" for _ in row]
            return ["" for _ in row]

        st.dataframe(df_total.style.apply(style_total_row, axis=1), use_container_width=True)

        bu = calculate_bilans_uspeh(df, pravila_file)
        bs = calculate_bilans_sostojba(df, pravila_file)
        try:
            cf = calculate_cash_flow(df, pravila_file, bu, bs)
        except Exception as cash_error:
            cf = None
            st.warning(f"Cash Flow ne e vcitan: {cash_error}")
        try:
            kpi = calculate_kpi(pravila_file, bu, bs)
        except Exception as kpi_error:
            kpi = None
            st.warning(f"KPI ne e vcitan: {kpi_error}")

        bu["Iznos"] = bu["Iznos"].round(0).astype(int)
        bs["Tekovna godina"] = bs["Tekovna godina"].round(0).astype(int)
        bs["Prethodna godina"] = bs["Prethodna godina"].round(0).astype(int)
        if cf is not None:
            cf["Iznos"] = cf["Iznos"].round(0).astype(int)

        aktiva = bs.loc[bs["AOP"] == "063", "Tekovna godina"].sum()
        pasiva = bs.loc[bs["AOP"] == "111", "Tekovna godina"].sum()
        razlika = aktiva - pasiva
        aktiva_prev = bs.loc[bs["AOP"] == "063", "Prethodna godina"].sum()
        pasiva_prev = bs.loc[bs["AOP"] == "111", "Prethodna godina"].sum()
        razlika_prev = aktiva_prev - pasiva_prev

        st.subheader("✅ Kontrolni proverki")
        if razlika == 0:
            st.success(f"Tekovna godina: Aktiva = Pasiva ✅ ({aktiva:,.0f})")
        else:
            st.error(f"Tekovna godina: Razlika ❌ {razlika:,.0f}")
        if razlika_prev == 0:
            st.success(f"Prethodna godina: Aktiva = Pasiva ✅ ({aktiva_prev:,.0f})")
        else:
            st.error(f"Prethodna godina: Razlika ❌ {razlika_prev:,.0f}")

        tab1, tab2, tab3, tab4 = st.tabs(["Bilans na uspeh", "Bilans na sostojba", "Cash Flow", "KPI Dashboard"])
        with tab1:
            st.subheader("Bilans na uspeh po AOP")
            st.dataframe(bu.style.apply(style_formula_rows, axis=1), use_container_width=True)
        with tab2:
            st.subheader("Bilans na sostojba po AOP")
            st.dataframe(bs.style.apply(style_formula_rows, axis=1), use_container_width=True)
        with tab3:
            st.subheader("Cash Flow")
            if cf is not None:
                st.dataframe(cf.style.apply(style_cashflow_rows, axis=1), use_container_width=True)
            else:
                st.warning("Cash Flow ne e vcitan.")
        with tab4:

            st.subheader("KPI Dashboard")

            if kpi is not None:

                for category in kpi["Kategorija"].unique():

                    st.subheader(category)

                    category_df = kpi[
                        kpi["Kategorija"] == category
                    ].reset_index(drop=True)

                    cols = st.columns(4)

                    for i in range(len(category_df)):

                        row = category_df.iloc[i]

                        with cols[i % 4]:

                            st.metric(
                                label=str(row["Naziv"]),
                                value=str(row["Vrednost"])
                            )

                st.subheader("🧠 Автоматски коментари")

                for _, row in kpi.iterrows():

                    status = str(row.get("Status", ""))
                    naziv = str(row.get("Naziv", ""))
                    vrednost = str(row.get("Vrednost", ""))

                    if "🟢" in status:
                        st.success(f"{status} {naziv}: {vrednost} — показателот е во добра зона.")

                    elif "🔴" in status:
                        st.error(f"{status} {naziv}: {vrednost} — потребно е внимание и дополнителна анализа.")

                    elif "🟡" in status:
                        st.warning(f"{status} {naziv}: {vrednost} — показателот е во средна зона.")

                    else:
                        st.info(f"{naziv}: {vrednost}")

                with st.expander("📋 Детално објаснување на KPI"):

                    st.dataframe(
                        kpi,
                        use_container_width=True
                    )

            else:

                st.warning("KPI ne e vcitan.")
                
        export_sheets = {"Zaklucen list": df_total, "Bilans uspeh AOP": bu, "Bilans sostojba AOP": bs}
        if cf is not None:
            export_sheets["Cash Flow"] = cf
        if kpi is not None:
            export_sheets["KPI"] = kpi
        excel_data = export_excel(export_sheets)
        st.download_button("📥 Export Excel", data=excel_data, file_name="aop_finansiski_izvestai.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="export_excel_btn")

        pdf_bu = export_single_pdf(bu, "Bilans na uspeh", "bilans_na_uspeh")
        st.download_button("📄 Export PDF - Bilans na uspeh", data=pdf_bu, file_name="bilans_na_uspeh.pdf", mime="application/pdf", key="pdf_bu_btn")
        pdf_bs = export_single_pdf(bs, "Bilans na sostojba", "bilans_na_sostojba")
        st.download_button("📄 Export PDF - Bilans na sostojba", data=pdf_bs, file_name="bilans_na_sostojba.pdf", mime="application/pdf", key="pdf_bs_btn")
        if cf is not None:
            pdf_cf = export_cash_flow_pdf(cf)
            st.download_button("📄 Export PDF - Cash Flow", data=pdf_cf, file_name="cash_flow.pdf", mime="application/pdf", key="pdf_cf_btn")
    except Exception as e:
        st.error(f"Greska: {e}")
