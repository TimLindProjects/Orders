import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import BytesIO
from fpdf import FPDF
import matplotlib.pyplot as plt

# Functie: Sla een DataFrame op als afbeelding met aangepaste kolombreedtes en celkleuring inclusief opmaak
def save_table_image_with_coloring(df, filename, column_color_funcs=None):
    # df: verwacht dat de celwaarden reeds geformatteerd zijn als strings (met gewenste afronding)
    # Bepaal kolombreedtes op basis van de lengte van de tekst in elke cel en header
    col_widths = []
    for col in df.columns:
        header = str(col)
        max_len = len(header)
        for cell in df[col]:
            max_len = max(max_len, len(str(cell)))
        col_widths.append(max_len)
    total = sum(col_widths)
    col_widths = [w / total for w in col_widths]

    # Afmetingen van de figuur: gebaseerd op het aantal kolommen en rijen
    fig_width = max(8, len(df.columns) * 1.5)
    fig_height = max(2, (df.shape[0] + 1) * 0.6)
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis('tight')
    ax.axis('off')
    # Maak de tabel
    table = ax.table(cellText=df.values,
                     colLabels=df.columns,
                     cellLoc='center',
                     loc='center',
                     colWidths=col_widths)
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    # Pas celkleuringen toe
    nrows, ncols = df.shape
    for i in range(nrows + 1):
        for j in range(ncols):
            cell = table[i, j]
            if i == 0:
                cell.set_text_props(weight='bold')
                cell.set_facecolor("#d3d3d3")
            else:
                col_name = df.columns[j]
                cel_text = df.values[i - 1, j]
                if column_color_funcs and col_name in column_color_funcs:
                    try:
                        # Verwijder spaties en vervang komma door punt voor conversie
                        numeric_val = float(str(cel_text).replace(",", ".").strip())
                    except Exception:
                        numeric_val = cel_text
                    tekst_kleur = column_color_funcs[col_name](numeric_val)
                    if tekst_kleur:
                        cell.get_text().set_color(tekst_kleur)
    plt.tight_layout()
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    plt.close(fig)

# Hulpfunctie: formatteer een DataFrame voor presentatie
def format_dataframe(df, num_cols):
    df_formatted = df.copy()
    for col in num_cols:
        if col in df_formatted.columns:
            df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:.2f}".replace(".", ","))
    return df_formatted

# Pagina-configuratie
st.set_page_config(page_title="Orders Analyse Dashboard", layout="wide")
st.title("Orders Analyse Dashboard")

# Sidebar: Bestand uploaden en instellingen
st.sidebar.header("Upload Excel-bestand")
uploaded_file = st.sidebar.file_uploader("Upload een Excel-bestand (.xlsx)", type=["xlsx"])

# Sidebar: Groepeeroptie: Week, Maand en Jaar
group_option = st.sidebar.radio("Selecteer groeperingsniveau", options=["Week", "Maand", "Jaar"], index=1)

if uploaded_file is not None:
    try:
        # Vereiste kolommen
        required_columns = {"Basisstartterm.", "BasEindterm.", "Order", "Gepland totaal", "Werk. totaal"}
        xls = pd.ExcelFile(uploaded_file)
        df_valid = None
        for sheet in xls.sheet_names:
            df_sheet = xls.parse(sheet)
            if required_columns.issubset(df_sheet.columns):
                df_valid = df_sheet.copy()
                break
        if df_valid is None:
            st.error("Geen werkblad met de vereiste kolommen gevonden!")
        else:
            # Data voorbereiding en opschonen
            df_valid["Basisstartterm."] = pd.to_datetime(df_valid["Basisstartterm."], errors="coerce").dt.normalize()
            df_valid["BasEindterm."] = pd.to_datetime(df_valid["BasEindterm."], errors="coerce")
            for col in ["Werk. totaal", "Gepland totaal"]:
                if df_valid[col].dtype == object:
                    df_valid[col] = df_valid[col].str.replace(',', '.')
                    df_valid[col] = pd.to_numeric(df_valid[col], errors="coerce")
            # Gemiddelde berekening: vervang 0-waarden in "Werk. totaal"
            df_nonzero = df_valid[df_valid["Werk. totaal"] != 0]
            avg_werk = df_nonzero.groupby("Gepland totaal")["Werk. totaal"].mean().reset_index()
            avg_werk = avg_werk.rename(columns={"Werk. totaal": "Gemiddelde Werk. totaal"})
            avg_mapping = avg_werk.set_index("Gepland totaal")["Gemiddelde Werk. totaal"].to_dict()
            df_valid["Werk. totaal"] = df_valid.apply(
                lambda row: avg_mapping.get(row["Gepland totaal"], row["Werk. totaal"])
                if row["Werk. totaal"] == 0 else row["Werk. totaal"],
                axis=1
            )
            df_valid = df_valid.dropna(subset=["Basisstartterm."]).sort_values(by="Basisstartterm.")
            min_date = df_valid["Basisstartterm."].min().date()
            max_date = df_valid["Basisstartterm."].max().date()
            start_date, end_date = st.sidebar.slider("Selecteer datumbereik", min_date, max_date, (min_date, max_date))
            df_filtered = df_valid[(df_valid["Basisstartterm."].dt.date >= start_date) &
                                   (df_valid["Basisstartterm."].dt.date <= end_date)]
            # Extra kolommen en berekeningen
            df_filtered["Datum"] = df_filtered["Basisstartterm."].dt.strftime("%Y-%m-%d")
            df_filtered["Duur"] = (df_filtered["BasEindterm."] - df_filtered["Basisstartterm."]).dt.days
            if group_option == "Week":
                df_filtered["Groep"] = df_filtered["Basisstartterm."].dt.strftime("%Y-W%U")
            elif group_option == "Maand":
                df_filtered["Groep"] = df_filtered["Basisstartterm."].dt.to_period("M").astype(str)
            else:
                df_filtered["Groep"] = df_filtered["Basisstartterm."].dt.year

            # --- Visualisaties ---

            # Staafdiagram: Aantal Orders per groep
            orders_per_group = df_filtered.groupby("Groep")["Order"].count().reset_index(name="Aantal Orders")
            fig_bar = px.bar(
                orders_per_group,
                x="Groep",
                y="Aantal Orders",
                title=f"Aantal Orders per {group_option}",
                labels={"Groep": group_option, "Aantal Orders": "Aantal Orders"},
                color_discrete_sequence=px.colors.qualitative.Plotly
            )
            fig_bar.update_layout(height=500, autosize=True)
            st.plotly_chart(fig_bar, use_container_width=True)

            # Kostenberekeningen & grafieken
            costs_per_group = df_filtered.groupby("Groep").agg({
                "Gepland totaal": "sum",
                "Werk. totaal": "sum"
            }).reset_index()
            costs_per_group["Verschil"] = costs_per_group["Gepland totaal"] - costs_per_group["Werk. totaal"]
            work_adjusted = []
            work_colors = []
            standaard_color = px.colors.qualitative.Set1[1]
            for _, row in costs_per_group.iterrows():
                if row["Werk. totaal"] == 0:
                    work_adjusted.append(row["Gepland totaal"])
                    work_colors.append("green")
                else:
                    work_adjusted.append(row["Werk. totaal"])
                    work_colors.append(standaard_color)
            fig_costs = go.Figure()
            fig_costs.add_trace(
                go.Bar(
                    x=costs_per_group["Groep"],
                    y=costs_per_group["Gepland totaal"],
                    name="Gepland totaal",
                    marker_color=px.colors.qualitative.Set1[0]
                )
            )
            fig_costs.add_trace(
                go.Bar(
                    x=costs_per_group["Groep"],
                    y=work_adjusted,
                    name="Werk. totaal",
                    marker_color=work_colors
                )
            )
            fig_costs.add_trace(
                go.Bar(
                    x=costs_per_group["Groep"],
                    y=costs_per_group["Verschil"],
                    name="Verschil",
                    marker_color=px.colors.qualitative.Set1[2]
                )
            )
            fig_costs.update_layout(
                title=f"Kosten Overzicht per {group_option}",
                barmode="group",
                yaxis=dict(autorange=True),
                height=500,
                autosize=True
            )
            st.plotly_chart(fig_costs, use_container_width=True)
            costs_per_group["Werk aangepast"] = work_adjusted
            costs_per_group = costs_per_group.sort_values(by="Groep")
            costs_per_group["Cumulatief werkelijk"] = costs_per_group["Werk aangepast"].cumsum()
            fig_line = go.Figure()
            fig_line.add_trace(
                go.Scatter(
                    x=costs_per_group["Groep"],
                    y=costs_per_group["Cumulatief werkelijk"],
                    mode="lines+markers",
                    name="Totale werkelijke kosten",
                    line=dict(color="blue")
                )
            )
            fig_line.update_layout(
                title="Cumulatieve totale werkelijke kosten",
                yaxis=dict(range=[0, max(costs_per_group["Cumulatief werkelijk"]) * 1.1]),
                height=500,
                autosize=True
            )
            st.plotly_chart(fig_line, use_container_width=True)

            # Order Kosten Overzicht
            totaal_gepland = costs_per_group["Gepland totaal"].sum()
            totaal_werkelijk = sum(work_adjusted)
            totaal_verschil = totaal_gepland - totaal_werkelijk
            st.subheader("Order Kosten Overzicht")
            col1, col2, col3 = st.columns(3)
            col1.metric("Totaal Gepland", f"EUR {totaal_gepland:,.2f}")
            col2.metric("Totaal Werkelijk", f"EUR {totaal_werkelijk:,.2f}")
            col3.metric("Verschil", f"EUR {totaal_verschil:,.2f}")

            # Toon de Gemiddelde Werk. totaal per Gepland totaal tabel
            st.markdown("### Gemiddelde Werk. totaal per Gepland totaal")
            st.table(avg_werk)

            # Details Overzicht: bouw de export-tabel
            df_filtered["Verschil"] = df_filtered["Gepland totaal"] - df_filtered["Werk. totaal"]
            if "Korte tekst" in df_filtered.columns:
                df_export = df_filtered[['Datum', 'Korte tekst', 'Order', 'Gepland totaal',
                                         'Werk. totaal', 'Verschil', 'Duur']].copy()
                df_export = df_export.rename(columns={'Korte tekst': 'Omschrijving'})
            else:
                df_export = df_filtered[['Datum', 'Order', 'Gepland totaal',
                                         'Werk. totaal', 'Verschil', 'Duur']].copy()
                df_export.insert(1, 'Omschrijving', "Geen korte tekst beschikbaar.")
            df_export = df_export[['Datum', 'Omschrijving', 'Order', 'Gepland totaal', 'Werk. totaal', 'Verschil', 'Duur']]
            df_export = df_export.rename(columns={"Duur": "Duur in dagen"})
            num_cols = ["Gepland totaal", "Werk. totaal", "Verschil", "Duur in dagen"]
            df_export[num_cols] = df_export[num_cols].round(2)
            df_export_formatted = format_dataframe(df_export, num_cols)

            st.markdown("### Details Overzicht", unsafe_allow_html=True)
            styled_export = df_export_formatted.style.applymap(
                lambda x: "color: red;" if float(x.replace(",", ".").strip()) < 0 else "",
                subset=["Verschil"]
            )
            st.table(styled_export)

            # --- Export naar PDF ---
            if st.button("Export naar PDF"):
                # Exporteer grafieken als afbeeldingen (zorg dat kaleido geïnstalleerd is)
                fig_bar.write_image("bar_chart.png")
                fig_costs.write_image("costs_chart.png")
                fig_line.write_image("line_chart.png")

                # Exporteer de tabellen als afbeeldingen:
                save_table_image_with_coloring(
                    df_export_formatted,
                    "details_table.png",
                    column_color_funcs={
                        "Verschil": lambda x: "red" if ((isinstance(x, str) and float(x.replace(",", ".").strip()) < 0)
                                                         or (isinstance(x, (int, float)) and x < 0)) else "black"
                    }
                )
                avg_table_formatted = format_dataframe(avg_werk, ["Gemiddelde Werk. totaal"])
                save_table_image_with_coloring(
                    avg_table_formatted,
                    "average_table.png"
                )

                # Bouw PDF-document met FPDF
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, "Orders Analyse Dashboard Export", ln=1, align="C")
                pdf.ln(10)
                pdf.set_font("Arial", "", 12)
                pdf.cell(0, 10, f"Aantal Orders per {group_option}", ln=1)
                pdf.image("bar_chart.png", x=10, y=40, w=pdf.w - 20)
                pdf.ln(85)
                pdf.add_page()
                pdf.cell(0, 10, f"Kosten Overzicht per {group_option}", ln=1)
                pdf.image("costs_chart.png", x=10, y=30, w=pdf.w - 20)
                pdf.ln(85)
                pdf.add_page()
                pdf.cell(0, 10, "Cumulatieve totale werkelijke kosten", ln=1)
                pdf.image("line_chart.png", x=10, y=30, w=pdf.w - 20)
                pdf.ln(85)
                pdf.add_page()
                pdf.cell(0, 10, "Details Overzicht", ln=1)
                pdf.image("details_table.png", x=10, y=20, w=pdf.w - 20)
                pdf.add_page()
                pdf.cell(0, 10, "Gemiddelde Werk. totaal per Gepland totaal", ln=1)
                pdf.image("average_table.png", x=10, y=30, w=pdf.w - 20)
                pdf.ln(85)
                pdf.add_page()
                pdf.cell(0, 10, "Order Kosten Overzicht", ln=1)
                pdf.cell(0, 10, f"Totaal Gepland: EUR {totaal_gepland:,.2f}", ln=1)
                pdf.cell(0, 10, f"Totaal Werkelijk: EUR {totaal_werkelijk:,.2f}", ln=1)
                pdf.cell(0, 10, f"Verschil: EUR {totaal_verschil:,.2f}", ln=1)
                pdf.ln(10)
                pdf_buffer = BytesIO()
                pdf.output(pdf_buffer)
                pdf_data = pdf_buffer.getvalue()
                pdf_buffer.close()

                st.download_button(
                    label="Download PDF",
                    data=pdf_data,
                    file_name="orders_analyse_export.pdf",
                    mime="application/pdf"
                )

                # Verwijder de tijdelijke PNG-bestanden
                for filename in ["bar_chart.png", "costs_chart.png", "line_chart.png", "details_table.png", "average_table.png"]:
                    try:
                        if os.path.exists(filename):
                            os.remove(filename)
                    except Exception as rem_err:
                        st.error(f"Fout bij verwijderen van {filename}: {rem_err}")

    except Exception as e:
        st.error(f"Fout opgetreden: {e}")
else:
    st.info("Upload een Excel-bestand om te beginnen.")
