import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import re
from io import BytesIO
from fpdf import FPDF
import tempfile
import os

st.set_page_config(layout="wide")
st.title("CPC Performance Report")

uploaded_file = st.file_uploader("Upload Excel or CSV File", type=["xlsx", "csv"])

COLUMN_MAPPING = {
    "OUTLET": "Outlet",
    "PO REF NO": "PO Number",
    "PO DATE": "PO Date",
    "PO VALUE": "PO Value"
}

COLORS = px.colors.qualitative.Bold


def clean_dataframe(df):
    df = df.rename(columns={k: v for k, v in COLUMN_MAPPING.items() if k in df.columns})
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    if "Outlet" in df.columns:
        df['Outlet Group'] = df['Outlet'].apply(lambda x: re.split(r"[-\s\d]", str(x))[0].strip().upper())
    return df


def save_chart_as_image(fig):
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    fig.write_image(temp_file.name)
    return temp_file.name


class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, self.title, ln=True, align='C')
        self.ln(5)

    def add_image(self, path, title):
        self.add_page()
        self.set_font("Arial", size=12)
        self.cell(0, 10, title, ln=True)
        self.image(path, x=10, y=25, w=190)

    def add_matrix_table(self, df, title):
        self.add_page()
        self.set_font("Arial", 'B', 11)
        self.cell(0, 10, title, ln=True)

        col_width = (self.w - 2 * self.l_margin) / len(df.columns)
        row_height = 6

        self.set_font("Arial", 'B', 10)
        for col in df.columns:
            self.cell(col_width, row_height, str(col), border=1, align='C')
        self.ln()

        self.set_font("Arial", '', 9)
        for idx, row in df.iterrows():
            self.cell(col_width, row_height, str(idx), border=1)
            for item in row:
                formatted = f"{item:,.2f}" if isinstance(item, float) else str(item)
                self.cell(col_width, row_height, formatted, border=1, align='R')
            self.ln()


def generate_pdf(sheet_name, value_img, count_img, matrix_df):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_title(f"{sheet_name} Report")

    pdf.add_image(value_img, f"PO Value – {sheet_name}")
    pdf.add_image(count_img, f"PO Count – {sheet_name}")
    pdf.add_matrix_table(matrix_df, f"Matrix Report – {sheet_name}")

    buffer = BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer


def plot_bar_chart(df_grouped, title, yaxis_title, sheet_name, color_palette, is_currency=True):
    fig = go.Figure()
    for i, month in enumerate(df_grouped.columns):
        fig.add_trace(go.Bar(
            x=df_grouped.index,
            y=df_grouped[month],
            name=month,
            marker_color=color_palette[i % len(color_palette)],
            text=[
                f"\u20B9 {val:,.2f}" if is_currency and isinstance(val, (int, float)) else f"{val}"
                for val in df_grouped[month]
            ],
            hovertemplate='%{x}<br>%{text}<extra>%{name}</extra>',
        ))

    fig.update_layout(
        title=dict(text=f"{title} - {sheet_name}"),
        xaxis=dict(title="Outlet Group"),
        yaxis=dict(title=yaxis_title),
        barmode='group',
        height=400
    )
    return fig


def process_sheet(df, sheet_name="Sheet"):
    df = clean_dataframe(df)
    required_cols = {'PO Number', 'PO Value', 'PO Date', 'Outlet', 'Outlet Group'}
    if not required_cols.issubset(df.columns):
        return

    df['PO Date'] = pd.to_datetime(df['PO Date'], errors='coerce', dayfirst=True)
    df = df.dropna(subset=['PO Date'])
    df['MonthPeriod'] = df['PO Date'].dt.to_period("M")
    df['Month'] = df['MonthPeriod'].dt.strftime("%b'%y")
    month_order_df = df[['Month', 'MonthPeriod']].drop_duplicates().sort_values('MonthPeriod')
    month_order = month_order_df['Month'].tolist()

    st.sidebar.markdown(f"### Filter Months – {sheet_name}")
    selected_months = st.sidebar.multiselect("Select Months", options=month_order, default=month_order)
    df = df[df['Month'].isin(selected_months)]
    if df.empty:
        st.warning(f"No data for selected months in {sheet_name}.")
        return

    filtered_months = month_order_df[month_order_df['Month'].isin(selected_months)]['Month']

    # PO Value Chart
    value_grouped = df.groupby(['Outlet Group', 'Month'])['PO Value'].sum().unstack().fillna(0)
    value_grouped = value_grouped[filtered_months]
    st.subheader(f"\U0001F4B0 PO Value – {sheet_name}")
    value_fig = plot_bar_chart(value_grouped, "PO Value", "Value (\u20B9)", sheet_name, COLORS, is_currency=True)
    st.plotly_chart(value_fig, use_container_width=True)

    # PO Count Chart
    count_grouped = df.groupby(['Outlet Group', 'Month'])['PO Number'].nunique().unstack().fillna(0)
    count_grouped = count_grouped[filtered_months]
    st.subheader(f"\U0001F522 PO Count – {sheet_name}")
    count_fig = plot_bar_chart(count_grouped, "PO Count", "Number of POs", sheet_name, COLORS, is_currency=False)
    st.plotly_chart(count_fig, use_container_width=True)

    # Matrix Report
    st.subheader(f"\U0001F4CB Matrix Report – {sheet_name}")
    subcategory_col = next((col for col in df.columns if 'SUB' in col.upper()), None)
    group_cols = [subcategory_col] if subcategory_col else []

    matrix_count = df.groupby(group_cols + ['Month'])['PO Number'].nunique().unstack().fillna(0)
    matrix_count = matrix_count[filtered_months]
    matrix_count['Total Count'] = matrix_count.sum(axis=1)

    matrix_value = df.groupby(group_cols + ['Month'])['PO Value'].sum().unstack().fillna(0)
    matrix_value = matrix_value[filtered_months]
    matrix_value['Total Value'] = matrix_value.sum(axis=1)

    matrix_combined = pd.concat([matrix_count, matrix_value], axis=1, keys=['PO No', 'PO Value'])
    matrix_combined.columns = [' '.join(col).strip() for col in matrix_combined.columns.values]

    st.dataframe(matrix_combined.style.format({
        col: "\u20B9 {:,.2f}" if "Value" in col else "{:,.0f}"
        for col in matrix_combined.columns
    }), use_container_width=True)

    # PDF download
    if st.button(f"\U0001F4C4 Download {sheet_name} Report as PDF"):
        value_img_path = save_chart_as_image(value_fig)
        count_img_path = save_chart_as_image(count_fig)
        pdf_buffer = generate_pdf(sheet_name, value_img_path, count_img_path, matrix_combined)
        st.download_button(
            label=f"\U0001F4E5 Save {sheet_name} Report PDF",
            data=pdf_buffer,
            file_name=f"{sheet_name}_Report.pdf",
            mime="application/pdf"
        )
        os.unlink(value_img_path)
        os.unlink(count_img_path)


if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            process_sheet(df, "CSV")
        else:
            all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
            for sheet_name, df in all_sheets.items():
                process_sheet(df, sheet_name)
    except Exception as e:
        st.error(f"\u274C Unexpected error: {e}")
