import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from fpdf import FPDF
from io import BytesIO
import base64
import re
from datetime import datetime

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

pdf_buffer = BytesIO()
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=10)


def clean_dataframe(df):
    df = df.rename(columns={k: v for k, v in COLUMN_MAPPING.items() if k in df.columns})
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    if "Outlet" in df.columns:
        df['Outlet Group'] = df['Outlet'].apply(lambda x: re.split(r"[-\s\d]", str(x))[0].strip().upper())
    return df


def plot_bar_chart(df_grouped, title, yaxis_title, sheet_name, color_palette, is_currency=True):
    fig = go.Figure()
    for i, month in enumerate(df_grouped.columns):
        fig.add_trace(go.Bar(
            x=df_grouped.index,
            y=df_grouped[month],
            name=month,
            marker_color=color_palette[i % len(color_palette)],
            text=[
                f"₹ {val:,.2f}" if is_currency and isinstance(val, (int, float)) else f"{val}"
                for val in df_grouped[month]
            ],
            hovertemplate='%{x}<br>%{text}<extra>%{name}</extra>',
        ))

    all_values = df_grouped.values.flatten()
    max_y = all_values.max()
    min_y = all_values.min()
    y_pad = max_y * 0.05

    fig.update_layout(
        title=dict(text=f"{title} - {sheet_name}", font=dict(color='white')),
        xaxis=dict(title=dict(text="Outlet Group", font=dict(color='white')), tickfont=dict(color='white')),
        yaxis=dict(
            title=dict(text=yaxis_title, font=dict(color='white')),
            tickfont=dict(color='white'),
            range=[max(0, min_y - y_pad), max_y + y_pad]
        ),
        legend=dict(font=dict(color='white'), title_font=dict(color='white')),
        font=dict(family="Segoe UI", size=12),
        plot_bgcolor="grey",
        paper_bgcolor="grey",
        height=500,
        margin=dict(l=20, r=20, t=40, b=40),
        barmode='group'
    )
    return fig


def add_to_pdf(text):
    clean_text = str(text).replace("\u2013", "-").replace("–", "-")
    pdf.multi_cell(0, 10, clean_text)


def process_sheet(df, sheet_name="Sheet"):
    global pdf
    df = clean_dataframe(df)
    required_cols = {'PO Number', 'PO Value', 'PO Date', 'Outlet', 'Outlet Group'}
    if not required_cols.issubset(df.columns):
        return

    try:
        df['PO Date'] = pd.to_datetime(df['PO Date'], errors='coerce', dayfirst=True)
        df = df.dropna(subset=['PO Date'])
        df['MonthPeriod'] = df['PO Date'].dt.to_period("M")
        df['Month'] = df['MonthPeriod'].dt.strftime("%b'%y")

        month_order_df = df[['Month', 'MonthPeriod']].drop_duplicates().sort_values('MonthPeriod')
        month_order = month_order_df['Month'].tolist()

        st.sidebar.markdown(f"### 📅 Filter Months – {sheet_name}")
        selected_months = st.sidebar.multiselect("Select Months", options=month_order, default=month_order)
        df = df[df['Month'].isin(selected_months)]

        if df.empty:
            st.warning(f"⚠ No data for selected months in {sheet_name}.")
            return

        filtered_months = month_order_df[month_order_df['Month'].isin(selected_months)]['Month']
        title_months = ', '.join(filtered_months)

        value_grouped = df.groupby(['Outlet Group', 'Month'])['PO Value'].sum().unstack().fillna(0)
        value_grouped = value_grouped[filtered_months]
        st.subheader(f"💰 PO Value – {sheet_name}")
        fig_val = plot_bar_chart(value_grouped, "PO Value", "Value (₹)", sheet_name, COLORS, is_currency=True)
        st.plotly_chart(fig_val, use_container_width=True)

        count_grouped = df.groupby(['Outlet Group', 'Month'])['PO Number'].nunique().unstack().fillna(0)
        count_grouped = count_grouped[filtered_months]
        st.subheader(f"🖗 PO Count – {sheet_name}")
        fig_cnt = plot_bar_chart(count_grouped, "PO Count", "Number of POs", sheet_name, COLORS, is_currency=False)
        st.plotly_chart(fig_cnt, use_container_width=True)

        st.subheader(f"📋 Matrix Report – {sheet_name}")
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
            col: "₹ {:,.2f}" if "Value" in col else "{:,.0f}"
            for col in matrix_combined.columns
        }), use_container_width=True)

        add_to_pdf(f"CPC Performance Report - {sheet_name} - {title_months}")
        add_to_pdf(matrix_combined.to_string())

    except Exception as e:
        st.error(f"Error in {sheet_name}: {e}")


if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            process_sheet(df, "CSV")
        else:
            all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
            for sheet_name, df in all_sheets.items():
                process_sheet(df, sheet_name)

        pdf_output = BytesIO()
        pdf.output(pdf_output)
        b64 = base64.b64encode(pdf_output.getvalue()).decode()
        st.markdown(f"""
            ### 🔽 Download Report
            <a href="data:application/octet-stream;base64,{b64}" download="CPC_Report_{datetime.now().strftime('%b_%Y')}.pdf">
                <button style='padding:10px;background-color:green;color:white;border:none;border-radius:5px;'>
                    🔧 Download PDF
                </button>
            </a>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Unexpected error: {e}")