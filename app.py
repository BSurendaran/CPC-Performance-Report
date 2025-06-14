import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import re
from fpdf import FPDF
import io

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
        title=dict(text=f"{title} – {sheet_name}", font=dict(color='white')),
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

def generate_pdf_from_dataframe(df: pd.DataFrame, title: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", size=14)
    safe_title = title.encode('latin-1', 'replace').decode('latin-1')  # Avoid encoding errors
    pdf.cell(200, 10, txt=safe_title, ln=True, align='C')
    pdf.ln(10)

    pdf.set_font("Arial", size=10)
    col_width = pdf.w / (len(df.columns) + 1)
    row_height = 8

    # Header
    for col in df.columns:
        safe_col = str(col).encode('latin-1', 'replace').decode('latin-1')
        pdf.cell(col_width, row_height, txt=safe_col, border=1)
    pdf.ln(row_height)

    # Rows
    for row in df.itertuples(index=True):
        pdf.cell(col_width, row_height, txt=str(row.Index), border=1)
        for val in row[1:]:
            safe_val = str(val).encode('latin-1', 'replace').decode('latin-1')
            pdf.cell(col_width, row_height, txt=safe_val, border=1)
        pdf.ln(row_height)

    output = io.BytesIO()
    pdf.output(output)
    output.seek(0)
    return output.read()

def process_sheet(df, sheet_name="Sheet"):
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

        # 💰 PO Value Chart
        value_grouped = df.groupby(['Outlet Group', 'Month'])['PO Value'].sum().unstack().fillna(0)
        value_grouped = value_grouped[filtered_months]
        st.subheader(f"💰 PO Value – {sheet_name}")
        st.plotly_chart(
            plot_bar_chart(value_grouped, "PO Value", "Value (₹)", sheet_name, COLORS, is_currency=True),
            use_container_width=True
        )

        # 🔢 PO Count Chart
        count_grouped = df.groupby(['Outlet Group', 'Month'])['PO Number'].nunique().unstack().fillna(0)
        count_grouped = count_grouped[filtered_months]
        st.subheader(f"🔢 PO Count – {sheet_name}")
        st.plotly_chart(
            plot_bar_chart(count_grouped, "PO Count", "Number of POs", sheet_name, COLORS, is_currency=False),
            use_container_width=True
        )

        # 📋 Matrix Report
        st.subheader(f"📋 Matrix Report – {sheet_name}")
        subcategory_col = next((col for col in df.columns if 'SUB' in col.upper()), None)
        group_cols = []
        if subcategory_col:
            group_cols.append(subcategory_col)

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

        # 📥 PDF Download Button
        title = f"Matrix Report – {sheet_name} – {', '.join(filtered_months)}"
        pdf_bytes = generate_pdf_from_dataframe(matrix_combined, title=title)

        st.download_button(
            label="📥 Download Matrix Report as PDF",
            data=pdf_bytes,
            file_name=f"{sheet_name}_Matrix_Report_{'_'.join(filtered_months)}.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Error in {sheet_name}: {e}")

# 📁 File Upload Handling
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
        st.error(f"❌ Unexpected error: {e}")
