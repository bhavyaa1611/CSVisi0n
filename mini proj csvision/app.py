import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
import tempfile
import os
import seaborn as sns
from matplotlib.figure import Figure
from pandas.api.types import is_numeric_dtype, is_string_dtype

st.set_page_config(layout="wide", page_title="CSVisi0n - Quick Data Analyzer")

st.title("CSVisi0n — CSV → Excel + Visualizations + AI Summary")
st.caption("Upload CSV, inspect, visualize, get a textual summary and export an Excel report.")

uploaded = st.file_uploader("Upload CSV file", type=["csv", "txt"])
if not uploaded:
    st.info("Upload a CSV to start.")
    st.stop()

@st.cache_data
def load_csv(file):
    try:
        return pd.read_csv(file)
    except Exception:
        file.seek(0)
        return pd.read_csv(file, encoding='latin1', error_bad_lines=False)

df = load_csv(uploaded)
st.success(f"Loaded {len(df)} rows and {len(df.columns)} columns")

# Preview
with st.expander("Data preview & basic info", expanded=True):
    st.dataframe(df.head(100))
    st.write("Columns:", list(df.columns))
    st.write("Shape:", df.shape)

# Summary stats
with st.expander("Summary (dtypes + missing)", expanded=False):
    summary = pd.DataFrame({
        "dtype": df.dtypes.astype(str),
        "missing": df.isna().sum()
    })
    st.dataframe(summary)

# Helper plotting functions
def plot_numeric_hist(col):
    fig = Figure(figsize=(4,3))
    ax = fig.subplots()
    sns.histplot(df[col].dropna(), ax=ax, kde=True)
    ax.set_title(f"Histogram — {col}")
    return fig

def plot_categorical_bar(col):
    vc = df[col].value_counts().head(20)
    fig = Figure(figsize=(4,3))
    ax = fig.subplots()
    vc.plot.bar(ax=ax)
    ax.set_title(f"Counts — {col}")
    return fig

def plot_corr_heatmap(numeric_df):
    corr = numeric_df.corr()
    fig = Figure(figsize=(6,5))
    ax = fig.subplots()
    sns.heatmap(corr, annot=False, cmap="vlag", center=0, ax=ax)
    ax.set_title("Correlation heatmap")
    return fig, corr

def plot_missing_map():
    fig = Figure(figsize=(8,2))
    ax = fig.subplots()
    sns.heatmap(df.isna().T, cbar=False, ax=ax)
    ax.set_title("Missing Values Map")
    return fig

numeric_cols = [c for c in df.columns if is_numeric_dtype(df[c])]
cat_cols = [c for c in df.columns if is_string_dtype(df[c]) or df[c].nunique() < 30]

st.subheader("Visualizations")
cols = st.columns(2)
with cols[0]:
    st.write("Numeric")
    if numeric_cols:
        for col in numeric_cols[:5]:
            st.pyplot(plot_numeric_hist(col))
with cols[1]:
    st.write("Categorical")
    if cat_cols:
        for col in cat_cols[:5]:
            st.pyplot(plot_categorical_bar(col))

# Correlation + missing
if numeric_cols:
    corr_fig, corr_df = plot_corr_heatmap(df[numeric_cols])
    st.pyplot(corr_fig)
st.pyplot(plot_missing_map())

# AI summary
def generate_summary(df):
    nrows, ncols = df.shape
    out = []
    out.append(f"Dataset has {nrows} rows and {ncols} columns.")
    out.append(f"Missing values: {df.isna().sum().sum()}")
    out.append(f"Numeric columns: {len(numeric_cols)}")
    out.append(f"Categorical columns: {len(cat_cols)}")
    return "\n".join(out)

summary_text = generate_summary(df)
st.subheader("AI Summary")
st.text_area("Summary", summary_text, height=200)

# Excel report
def create_excel_report(df):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    path = tmp.name
    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    df.to_excel(writer, sheet_name="raw_data", index=False)

    summary_sheet = writer.book.add_worksheet("summary")
    writer.sheets["summary"] = summary_sheet

    summary_sheet.write(0,0,"AI Summary")
    for i, line in enumerate(summary_text.split("\n"), start=1):
        summary_sheet.write(i,0,line)

    writer.save()
    return path

if st.button("Generate Excel Report"):
    excel_path = create_excel_report(df)
    with open(excel_path, "rb") as f:
        st.download_button("Download Excel", f, file_name="report.xlsx")

st.markdown("---")
st.caption("Built from scratch. Need more features? Just ask.")