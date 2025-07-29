import pandas as pd
import streamlit as st
import altair as alt
import io
import os
from prophet import Prophet
from datetime import datetime
import tempfile
import base64

st.set_page_config(page_title="Jarvis Excel AI", layout="wide")
st.title("🤖 Jarvis – Universal Excel AI")

# Upload multiple Excel files
uploaded_files = st.file_uploader("📥 Upload up to 3 Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_dataframes = []
    for uploaded_file in uploaded_files:
        try:
            # Handle multiple sheets
            sheet_names = pd.ExcelFile(uploaded_file).sheet_names
            selected_sheet = st.selectbox(f"Select Sheet from {uploaded_file.name}", sheet_names, key=uploaded_file.name)
            data = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

            # Convert object columns to string to avoid Arrow issues
            for col in data.select_dtypes(include='object').columns:
                data[col] = data[col].astype(str)

            all_dataframes.append((uploaded_file.name, data))

        except Exception as e:
            st.error(f"❌ Error loading {uploaded_file.name}: {e}")

    if all_dataframes:
        # Combine all files
        combined_data = pd.concat([df.assign(Source=name) for name, df in all_dataframes], ignore_index=True)
        st.success("✅ All files loaded and merged successfully!")

        # Detect columns
        text_columns = combined_data.select_dtypes(include='object').columns.tolist()
        year_columns = [col for col in combined_data.columns if col.startswith("CY")]

        # View data
        with st.expander("🔍 View Merged Data"):
            st.dataframe(combined_data)

        # Sidebar filters
        st.sidebar.header("🔽 Filters")
        selected_years = st.sidebar.multiselect("Select Year(s)", year_columns, default=year_columns[:1])

        filter_values = {}
        for col in text_columns:
            values = st.sidebar.multiselect(f"Filter: {col}", combined_data[col].unique(), key=col)
            if values:
                filter_values[col] = values

        # Apply filters
        filtered_data = combined_data.copy()
        for col, vals in filter_values.items():
            filtered_data = filtered_data[filtered_data[col].isin(vals)]

        # Totals per year
        st.subheader("📊 Total Volumes by Year")
        for year in selected_years:
            if year in filtered_data.columns:
                st.metric(label=f"Total in {year}", value=f"{int(filtered_data[year].sum()):,}")

        # Chart grouping
        group_col = st.selectbox("📎 Group chart by column", text_columns)
        if group_col and selected_years:
            chart_data = filtered_data[[group_col] + selected_years].copy()
            chart_data = chart_data.melt(id_vars=group_col, var_name='Year', value_name='Volume')

            chart = alt.Chart(chart_data).mark_bar().encode(
                x='Year:N',
                y='Volume:Q',
                color='Year:N',
                column=alt.Column(group_col + ':N')
            ).properties(height=300)

            st.altair_chart(chart, use_container_width=True)

        # Forecasting section
        st.subheader("🔮 Forecasting with Prophet")
        forecast_col = st.selectbox("Select a year column to forecast", year_columns)
        if forecast_col:
            try:
                df_forecast = filtered_data[[forecast_col]].copy()
                df_forecast = df_forecast.groupby(filtered_data.index).sum().reset_index()
                df_forecast.columns = ['ds', 'y']
                df_forecast['ds'] = pd.date_range(start='2023-01-01', periods=len(df_forecast), freq='Y')

                model = Prophet()
                model.fit(df_forecast)
                future = model.make_future_dataframe(periods=3, freq='Y')
                forecast = model.predict(future)

                st.line_chart(forecast[['ds', 'yhat']].set_index('ds'))
            except Exception as e:
                st.error(f"❌ Forecasting error: {e}")

        # Download CSV
        st.download_button(
            label="📥 Download Filtered CSV",
            data=filtered_data.to_csv(index=False).encode('utf-8'),
            file_name="filtered_combined_data.csv",
            mime="text/csv"
        )

else:
    st.info("Upload up to 3 Excel files to get started!")
