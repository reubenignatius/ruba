import pandas as pd
import streamlit as st
import plotly.express as px
from streamlit_plotly_events import plotly_events # New library for better event handling

# --- Configuration and Data Loading ---

# 1. Update the file path and sheet name based on your input
FILE_PATH = 'C:/Users/REUBEN/Documents/Dummy Data.xlsx'
SHEET_NAME = 'Sh1' # Check spelling and capitalization of this sheet name in Excel

# --- SESSION STATE INITIALIZATION ---
# Initialize session state variables to hold chart click selections
if 'region_click_filter' not in st.session_state:
    st.session_state.region_click_filter = []
if 'category_click_filter' not in st.session_state:
    st.session_state.category_click_filter = []
# --- END SESSION STATE ---

# Use Streamlit's cache decorator to load the data efficiently (only loads once)
@st.cache_data
def load_data(file_path, sheet_name):
    """Loads data from the Excel file into a pandas DataFrame."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        return df
    except FileNotFoundError:
        st.error(f"Error: Excel file not found at path: {file_path}")
        return pd.DataFrame()
    except KeyError:
        # --- IMPROVED ERROR HANDLING FOR SHEET NAME ---
        try:
            xlsx = pd.ExcelFile(file_path, engine='openpyxl')
            available_sheets = ", ".join([f"**'{s}'**" for s in xlsx.sheet_names])
            st.error(
                f"🚨 **Sheet Not Found Error**:\n\n"
                f"The worksheet named **'{sheet_name}'** was not found in the Excel file. "
                f"Please check the spelling and capitalization in the `SHEET_NAME` variable.\n\n"
                f"**Available Sheets:** {available_sheets}"
            )
        except Exception:
             st.error(f"Error: Sheet name '{sheet_name}' not found. Please check the sheet name.")
        # --- END IMPROVED ERROR HANDLING ---
        return pd.DataFrame()
    except Exception as e:
        st.error(f"An unexpected error occurred while loading data: {e}")
        return pd.DataFrame()

df = load_data(FILE_PATH, SHEET_NAME)

# Check if DataFrame is empty due to an error
if df.empty:
    st.stop() # Stop the script if data loading failed

# --- 2. Dashboard Layout and Title ---
st.set_page_config(layout="wide")
st.title("📊 Reuben's Excel Sales Data Dashboard (Interactive)")

# --- 3. Interactive Filter (Sidebar) ---
st.sidebar.header("Filter Options")
st.sidebar.markdown("---")

# Function to reset chart filters
def reset_chart_filters():
    st.session_state.region_click_filter = []
    st.session_state.category_click_filter = []

# Button to clear chart filters
st.sidebar.button("Clear Chart Filters", on_click=reset_chart_filters, help="Removes any filtering applied by clicking on the pie charts.")
st.sidebar.markdown("---")

st.sidebar.info("Sidebar selections filter all data. Chart clicks below can be used to apply temporary, single-value filters.")


# Initialize the filtered DataFrame to the full dataset
df_filtered = df.copy()

# --- NEW FILTER APPLICATION LOGIC (Explicit Priority) ---

# 1. Handle Region Filter (Chart Click Priority)
region_options = df['Region'].unique()
if st.session_state.region_click_filter:
    selected_regions = st.session_state.region_click_filter
    st.sidebar.markdown(f"**Region Filter (Chart Override):** `{selected_regions[0]}`")
    st.sidebar.multiselect('Select Region:', options=region_options, default=selected_regions, disabled=True)
else:
    selected_regions = st.sidebar.multiselect(
        'Select Region:',
        options=region_options,
        default=region_options.tolist()
    )

# 2. Handle Category Filter (Chart Click Priority)
category_options = df['Category'].unique()
if st.session_state.category_click_filter:
    selected_categories = st.session_state.category_click_filter
    st.sidebar.markdown(f"**Category Filter (Chart Override):** `{selected_categories[0]}`")
    st.sidebar.multiselect('Select Category:', options=category_options, default=selected_categories, disabled=True)
else:
    selected_categories = st.sidebar.multiselect(
        'Select Category:',
        options=category_options,
        default=category_options.tolist()
    )
    
# 3. Handle remaining filters using ONLY multiselect
selected_suppliers = st.sidebar.multiselect(
    'Select Supplier:',
    options=df['Supplier'].unique(),
    default=df['Supplier'].unique().tolist()
)

selected_items = st.sidebar.multiselect(
    'Select Item:',
    options=df['Item'].unique(),
    default=df['Item'].unique().tolist()
)

selected_months = st.sidebar.multiselect(
    'Select Month:',
    options=df['Month'].unique(),
    default=df['Month'].unique().tolist()
)

# 4. Apply all collected filters to the filtered DataFrame
if selected_regions:
    df_filtered = df_filtered[df_filtered['Region'].isin(selected_regions)]
if selected_categories:
    df_filtered = df_filtered[df_filtered['Category'].isin(selected_categories)]
if selected_suppliers:
    df_filtered = df_filtered[df_filtered['Supplier'].isin(selected_suppliers)]
if selected_items:
    df_filtered = df_filtered[df_filtered['Item'].isin(selected_items)]
if selected_months:
    df_filtered = df_filtered[df_filtered['Month'].isin(selected_months)]

# --- END NEW FILTER APPLICATION LOGIC ---

if df_filtered.empty and not df.empty:
    st.warning("No data matches the current filter selection. Please adjust your filters.")
    st.stop() # Stop further execution if no data matches filters

# --- 4. Key Performance Indicators (KPIs) ---
st.markdown("---")
col1, col2, col3 = st.columns(3)

if 'Sales Total' in df_filtered.columns:
    total_sales = df_filtered['Sales Total'].sum()
    col1.metric("Total Sales", f"₹{total_sales:,.2f}")

if 'Margin' in df_filtered.columns:
    total_margin = df_filtered['Margin'].sum()
    col2.metric("Total Margin", f"₹{total_margin:,.2f}")

if 'Sales Price' in df_filtered.columns and 'Sales Qty' in df_filtered.columns:
    sales_qty_sum = df_filtered['Sales Qty'].sum()
    avg_sales_price = (df_filtered['Sales Total'].sum() / sales_qty_sum) if sales_qty_sum > 0 else 0
    col3.metric("Avg. Sales Price", f"₹{avg_sales_price:.2f}")


# --- 5. Monthly Sales Trend (Time Series Chart) ---
st.markdown("---")
st.subheader("Time Series Trend")

if 'Month' in df_filtered.columns and 'Sales Total' in df_filtered.columns:
    sales_trend = df_filtered.groupby('Month')['Sales Total'].sum().reset_index()

    fig_line = px.line(
        sales_trend,
        x='Month',
        y='Sales Total',
        title='Monthly Sales Total Trend',
        markers=True,
        template='seaborn'
    )
    fig_line.update_traces(
        mode='lines+markers+text', 
        text=sales_trend['Sales Total'].round(0), 
        texttemplate='₹%{text:,.0f}',           
        textposition='top center'
    )
    st.plotly_chart(fig_line, use_container_width=True)
else:
    st.warning("Required columns ('Month', 'Sales Total') for Monthly Sales Trend are missing.")


# --- 6. Bar Charts (Distribution & Performance) ---
st.markdown("---")
st.subheader("Distribution and Performance Analysis")
chart_col1, chart_col2 = st.columns(2)

# Bar Chart 1: Sales Total by Category
with chart_col1:
    if 'Category' in df_filtered.columns and 'Sales Total' in df_filtered.columns:
        sales_by_category = df_filtered.groupby('Category')['Sales Total'].sum().reset_index()

        fig_bar = px.bar(
            sales_by_category,
            x='Category',
            y='Sales Total',
            title='Sales Total by Product Category',
            color='Category',
            template='seaborn'
        )
        fig_bar.update_traces(
            text=sales_by_category['Sales Total'].round(0), 
            texttemplate='₹%{text:,.0f}', 
            textposition='outside'
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.warning("Required columns ('Category', 'Sales Total') for Category Sales Bar Chart are missing.")

# Bar Chart 2: Top 10 Items by Margin
with chart_col2:
    if 'Item' in df_filtered.columns and 'Margin' in df_filtered.columns:
        top_items_margin = df_filtered.groupby('Item')['Margin'].sum().nlargest(10).reset_index()
        
        fig_top_items = px.bar(
            top_items_margin,
            x='Margin',
            y='Item',
            title='Top 10 Items by Margin',
            orientation='h', 
            color='Margin',
            color_continuous_scale=px.colors.sequential.Plotly3,
            template='seaborn'
        )
        fig_top_items.update_layout(yaxis={'categoryorder':'total ascending'})
        
        fig_top_items.update_traces(
            text=top_items_margin['Margin'].round(0), 
            texttemplate='₹%{text:,.0f}', 
            textposition='outside'
        )

        st.plotly_chart(fig_top_items, use_container_width=True)
    else:
        st.warning("Required columns ('Item', 'Margin') for Top 10 Items chart are missing.")


# --- 7. Regional Comparison Chart ---
st.markdown("---")
st.subheader("Regional Performance Comparison (Sales vs. Margin)")

if 'Region' in df_filtered.columns and 'Sales Total' in df_filtered.columns and 'Margin' in df_filtered.columns:
    regional_summary = df_filtered.groupby('Region').agg(
        {'Sales Total': 'sum', 'Margin': 'sum'}
    ).reset_index()
    
    regional_summary_melted = regional_summary.melt(
        id_vars='Region',
        value_vars=['Sales Total', 'Margin'],
        var_name='Metric',
        value_name='Amount'
    )

    fig_comparison = px.bar(
        regional_summary_melted,
        x='Region',
        y='Amount',
        color='Metric',
        barmode='group', 
        title='Sales Total vs. Margin by Region',
        template='seaborn'
    )
    
    fig_comparison.update_traces(
        text=regional_summary_melted['Amount'].round(0),
        texttemplate='₹%{text:,.0f}',
        textposition='outside'
    )
    
    st.plotly_chart(fig_comparison, use_container_width=True)
else:
    st.warning("Required columns ('Region', 'Sales Total', 'Margin') for Regional Comparison Chart are missing.")


# --- 8. Pie Charts (Contribution Analysis) ---
st.markdown("---")
st.subheader("Contribution Analysis (Click a slice to filter all charts!)")
pie_col1, pie_col2 = st.columns(2)

# --- NEW SECTION 8.5: CHART INTERACTION LOGIC ---

# PIE CHART 1: Region Wise Sales % (Clickable)
with pie_col1:
    if 'Region' in df_filtered.columns and 'Sales Total' in df_filtered.columns:
        # Use the original unfiltered data (df) here so the pie chart is static and always shows the full distribution.
        region_sales = df.groupby('Region')['Sales Total'].sum().reset_index()

        fig_pie_region_sales = px.pie(
            region_sales,
            values='Sales Total',
            names='Region',
            title='Region Wise Sales Percentage (Click to Filter)',
            template='seaborn',
            hole=0.3 
        )
        fig_pie_region_sales.update_traces(textinfo='label+percent')
        
        # CAPTURE CLICK EVENT
        selected_region = plotly_events(
            fig_pie_region_sales,
            override_height=fig_pie_region_sales.layout.height,
            key="region_pie_click"
        )
        
        # Update session state based on click
        if selected_region:
            # Use 'pointIndex' to reliably map back to the 'Region' name in the DataFrame
            if selected_region[0] and 'pointIndex' in selected_region[0]:
                point_index = selected_region[0]['pointIndex']
                # Retrieve the actual region name from the DataFrame used to generate the chart
                clicked_region_name = region_sales.iloc[point_index]['Region']
                
                # Logic to apply/clear filter
                if not st.session_state.region_click_filter or st.session_state.region_click_filter[0] != clicked_region_name:
                    st.session_state.region_click_filter = [clicked_region_name]
                    st.rerun() # Rerun the script to apply filter
                elif st.session_state.region_click_filter[0] == clicked_region_name:
                     # Clear filter if the same slice is clicked again
                     reset_chart_filters()
                     st.rerun()

    else:
        st.warning("Required columns ('Region', 'Sales Total') for Region Sales Pie Chart are missing.")

# PIE CHART 2: Category Wise Margin (Clickable)
with pie_col2:
    if 'Category' in df_filtered.columns and 'Margin' in df_filtered.columns:
        # Use the original unfiltered data (df) here so the pie chart is static and always shows the full distribution.
        category_margin = df.groupby('Category')['Margin'].sum().reset_index()

        fig_pie_category_margin = px.pie(
            category_margin,
            values='Margin',
            names='Category',
            title='Category Wise Margin Distribution (Click to Filter)',
            template='seaborn',
            hole=0.3
        )
        fig_pie_category_margin.update_traces(textinfo='label+percent')
        
        # CAPTURE CLICK EVENT
        selected_category = plotly_events(
            fig_pie_category_margin,
            override_height=fig_pie_category_margin.layout.height,
            key="category_pie_click"
        )
        
        # Update session state based on click
        if selected_category:
            # Use 'pointIndex' to reliably map back to the 'Category' name in the DataFrame
            if selected_category[0] and 'pointIndex' in selected_category[0]:
                point_index = selected_category[0]['pointIndex']
                # Retrieve the actual category name from the DataFrame used to generate the chart
                clicked_category_name = category_margin.iloc[point_index]['Category']

                # Logic to apply/clear filter
                if not st.session_state.category_click_filter or st.session_state.category_click_filter[0] != clicked_category_name:
                    st.session_state.category_click_filter = [clicked_category_name]
                    st.rerun() # Rerun the script to apply filter
                elif st.session_state.category_click_filter[0] == clicked_category_name:
                     # Clear filter if the same slice is clicked again
                     reset_chart_filters()
                     st.rerun()

    else:
        st.warning("Required columns ('Category', 'Margin') for Category Margin Pie Chart are missing.")

# --- END CHART INTERACTION LOGIC ---


# --- 9. Data Table Preview ---
st.markdown("---")
st.subheader("Raw Data Preview")
st.dataframe(df_filtered)

# --- How to Run This Script ---
# IMPORTANT: This update requires a new library for event handling!
# 1. Install the new library:
#    pip install streamlit-plotly-events
# 2. Save this code as 'dashboard_app.py'.
# 3. Run the application from your command prompt:
#    streamlit run dashboard_app.py
