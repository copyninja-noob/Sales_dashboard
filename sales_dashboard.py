import pandas as pd
import streamlit as st
import os
from datetime import datetime
import plotly.graph_objects as go

# Load Excel file
excel_path = 'sales_data.xlsx'
if os.path.exists(excel_path):
    df = pd.read_excel(excel_path, engine='openpyxl')
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

    def get_fin_year(d):
        y = d.year
        return f"{y-1}-{str(y)[-2:]}" if d.month < 4 else f"{y}-{str(y+1)[-2:]}"
    
    df['Financial Year'] = df['Date'].apply(get_fin_year)

    def get_fin_week(d):
        fy_start = datetime(d.year if d.month >= 4 else d.year - 1, 4, 1)
        return ((d - fy_start).days // 7) + 1

    df['Week'] = df['Date'].apply(get_fin_week)
else:
    st.error("Error: sales_data.xlsx file not found!")
    st.stop()

def indian_format(x):
    try:
        x = int(x)
    except:
        return x
    s = str(x)[::-1]
    groups = [s[:3]]
    s = s[3:]
    while s:
        groups.append(s[:2])
        s = s[2:]
    return ','.join(g[::-1] for g in groups[::-1])

# Streamlit app
st.set_page_config(layout="wide")
st.title("📊 Weekly Sales Dashboard")

# Sidebar filters
with st.sidebar:
    st.header("Filters")
    fy_options = sorted(df['Financial Year'].unique(), reverse=True)  # Sort descending
    
    if len(fy_options) < 2:
        st.error("Need at least 2 financial years of data for comparison")
        st.stop()
    
    # First financial year (more recent)
    fy1 = st.selectbox(
        "Select First (More Recent) Financial Year", 
        fy_options,
        index=0  # Default to most recent
    )
    
    # Second financial year (older than first)
    available_fy2 = [fy for fy in fy_options if fy < fy1]
    if not available_fy2:
        st.error("No older financial years available for comparison")
        st.stop()
        
    fy2 = st.selectbox(
        "Select Second (Older) Financial Year", 
        available_fy2,
        index=0  # Default to most recent older year
    )

def build_card(dataframe, title, domain_filter=None):
    if domain_filter:
        filtered = dataframe[dataframe['Domain'] == domain_filter]
    else:
        filtered = dataframe.copy()

    # Check if we have data for the selected financial years
    if fy1 not in filtered['Financial Year'].unique() or fy2 not in filtered['Financial Year'].unique():
        return

    df_y1 = filtered[filtered['Financial Year'] == fy1].groupby('Week', as_index=False)['Revenue'].sum()
    df_y2 = filtered[filtered['Financial Year'] == fy2].groupby('Week', as_index=False)['Revenue'].sum()

    df_y1.rename(columns={'Revenue': f'Revenue_{fy1}'}, inplace=True)
    df_y2.rename(columns={'Revenue': f'Revenue_{fy2}'}, inplace=True)

    merged = pd.merge(df_y1, df_y2, on='Week', how='outer').fillna(0).sort_values('Week')

    rev1_col = f'Revenue_{fy1}'
    rev2_col = f'Revenue_{fy2}'
    total_rev2 = merged[rev2_col].sum()
    total_rev1 = merged[rev1_col].sum()

    merged['Variation (%)'] = merged.apply(
        lambda row: ((row[rev1_col] - row[rev2_col]) * 100 / total_rev2) if total_rev2 else 0,
        axis=1
    ).round(2)

    merged['Variation in Amount'] = (merged[rev1_col] - merged[rev2_col]).apply(indian_format)
    merged[f'{rev1_col}_formatted'] = merged[rev1_col].apply(indian_format)
    merged[f'{rev2_col}_formatted'] = merged[rev2_col].apply(indian_format)

    total_var = ((total_rev1 - total_rev2) * 100 / total_rev2) if total_rev2 else 0
    total_amount_variation = indian_format(total_rev1 - total_rev2)

    total_row = pd.DataFrame([{
        'Week': 'Total',
        f'{fy1} Revenue': indian_format(total_rev1),
        f'{fy2} Revenue': indian_format(total_rev2),
        'Variation (%)': round(total_var, 2),
        'Variation in Amount': total_amount_variation
    }])

    # Display the card - starts collapsed
    with st.expander(title, expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            # Summary table
            st.dataframe(
                total_row.style
                .apply(lambda x: ['background-color: #d1e7dd' if x['Variation (%)'] > 0 
                                 else 'background-color: #f8d7da' for i, v in enumerate(x)], 
                       axis=1)
                .format({'Variation (%)': "{:.2f}%"}),
                use_container_width=True,
                hide_index=True
            )

        with col2:
            # Key metrics
            st.metric(f"Total {fy1} Revenue", indian_format(total_rev1))
            st.metric(f"Total {fy2} Revenue", indian_format(total_rev2))
            st.metric("Total Variation", f"{total_var:.2f}%", 
                     delta_color="inverse" if total_var < 0 else "normal")

        # Cumulative trend chart
        merged[f'Cumulative_{rev1_col}'] = merged[rev1_col].cumsum()
        merged[f'Cumulative_{rev2_col}'] = merged[rev2_col].cumsum()

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=merged['Week'], y=merged[f'Cumulative_{rev1_col}'], 
                               mode='lines', name=fy1, line=dict(color='blue')))
        fig.add_trace(go.Scatter(x=merged['Week'], y=merged[f'Cumulative_{rev2_col}'], 
                               mode='lines', name=fy2, line=dict(color='orange')))
        fig.update_layout(
            title="Cumulative Weekly Sales Trend",
            xaxis_title="Week",
            yaxis_title="Revenue",
            height=300,
            template="plotly_white",
            legend=dict(x=0, y=1.1, orientation='h'),
            margin=dict(l=20, r=20, t=40, b=20),
            yaxis_tickformat=','
        )
        st.plotly_chart(fig, use_container_width=True)

        # Detailed table
        display_df = merged[['Week', f'{rev1_col}_formatted', f'{rev2_col}_formatted', 'Variation (%)', 'Variation in Amount']]
        display_df.columns = ['Week', f'{fy1} Revenue', f'{fy2} Revenue', 'Variation (%)', 'Variation in Amount']

        st.dataframe(
            display_df.style
            .applymap(lambda x: 'color: green' if isinstance(x, (int, float)) and x > 0 else 
                     ('color: red' if isinstance(x, (int, float)) and x < 0 else ''),
            subset=['Variation (%)']
            )
            .format({'Variation (%)': "{:.2f}%"}),
            use_container_width=True,
            hide_index=True
        )

# Main content
st.write("Click on the sections below to expand and view the sales data:")

# Define the exact order of domains as requested
DOMAIN_ORDER = [
    None,  # For "All Domains Combined"
    "Training",
    "Tech Alumni",
    "Whatsapp API Business",
    "G-Suite",
    "Tech Assist Recruitment",
    "Other",
    "Consulting"
]

# Create cards in the specified order
for domain in DOMAIN_ORDER:
    if domain is None:
        build_card(df, "🌍 All Domains Combined")
    else:
        # Only show domain if it exists in the data (silently skip if not)
        if domain in df['Domain'].unique():
            build_card(df, f"📂 {domain}", domain_filter=domain)

