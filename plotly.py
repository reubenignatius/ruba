import pandas as pd
import plotly.express as px

# 1. Prepare your data (Example: Monthly Revenue Data)
data = {
    'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
    'Revenue': [15000, 18000, 16500, 20000, 22000],
    'Target': [17000, 17000, 17000, 19000, 19000]
}
df = pd.DataFrame(data)

# 2. Create the Plotly Figure object
revenue_fig = px.bar(
    df,
    x='Month',
    y='Revenue',
    title='Monthly Revenue Performance',
    # Customize the figure appearance
    color='Revenue',
    color_continuous_scale=px.colors.sequential.Plotly3
)

# Add a line for the target (demonstrates layering)
revenue_fig.add_scatter(
    x=df['Month'],
    y=df['Target'],
    mode='lines+markers',
    name='Budget Target',
    line=dict(color='red', width=3)
)