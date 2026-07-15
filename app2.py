import streamlit as st
import pandas as pd

# 1. Page Setup
st.set_page_config(page_title="ProCalc | Digital QS", layout="wide")
st.title("Cost Analysis")
st.markdown("Interactive Bill of Quantities dashboard with real-time risk variance.")

# 2. Sidebar Controls (The "What-If" Scenario Builders)
st.sidebar.header("Market Risk Factors")
st.sidebar.markdown("Adjust the sliders to simulate market volatility.")
material_inflation = st.sidebar.slider("Material Inflation (%)", 0.0, 20.0, 5.0) / 100
labor_variance = st.sidebar.slider("Labor Productivity Variance (%)", -10.0, 20.0, 2.0) / 100

# 3. Core Data (In the future, we will link this to a Glodon TAS/TRB export parser)
if 'boq_data' not in st.session_state:
    st.session_state.boq_data = pd.DataFrame({
        "Item": ["Concrete C30", "High-Yield Rebar", "Formwork", "Internal Blockwork"],
        "Category": ["Substructure", "Substructure", "Substructure", "Superstructure"],
        "Unit": ["m3", "ton", "m2", "m2"],
        "Quantity": [150.0, 12.5, 450.0, 200.0],
        "Base_Rate": [85.0, 950.0, 12.0, 15.0] # Represented in USD for international portfolio
    })

st.subheader("1. Interactive Bill of Quantities")
st.info("💡 Edit the Quantities or Base Rates directly in the table below to see the forecast update.")

# 4. The Interactive Data Grid
edited_boq = st.data_editor(
    st.session_state.boq_data,
    column_config={
        "Category": st.column_config.SelectboxColumn(
            "Category", options=["Substructure", "Superstructure", "MEP", "Finishes"]
        ),
        "Base_Rate": st.column_config.NumberColumn("Base Rate ($)", format="$ %.2f"),
        "Quantity": st.column_config.NumberColumn("Quantity", format="%.2f")
    },
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic" # Allows users to add/delete rows on the fly
)

# 5. Calculation Engine
# Calculate adjusted rates based on the sidebar sliders
edited_boq["Adjusted_Rate"] = edited_boq["Base_Rate"] * (1 + material_inflation + labor_variance)
edited_boq["Total_Cost"] = edited_boq["Quantity"] * edited_boq["Adjusted_Rate"]

base_total = (edited_boq["Quantity"] * edited_boq["Base_Rate"]).sum()
forecast_total = edited_boq["Total_Cost"].sum()
variance = forecast_total - base_total

# 6. Top-Level Metrics
st.subheader("2. Commercial Summary")
col1, col2, col3 = st.columns(3)
col1.metric("Baseline Budget", f"${base_total:,.2f}")
col2.metric("Forecasted Cost", f"${forecast_total:,.2f}", f"${variance:,.2f} Overrun Risk", delta_color="inverse")

# 7. Visual Breakdown
st.subheader("3. Cost Distribution")
category_costs = edited_boq.groupby("Category")["Total_Cost"].sum().reset_index()
st.bar_chart(category_costs.set_index("Category"))