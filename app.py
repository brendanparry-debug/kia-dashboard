import streamlit as str
import pandas as pd
import io

# Set page config for a professional workspace layout
str.set_page_config(page_title="Kia Inventory Reconciliation Dashboard", layout="wide")

str.title("🚘 Kia Portal vs. Hub Website Reconciliation")
str.write("Upload your Excel reports below to find unentered inventory, ghost listings, and status mismatches.")

# Sidebar Instructions & Column references
with str.sidebar:
    str.header("📋 Expected Data Layout")
    str.subheader("1. Kia Portal (True Inventory)")
    str.caption("4 Columns: [Status, VIN #, Year, Description]")
    str.caption("Statuses: On order, On Water, Units in Transit, Dealer Stock")
    
    str.subheader("2. Hub (Website Inventory)")
    str.caption("6 Columns: [Stock #, VIN #, Year, Model, Trim, Stock Status]")
    str.caption("Stock Statuses: S (In Stock) or O (Inbound/Ordered)")

# 1. File Upload Fields
col_upload1, col_upload2 = str.columns(2)

with col_upload1:
    portal_file = str.file_uploader("Upload 'Kia Portal' Excel File", type=["xlsx", "xls"])

with col_upload2:
    hub_file = str.file_uploader("Upload 'Hub' Excel File", type=["xlsx", "xls"])

# 2. Processing Logic
if portal_file and hub_file:
    try:
        # Load sheets (Assuming data starts on first sheet, header on row 0)
        df_portal = pd.read_excel(portal_file)
        df_hub = pd.read_excel(hub_file)
        
        # Standardize column naming based on positional index to prevent naming syntax errors
        df_portal.columns = ['Status', 'VIN', 'Year', 'Description'] + list(df_portal.columns[4:])
        df_hub.columns = ['Stock_No', 'VIN', 'Year', 'Model', 'Trim', 'Stock_Status'] + list(df_hub.columns[6:])
        
        # Clean VIN spaces and make uppercase for flawless matching
        df_portal['VIN'] = df_portal['VIN'].astype(str).str.strip().str.upper()
        df_hub['VIN'] = df_hub['VIN'].astype(str).str.strip().str.upper()
        
        # Clean statuses
        df_portal['Status'] = df_portal['Status'].astype(str).str.strip()
        df_hub['Stock_Status'] = df_hub['Stock_Status'].astype(str).str.strip().str.upper()

        # Isolate VIN arrays
        portal_vins = set(df_portal['VIN'].unique())
        hub_vins = set(df_hub['VIN'].unique())

        # --- AUDIT CRITERIA 1: Missing from Hub (In Portal but not in Hub) ---
        missing_from_hub = df_portal[~df_portal['VIN'].isin(hub_vins)].copy()
        
        # --- AUDIT CRITERIA 2: Ghost Inventory (In Hub but not in Portal) ---
        ghost_inventory = df_hub[~df_hub['VIN'].isin(portal_vins)].copy()
        
        # --- AUDIT CRITERIA 3: Status Mismatches ---
        # Merge datasets on VIN to cross-reference columns directly
        merged = pd.merge(df_portal, df_hub, on='VIN', suffixes=('_portal', '_hub'))
        
        # Condition A: Dealer Stock in Portal, but marked "O" in Hub
        mismatch_dealer_stock = merged[(merged['Status'] == 'Dealer Stock') & (merged['Stock_Status'] == 'O')]
        
        # Condition B: Inbound (On Order, On Water, Units in Transit) in Portal, but marked "S" in Hub
        inbound_statuses = ['On order', 'On Water', 'Units in Transit']
        mismatch_inbound = merged[(merged['Status'].isin(inbound_statuses)) & (merged['Stock_Status'] == 'S')]
        
        # Combine all status anomalies
        status_mismatches = pd.concat([mismatch_dealer_stock, mismatch_inbound]).copy()
        
        # Clean up columns for the mismatch display
        if not status_mismatches.empty:
            status_mismatches = status_mismatches[[
                'VIN', 'Year_portal', 'Description', 'Status', 'Stock_No', 'Model', 'Trim', 'Stock_Status'
            ]].rename(columns={'Year_portal': 'Year', 'Status': 'Portal Status', 'Stock_Status': 'Hub Status'})

        # --- 3. DISPLAY DASHBOARD METRICS ---
        str.markdown("---")
        m_col1, m_col2, m_col3 = str.columns(3)
        m_col1.metric("🚨 Missing from Hub (Unentered)", len(missing_from_hub))
        m_col2.metric("👻 Ghost Inventory (Remove)", len(ghost_inventory))
        m_col3.metric("⚠️ Status Mismatches", len(status_mismatches))
        str.markdown("---")

        # --- 4. EXCEL EXPORT BUTTON CREATION ---
        # Create a buffer to save Excel workbook streams into memory
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            missing_from_hub.to_excel(writer, sheet_name='Missing from Hub', index=False)
            ghost_inventory.to_excel(writer, sheet_name='Ghost Listings (Remove)', index=False)
            status_mismatches.to_excel(writer, sheet_name='Status Mismatches', index=False)
        
        output_buffer.seek(0)
        
        # One-Click download widget
        str.download_button(
            label="📥 Download Flagged Errors (Excel File)",
            data=output_buffer,
            file_name="Kia_Inventory_Audit_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # --- 5. TABS FOR VISUALIZING ON-SCREEN ---
        tab1, tab2, tab3 = str.tabs([
            "🚨 Missing From Hub", 
            "👻 Ghost Listings (Needs Deletion)", 
            "⚠️ Status Mismatches"
        ])
        
        with tab1:
            str.subheader("Cars in Portal that aren't on the Website")
            if not missing_from_hub.empty:
                str.dataframe(missing_from_hub, use_container_width=True)
            else:
                str.success("Perfect! All factory inventory has been entered in the Hub.")
                
        with tab2:
            str.subheader("Cars on Website that are no longer in Portal allocation")
            if not ghost_inventory.empty:
                str.dataframe(ghost_inventory, use_container_width=True)
            else:
                str.success("Clean sweep! No stale/ghost cars found on the website.")
                
        with tab3:
            str.subheader("Status Misalignment Alerts")
            if not status_mismatches.empty:
                str.dataframe(status_mismatches, use_container_width=True)
                str.caption("**Note:** Rules flagged: 'Dealer Stock' cannot be status 'O'; 'On order/On Water/Units in Transit' cannot be status 'S'.")
            else:
                str.success("Awesome! All shipping phases match your website's 'S' and 'O' designations perfectly.")

    except Exception as e:
        str.error(f"Error parsing sheets. Please check that column ordering matches the required structure. Detailed Error: {e}")
else:
    str.info("💡 Please upload both your Kia Portal and Hub Excel files above to begin the audit.")
