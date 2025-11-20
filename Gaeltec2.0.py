# dashboard_mapped.py
import streamlit as st
import pandas as pd
import plotly.express as px
import re
import geopandas as gpd
import pydeck as pdk
import os
import glob
from PIL import Image
from io import BytesIO
import base64
from streamlit_plotly_events import plotly_events

# --- Page config for wide layout ---
st.set_page_config(
    page_title="Gaeltec Dashboard",
    layout="wide",  # <-- makes the dashboard wider
    initial_sidebar_state="expanded"
)

def sanitize_sheet_name(name: str) -> str:

    name = str(name)
    # Replace any invalid character with underscore
    name = re.sub(r'[^A-Za-z0-9 _-]', '_', name)
    # Truncate to 31 chars
    return name[:31]

# --- MAPPINGS ---

# --- Project Manager Mapping ---
project_mapping = {
    "Jonathon Mcclung": ["Ayrshire", "PCB"],
    "Gary MacDonald": ["Ayrshire", "LV"],
    "Jim Gaffney": ["Lanark", "PCB"],
    "Calum Thomson": ["Ayrshire", "Connections"],
    "Calum Thomsom": ["Ayrshire", "Connections"],
    "Calum Thompson": ["Ayrshire", "Connections"],
    "Andrew Galt": ["Ayrshire", "-"],
    "Henry Gordon": ["Ayrshire", "-"],
    "Jonathan Douglas": ["Ayrshire", "11 kV"],
    "Jonathon Douglas": ["Ayrshire", "11 kV"],
    "Matt": ["Lanark", ""],
    "Lee Fraser": ["Ayrshire", "Connections"],
    "Lee Frazer": ["Ayrshire", "Connections"],
    "Mark": ["Lanark", "Connections"],
    "Mark Nicholls": ["Ayrshire", "Connections"],
    "Cameron Fleming": ["Lanark", "Connections"],
    "Ronnie Goodwin": ["Lanark", "Connections"],
    "Ian Young": ["Ayrshire", "Connections"],
    "Matthew Watson": ["Lanark", "Connections"],
    "Aileen Brese": ["Ayrshire", "Connections"],
    "Mark McGoldrick": ["Lanark", "Connections"]
}

# --- Region Mapping ---
mapping_region = {
    "Newmilns": ["Irvine Valley"],
    "New Cumnock": ["New Cumnock"],
    "Kilwinning": ["Kilwinning"],
    "Stewarton": ["Irvine Valley"],
    "Kilbirnie": ["Kilbirnie and Beith"],
    "Coylton": ["Ayr East"],
    "Irvine": ["Irvine Valley", "Irvine East", "Irvine West"],
    "TROON": ["Troon"],
    "Ayr": ["Ayr East", "Ayr North", "Ayr West"],
    "Maybole": ["Maybole, North Carrick and Coylton"],
    "Clerkland": ["Irvine Valley"],
    "Glengarnock": ["Kilbirnie and Beith"],
    "Ayrshire": ["North Coast and Cumbraes","Prestwick", "Saltcoats and Stevenston", "Troon", "Ayr East", "Ayr North",
                 "Ayr West","Annick","Ardrossan and Arran","Dalry and West Kilbride","Girvan and South Carrick","Irvine East",
                 "Irvine Valley","Irvine West","Kilbirnie and Beith","Kilmarnock East and Hurlford","Kilmarnock North",
                 "Kilmarnock South","Kilmarnock West and Crosshouse","Kilwinning","Kyle","Maybole, North Carrick and Coylton",
                 "Ayr, Carrick and Cumnock","East_Ayrshire","North_Ayrshre","South_Ayrshre","Doon Valley"],
    "Lanark": ["Abronhill, Kildrum and the Village","Airdrie Central","Airdrie North","Airdrie South","Avondale and Stonehouse",
               "Ballochmyle","Bellshill","Blantyre","Bothwell and Uddingston","Cambuslang East","Cambuslang West",
               "Clydesdale East","Clydesdale North","Clydesdale South","Clydesdale West","Coatbridge North and Glenboig",
               "Coatbridge South","Coatbridge West","Cumbernauld North","Cumbernauld South",
               "East Kilbride Central North","East Kilbride Central South","East Kilbride East","East Kilbride South",
               "East Kilbride West","Fortissat","Hamilton North and East","Hamilton South","Hamilton West and Earnock",
               "Mossend and Holytown","Motherwell North","Motherwell South East and Ravenscraig","Motherwell West",
               "Rutherglen Central and North","Rutherglen South","Strathkelvin","Thorniewood","Wishaw","Larkhall",
               "Airdrie and Shotts","Cumbernauld, Kilsyth and Kirkintilloch East","East Kilbride, Strathaven and Lesmahagow",
               "Lanark and Hamilton East","Motherwell and Wishaw","North_Lanarkshire","South_Lanarkshire"]
}

# --- File Project Mapping ---
file_project_mapping = {
    "33kv refurb": ["Ayrshire", "33kv Refurb"],
    "connections": ["Ayrshire", "Connections"],
    "storms": ["Ayrshire", "Storms"],
    "11kv refurb": ["Ayrshire", "11kv Refurb"],
    "aurs road": ["Ayrshire", "Aurs Road"],
    "spen labour": ["Ayrshire", "SPEN Labour"],
    "lvhi5": ["Ayrshire", "LV"],
    "pcb": ["Ayrshire", "PCB"],
    "lanark": ["Lanark", ""],
    "11kv refur": ["Lanark", "11kv Refurb"],
    "lv & esqcr": ["Lanark", "LV"],
    "11kv rebuilt": ["Lanark", "11kV Rebuilt"],
    "33kv rebuilt": ["Lanark", "33kV Rebuilt"]
}

# --- Pole Mappings (dictionary style, includes new additions) ---
pole_keys = {
    "9x220 BIOCIDE LV POLE": "9m B",
    "9x275 BIOCIDE LV POLE": "9s B",
    "9x220 CREOSOTE LV POLE": "9m",
    "9x275 CREOSOTE LV POLE": "9s",
    "9x220 HV SINGLE POLE": "9m",
    "9x275 HV SINGLE POLE": "9s",
    "9x295 HV SINGLE POLE": "9es",
    "9x315 HV SINGLE POLE": "9esp",
    "10x230 BIOCIDE LV POLE": "10m B",
    "10x230 HV SINGLE POLE": "10m",
    "10x285 BIOCIDE LV POLE": "10s B",
    "10x285 H POLE HV Creosote": "10s",
    "10x285 HV SINGLE POLE": "10s",
    "10x305 HV SINGLE POLE": "10es",
    "11x295 HV SINGLE POLE": "11s",
    "11x295 H POLE HV Creosote": "11s",
    "11x295 BIOCIDE LV POLE": "11sB",
    "12x250 BIOCIDE LV POLE": "12m B",
    "12x305 BIOCIDE LV POLE": "12s B",
    "12x250 CREOSOTE LV POLE": "12m",
    "12x305 CREOSOTE LV POLE": "12s",
    "12x305 H POLE HV Creosote":"12s",
    "12x250 HV SINGLE POLE": "12m",
    "12x305 HV SINGLE POLE": "12s",
    "12x325 HV SINGLE POLE": "12es",
    "12x345 HV SINGLE POLE": "12esp",
    "13x260 BIOCIDE LV POLE": "13m B",
    "13x320 BIOCIDE LV POLE": "13s B",
    "13x260 CREOSOTE LV POLE": "13m",
    "13x320 CREOSOTE LV POLE": "13s",
    "13x260 HV SINGLE POLE": "13m",
    "13x320 HV SINGLE POLE": "13s",
    "13x340 HV SINGLE POLE": "13es",
    "13x365 HV SINGLE POLE": "13esp",
    "14x275 BIOCIDE LV POLE": "14m B",
    "14x335 BIOCIDE LV POLE": "14s B",
    "14x275 CREOSOTE LV POLE": "14m",
    "14x335 CREOSOTE LV POLE": "14s",
    "14x275 HV SINGLE POLE": "14m",
    "14x335 HV SINGLE POLE": "14s",
    "14x355 HV SINGLE POLE": "14es",
    "14x375 HV SINGLE POLE": "14esp",
    "16x305 BIOCIDE LV POLE": "16m B",
    "16x365 BIOCIDE LV POLE": "16s B",
    "16x305 CREOSOTE LV POLE": "16m",
    "16x365 CREOSOTE LV POLE": "16s",
    "16x305 HV SINGLE POLE": "16m",
    "16x365 HV SINGLE POLE": "16s",
    "16x385 HV SINGLE POLE": "16es",
    "16x405 HV SINGLE POLE": "16esp",
}

# --- Equipment / Conductor Mappings ---
equipment_keys = {
    "Hazel - 50mm² AAAC bare (1000m drums)": "Hazel 50mm²",
    "Oak - 100mm² AAAC bare (1000m drums)": "Oak 100mm²",
    "Ash - 150mm² AAAC bare (1000m drums)": "Ash 150mm²",
    "Poplar - 200mm² AAAC bare (1000m drums)": "Poplar 200mm²",
    "Upas - 300mm² AAAC bare (1000m drums)": "Upas 300mm²",
    "Poplar OPPC - 200mm² AAAC equivalent bare": "Poplar OPPC 200mm²",
    "Upas OPPC - 300mm² AAAC equivalent bare": "Upas OPPC 300mm²",
    # ACSR
    "Gopher - 25mm² ACSR bare (1000m drums)": "Gopher 25mm²",
    "Caton - 25mm² Compacted ACSR bare (1000m drums)": "Caton 25mm²",
    "Rabbit - 50mm² ACSR bare (1000m drums)": "Rabbit 50mm²",
    "Wolf - 150mm² ACSR bare (1000m drums)": "Wolf 150mm²",
    "Horse - 70mm² ACSR bare": "Horse 70mm²",
    "Dog - 100mm² ACSR bare (1000m drums)": "Dog 100mm²",
    "Dingo - 150mm² ACSR bare (1000m drums)": "Dingo 150mm²",
    # Copper
    "Hard Drawn Copper 16mm² ( 3/2.65mm ) (500m drums)": "Copper 16mm²",
    "Hard Drawn Copper 32mm² ( 3/3.75mm ) (1000m drums)": "Copper 32mm²",
    "Hard Drawn Copper 70mm² (500m drums)": "Copper 70mm²",
    "Hard Drawn Copper 100mm² (500m drums)": "Copper 100mm²",
    # PVC covered
    "35mm² Copper (Green / Yellow PVC covered) (50m drums)": "Copper 35mm² GY PVC",
    "70mm² Copper (Green / Yellow PVC covered) (50m drums)": "Copper 70mm² GY PVC",
    "35mm² Copper (Blue PVC covered) (50m drums)": "Copper 35mm² Blue PVC",
    "70mm² Copper (Blue PVC covered) (50m drums)": "Copper 70mm² Blue PVC",
    # Double insulated
    "35mm² Double Insulated (Brown) (50m drums)": "Double Insulated 35mm² Brown",
    "35mm² Double Insulated (Blue) (50m drums)": "Double Insulated 35mm² Blue",
    "70mm² Double Insulated (Brown) (50m drums)": "Double Insulated 70mm² Brown",
    "70mm² Double Insulated (Blue) (50m drums)": "Double Insulated 70mm² Blue",
    "120mm² Double Insulated (Brown) (50m drums)": "Double Insulated 120mm² Brown",
    "120mm² Double Insulated (Blue) (50m drums)": "Double Insulated 120mm² Blue",
    # LV cables
    "LV Cable 1ph 4mm Concentric (250m drums)": "LV 1ph 4mm Concentric",
    "LV Cable 1ph 25mm CNE (250m drums)": "LV 1ph 25mm CNE",
    "LV Cable 1ph 25mm SNE (100m drums)": "LV 1ph 25mm SNE",
    "LV Cable 1ph 35mm CNE (250m drums)": "LV 1ph 35mm CNE",
    "LV Cable 1ph 35mm SNE (100m drums)": "LV 1ph 35mm SNE",
    "LV Cable 3ph 35mm Cu Split Con (250m drums)": "LV 3ph 35mm Cu Split Con",
    "LV Cable 3ph 35mm SNE (250m drums)": "LV 3ph 35mm SNE",
    "LV Cable 3ph 35mm CNE (250m drums)": "LV 3ph 35mm CNE",
    "LV Cable 3ph 35mm CNE Al (LSOH) (250m drums)": "LV 3ph 35mm CNE Al LSOH",
    "LV Cable 3c 95mm W/F (250m drums)": "LV 3c 95mm W/F",
    "LV Cable 3c 185mm W/F (250m drums)": "LV 3c 185mm W/F",
    "LV Cable 3c 300mm W/F (250m drums)": "LV 3c 300mm W/F",
    "LV Cable 4c 95mm W/F (250m drums)": "LV 4c 95mm W/F",
    "LV Cable 4c 185mm W/F (250m drums)": "LV 4c 185mm W/F",
    "LV Cable 4c 240mm W/F (250m drums)": "LV 4c 240mm W/F",
    "LV Marker Tape (365m roll)": "LV Marker Tape",
    # 11kV
    "11kv Cable 95mm 3c Poly (250m drums)": "11kV 3c 95mm Poly",
    "11kv Cable 185mm 3c Poly (250m drums)": "11kV 3c 185mm Poly",
    "11kv Cable 300mm 3c Poly (250m drums)": "11kV 3c 300mm Poly",
    "11kv Cable 95mm 1c Poly (250m drums)": "11kV 1c 95mm Poly",
    "11kv Cable 185mm 1c Poly (250m drums)": "11kV 1c 185mm Poly",
    "11kv Cable 300mm 1c Poly (250m drums)": "11kV 1c 300mm Poly",
    "11kV Marker Tape (40m roll)": "11kV Marker Tape"
}

# --- Transformer Mappings ---
transformer_keys = {
    "Transformer 1ph 50kVA": "TX 1ph (50kVA)",
    "Transformer 3ph 50kVA": "TX 3ph (50kVA)",
    "Transformer 1ph 100kVA": "TX 1ph (100kVA)",
    "Transformer 1ph 25kVA": "TX 1ph (25kVA)",
    "Transformer 3ph 200kVA": "TX 3ph (200kVA)",
    "Transformer 3ph 100kVA": "TX 3ph (100kVA)"
}

# --- Gradient background ---
gradient_bg = """
<style>
    .stApp {
        background: linear-gradient(
            90deg,
            rgba(41, 28, 66, 1) 10%, 
            rgba(36, 57, 87, 1) 35%
        );
        color: white;
    }
</style>
"""
st.markdown(gradient_bg, unsafe_allow_html=True)

# --- Load logos ---
logo_left = Image.open(r"C:\Users\Xavier.Mascarenhas\OneDrive - Gaeltec Utilities Ltd\Desktop\Gaeltec\06_Programs\Dashboard\Images\GaeltecImage.png").resize((80, 80))
logo_right = Image.open(r"C:\Users\Xavier.Mascarenhas\OneDrive - Gaeltec Utilities Ltd\Desktop\Gaeltec\06_Programs\Dashboard\Images\SPEN.png").resize((160, 80))

# --- Header layout ---
col1, col2, col3 = st.columns([1, 4, 1])
with col1: st.image(logo_left)
with col2: st.markdown("<h1 style='text-align:center; margin:0;'>Gaeltec Utilities.UK</h1>", unsafe_allow_html=True)
with col3: st.image(logo_right)
st.markdown("<h1>📊 Data Management Dashboard</h1>", unsafe_allow_html=True)

# -------------------------------
# --- File Upload & Initial DF ---
# -------------------------------
# --- Upload Aggregated Parquet file ---
aggregated_file = st.file_uploader("Upload aggregated Parquet file", type=["parquet"])
if aggregated_file is not None:
    df = pd.read_parquet(aggregated_file)
    df.columns = df.columns.str.strip().str.lower()  # normalize columns

    if 'datetouse' in df.columns:
        df['datetouse'] = pd.to_datetime(df['datetouse'], errors='coerce')
        df = df.dropna(subset=['datetouse'])
        df['datetouse'] = df['datetouse'].dt.normalize()

# --- Upload Resume Parquet file (for %Complete pie chart) ---
resume_file = st.file_uploader("Upload resume Parquet file", type=["parquet"])
if resume_file is not None:
    resume_df = pd.read_parquet(resume_file)
    resume_df.columns = resume_df.columns.str.strip().str.lower()  # normalize columns

    # -------------------------------
    # --- Sidebar Filters ---
    # -------------------------------
    st.sidebar.header("Filter Options")

    def multi_select_filter(col_name, label, df, parent_filter=None):
        """Helper for multiselect filter, handles 'All' selection."""
        if col_name not in df.columns:
            return ["All"], df
        temp_df = df.copy()
        if parent_filter is not None and "All" not in parent_filter[1]:
            temp_df = temp_df[temp_df[parent_filter[0]].isin(parent_filter[1])]
        options = ["All"] + sorted(temp_df[col_name].dropna().unique())
        selected = st.sidebar.multiselect(label, options, default=["All"])
        if "All" not in selected:
            temp_df = temp_df[temp_df[col_name].isin(selected)]
        return selected, temp_df

    selected_shire, filtered_df = multi_select_filter('shire', "Select Shire", df)
    selected_project, filtered_df = multi_select_filter('project', "Select Project", filtered_df,
                                                        parent_filter=('shire', selected_shire))
    selected_pm, filtered_df = multi_select_filter('projectmanager', "Select Project Manager", filtered_df,
                                                   parent_filter=('shire', selected_shire))
    selected_segment, filtered_df = multi_select_filter('segmentcode', "Select Segment Code", filtered_df)
    selected_type, filtered_df = multi_select_filter('type', "Select Type", filtered_df)

    # -------------------------------
    # --- Date Filter ---
    # -------------------------------
    filter_type = st.sidebar.selectbox("Filter by Date", ["Single Day", "Week", "Month", "Year", "Custom Range"])
    date_range_str = ""
    if 'datetouse' in filtered_df.columns:
        if filter_type == "Single Day":
            date_selected = st.sidebar.date_input("Select date")
            filtered_df = filtered_df[filtered_df['datetouse'] == pd.Timestamp(date_selected)]
            date_range_str = str(date_selected)
        elif filter_type == "Week":
            week_start = st.sidebar.date_input("Week start date")
            week_end = week_start + pd.Timedelta(days=6)
            filtered_df = filtered_df[(filtered_df['datetouse'] >= pd.Timestamp(week_start)) &
                                      (filtered_df['datetouse'] <= pd.Timestamp(week_end))]
            date_range_str = f"{week_start} to {week_end}"
        elif filter_type == "Month":
            month_selected = st.sidebar.date_input("Pick any date in month")
            filtered_df = filtered_df[(filtered_df['datetouse'].dt.month == month_selected.month) &
                                      (filtered_df['datetouse'].dt.year == month_selected.year)]
            date_range_str = month_selected.strftime("%B %Y")
        elif filter_type == "Year":
            year_selected = st.sidebar.number_input("Select year", min_value=2000, max_value=2100, value=2025)
            filtered_df = filtered_df[filtered_df['datetouse'].dt.year == year_selected]
            date_range_str = str(year_selected)
        elif filter_type == "Custom Range":
            start_date = st.sidebar.date_input("Start date")
            end_date = st.sidebar.date_input("End date")
            filtered_df = filtered_df[(filtered_df['datetouse'] >= pd.Timestamp(start_date)) &
                                      (filtered_df['datetouse'] <= pd.Timestamp(end_date))]
            date_range_str = f"{start_date} to {end_date}"

    # -------------------------------
    # --- Total & Variation Display ---
    # -------------------------------
    total_sum, variation_sum = 0, 0
    if 'total' in filtered_df.columns:
        total_series = pd.to_numeric(filtered_df['total'].astype(str).str.replace(" ", "").str.replace(",", ".", regex=False),
                                     errors='coerce')
        total_sum = total_series.sum(skipna=True)
        if 'orig' in filtered_df.columns:
            orig_series = pd.to_numeric(filtered_df['orig'].astype(str).str.replace(" ", "").str.replace(",", ".", regex=False),
                                        errors='coerce')
            variation_sum = (total_series - orig_series).sum(skipna=True)

    formatted_total = f"{total_sum:,.2f}".replace(",", " ").replace(".", ",")
    formatted_variation = f"{variation_sum:,.2f}".replace(",", " ").replace(".", ",")

    # Money logo
    money_logo_path = r"C:\Users\Xavier.Mascarenhas\OneDrive - Gaeltec Utilities Ltd\Desktop\Gaeltec\06_Programs\Dashboard\Images\Pound.png"
    money_logo = Image.open(money_logo_path).resize((40, 40))
    buffered = BytesIO()
    money_logo.save(buffered, format="PNG")
    money_logo_base64 = base64.b64encode(buffered.getvalue()).decode()

    # Display Total & Variation
    col_top_left, col_top_right = st.columns([1, 1])
    with col_top_left:
        st.markdown(
            f"""
            <div style='display:flex; flex-direction:column; gap:4px;'>
                <div style='display:flex; align-items:center; gap:10px;'>
                    <h2 style='color:#32CD32; margin:0; font-size:36px;'><b>Total:</b> {formatted_total}</h2>
                    <img src='data:image/png;base64,{money_logo_base64}' width='40' height='40'/>
                </div>
                <div style='display:flex; align-items:center; gap:8px;'>
                    <h2 style='color:#32CD32; font-size:25px; margin:0;'><b>Variation:</b> {formatted_variation}</h2>
                    <img src='data:image/png;base64,{money_logo_base64}' width='28' height='28'/>
                </div>
                <p style='text-align:left; font-size:14px; margin-top:4px;'>
                    ({date_range_str}, Shires: {selected_shire}, Projects: {selected_project}, PMs: {selected_pm})
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )
    with col_top_right:
        st.markdown("<h3 style='text-align:center; color:white;'>Works Complete </h3>", unsafe_allow_html=True)



        # --- Top-right Pie Chart: % Complete ---
        try:
            # Ensure resume_df exists
            if 'resume_df' in locals():

                # Normalize both columns to lowercase strings without extra spaces
                filtered_segments = filtered_df['segment'].dropna().astype(str).str.strip().str.lower().unique()
                resume_df['section'] = resume_df['section'].dropna().astype(str).str.strip().str.lower()

                # Check if necessary columns exist in resume_df
                if {'section', '%complete'}.issubset(resume_df.columns):

                    # Filter resume to only include relevant sections
                    resume_filtered = resume_df[resume_df['section'].isin(filtered_segments)]

                    if not resume_filtered.empty:
                        avg_complete = resume_filtered['%complete'].mean()
                        avg_complete = min(max(avg_complete, 0), 100)  # clamp 0-100

                        # Pie chart data
                        pie_data = pd.DataFrame({
                            'Status': ['Completed', 'Done or Remaining'],
                            'Value': [avg_complete, 100 - avg_complete]
                        })

                        # Plot pie chart
                        fig_pie = px.pie(
                            pie_data,
                            names='Status',
                            values='Value',
                            color='Status',
                            color_discrete_map={'Completed': 'green', 'Done or Remaining': 'red'},
                            hole=0.6
                        )
                        fig_pie.update_traces(
                            textinfo='percent+label',
                            textfont_size=20
                        )
                        fig_pie.update_layout(
                            title_text="",
                            title_font_size=20,
                            font=dict(color='white'),
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(0,0,0,0)',
                            showlegend=True,
                            legend=dict(font=dict(color='white'))
                        )

                        # Display in top-right column
                        if 'col_top_right' in locals():
                            col_top_right.plotly_chart(fig_pie, use_container_width=True)
                        else:
                            st.plotly_chart(fig_pie, use_container_width=True)

                    else:
                        st.info("No matching sections found for the selected filters to generate % completion chart.")

        except Exception as e:
            st.warning(f"Could not generate % Complete pie chart: {e}")


    # -------------------------------
    # --- Map Section ---
    # -------------------------------
    col_map, col_desc = st.columns([2, 1])
    with col_map:
        st.header("🗺️ Regional Map View")
        folder_path = r"C:\Users\Xavier.Mascarenhas\OneDrive - Gaeltec Utilities Ltd\Desktop\Gaeltec\06_Programs\Dashboard\Maps"
        file_list = glob.glob(os.path.join(folder_path, "*.json"))

        if not file_list:
            st.error(f"No JSON files found in folder: {folder_path}")
        else:
            gdf_list = [gpd.read_file(file) for file in file_list]
            combined_gdf = gpd.GeoDataFrame(pd.concat(gdf_list, ignore_index=True), crs=gdf_list[0].crs)

            if "region" in filtered_df.columns:
                active_regions = filtered_df["region"].dropna().unique().tolist()
                wards_to_select = []
                for region in active_regions:
                    if region in mapping_region:
                        wards_to_select.extend(mapping_region[region])
                    else:
                        wards_to_select.append(region)
                wards_to_select = list(set(wards_to_select))
                areas_of_interest = combined_gdf[combined_gdf["WD13NM"].isin(wards_to_select)]
            else:
                areas_of_interest = pd.DataFrame()

            if not areas_of_interest.empty:
                areas_of_interest["geometry_simplified"] = areas_of_interest.geometry.simplify(tolerance=0.01)
                centroid = areas_of_interest.geometry_simplified.centroid.unary_union.centroid

                # Red flag
                flag_data = pd.DataFrame({"lon": [centroid.x], "lat": [centroid.y], "icon_name": ["red_flag"]})
                icon_mapping = {
                    "red_flag": {
                        "url": "https://upload.wikimedia.org/wikipedia/commons/thumb/3/3e/Red_flag_icon.svg/128px-Red_flag_icon.png",
                        "width": 128, "height": 128, "anchorY": 128
                    }
                }

                polygon_layer = pdk.Layer(
                    "GeoJsonLayer",
                    areas_of_interest["geometry_simplified"].__geo_interface__,
                    stroked=True,
                    filled=True,
                    get_fill_color=[160, 120, 80, 200],
                    get_line_color=[0, 0, 0],
                    pickable=True
                )

                flag_layer = pdk.Layer(
                    "IconLayer",
                    data=flag_data,
                    get_icon="icon_name",
                    get_size=4,
                    size_scale=15,
                    get_position='[lon, lat]',
                    pickable=True,
                    icon_mapping=icon_mapping
                )

                view_state = pdk.ViewState(latitude=centroid.y, longitude=centroid.x, zoom=8, pitch=0)

                st.pydeck_chart(
                    pdk.Deck(
                        layers=[polygon_layer, flag_layer],
                        initial_view_state=view_state,
                        map_style="mapbox://styles/mapbox/outdoors-v11"
                    )
                )
            else:
                st.info("No matching regions found for the selected filters.")


    with col_desc:
        st.markdown("<h3 style='color:white;'>Projects & Segments Overview</h3>", unsafe_allow_html=True)

        if 'project' in filtered_df.columns and 'segmentcode' in filtered_df.columns:
            projects = filtered_df['project'].dropna().unique()
            if len(projects) == 0:
                st.info("No projects found for the selected filters.")
            else:
                for proj in sorted(projects):
                    segments = filtered_df[filtered_df['project'] == proj]['segmentcode'].dropna().unique()
                
                    # Use expander to make segment list scrollable
                    with st.expander(f"Project: {proj} ({len(segments)} segments)"):
                        if len(segments) > 0:
                            # Scrollable container for segments
                            st.markdown(
                                "<div style='max-height:150px; overflow-y:auto; padding:5px; border:1px solid #444;'>"
                                + "<br>".join(segments.astype(str))
                                + "</div>",
                                unsafe_allow_html=True
                            )
                        else:
                            st.write("No segment codes for this project.")
        else:
            st.info("Project or Segment Code columns not found in the data.")

    # -------------------------------
    # --- Mapping Bar Charts + Drill-down + Excel Export ---
    # -------------------------------
    st.header("📊 Mapping Charts")
    convert_to_miles = st.checkbox("Convert Equipment/Conductor Length to Miles")

    categories = [
        ("Poles", pole_keys, "Quantity"),
        ("Equipment / Conductor", equipment_keys, "Length (Km)"),
        ("Transformers", transformer_keys, "Quantity")
    ]

    for cat_name, keys, y_label in categories:
        if 'item' in filtered_df.columns and 'mapped' in filtered_df.columns:
            pattern = '|'.join([re.escape(k) for k in keys.keys()])
            mask = filtered_df['item'].astype(str).str.contains(pattern, case=False, na=False)
            sub_df = filtered_df[mask]

            if not sub_df.empty:
                # Aggregate
                if 'qsub' in sub_df.columns:
                    sub_df['qsub_clean'] = pd.to_numeric(sub_df['qsub'].astype(str).str.replace(" ", "").str.replace(",", ".", regex=False), errors='coerce')
                    bar_data = sub_df.groupby('mapped')['qsub_clean'].sum().reset_index()
                    bar_data.columns = ['Mapped', 'Total']
                else:
                    bar_data = sub_df['mapped'].value_counts().reset_index()
                    bar_data.columns = ['Mapped', 'Total']

                y_axis_label = y_label
                if cat_name == "Equipment / Conductor" and convert_to_miles:
                    bar_data['Total'] = bar_data['Total'] * 0.621371
                    y_axis_label = "Length (Miles)"

                # Plot
                fig = px.bar(bar_data, x='Mapped', y='Total', color='Total', text='Total', title=f"{cat_name} Overview",
                             color_continuous_scale=['rgba(128,0,128,1)','rgba(147,112,219,1)','rgba(186,85,211,1)','rgba(221,160,221,1)'],
                             labels={'Mapped': 'Mapping', 'Total': y_axis_label})

                fig.update_layout(plot_bgcolor='rgba(0,0,0,1)', paper_bgcolor='rgba(0,0,0,1)', font=dict(color='white'), coloraxis_showscale=False)
                click = plotly_events(fig, click_event=True)
                st.plotly_chart(fig, use_container_width=True)

                # --- Drill-down + Excel Export ---
                if click:
                    clicked_mapping = click[0]["x"]
                    st.subheader(f"Details for: **{clicked_mapping}**")
                    selected_rows = sub_df[sub_df['mapped'] == clicked_mapping].copy()
                    selected_rows = selected_rows.loc[:, ~selected_rows.columns.duplicated()]
                    if 'datetouse' in selected_rows.columns:
                        selected_rows['datetouse'] = pd.to_datetime(selected_rows['datetouse'], errors='coerce').dt.date

                    extra_cols = ['pole', 'projectmanager', 'qsub', 'project', 'shire', 'segmentdesc', 'sourcefile']
                    display_cols = ['mapped', 'datetouse'] + extra_cols
                    st.dataframe(selected_rows[display_cols], use_container_width=True)

                    # --- Excel export ---
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        for bar_value in bar_data['Mapped']:
                            df_bar = sub_df[sub_df['mapped'] == bar_value].copy()
                            df_bar = df_bar.loc[:, ~df_bar.columns.duplicated()]
                            if 'datetouse' in df_bar.columns:
                                df_bar['datetouse'] = pd.to_datetime(df_bar['datetouse'], errors='coerce').dt.date
                            cols_to_include = ['mapped', 'datetouse'] + extra_cols
                            df_bar = df_bar[cols_to_include]

                            # Sanitize sheet name
                            sheet_name = sanitize_sheet_name(str(bar_value))
                            df_bar.to_excel(writer, sheet_name=sheet_name, index=False)

                    buffer.seek(0)
                    st.download_button(
                        f"📥 Download Excel: {cat_name} Details",
                        buffer,
                        file_name=f"{cat_name}_Details.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )