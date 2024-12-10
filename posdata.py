# Import necessary libraries
import streamlit as st
import pandas as pd
import plotly.express as px
from google.oauth2.service_account import Credentials
import gspread
from io import BytesIO
from datetime import datetime, timedelta

# Streamlit configuration
st.set_page_config(
    page_title="POS Dashboard - Advanced Visualization",
    layout="wide",
    page_icon="🚌"
)

# Styled Title Section
st.markdown(f"""
    <div style="border: 4px solid #4CAF50; padding: 15px; border-radius: 15px; background-color: #E3F2FD; text-align: center;">
        <h1 style="color: #FF5722; font-family: 'Segoe UI', sans-serif; font-size: 45px; margin-bottom: 5px;">
            POS Dashboard - TNSTC[KUM]LTD., KUMBAKONAM
        </h1>
        <h2 style="color: #4CAF50; font-family: 'Arial', sans-serif; font-size: 36px; margin-bottom: 0;">
            Data as of {(datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")}
        </h2>
    </div>
    """, unsafe_allow_html=True)

# Branch-to-region mapping
branch_to_region = {
    "KM1": "Kumbakonam", "KM2": "Kumbakonam", "TYR": "Kumbakonam", "TJ1": "Kumbakonam", "TJ2": "Kumbakonam",
    "OND": "Kumbakonam", "PKT": "Kumbakonam", "PVR": "Kumbakonam",
    "RFT": "Trichy", "DCN": "Trichy", "TVK": "Trichy", "LAL": "Trichy", "MCR": "Trichy", "TMF": "Trichy",
    "CNT": "Trichy", "MNP": "Trichy", "TKI": "Trichy", "PBR": "Trichy", "JKM": "Trichy", "ALR": "Trichy",
    "UPM": "Trichy", "TRR": "Trichy", "KNM": "Trichy",
    "KR1": "Karur", "KR2": "Karur", "AKI": "Karur", "KLI": "Karur", "MSI": "Karur",
    "PDK": "Pudukottai", "ATQ": "Pudukottai", "ALD": "Pudukottai", "PTK": "Pudukottai", "TRY": "Pudukottai",
    "ILP": "Pudukottai", "GVK": "Pudukottai", "PON": "Pudukottai",
    "KKD": "Karaikudi", "TPR": "Karaikudi", "MDU": "Karaikudi", "SVG": "Karaikudi", "DVK": "Karaikudi",
    "RNM": "Karaikudi", "RNT": "Karaikudi", "RMM": "Karaikudi", "KMD": "Karaikudi", "MDK": "Karaikudi",
    "PMK": "Karaikudi",
    "NGT": "Nagappattinam", "KKL": "Nagappattinam", "PYR": "Nagappattinam", "MLD": "Nagappattinam",
    "SHY": "Nagappattinam", "CDM": "Nagappattinam", "TVR": "Nagappattinam", "NLM": "Nagappattinam",
    "MNG": "Nagappattinam", "TTP": "Nagappattinam"
}

# Function to load and preprocess data
def load_google_sheet(sheet_url, gid, json_keyfile_dict):
    """
    Load and preprocess data from a Google Sheet.
    """
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(json_keyfile_dict, scopes=scopes)
        gc = gspread.authorize(credentials)
        sheet = gc.open_by_url(sheet_url)
        worksheets = {ws.id: ws for ws in sheet.worksheets()}
        worksheet = worksheets.get(int(gid))
        if not worksheet:
            st.error(f"No worksheet found with GID {gid}.")
            return None

        raw_data = worksheet.get_all_values()
        if not raw_data or len(raw_data) < 2:
            st.error("The worksheet has insufficient data or is empty.")
            return None

        headers = raw_data[0]
        data_rows = raw_data[1:]
        data = pd.DataFrame(data_rows, columns=headers)
        data.dropna(how="all", inplace=True)
        data.replace("", None, inplace=True)
        data.columns = [col.strip() for col in data.columns]

        if "BRANCH" in data.columns:
            data["REGION"] = data["BRANCH"].map(branch_to_region)
        else:
            st.warning("BRANCH column not found in the data.")
            data["REGION"] = None

        numeric_cols = [
            "POS MOF Total", "POS MOF Issued", "MOF POS %", "POS TOWN Total",
            "POS TOWN  Issued", "TOWN POS %", "OVER ALL MOF+TOWN Total",
            "OVER ALL MOF+TOWN POS Issued", "OVER ALL MOF+TOWN POS %",
            "MOF POS Tickets", "MOF Pre Printed  Tickets", "TOWN POS Tickets", "TOWN Pre Printed  Tickets"
        ]
        for col in numeric_cols:
            if col in data.columns:
                data[col] = pd.to_numeric(data[col], errors="coerce")

        if "POS IMPLEMENTED DATE" in data.columns:
            data["POS IMPLEMENTED DATE"] = pd.to_datetime(data["POS IMPLEMENTED DATE"], errors="coerce")

        return data
    except Exception as e:
        st.error(f"Error loading Google Sheet: {e}")
        return None

google_credentials = {
    "type": "service_account",
    "project_id": "posdashboard-442201",
    "private_key_id": "78a8b4fe3f25a90176067dd6b46f45f2abbf87f8",
     "private_key": (
        "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDEtdSLGwNIJtXc\nCBCbf8L/oA8F9a4P+sBkz8lRZQcAPo2GPASPNXcg4GRNncGi0L2L2cMrGh9IYE6K\nH4SFvutG1Xe1WEbFi0U0/j+2V0ULpxENLPhJmnkBanMts6APd9WvPdxFagT7NUvq\nYIKGvwrpawsq5fwgjmuiKDWgzn7XaIB5duUwAReMbbXOfDoZKp7l3dp5gyXEzAxD\neMVhd2La7agblmzrRP0E+32ct4hPyLrGcBDkcwjG559Zmk+q5/2MbhDrIS0P5ux7\n+qftVUN5WwxE9UrJoZXHoSxKz96A8aHJMevK55zPEP1JjTR8LPsWNW3jl1IU8KyT\nLZyLz52dAgMBAAECggEAGMZmnW81XnDXDkONGpX1+y3Eu/VBoIKY7mQwkLSpiTws\n+qNJgbiep0CLxtjKgBVL94wMOuarL6RC5WjrJHszdR7tWPo5ePzoUQXVWW5WubfS\nJneA/VLcJbOrYQNWx/afA6zGDDoPpDdbYeAY4HFkUES/9H2138CASZKd5Sx3fkKo\nKcB497f7YpHg/mgN07wxhB29QpN35E5UB3XlAOdBGk0vIJ9V3Lvl/OHBc+EiXkzV\nqzQLI9kgIdOAWJ8o3NaGx6te2vy5TUzT0uWOL7/eOvDpI6UV9yy6K7xrg9XtSpu8\nY6DkUEMdCLlxVXcfM2+SZE5XiJ86E17njIuuLV5L4wKBgQD5jTeMqL7ptQV1t1J/\n9IvDg2jxuyJDFo77ifkklY3Sefe4j4ZNNv0799iVAmAMf1nUYk0bp+4Xns2AQjCx\nAKKYPOx6YfiaMorALkmrH5mXYX6o2Tgv0BnVNoifjuHPhSm2LEbzgUeDhAgEZN2i\nOJb684KJbhe15GPKEp44RDmcSwKBgQDJyxFiuerVXYri9c4mLxUJHyBWCN5Yib+H\nQj6UyN1r61girPMysC2jWJpZ4oU7nv/6IXwUBZQ3o4/fnxC02y3vRz3GFJkMtcqs\no+kfFgbKUqkH4nSsJBjZQW1DfC+xIIBdT8k7oD2/Ux119R/9NLNnD7iOvFKaC+7Y\n1RQHfMwstwKBgEsa4TkIIE0eGgKPpdi0tMum5RK7i1g9ldLGd6E3EXPjGVcGexkK\nD7TYpupRyK56NYLiAurr45BgTuDnCth6pHTFATbj/XoK9A9a3vkNjaAty3ztwydA\nrkWpH/1Fd1iJb0BQmxn2Mpu2RONtp/aGqYnld8f8xk4L6qyKZevxPJV5AoGBALSk\noM+sd1jCAI7kVMNB6qbbwmrCTakcxuQinTs8BVuStrdz89IwfOp5atOEQJj64VPd\nneGejOyx8x3Qm3gLrbdCIz6rOcdzBhg+M3aslS+Rh9eTFbb0KXpzY4jCJz99ROxD\nfHVwIVag5QKviQ92mhNss16zn45fmFVrih6ZzX1JAoGAaOChj3CbcuDmSYxfcV9C\nGYdFBvYwCCWeYMd6EWL+V12dV7Sqf0EpETsNQlvbcTweNWoo04fdURByV/t6pEkx\niaNhclcH7iCSu7t0mZVqo2QaUVqErNpVh4+pqQUplCwopLCQeJ3SvRX4OJr6VRha\noULq6+BNyVLozX4LdNPMutI=\n-----END PRIVATE KEY-----\n"
    ),
    "client_email": "datapos@posdashboard-442201.iam.gserviceaccount.com",
    "client_id": "116030208639366774536",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/datapos@posdashboard-442201.iam.gserviceaccount.com",
}

sheet_url = "https://docs.google.com/spreadsheets/d/1zjn_8Qi0RAdffzsfOmYWK3cqqVGdveaIrDABzB6BcRk/edit"
sheet_gid = "1150984969"

data = load_google_sheet(sheet_url, sheet_gid, google_credentials)

if data is not None:
    st.sidebar.markdown("### Filters")
    regions = st.sidebar.multiselect(
        "Select Regions", 
        options=data["REGION"].dropna().unique(), 
        default=data["REGION"].dropna().unique()
    )

    filtered_data = data[data["REGION"].isin(regions)]

    st.markdown("### Key Metrics")
    for region in filtered_data["REGION"].unique():
        region_data = filtered_data[filtered_data["REGION"] == region]
        st.markdown(f"#### Region: {region}")
        st.write(region_data[[
            "BRANCH", "POS MOF Total", "POS MOF Issued", "MOF POS %",
            "POS TOWN Total", "POS TOWN  Issued", "TOWN POS %"
        ]])

        # MOF POS Bar Chart
        st.markdown(f"##### MOF POS Services for {region}")
        mof_bar_fig = px.bar(
            region_data,
            x="BRANCH",
            y=["POS MOF Total", "POS MOF Issued"],
            title=f"MOF POS Services in {region}",
            labels={"value": "POS Metrics", "BRANCH": "Branch"},
            barmode="group",
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        st.plotly_chart(mof_bar_fig, use_container_width=True)
        st.markdown(f"**MOF POS %**: {region_data['MOF POS %'].mean():.2f}%")

        # Town POS Bar Chart
        st.markdown(f"##### Town POS Services for {region}")
        town_bar_fig = px.bar(
            region_data,
            x="BRANCH",
            y=["POS TOWN Total", "POS TOWN  Issued"],
            title=f"Town POS Services in {region}",
            labels={"value": "POS Metrics", "BRANCH": "Branch"},
            barmode="group",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        st.plotly_chart(town_bar_fig, use_container_width=True)
        st.markdown(f"**Town POS %**: {region_data['TOWN POS %'].mean():.2f}%")

        # Pie Charts for MOF Ticket Sales
        st.markdown(f"##### MOF Ticket Sales for {region}")
        mof_tickets_fig = px.pie(
            region_data,
            names="BRANCH",
            values="MOF POS Tickets",
            title=f"MOF POS Tickets Distribution",
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        st.plotly_chart(mof_tickets_fig, use_container_width=True)

        mof_preprint_fig = px.pie(
            region_data,
            names="BRANCH",
            values="MOF Pre Printed  Tickets",
            title=f"MOF Pre-Printed Tickets Distribution",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        st.plotly_chart(mof_preprint_fig, use_container_width=True)

        # Pie Charts for Town Ticket Sales
        st.markdown(f"##### Town Ticket Sales for {region}")
        town_tickets_fig = px.pie(
            region_data,
            names="BRANCH",
            values="TOWN POS Tickets",
            title=f"Town POS Tickets Distribution",
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        st.plotly_chart(town_tickets_fig, use_container_width=True)

        town_preprint_fig = px.pie(
            region_data,
            names="BRANCH",
            values="TOWN Pre Printed  Tickets",
            title=f"Town Pre-Printed Tickets Distribution",
            color_discrete_sequence=px.colors.qualitative.Set1
        )
        st.plotly_chart(town_preprint_fig, use_container_width=True)

    # Download Filtered Data
    st.sidebar.markdown("### Download Filtered Data")
    if not filtered_data.empty:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            filtered_data.to_excel(writer, index=False, sheet_name="Filtered Data")
        buffer.seek(0)
        st.sidebar.download_button(
            label="Download Filtered Data (Excel)",
            data=buffer,
            file_name="filtered_pos_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Failed to load data.")
