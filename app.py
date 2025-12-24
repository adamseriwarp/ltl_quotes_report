import streamlit as st
import pandas as pd
from datetime import datetime
from drive_client import DriveClient
from report_generator import generate_report, generate_lanes_report, generate_regions_report, clear_csv_cache

st.set_page_config(
    page_title="WARP Freight Quotes Report",
    page_icon="📊",
    layout="wide"
)

# Password protection
def check_password():
    """Returns True if the user has entered the correct password."""

    # Check if password is configured in secrets
    if "password" not in st.secrets:
        # No password configured - allow access (for local dev)
        return True

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.info("Please enter the password to access this report.")
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Incorrect password. Please try again.")
        return False
    else:
        # Password correct
        return True

if not check_password():
    st.stop()

# Custom CSS for alternating week colors
st.markdown("""
<style>
    .main-header {
        background-color: #00B050;
        color: white;
        padding: 15px;
        text-align: center;
        font-size: 24px;
        font-weight: bold;
        border-radius: 5px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">WARP FREIGHT QUOTES RETURNED</div>', unsafe_allow_html=True)

@st.cache_resource
def get_drive_client():
    """Cache the Drive client connection."""
    return DriveClient()

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_report(_client, num_weeks: int) -> tuple:
    """Load and cache the report data."""
    return generate_report(_client, num_weeks=num_weeks)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_lanes_report(_client, num_weeks: int) -> tuple:
    """Load and cache the lanes report data."""
    return generate_lanes_report(_client, num_weeks=num_weeks)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_regions_report(_client, num_weeks: int) -> tuple:
    """Load and cache the regions report data."""
    return generate_regions_report(_client, num_weeks=num_weeks)

def main():
    # Sidebar controls
    st.sidebar.header("📅 Report Options")

    num_weeks = st.sidebar.slider("Number of weeks:", min_value=1, max_value=8, value=4)

    # Connect to Drive
    with st.spinner("Connecting to Google Drive..."):
        try:
            client = get_drive_client()
            st.sidebar.success("✓ Connected to Google Drive")
        except Exception as e:
            st.error(f"Failed to connect to Google Drive: {e}")
            return

    # Load data button
    if st.sidebar.button("🔄 Refresh Data", type="primary"):
        st.cache_data.clear()
        clear_csv_cache()  # Also clear the in-memory CSV cache
        st.rerun()

    # Load report
    with st.spinner(f"Loading data for {num_weeks} weeks..."):
        try:
            report_df, weeks = load_report(client, num_weeks)
        except Exception as e:
            st.error(f"Error loading report: {e}")
            return

    if report_df.empty:
        st.warning("No data found.")
        return

    st.sidebar.write(f"Weeks loaded: {weeks}")

    latest_week = weeks[-1]  # Most recent week

    # Get TOTAL row stats for metrics
    total_row = report_df[report_df['Customers'] == 'TOTAL'].iloc[0] if 'TOTAL' in report_df['Customers'].values else None

    # Display metrics for latest week
    st.subheader(f"📈 {latest_week} Summary")
    col1, col2, col3, col4 = st.columns(4)
    if total_row is not None:
        col1.metric("Total Quotes", f"{int(total_row[f'{latest_week}_Total Quotes']):,}")
        col2.metric("Rated Quotes", f"{int(total_row[f'{latest_week}_Rated']):,}")
        col3.metric("Booked", f"{int(total_row[f'{latest_week}_Booked']):,}")
        col4.metric("% Rated", f"{total_row[f'{latest_week}_% Rated']:.2f}%")

    st.divider()

    # Display the report table
    st.subheader("📊 Quotes by Customer")

    # Build HTML table with proper hierarchical headers
    def build_customer_html_table(report_df, weeks):
        # Header row 1: Week labels spanning 4 columns each
        header1 = '<tr><th rowspan="2" style="background-color: #F2F2F2; padding: 8px; border: 1px solid #ddd;">Customers</th>'
        for i, week in enumerate(weeks):
            bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
            header1 += f'<th colspan="4" style="background-color: {bg_color}; padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;">{week}</th>'
        header1 += '</tr>'

        # Header row 2: Sub-columns
        header2 = '<tr>'
        for i, week in enumerate(weeks):
            bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
            for col in ['Booked', 'Rated', 'Total Quotes', '% Rated']:
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: center; font-size: 12px;">{col}</th>'
        header2 += '</tr>'

        # Data rows
        rows_html = ''
        for idx, row in report_df.iterrows():
            is_total = row['Customers'] == 'TOTAL'
            row_style = 'background-color: #D9D9D9; font-weight: bold;' if is_total else ''
            rows_html += f'<tr style="{row_style}">'
            rows_html += f'<td style="padding: 6px; border: 1px solid #ddd; {row_style}">{row["Customers"]}</td>'

            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                if is_total:
                    bg_color = '#D9D9D9'
                cell_style = f'background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: right;'

                booked = int(row[f'{week}_Booked'])
                rated = int(row[f'{week}_Rated'])
                total = int(row[f'{week}_Total Quotes'])
                pct = row[f'{week}_% Rated']

                rows_html += f'<td style="{cell_style}">{booked:,}</td>'
                rows_html += f'<td style="{cell_style}">{rated:,}</td>'
                rows_html += f'<td style="{cell_style}">{total:,}</td>'
                rows_html += f'<td style="{cell_style}">{pct:.2f}%</td>'
            rows_html += '</tr>'

        return f'''
        <div style="overflow-x: auto;">
            <table style="border-collapse: collapse; width: 100%; font-size: 14px;">
                <thead>{header1}{header2}</thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        '''

    st.markdown(build_customer_html_table(report_df, weeks), unsafe_allow_html=True)

    # Download buttons for customer report
    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        csv = report_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Customer Report CSV",
            data=csv,
            file_name=f"ltl_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

    with col2:
        st.info("For Excel with formatting, use the command-line report generator.")

    # ========== LANES QUOTED REPORT ==========
    st.divider()
    st.markdown('<div class="main-header">LANES QUOTED</div>', unsafe_allow_html=True)

    # Load lanes report using same weeks
    with st.spinner("Loading lanes data..."):
        try:
            lanes_df, lanes_weeks = load_lanes_report(client, num_weeks)
        except Exception as e:
            st.error(f"Error loading lanes report: {e}")
            lanes_df = pd.DataFrame()

    if not lanes_df.empty:
        # Display lanes table
        st.subheader("📊 Rated Quotes by Lane (Airport-to-Airport)")

        # Build HTML table with proper hierarchical headers
        def build_lanes_html_table(df, weeks):
            header1 = '<tr><th rowspan="2" style="background-color: #F2F2F2; padding: 8px; border: 1px solid #ddd;">Lanes</th>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header1 += f'<th colspan="2" style="background-color: {bg_color}; padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;">{week}</th>'
            header1 += '</tr>'

            header2 = '<tr>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: center; font-size: 12px;">Total</th>'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: center; font-size: 12px;">%Change</th>'
            header2 += '</tr>'

            rows_html = ''
            for idx, row in df.iterrows():
                rows_html += '<tr>'
                rows_html += f'<td style="padding: 6px; border: 1px solid #ddd;">{row["Lanes"]}</td>'
                for i, week in enumerate(weeks):
                    bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                    cell_style = f'background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: right;'
                    total = row.get(f'{week}_Total', 0)
                    pct_change = row.get(f'{week}_%Change', None)
                    total_val = int(total) if pd.notna(total) else 0
                    pct_val = f"{pct_change:+.0f}%" if pd.notna(pct_change) else "-"
                    rows_html += f'<td style="{cell_style}">{total_val:,}</td>'
                    rows_html += f'<td style="{cell_style}">{pct_val}</td>'
                rows_html += '</tr>'

            return f'''
            <div style="overflow-x: auto; max-height: 600px; overflow-y: auto;">
                <table style="border-collapse: collapse; width: 100%; font-size: 14px;">
                    <thead style="position: sticky; top: 0;">{header1}{header2}</thead>
                    <tbody>{rows_html}</tbody>
                </table>
            </div>
            '''

        st.markdown(build_lanes_html_table(lanes_df, lanes_weeks), unsafe_allow_html=True)

        # Download button for lanes report
        st.divider()
        lanes_csv = lanes_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Lanes Report CSV",
            data=lanes_csv,
            file_name=f"lanes_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    else:
        st.warning("No lanes data found for the selected weeks.")

    # ========== REGIONS QUOTED REPORT ==========
    st.divider()
    st.markdown('<div class="main-header">REGIONS QUOTED</div>', unsafe_allow_html=True)

    # Load regions report using same weeks
    with st.spinner("Loading regions data..."):
        try:
            regions_df, regions_weeks = load_regions_report(client, num_weeks)
        except Exception as e:
            st.error(f"Error loading regions report: {e}")
            regions_df = pd.DataFrame()

    if not regions_df.empty:
        # Display regions table
        st.subheader("📊 Rated Quotes by Region (Region-to-Region)")

        # Build HTML table with proper hierarchical headers
        def build_regions_html_table(df, weeks):
            header1 = '<tr><th rowspan="2" style="background-color: #F2F2F2; padding: 8px; border: 1px solid #ddd;">Regions</th>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header1 += f'<th colspan="2" style="background-color: {bg_color}; padding: 8px; border: 1px solid #ddd; text-align: center; font-weight: bold;">{week}</th>'
            header1 += '</tr>'

            header2 = '<tr>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: center; font-size: 12px;">Total</th>'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: center; font-size: 12px;">%Change</th>'
            header2 += '</tr>'

            rows_html = ''
            for idx, row in df.iterrows():
                rows_html += '<tr>'
                rows_html += f'<td style="padding: 6px; border: 1px solid #ddd;">{row["Regions"]}</td>'
                for i, week in enumerate(weeks):
                    bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                    cell_style = f'background-color: {bg_color}; padding: 6px; border: 1px solid #ddd; text-align: right;'
                    total = row.get(f'{week}_Total', 0)
                    pct_change = row.get(f'{week}_%Change', None)
                    total_val = int(total) if pd.notna(total) else 0
                    pct_val = f"{pct_change:+.0f}%" if pd.notna(pct_change) else "-"
                    rows_html += f'<td style="{cell_style}">{total_val:,}</td>'
                    rows_html += f'<td style="{cell_style}">{pct_val}</td>'
                rows_html += '</tr>'

            return f'''
            <div style="overflow-x: auto; max-height: 600px; overflow-y: auto;">
                <table style="border-collapse: collapse; width: 100%; font-size: 14px;">
                    <thead style="position: sticky; top: 0;">{header1}{header2}</thead>
                    <tbody>{rows_html}</tbody>
                </table>
            </div>
            '''

        st.markdown(build_regions_html_table(regions_df, regions_weeks), unsafe_allow_html=True)

        # Download button for regions report
        st.divider()
        regions_csv = regions_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Regions Report CSV",
            data=regions_csv,
            file_name=f"regions_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    else:
        st.warning("No regions data found for the selected weeks.")

if __name__ == "__main__":
    main()

