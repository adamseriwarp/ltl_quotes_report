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
    st.subheader(f"📈 Week {latest_week} Summary")
    col1, col2, col3, col4 = st.columns(4)
    if total_row is not None:
        col1.metric("Total Quotes", f"{int(total_row[f'{latest_week}_Total Quotes']):,}")
        col2.metric("Rated Quotes", f"{int(total_row[f'{latest_week}_Rated']):,}")
        col3.metric("Booked", f"{int(total_row[f'{latest_week}_Booked']):,}")
        col4.metric("% Rated", f"{total_row[f'{latest_week}_% Rated']:.2f}%")

    st.divider()

    # Display the report table
    st.subheader("📊 Quotes by Customer")

    # Create multi-index columns for hierarchical headers (Week -> Booked, Rated, etc.)
    customers = report_df['Customers'].values

    # Build new dataframe with MultiIndex columns
    new_columns = [('', 'Customers')]
    for week in weeks:
        new_columns.append((str(week), 'Booked'))
        new_columns.append((str(week), 'Rated'))
        new_columns.append((str(week), 'Total Quotes'))
        new_columns.append((str(week), '% Rated'))

    multi_index = pd.MultiIndex.from_tuples(new_columns)

    # Rebuild data with new column structure
    data = []
    for idx, row in report_df.iterrows():
        new_row = [row['Customers']]
        for week in weeks:
            new_row.append(row[f'{week}_Booked'])
            new_row.append(row[f'{week}_Rated'])
            new_row.append(row[f'{week}_Total Quotes'])
            pct = row[f'{week}_% Rated']
            new_row.append(f"{pct:.2f}%")
        data.append(new_row)

    display_df = pd.DataFrame(data, columns=multi_index)

    # Apply styling with alternating week colors
    def style_table(df):
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        colors = ['background-color: #F2F2F2', 'background-color: #FFFFFF']

        for i, week in enumerate(weeks):
            color = colors[i % 2]
            for sub_col in ['Booked', 'Rated', 'Total Quotes', '% Rated']:
                if (str(week), sub_col) in df.columns:
                    styles[(str(week), sub_col)] = color

        # Bold the TOTAL row
        styles.iloc[0] = 'background-color: #D9D9D9; font-weight: bold'
        return styles

    styled_df = display_df.style.apply(lambda _: style_table(display_df), axis=None)
    st.dataframe(styled_df, use_container_width=True, height=600)

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

        # Build display dataframe with MultiIndex columns
        lanes_columns = [('', 'Lanes')]
        for week in lanes_weeks:
            lanes_columns.append((str(week), 'Total'))
            lanes_columns.append((str(week), '%Change'))

        lanes_multi_index = pd.MultiIndex.from_tuples(lanes_columns)

        # Rebuild data with new column structure
        lanes_data = []
        for idx, row in lanes_df.iterrows():
            new_row = [row['Lanes']]
            for week in lanes_weeks:
                total = row.get(f'{week}_Total', 0)
                pct_change = row.get(f'{week}_%Change', None)
                new_row.append(int(total) if pd.notna(total) else 0)
                if pd.notna(pct_change):
                    new_row.append(f"{pct_change:+.0f}%")
                else:
                    new_row.append("-")
            lanes_data.append(new_row)

        lanes_display_df = pd.DataFrame(lanes_data, columns=lanes_multi_index)

        # Apply styling with alternating week colors
        def style_lanes_table(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            colors = ['background-color: #F2F2F2', 'background-color: #FFFFFF']

            for i, week in enumerate(lanes_weeks):
                color = colors[i % 2]
                for sub_col in ['Total', '%Change']:
                    if (str(week), sub_col) in df.columns:
                        styles[(str(week), sub_col)] = color

            return styles

        styled_lanes_df = lanes_display_df.style.apply(lambda _: style_lanes_table(lanes_display_df), axis=None)
        st.dataframe(styled_lanes_df, use_container_width=True, height=600)

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

        # Build display dataframe with MultiIndex columns
        regions_columns = [('', 'Regions')]
        for week in regions_weeks:
            regions_columns.append((str(week), 'Total'))
            regions_columns.append((str(week), '%Change'))

        regions_multi_index = pd.MultiIndex.from_tuples(regions_columns)

        # Rebuild data with new column structure
        regions_data = []
        for idx, row in regions_df.iterrows():
            new_row = [row['Regions']]
            for week in regions_weeks:
                total = row.get(f'{week}_Total', 0)
                pct_change = row.get(f'{week}_%Change', None)
                new_row.append(int(total) if pd.notna(total) else 0)
                if pd.notna(pct_change):
                    new_row.append(f"{pct_change:+.0f}%")
                else:
                    new_row.append("-")
            regions_data.append(new_row)

        regions_display_df = pd.DataFrame(regions_data, columns=regions_multi_index)

        # Apply styling with alternating week colors
        def style_regions_table(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            colors = ['background-color: #F2F2F2', 'background-color: #FFFFFF']

            for i, week in enumerate(regions_weeks):
                color = colors[i % 2]
                for sub_col in ['Total', '%Change']:
                    if (str(week), sub_col) in df.columns:
                        styles[(str(week), sub_col)] = color

            return styles

        styled_regions_df = regions_display_df.style.apply(lambda _: style_regions_table(regions_display_df), axis=None)
        st.dataframe(styled_regions_df, use_container_width=True, height=600)

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

