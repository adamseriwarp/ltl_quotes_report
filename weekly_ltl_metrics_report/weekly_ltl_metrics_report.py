import streamlit as st
import pandas as pd
from datetime import datetime
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from shared.drive_client import DriveClient
from weekly_ltl_metrics_report.report_generator import (
    generate_report, generate_lanes_report, generate_regions_report,
    generate_monthly_report, clear_csv_cache, get_available_weeks, get_month_from_week
)

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
        # Check if password key exists before accessing
        if "password" not in st.session_state:
            return
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
def load_available_weeks(_client, max_weeks: int = 12) -> list[dict]:
    """Load and cache available weeks from Drive."""
    return get_available_weeks(_client, max_weeks=max_weeks)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_report(_client, selected_week_labels: tuple[str, ...]) -> tuple:
    """Load and cache the report data for selected weeks."""
    # We need to re-fetch the week folders based on labels (cache-friendly)
    all_weeks = load_available_weeks(_client, max_weeks=52)
    selected_weeks = [w for w in all_weeks if w['label'] in selected_week_labels]
    return generate_report(_client, selected_weeks=selected_weeks)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_monthly_report(_client, selected_week_labels: tuple[str, ...]) -> tuple:
    """Load and cache the monthly aggregated report data for selected weeks."""
    all_weeks = load_available_weeks(_client, max_weeks=52)
    selected_weeks = [w for w in all_weeks if w['label'] in selected_week_labels]
    return generate_monthly_report(_client, selected_weeks=selected_weeks)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_lanes_report(_client, selected_week_labels: tuple[str, ...]) -> tuple:
    """Load and cache the lanes report data for selected weeks."""
    all_weeks = load_available_weeks(_client)
    selected_weeks = [w for w in all_weeks if w['label'] in selected_week_labels]
    return generate_lanes_report(_client, selected_weeks=selected_weeks)

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_regions_report(_client, selected_week_labels: tuple[str, ...]) -> tuple:
    """Load and cache the regions report data for selected weeks."""
    all_weeks = load_available_weeks(_client)
    selected_weeks = [w for w in all_weeks if w['label'] in selected_week_labels]
    return generate_regions_report(_client, selected_weeks=selected_weeks)

def main():
    # Sidebar controls
    st.sidebar.header("📅 Report Options")

    # Connect to Drive
    with st.spinner("Connecting to Google Drive..."):
        try:
            client = get_drive_client()
            st.sidebar.success("✓ Connected to Google Drive")
        except Exception as e:
            st.error(f"Failed to connect to Google Drive: {e}")
            return

    # Load available weeks for the multiselect (up to 52 weeks / 1 year)
    with st.spinner("Loading available weeks..."):
        try:
            available_weeks = load_available_weeks(client, max_weeks=52)
        except Exception as e:
            st.error(f"Error loading available weeks: {e}")
            return

    # Create week label options (chronological order, oldest first)
    week_labels = [w['label'] for w in available_weeks]

    if not week_labels:
        st.warning("No weeks available.")
        return

    # Build year and month options from available weeks
    years_available = sorted(set(w['year'] for w in available_weeks))

    # Build month options (e.g., "Jan 2025", "Feb 2025")
    month_options = {}
    for w in available_weeks:
        _, _, month_label = get_month_from_week(w['week'], w['year'])
        if month_label not in month_options:
            month_options[month_label] = []
        month_options[month_label].append(w['label'])

    # Sort month options chronologically
    month_names_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    sorted_months = sorted(month_options.keys(),
                          key=lambda x: (int(x.split()[1]), month_names_order.index(x.split()[0])))

    # Quick filter selection method
    st.sidebar.subheader("📅 Quick Filters")
    filter_method = st.sidebar.radio(
        "Select by:",
        options=["Last N weeks", "Year", "Month"],
        horizontal=True,
        help="Choose how to quickly select weeks"
    )

    # Initialize session state
    if 'selected_weeks' not in st.session_state:
        st.session_state.selected_weeks = week_labels[-4:] if len(week_labels) >= 4 else week_labels
    if 'last_filter_method' not in st.session_state:
        st.session_state.last_filter_method = filter_method
    if 'last_filter_value' not in st.session_state:
        st.session_state.last_filter_value = None

    quick_selection = []

    if filter_method == "Last N weeks":
        max_slider = min(len(week_labels), 52)
        num_weeks_slider = st.sidebar.slider(
            "Number of weeks:",
            min_value=1,
            max_value=max_slider,
            value=min(4, max_slider),
            help="Select the most recent N weeks"
        )
        quick_selection = week_labels[-num_weeks_slider:]
        current_filter_value = f"slider_{num_weeks_slider}"

    elif filter_method == "Year":
        selected_year = st.sidebar.selectbox(
            "Select year:",
            options=years_available,
            index=len(years_available) - 1,  # Default to most recent year
            help="Select all weeks from a specific year"
        )
        quick_selection = [w['label'] for w in available_weeks if w['year'] == selected_year]
        current_filter_value = f"year_{selected_year}"

    elif filter_method == "Month":
        selected_month = st.sidebar.selectbox(
            "Select month:",
            options=sorted_months,
            index=len(sorted_months) - 1,  # Default to most recent month
            help="Select all weeks from a specific month"
        )
        quick_selection = month_options.get(selected_month, [])
        current_filter_value = f"month_{selected_month}"

    # Update selection if filter method or value changed
    if (st.session_state.last_filter_method != filter_method or
        st.session_state.last_filter_value != current_filter_value):
        st.session_state.last_filter_method = filter_method
        st.session_state.last_filter_value = current_filter_value
        st.session_state.selected_weeks = quick_selection

    # Multiselect for fine-tuning
    st.sidebar.subheader("🔧 Fine-tune")
    selected_week_labels = st.sidebar.multiselect(
        "Selected weeks:",
        options=week_labels,
        default=st.session_state.selected_weeks,
        help="Add or remove specific weeks"
    )

    # Update session state with current selection
    st.session_state.selected_weeks = selected_week_labels

    if not selected_week_labels:
        st.warning("Please select at least one week.")
        return

    # Load data button
    if st.sidebar.button("🔄 Refresh Data", type="primary"):
        st.cache_data.clear()
        clear_csv_cache()  # Also clear the in-memory CSV cache
        st.rerun()

    # Convert to tuple for caching (lists are not hashable)
    selected_week_labels_tuple = tuple(selected_week_labels)

    # Decide: weekly view (<= 12 weeks) or monthly aggregation (> 12 weeks)
    WEEKLY_THRESHOLD = 12
    use_monthly = len(selected_week_labels) > WEEKLY_THRESHOLD

    # Show warning for large selections
    if len(selected_week_labels) > 30:
        st.sidebar.warning(f"⚠️ Loading {len(selected_week_labels)} weeks may take a while...")

    # Load report
    if use_monthly:
        loading_msg = f"Loading {len(selected_week_labels)} weeks and aggregating by month... This may take a moment."
        with st.spinner(loading_msg):
            try:
                report_df, period_labels = load_monthly_report(client, selected_week_labels_tuple)
                view_mode = "monthly"
            except Exception as e:
                st.error(f"Error loading report: {e}")
                import traceback
                st.code(traceback.format_exc())
                return
    else:
        with st.spinner(f"Loading data for {len(selected_week_labels)} weeks..."):
            try:
                report_df, period_labels = load_report(client, selected_week_labels_tuple)
                view_mode = "weekly"
            except Exception as e:
                st.error(f"Error loading report: {e}")
                import traceback
                st.code(traceback.format_exc())
                return

    if report_df.empty:
        st.warning("No data found.")
        return

    # Show mode indicator
    if use_monthly:
        st.sidebar.info(f"📅 **Monthly view** ({len(selected_week_labels)} weeks → {len(period_labels)} months)")
    else:
        st.sidebar.write(f"Weeks loaded: {period_labels}")

    # Customer filter
    st.sidebar.header("🏢 Customer Filter")
    all_customers = sorted([c for c in report_df['Customers'].unique() if c != 'TOTAL'])
    selected_customers = st.sidebar.multiselect(
        "Select customers:",
        options=all_customers,
        default=all_customers,
        help="Filter the report to show only selected customers"
    )

    # Filter the report based on selected customers (always keep TOTAL row)
    if selected_customers:
        display_df = report_df[
            (report_df['Customers'].isin(selected_customers)) |
            (report_df['Customers'] == 'TOTAL')
        ]
    else:
        display_df = report_df[report_df['Customers'] == 'TOTAL']  # Show only TOTAL if none selected

    latest_period = period_labels[-1]  # Most recent period (week or month)

    # Get TOTAL row stats for metrics
    total_row = report_df[report_df['Customers'] == 'TOTAL'].iloc[0] if 'TOTAL' in report_df['Customers'].values else None

    # Display metrics for latest period
    period_type = "Month" if use_monthly else "Week"
    st.subheader(f"📈 {latest_period} Summary")
    col1, col2, col3, col4 = st.columns(4)
    if total_row is not None:
        col1.metric("Total Quotes", f"{int(total_row[f'{latest_period}_Total Quotes']):,}")
        col2.metric("Rated Quotes", f"{int(total_row[f'{latest_period}_Rated']):,}")
        col3.metric("Booked", f"{int(total_row[f'{latest_period}_Booked']):,}")
        col4.metric("% Rated", f"{total_row[f'{latest_period}_% Rated']:.2f}%")

    st.divider()

    # Display the report table
    table_title = "📊 Quotes by Customer (Monthly)" if use_monthly else "📊 Quotes by Customer"
    st.subheader(table_title)

    # Build HTML table with proper hierarchical headers
    def build_customer_html_table(report_df, periods):
        # Header row 1: Period labels spanning 4 columns each
        header1 = '<tr><th rowspan="2" style="background-color: #F2F2F2; padding: 8px; border: 1px solid #999; border-left: 2px solid #666;">Customers</th>'
        for i, period in enumerate(periods):
            bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
            header1 += f'<th colspan="4" style="background-color: {bg_color}; padding: 8px; border: 1px solid #999; border-left: 2px solid #666; text-align: center; font-weight: bold;">{period}</th>'
        header1 += '</tr>'

        # Header row 2: Sub-columns
        header2 = '<tr>'
        for i, period in enumerate(periods):
            bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
            for j, col in enumerate(['Booked', 'Rated', 'Total Quotes', '% Rated']):
                left_border = 'border-left: 2px solid #666;' if j == 0 else 'border-left: 1px solid #999;'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; {left_border} text-align: center; font-size: 12px;">{col}</th>'
        header2 += '</tr>'

        # Data rows
        rows_html = ''
        for idx, row in report_df.iterrows():
            is_total = row['Customers'] == 'TOTAL'
            row_style = 'background-color: #D9D9D9; font-weight: bold;' if is_total else ''
            rows_html += f'<tr style="{row_style}">'
            rows_html += f'<td style="padding: 6px; border: 1px solid #999; border-left: 2px solid #666; {row_style}">{row["Customers"]}</td>'

            for i, period in enumerate(periods):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                if is_total:
                    bg_color = '#D9D9D9'

                booked = int(row[f'{period}_Booked'])
                rated = int(row[f'{period}_Rated'])
                total = int(row[f'{period}_Total Quotes'])
                pct = row[f'{period}_% Rated']

                # First column of each period gets thick left border
                rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; border-left: 2px solid #666; text-align: right;">{booked:,}</td>'
                rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: right;">{rated:,}</td>'
                rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: right;">{total:,}</td>'
                rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: right;">{pct:.2f}%</td>'
            rows_html += '</tr>'

        return f'''
        <div style="overflow-x: auto;">
            <table style="border-collapse: collapse; width: 100%; font-size: 14px; border: 2px solid #666;">
                <thead>{header1}{header2}</thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        '''

    st.markdown(build_customer_html_table(display_df, period_labels), unsafe_allow_html=True)

    # Download buttons for customer report
    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        csv = display_df.to_csv(index=False)
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
            lanes_df, lanes_weeks = load_lanes_report(client, selected_week_labels_tuple)
        except Exception as e:
            st.error(f"Error loading lanes report: {e}")
            lanes_df = pd.DataFrame()

    if not lanes_df.empty:
        # Display lanes table
        st.subheader("📊 Rated Quotes by Lane (Airport-to-Airport)")

        # Build HTML table with proper hierarchical headers
        def build_lanes_html_table(df, weeks):
            header1 = '<tr><th rowspan="2" style="background-color: #F2F2F2; padding: 8px; border: 1px solid #999; border-left: 2px solid #666;">Lanes</th>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header1 += f'<th colspan="2" style="background-color: {bg_color}; padding: 8px; border: 1px solid #999; border-left: 2px solid #666; text-align: center; font-weight: bold;">{week}</th>'
            header1 += '</tr>'

            header2 = '<tr>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; border-left: 2px solid #666; text-align: center; font-size: 12px;">Total</th>'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: center; font-size: 12px;">%Change</th>'
            header2 += '</tr>'

            rows_html = ''
            for idx, row in df.iterrows():
                rows_html += '<tr>'
                rows_html += f'<td style="padding: 6px; border: 1px solid #999; border-left: 2px solid #666;">{row["Lanes"]}</td>'
                for i, week in enumerate(weeks):
                    bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                    total = row.get(f'{week}_Total', 0)
                    pct_change = row.get(f'{week}_%Change', None)
                    total_val = int(total) if pd.notna(total) else 0
                    pct_val = f"{pct_change:+.0f}%" if pd.notna(pct_change) else "-"
                    rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; border-left: 2px solid #666; text-align: right;">{total_val:,}</td>'
                    rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: right;">{pct_val}</td>'
                rows_html += '</tr>'

            return f'''
            <div style="overflow-x: auto; max-height: 600px; overflow-y: auto;">
                <table style="border-collapse: collapse; width: 100%; font-size: 14px; border: 2px solid #666;">
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
            regions_df, regions_weeks = load_regions_report(client, selected_week_labels_tuple)
        except Exception as e:
            st.error(f"Error loading regions report: {e}")
            regions_df = pd.DataFrame()

    if not regions_df.empty:
        # Display regions table
        st.subheader("📊 Rated Quotes by Region (Region-to-Region)")

        # Build HTML table with proper hierarchical headers
        def build_regions_html_table(df, weeks):
            header1 = '<tr><th rowspan="2" style="background-color: #F2F2F2; padding: 8px; border: 1px solid #999; border-left: 2px solid #666;">Regions</th>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header1 += f'<th colspan="2" style="background-color: {bg_color}; padding: 8px; border: 1px solid #999; border-left: 2px solid #666; text-align: center; font-weight: bold;">{week}</th>'
            header1 += '</tr>'

            header2 = '<tr>'
            for i, week in enumerate(weeks):
                bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; border-left: 2px solid #666; text-align: center; font-size: 12px;">Total</th>'
                header2 += f'<th style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: center; font-size: 12px;">%Change</th>'
            header2 += '</tr>'

            rows_html = ''
            for idx, row in df.iterrows():
                rows_html += '<tr>'
                rows_html += f'<td style="padding: 6px; border: 1px solid #999; border-left: 2px solid #666;">{row["Regions"]}</td>'
                for i, week in enumerate(weeks):
                    bg_color = '#E8E8E8' if i % 2 == 0 else '#FFFFFF'
                    total = row.get(f'{week}_Total', 0)
                    pct_change = row.get(f'{week}_%Change', None)
                    total_val = int(total) if pd.notna(total) else 0
                    pct_val = f"{pct_change:+.0f}%" if pd.notna(pct_change) else "-"
                    rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; border-left: 2px solid #666; text-align: right;">{total_val:,}</td>'
                    rows_html += f'<td style="background-color: {bg_color}; padding: 6px; border: 1px solid #999; text-align: right;">{pct_val}</td>'
                rows_html += '</tr>'

            return f'''
            <div style="overflow-x: auto; max-height: 600px; overflow-y: auto;">
                <table style="border-collapse: collapse; width: 100%; font-size: 14px; border: 2px solid #666;">
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

