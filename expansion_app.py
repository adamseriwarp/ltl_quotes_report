import streamlit as st

st.set_page_config(
    page_title="WARP Expansion Analysis",
    page_icon="🗺️",
    layout="wide"
)

st.markdown('<div style="background-color: #4472C4; color: white; padding: 15px; text-align: center; font-size: 24px; font-weight: bold; border-radius: 5px; margin-bottom: 20px;">ZIP CODE EXPANSION ANALYSIS</div>', unsafe_allow_html=True)

try:
    import pandas as pd
    import numpy as np
    from shapely.strtree import STRtree
    from drive_client import DriveClient
    from report_generator import load_csvs_from_folder, load_zip_to_airport_mapping, load_airport_to_region_mapping, get_airport_code, get_region, clear_csv_cache
    import re
except Exception as e:
    st.error(f"Import error: {e}")
    st.stop()

CENTROIDS_FILE = "zip_centroids.csv"

@st.cache_resource
def get_drive_client():
    """Cache the Drive client connection."""
    return DriveClient()

@st.cache_resource
def load_zip_centroids():
    """Load pre-computed ZIP code centroids."""
    print("Loading ZIP code centroids...")
    df = pd.read_csv(CENTROIDS_FILE, dtype={'zip_code': str})
    df.set_index('zip_code', inplace=True)
    print(f"Loaded {len(df)} ZIP code centroids")
    return df

@st.cache_data(ttl=300)
def load_quotes_data(_client, selected_year_weeks: tuple) -> pd.DataFrame:
    """Load quote data for selected year-week combinations.

    Args:
        selected_year_weeks: tuple of (year, week) tuples to load
    """
    folders = _client.search_folders("Quotes")

    # Parse week folders
    week_pattern = re.compile(r'^W(\d{2})(\d{2})\s+Quotes$')
    week_folders = []
    for folder in folders:
        match = week_pattern.match(folder['name'])
        if match:
            week_num = int(match.group(1))
            year = 2000 + int(match.group(2))
            week_folders.append({
                'id': folder['id'],
                'name': folder['name'],
                'week': week_num,
                'year': year
            })

    # Sort by (year, week)
    week_folders = sorted(week_folders, key=lambda x: (x['year'], x['week']), reverse=True)

    all_data = []
    for wf in week_folders:
        if (wf['year'], wf['week']) in selected_year_weeks:
            print(f"Loading {wf['name']}...")
            df = load_csvs_from_folder(_client, wf['id'], wf['name'])
            if not df.empty:
                df['week'] = wf['week']
                df['year'] = wf['year']
                all_data.append(df)

    if not all_data:
        return pd.DataFrame()

    return pd.concat(all_data, ignore_index=True)

def get_available_year_weeks(client) -> list[tuple]:
    """Get list of available (year, week) tuples from Drive, sorted chronologically."""
    folders = client.search_folders("Quotes")
    week_pattern = re.compile(r'^W(\d{2})(\d{2})\s+Quotes$')
    year_weeks = []
    for folder in folders:
        match = week_pattern.match(folder['name'])
        if match:
            week_num = int(match.group(1))
            year = 2000 + int(match.group(2))
            year_weeks.append((year, week_num))
    return sorted(set(year_weeks))

def analyze_expansion_opportunities(df: pd.DataFrame, centroids_df: pd.DataFrame,
                                     zip_mapping: dict, region_mapping: dict) -> tuple:
    """Analyze unserviced zip codes for expansion opportunities.

    Returns: (results_df, serviced_zips, total_quotes_by_airport, total_quotes_by_region)
    """
    from shapely.geometry import Point

    # Identify rated vs unrated quotes
    df['is_rated'] = df['rate'].notna() & (df['rate'].astype(str).str.strip() != '')

    # Calculate total quotes by airport code and region (for percentage calculation)
    # We look at pickup zip to determine airport/region for each quote
    df['pickup_zip_clean'] = df['pickup Zip'].astype(str).str.strip().str[:5]
    df['pickup_airport'] = df['pickup_zip_clean'].apply(lambda x: get_airport_code(x, zip_mapping))
    df['pickup_region'] = df['pickup_airport'].apply(lambda x: get_region(x, region_mapping))

    total_quotes_by_airport = df.groupby('pickup_airport').size().to_dict()
    total_quotes_by_region = df.groupby('pickup_region').size().to_dict()

    # Get all unique zip codes from rated quotes (these are "serviced")
    rated_df = df[df['is_rated']]
    serviced_zips = set()
    for col in ['pickup Zip', 'dropoff Zip']:
        if col in rated_df.columns:
            serviced_zips.update(rated_df[col].dropna().astype(str).str.strip().str[:5].unique())

    # Get unrated quotes
    unrated_df = df[~df['is_rated']]

    if unrated_df.empty:
        return pd.DataFrame(), serviced_zips, total_quotes_by_airport, total_quotes_by_region

    # Count quotes per zip (combining origin and destination)
    zip_counts = {}
    for col in ['pickup Zip', 'dropoff Zip']:
        if col in unrated_df.columns:
            for zip_code in unrated_df[col].dropna().astype(str).str.strip().str[:5]:
                if zip_code and len(zip_code) == 5 and zip_code.isdigit():
                    if zip_code not in serviced_zips:  # Only unserviced zips
                        zip_counts[zip_code] = zip_counts.get(zip_code, 0) + 1

    if not zip_counts:
        return pd.DataFrame(), serviced_zips

    # Create DataFrame of unserviced zips with counts
    results = pd.DataFrame([
        {'zip_code': z, 'quote_count': c}
        for z, c in zip_counts.items()
    ])

    # Build list of serviced centroids from CSV
    serviced_centroids = []
    serviced_zip_list = []
    for z in serviced_zips:
        if z in centroids_df.index:
            row = centroids_df.loc[z]
            serviced_centroids.append(Point(row['centroid_x'], row['centroid_y']))
            serviced_zip_list.append(z)

    if serviced_centroids:
        # Build R-tree spatial index for fast nearest neighbor queries
        tree = STRtree(serviced_centroids)

        distances = []
        nearest_zips = []
        for zip_code in results['zip_code']:
            if zip_code in centroids_df.index:
                row = centroids_df.loc[zip_code]
                unserviced_centroid = Point(row['centroid_x'], row['centroid_y'])
                # Find nearest serviced centroid using spatial index
                nearest_idx = tree.nearest(unserviced_centroid)
                min_dist = unserviced_centroid.distance(serviced_centroids[nearest_idx])
                nearest_zip = serviced_zip_list[nearest_idx]
                # Convert from meters to miles (coordinates are in EPSG:2163 meters)
                distances.append(min_dist / 1609.34)
                nearest_zips.append(nearest_zip)
            else:
                distances.append(np.nan)
                nearest_zips.append(None)
        results['distance_miles'] = distances
        results['nearest_serviced_zip'] = nearest_zips
    else:
        results['distance_miles'] = np.nan
        results['nearest_serviced_zip'] = None

    # Add airport code and region
    results['airport_code'] = results['zip_code'].apply(lambda x: get_airport_code(x, zip_mapping))
    results['region'] = results['airport_code'].apply(lambda x: get_region(x, region_mapping))

    # Sort by quote count descending
    results = results.sort_values('quote_count', ascending=False)

    return results, serviced_zips, total_quotes_by_airport, total_quotes_by_region

def main():
    st.sidebar.header("🔧 Analysis Options")
    
    # Connect to Drive
    with st.spinner("Connecting to Google Drive..."):
        try:
            client = get_drive_client()
            st.sidebar.success("✓ Connected to Google Drive")
        except Exception as e:
            st.error(f"Failed to connect to Google Drive: {e}")
            return
    
    # Refresh button
    if st.sidebar.button("🔄 Refresh Data", type="primary"):
        st.cache_data.clear()
        clear_csv_cache()
        st.rerun()

    # Get available year-weeks
    available_year_weeks = get_available_year_weeks(client)
    if not available_year_weeks:
        st.error("No quote folders found")
        return

    # Create display labels for the dropdown (e.g., "W01 2025", "W52 2024")
    year_week_labels = [f"W{w:02d} {y}" for y, w in available_year_weeks]

    # Week range selector - default to last 4 weeks for faster loading
    st.sidebar.subheader("📅 Week Range")

    default_from_idx = max(len(available_year_weeks) - 4, 0)
    from_selection = st.sidebar.selectbox(
        "From:",
        options=year_week_labels,
        index=default_from_idx
    )
    to_selection = st.sidebar.selectbox(
        "To:",
        options=year_week_labels,
        index=len(year_week_labels) - 1
    )

    # Get indices and create range
    from_idx = year_week_labels.index(from_selection)
    to_idx = year_week_labels.index(to_selection)
    if from_idx > to_idx:
        from_idx, to_idx = to_idx, from_idx  # Swap if reversed

    selected_year_weeks = tuple(available_year_weeks[from_idx:to_idx + 1])

    # Filter controls
    st.sidebar.subheader("🎯 Filters")
    min_quotes = st.sidebar.slider("Minimum Quote Count:", min_value=1, max_value=500, value=10)
    max_distance = st.sidebar.slider("Maximum Distance to Serviced ZIP (miles):",
                                      min_value=1, max_value=300, value=60)

    # Run Analysis button
    if not st.sidebar.button("🚀 Run Analysis", type="primary"):
        st.info("👈 Adjust filters in the sidebar and click **Run Analysis** to start.")
        return

    # Load ZIP centroids
    with st.spinner("Loading ZIP code centroids..."):
        try:
            centroids_df = load_zip_centroids()
        except Exception as e:
            st.error(f"Failed to load ZIP centroids: {e}")
            return

    # Load data
    with st.spinner(f"Loading quotes for {from_selection} to {to_selection}..."):
        try:
            quotes_df = load_quotes_data(client, selected_year_weeks)
        except Exception as e:
            st.error(f"Failed to load quotes: {e}")
            return

    if quotes_df.empty:
        st.warning("No quote data found for selected weeks")
        return

    # Load mappings
    zip_mapping = load_zip_to_airport_mapping()
    region_mapping = load_airport_to_region_mapping()

    # Analyze
    with st.spinner("Analyzing expansion opportunities..."):
        try:
            results, serviced_zips, total_quotes_by_airport, total_quotes_by_region = analyze_expansion_opportunities(quotes_df, centroids_df, zip_mapping, region_mapping)
        except Exception as e:
            import traceback
            st.error(f"Analysis failed: {e}")
            st.code(traceback.format_exc())
            return

    if results.empty:
        st.warning("No unserviced zip codes found")
        return

    # Apply filters
    filtered = results[
        (results['quote_count'] >= min_quotes) &
        (results['distance_miles'].notna()) &
        (results['distance_miles'] <= max_distance)
    ].copy()

    # Display metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Unserviced ZIPs", len(results))
    with col2:
        st.metric("ZIPs After Filter", len(filtered))
    with col3:
        st.metric("Total Unrated Quotes", results['quote_count'].sum())
    with col4:
        if not filtered.empty:
            st.metric("Avg Distance (mi)", f"{filtered['distance_miles'].mean():.1f}")
        else:
            st.metric("Avg Distance (mi)", "N/A")

    st.markdown("---")

    # Display table
    if not filtered.empty:
        st.subheader(f"📊 Top Expansion Opportunities ({len(filtered)} ZIP codes)")

        # Format for display
        display_df = filtered[['zip_code', 'quote_count', 'nearest_serviced_zip', 'distance_miles', 'airport_code', 'region']].copy()
        display_df.columns = ['ZIP Code', 'Quote Count', 'Nearest Serviced ZIP', 'Distance (mi)', 'Airport Code', 'Region']
        display_df['Distance (mi)'] = display_df['Distance (mi)'].round(1)

        st.dataframe(display_df, use_container_width=True, hide_index=True)

        # Export button
        csv = filtered.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name="expansion_opportunities.csv",
            mime="text/csv"
        )

        # Summary by airport code
        st.subheader("📈 Summary by Airport Code")
        airport_summary = filtered.groupby('airport_code').agg({
            'zip_code': 'count',
            'quote_count': 'sum',
            'distance_miles': 'mean'
        }).reset_index()
        airport_summary.columns = ['Airport Code', 'ZIP Count', 'Unserviced Quotes', 'Avg Distance (mi)']
        # Add total quotes and % not serviced
        airport_summary['Total Quotes'] = airport_summary['Airport Code'].map(total_quotes_by_airport).fillna(0).astype(int)
        airport_summary['% Not Serviced'] = (airport_summary['Unserviced Quotes'] / airport_summary['Total Quotes'] * 100).round(1)
        airport_summary = airport_summary.sort_values('Unserviced Quotes', ascending=False)
        airport_summary['Avg Distance (mi)'] = airport_summary['Avg Distance (mi)'].round(1)
        # Reorder columns
        airport_summary = airport_summary[['Airport Code', 'ZIP Count', 'Unserviced Quotes', 'Total Quotes', '% Not Serviced', 'Avg Distance (mi)']]
        st.dataframe(airport_summary, use_container_width=True, hide_index=True)

        # Summary by region
        st.subheader("📈 Summary by Region")
        region_summary = filtered.groupby('region').agg({
            'zip_code': 'count',
            'quote_count': 'sum',
            'distance_miles': 'mean'
        }).reset_index()
        region_summary.columns = ['Region', 'ZIP Count', 'Unserviced Quotes', 'Avg Distance (mi)']
        # Add total quotes and % not serviced
        region_summary['Total Quotes'] = region_summary['Region'].map(total_quotes_by_region).fillna(0).astype(int)
        region_summary['% Not Serviced'] = (region_summary['Unserviced Quotes'] / region_summary['Total Quotes'] * 100).round(1)
        region_summary = region_summary.sort_values('Unserviced Quotes', ascending=False)
        region_summary['Avg Distance (mi)'] = region_summary['Avg Distance (mi)'].round(1)
        # Reorder columns
        region_summary = region_summary[['Region', 'ZIP Count', 'Unserviced Quotes', 'Total Quotes', '% Not Serviced', 'Avg Distance (mi)']]
        st.dataframe(region_summary, use_container_width=True, hide_index=True)
    else:
        st.info("No ZIP codes match the current filters. Try adjusting the minimum quote count or maximum distance.")

if __name__ == "__main__":
    main()

