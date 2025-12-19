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
def load_quotes_data(_client, selected_weeks: tuple) -> pd.DataFrame:
    """Load quote data for selected weeks."""
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
        if wf['week'] in selected_weeks:
            print(f"Loading {wf['name']}...")
            df = load_csvs_from_folder(_client, wf['id'], wf['name'])
            if not df.empty:
                df['week'] = wf['week']
                all_data.append(df)
    
    if not all_data:
        return pd.DataFrame()
    
    return pd.concat(all_data, ignore_index=True)

def get_available_weeks(client) -> list[int]:
    """Get list of available weeks from Drive."""
    folders = client.search_folders("Quotes")
    week_pattern = re.compile(r'^W(\d{2})(\d{2})\s+Quotes$')
    weeks = []
    for folder in folders:
        match = week_pattern.match(folder['name'])
        if match:
            weeks.append(int(match.group(1)))
    return sorted(set(weeks))

def analyze_expansion_opportunities(df: pd.DataFrame, centroids_df: pd.DataFrame,
                                     zip_mapping: dict, region_mapping: dict) -> pd.DataFrame:
    """Analyze unserviced zip codes for expansion opportunities."""
    from shapely.geometry import Point

    # Identify rated vs unrated quotes
    df['is_rated'] = df['rate'].notna() & (df['rate'].astype(str).str.strip() != '')

    # Get all unique zip codes from rated quotes (these are "serviced")
    rated_df = df[df['is_rated']]
    serviced_zips = set()
    for col in ['pickup Zip', 'dropoff Zip']:
        if col in rated_df.columns:
            serviced_zips.update(rated_df[col].dropna().astype(str).str.strip().str[:5].unique())

    # Get unrated quotes
    unrated_df = df[~df['is_rated']]

    if unrated_df.empty:
        return pd.DataFrame(), serviced_zips

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
                # Convert from meters to km (coordinates are in EPSG:2163 meters)
                distances.append(min_dist / 1000)
                nearest_zips.append(nearest_zip)
            else:
                distances.append(np.nan)
                nearest_zips.append(None)
        results['distance_km'] = distances
        results['nearest_serviced_zip'] = nearest_zips
    else:
        results['distance_km'] = np.nan
        results['nearest_serviced_zip'] = None

    # Add airport code and region
    results['airport_code'] = results['zip_code'].apply(lambda x: get_airport_code(x, zip_mapping))
    results['region'] = results['airport_code'].apply(lambda x: get_region(x, region_mapping))

    # Sort by quote count descending
    results = results.sort_values('quote_count', ascending=False)

    return results, serviced_zips

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

    # Get available weeks
    available_weeks = get_available_weeks(client)
    if not available_weeks:
        st.error("No quote folders found")
        return

    # Week range selector - default to last 4 weeks for faster loading
    st.sidebar.subheader("📅 Week Range")
    default_from = max(max(available_weeks) - 3, min(available_weeks))
    min_week = st.sidebar.number_input("From Week:", min_value=min(available_weeks),
                                        max_value=max(available_weeks), value=default_from)
    max_week = st.sidebar.number_input("To Week:", min_value=min(available_weeks),
                                        max_value=max(available_weeks), value=max(available_weeks))

    selected_weeks = tuple(range(int(min_week), int(max_week) + 1))

    # Filter controls
    st.sidebar.subheader("🎯 Filters")
    min_quotes = st.sidebar.slider("Minimum Quote Count:", min_value=1, max_value=500, value=10)
    max_distance = st.sidebar.slider("Maximum Distance to Serviced ZIP (km):",
                                      min_value=1, max_value=500, value=100)

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
    with st.spinner(f"Loading quotes for weeks {min_week}-{max_week}..."):
        try:
            quotes_df = load_quotes_data(client, selected_weeks)
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
            results, serviced_zips = analyze_expansion_opportunities(quotes_df, centroids_df, zip_mapping, region_mapping)
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
        (results['distance_km'].notna()) &
        (results['distance_km'] <= max_distance)
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
            st.metric("Avg Distance (km)", f"{filtered['distance_km'].mean():.1f}")
        else:
            st.metric("Avg Distance (km)", "N/A")

    st.markdown("---")

    # Display table
    if not filtered.empty:
        st.subheader(f"📊 Top Expansion Opportunities ({len(filtered)} ZIP codes)")

        # Format for display
        display_df = filtered[['zip_code', 'quote_count', 'nearest_serviced_zip', 'distance_km', 'airport_code', 'region']].copy()
        display_df.columns = ['ZIP Code', 'Quote Count', 'Nearest Serviced ZIP', 'Distance (km)', 'Airport Code', 'Region']
        display_df['Distance (km)'] = display_df['Distance (km)'].round(1)

        st.dataframe(display_df, use_container_width=True, hide_index=True)

        # Export button
        csv = filtered.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name="expansion_opportunities.csv",
            mime="text/csv"
        )

        # Summary by region
        st.subheader("📈 Summary by Region")
        region_summary = filtered.groupby('region').agg({
            'zip_code': 'count',
            'quote_count': 'sum',
            'distance_km': 'mean'
        }).reset_index()
        region_summary.columns = ['Region', 'ZIP Count', 'Total Quotes', 'Avg Distance (km)']
        region_summary = region_summary.sort_values('Total Quotes', ascending=False)
        region_summary['Avg Distance (km)'] = region_summary['Avg Distance (km)'].round(1)
        st.dataframe(region_summary, use_container_width=True, hide_index=True)
    else:
        st.info("No ZIP codes match the current filters. Try adjusting the minimum quote count or maximum distance.")

if __name__ == "__main__":
    main()

