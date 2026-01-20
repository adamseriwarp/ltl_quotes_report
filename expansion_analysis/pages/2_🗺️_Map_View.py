"""
Expansion Opportunities Map - Visual map of unserviced ZIP codes and crossdock locations.
"""

import streamlit as st
import pandas as pd
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from pathlib import Path
import pyproj

# Page config
st.set_page_config(page_title="Expansion Map", page_icon="🗺️", layout="wide")

# Paths
SHARED_DIR = Path(__file__).parent.parent.parent / 'shared'
CENTROIDS_FILE = SHARED_DIR / "zip_centroids.csv"
CROSSDOCKS_FILE = Path(__file__).parent.parent / "warp_xdocks_with_accl.csv"

# Coordinate transformer: EPSG:2163 (US National Atlas) -> EPSG:4326 (lat/lng)
transformer = pyproj.Transformer.from_crs("EPSG:2163", "EPSG:4326", always_xy=True)


@st.cache_data
def load_zip_centroids_latlng():
    """Load ZIP centroids and convert to lat/lng."""
    df = pd.read_csv(CENTROIDS_FILE, dtype={'zip_code': str})
    
    # Convert coordinates from EPSG:2163 to lat/lng
    lngs, lats = transformer.transform(df['centroid_x'].values, df['centroid_y'].values)
    df['lat'] = lats
    df['lng'] = lngs
    
    return df.set_index('zip_code')[['lat', 'lng']]


@st.cache_data
def load_crossdocks_latlng():
    """Load crossdock locations with lat/lng from ZIP centroids."""
    df = pd.read_csv(CROSSDOCKS_FILE)
    
    # Extract ZIP from address
    def extract_zip(address):
        if pd.isna(address):
            return None
        parts = [p.strip() for p in str(address).split(',')]
        if parts:
            zip_code = parts[-1].strip()[:5]
            if zip_code.isdigit() and len(zip_code) == 5:
                return zip_code
        return None
    
    df['zip_code'] = df['Address'].apply(extract_zip)
    df = df.dropna(subset=['zip_code'])
    df = df.rename(columns={'Name': 'dock_name'})
    
    # Get lat/lng from centroids
    centroids = load_zip_centroids_latlng()
    df = df.merge(centroids, left_on='zip_code', right_index=True, how='left')
    df = df.dropna(subset=['lat', 'lng'])
    
    return df[['dock_name', 'zip_code', 'lat', 'lng']]


# Import shared functions from main app
import sys
import re
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from shared.drive_client import DriveClient
from weekly_ltl_metrics_report.report_generator import load_csvs_from_folder


@st.cache_resource
def get_drive_client():
    return DriveClient()


@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_available_year_weeks(_client) -> tuple:
    """Get list of available (year, week) tuples from Drive, sorted chronologically."""
    try:
        folders = _client.search_folders("Quotes")
    except Exception as e:
        st.error(f"Failed to connect to Google Drive: {e}")
        return [], {}

    week_pattern = re.compile(r'^W(\d{2})(\d{2})\s+Quotes$')
    year_weeks = []
    folder_map = {}  # Map (year, week) to folder info
    for folder in folders:
        match = week_pattern.match(folder['name'])
        if match:
            week_num = int(match.group(1))
            year = 2000 + int(match.group(2))
            year_weeks.append((year, week_num))
            folder_map[(year, week_num)] = folder
    return sorted(set(year_weeks)), folder_map


@st.cache_data
def load_quotes_data(_client, folder_ids):
    """Load quotes from selected folders using shared function."""
    all_quotes = []

    for folder_id, folder_name in folder_ids:
        print(f"Loading {folder_name}...")
        df = load_csvs_from_folder(_client, folder_id, folder_name)
        if not df.empty:
            all_quotes.append(df)

    if all_quotes:
        return pd.concat(all_quotes, ignore_index=True)
    return pd.DataFrame()


st.title("🗺️ Expansion Opportunities Map")
st.markdown("Visualize unserviced ZIP codes and crossdock locations")

# Load data
with st.spinner("Loading data..."):
    centroids_latlng = load_zip_centroids_latlng()
    crossdocks = load_crossdocks_latlng()

st.sidebar.header("🎯 Filters")

# Min quotes filter
min_quotes = st.sidebar.slider("Minimum Quotes", min_value=1, max_value=100, value=10)

# Max distance filter (miles)
max_distance = st.sidebar.slider("Max Distance to Crossdock (miles)", min_value=10, max_value=500, value=150)

# Load quote folders using the same method as main app
drive = get_drive_client()
available_weeks, folder_map = get_available_year_weeks(drive)

if not available_weeks:
    st.error("No quote folders found in Drive")
    st.stop()

# Week selection - format as "W## YYYY" for display
week_options = [f"W{w:02d} {y}" for y, w in available_weeks]
default_weeks = week_options[-2:] if len(week_options) >= 2 else week_options

selected_week_labels = st.sidebar.multiselect(
    "Select Weeks:",
    options=week_options,
    default=default_weeks
)

if not selected_week_labels:
    st.warning("Please select at least one week")
    st.stop()

# Convert selected labels back to folder info
selected_folder_ids = []
for label in selected_week_labels:
    # Parse "W## YYYY" format
    parts = label.split()
    week_num = int(parts[0][1:])
    year = int(parts[1])
    if (year, week_num) in folder_map:
        folder = folder_map[(year, week_num)]
        selected_folder_ids.append((folder['id'], folder['name']))

# Initialize session state for map data
if 'map_data' not in st.session_state:
    st.session_state.map_data = None
if 'map_crossdocks' not in st.session_state:
    st.session_state.map_crossdocks = None
if 'quotes_df_for_filter' not in st.session_state:
    st.session_state.quotes_df_for_filter = None

# Load data button (to populate customer filter)
if st.sidebar.button("📥 Load Data", type="secondary"):
    with st.spinner("Loading quotes..."):
        quotes_df = load_quotes_data(drive, tuple(selected_folder_ids))
        if not quotes_df.empty:
            st.session_state.quotes_df_for_filter = quotes_df
            st.sidebar.success(f"Loaded {len(quotes_df):,} quotes")
        else:
            st.error("No quote data loaded")

# Customer filter (only show if data is loaded)
selected_customers = []
if st.session_state.quotes_df_for_filter is not None:
    quotes_df = st.session_state.quotes_df_for_filter
    if 'customer' in quotes_df.columns:
        all_customers = sorted(quotes_df['customer'].dropna().unique().tolist())
        selected_customers = st.sidebar.multiselect(
            "Filter by Customer:",
            options=all_customers,
            default=[],
            placeholder="All customers"
        )

# Run button
if st.sidebar.button("🗺️ Generate Map", type="primary"):
    if st.session_state.quotes_df_for_filter is None:
        st.warning("Please click 'Load Data' first")
        st.stop()

    with st.spinner("Generating map..."):
        quotes_df = st.session_state.quotes_df_for_filter.copy()

        # Apply customer filter if selected
        if selected_customers:
            quotes_df = quotes_df[quotes_df['customer'].isin(selected_customers)]
            st.sidebar.caption(f"Filtered to {len(quotes_df):,} quotes from {len(selected_customers)} customer(s)")

        if quotes_df.empty:
            st.error("No quote data after filtering")
            st.stop()

        # Identify rated quotes (same logic as main app)
        quotes_df['is_rated'] = quotes_df['rate'].notna() & (quotes_df['rate'].astype(str).str.strip() != '')

        # Normalize ZIP codes (using same column names as main app: 'pickup Zip', 'dropoff Zip')
        quotes_df['pickup_zip_clean'] = quotes_df['pickup Zip'].astype(str).str.strip().str[:5]
        quotes_df['dropoff_zip_clean'] = quotes_df['dropoff Zip'].astype(str).str.strip().str[:5]

        # Get serviced ZIPs (where we have rated quotes)
        rated_df = quotes_df[quotes_df['is_rated']]
        serviced_zips = set()
        for col in ['pickup_zip_clean', 'dropoff_zip_clean']:
            serviced_zips.update(rated_df[col].dropna().unique())

        # Get unrated quotes
        unrated_df = quotes_df[~quotes_df['is_rated']]

        if unrated_df.empty:
            st.warning("No unrated quotes found in selected data")
            st.stop()

        # Count quotes by ZIP for unserviced locations
        pickup_counts = unrated_df[~unrated_df['pickup_zip_clean'].isin(serviced_zips)].groupby('pickup_zip_clean').size()
        dropoff_counts = unrated_df[~unrated_df['dropoff_zip_clean'].isin(serviced_zips)].groupby('dropoff_zip_clean').size()

        # Combine pickup and dropoff counts
        all_unserviced = pd.DataFrame({
            'pickup_quotes': pickup_counts,
            'dropoff_quotes': dropoff_counts
        }).fillna(0)
        all_unserviced['total_quotes'] = all_unserviced['pickup_quotes'] + all_unserviced['dropoff_quotes']
        all_unserviced = all_unserviced[all_unserviced['total_quotes'] >= min_quotes]

        # Add lat/lng for unserviced ZIPs
        all_unserviced = all_unserviced.merge(centroids_latlng, left_index=True, right_index=True, how='left')
        all_unserviced = all_unserviced.dropna(subset=['lat', 'lng'])

        # Calculate distance to nearest crossdock
        from shapely.geometry import Point
        from shapely import STRtree

        # Create crossdock points for distance calculation (in EPSG:2163 for accurate distance)
        centroids_raw = pd.read_csv(CENTROIDS_FILE, dtype={'zip_code': str}).set_index('zip_code')

        crossdock_points = []
        for _, row in crossdocks.iterrows():
            if row['zip_code'] in centroids_raw.index:
                x, y = centroids_raw.loc[row['zip_code'], ['centroid_x', 'centroid_y']]
                crossdock_points.append(Point(x, y))

        if crossdock_points:
            tree = STRtree(crossdock_points)

            def get_nearest_distance(zip_code):
                if zip_code in centroids_raw.index:
                    x, y = centroids_raw.loc[zip_code, ['centroid_x', 'centroid_y']]
                    pt = Point(x, y)
                    nearest = tree.nearest(pt)
                    dist_meters = pt.distance(crossdock_points[nearest])
                    return dist_meters * 0.000621371  # Convert to miles
                return None

            all_unserviced['distance_miles'] = [get_nearest_distance(z) for z in all_unserviced.index]
            all_unserviced = all_unserviced.dropna(subset=['distance_miles'])
            all_unserviced = all_unserviced[all_unserviced['distance_miles'] <= max_distance]

        # Store results in session state
        st.session_state.map_data = all_unserviced
        st.session_state.map_crossdocks = crossdocks

# Display map if data exists in session state
if st.session_state.map_data is not None and st.session_state.map_crossdocks is not None:
    all_unserviced = st.session_state.map_data
    crossdocks_display = st.session_state.map_crossdocks

    st.success(f"Found {len(all_unserviced)} unserviced ZIP codes meeting criteria")

    # Create map centered on US
    m = folium.Map(location=[39.8283, -98.5795], zoom_start=4, tiles='cartodbpositron')

    # Add crossdock markers (blue)
    for _, row in crossdocks_display.iterrows():
        folium.CircleMarker(
            location=[row['lat'], row['lng']],
            radius=8,
            color='#2563eb',
            fill=True,
            fillColor='#2563eb',
            fillOpacity=0.8,
            popup=f"<b>{row['dock_name']}</b><br>ZIP: {row['zip_code']}",
            tooltip=row['dock_name']
        ).add_to(m)

    # Add unserviced ZIP markers with clustering (orange)
    # Custom icon function to show sum of quotes instead of marker count
    icon_create_function = '''
    function(cluster) {
        var markers = cluster.getAllChildMarkers();
        var totalQuotes = 0;
        for (var i = 0; i < markers.length; i++) {
            totalQuotes += markers[i].options.quoteCount || 0;
        }

        // Format number with K suffix for thousands
        var displayNum = totalQuotes;
        if (totalQuotes >= 1000) {
            displayNum = (totalQuotes / 1000).toFixed(1) + 'K';
        }

        // Color based on total quotes
        var c = totalQuotes < 500 ? '#22c55e' :   // green
                totalQuotes < 2000 ? '#eab308' :   // yellow
                totalQuotes < 5000 ? '#f97316' :   // orange
                '#ef4444';                          // red

        var size = totalQuotes < 500 ? 30 :
                   totalQuotes < 2000 ? 40 :
                   totalQuotes < 5000 ? 50 : 60;

        return L.divIcon({
            html: '<div style="background-color:' + c + '; width:' + size + 'px; height:' + size + 'px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:white; font-weight:bold; font-size:11px; border:2px solid white; box-shadow:0 2px 5px rgba(0,0,0,0.3);">' + displayNum + '</div>',
            className: 'marker-cluster',
            iconSize: L.point(size, size)
        });
    }
    '''

    if len(all_unserviced) > 0:
        marker_cluster = MarkerCluster(
            name='Unserviced ZIPs',
            icon_create_function=icon_create_function,
            options={
                'maxClusterRadius': 50,
                'disableClusteringAtZoom': 8,
                'spiderfyOnMaxZoom': True
            }
        ).add_to(m)

        max_quotes = all_unserviced['total_quotes'].max()
        for zip_code, row in all_unserviced.iterrows():
            radius = 4 + (row['total_quotes'] / max_quotes) * 12  # Scale 4-16
            # Use Marker instead of CircleMarker to pass custom quoteCount option
            folium.Marker(
                location=[row['lat'], row['lng']],
                icon=folium.DivIcon(
                    html=f'<div style="background-color:#ea580c; width:{int(radius*2)}px; height:{int(radius*2)}px; border-radius:50%; border:1px solid #c2410c;"></div>',
                    icon_size=(int(radius*2), int(radius*2)),
                    icon_anchor=(int(radius), int(radius))
                ),
                popup=f"<b>ZIP: {zip_code}</b><br>Quotes: {int(row['total_quotes'])}<br>Distance: {row['distance_miles']:.1f} mi",
                tooltip=f"{zip_code}: {int(row['total_quotes'])} quotes",
                quoteCount=int(row['total_quotes'])  # Custom property for clustering
            ).add_to(marker_cluster)

    # Add legend
    legend_html = '''
    <div style="position: fixed; bottom: 50px; left: 50px; z-index: 1000; background-color: white;
                padding: 10px; border-radius: 5px; border: 2px solid gray;">
        <p><span style="color: #2563eb;">●</span> Crossdock</p>
        <p><span style="color: #ea580c;">●</span> Unserviced ZIP</p>
    </div>
    '''
    m.get_root().html.add_child(folium.Element(legend_html))

    # Display map (use key and returned_objects=[] to prevent rerun issues)
    st_folium(m, width=1200, height=600, key="expansion_map", returned_objects=[])

    # Show data table below
    st.subheader("📋 Unserviced ZIPs on Map")
    display_df = all_unserviced.reset_index().rename(columns={'index': 'zip_code'})
    display_df['total_quotes'] = display_df['total_quotes'].astype(int)
    display_df['distance_miles'] = display_df['distance_miles'].round(1)
    st.dataframe(
        display_df[['zip_code', 'total_quotes', 'distance_miles']].sort_values('total_quotes', ascending=False),
        hide_index=True
    )
else:
    st.info("👈 Adjust filters in the sidebar and click 'Generate Map' to visualize expansion opportunities")

