import dash
from dash import dcc, html, Input, Output, State
import pandas as pd
import plotly.express as px
import requests
import dash_leaflet as dl
import dash_bootstrap_components as dbc
from dash_extensions.javascript import assign

# Load the Excel file (get sheet names dynamically)
excel_file = "KS.xlsx"
xls = pd.ExcelFile(excel_file)
sheet_names = xls.sheet_names  # Get available sheets

# Function to get Latitude and Longitude from Postal Code using OneMap API
def get_lat_lon(postal_code):
    postal_code = str(postal_code)
    url = f"https://www.onemap.gov.sg/api/common/elastic/search?searchVal={postal_code}&returnGeom=Y&getAddrDetails=N"
    response = requests.get(url).json()
    
    if response["found"] > 0:
        lat = response["results"][0]["LATITUDE"]
        lon = response["results"][0]["LONGITUDE"]
        return float(lat), float(lon)
    return None, None

# Function to load a sheet and process it
def load_sheet(sheet_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Convert "Event Date" column to datetime format
    df["Event Date"] = pd.to_datetime(df["Event Date"], errors='coerce')
    
    df["Latitude"], df["Longitude"] = zip(*df["Postal Code"].apply(get_lat_lon))
    return df

# Initialize Dash app with Bootstrap
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# Heatmap function
geojson_heatmap = assign("""
function(feature, latlng) { 
    return L.circleMarker(latlng, { 
        radius: 8, 
        fillColor: 'red', 
        color: 'red', 
        weight: 1, 
        opacity: 1, 
        fillOpacity: 0.6 
    }); 
}
""")

# App layout
app.layout = dbc.Container([
    # Title
    dbc.Row([dbc.Col(html.H1("📊 Kampung Spirit Dashboard", className="text-center"), width=12)], className="mb-4"),

    # Sheet Selector
    dbc.Row([
        dbc.Col(html.Label("Select Data Sheet:"), width=3, className="text-right"),
        dbc.Col(dcc.Dropdown(
            id="sheet-selector",
            options=[{"label": sheet, "value": sheet} for sheet in sheet_names],
            value=sheet_names[0],
            clearable=False
        ), width=6)
    ], className="mb-4 justify-content-center"),

    # Date Filter
    dbc.Row([
        dbc.Col(html.Label("Filter by Event Date:"), width=3, className="text-right"),
        dbc.Col(dcc.DatePickerRange(
            id="date-filter",
            display_format="YYYY-MM-DD"
        ), width=6)
    ], className="mb-4 justify-content-center"),

    # Attendance & Attrition Rate Display
    dbc.Row(id="attendance-row", className="mb-4 justify-content-center"),

    # Averages Row
    dbc.Row(id="averages-row", className="mb-4 justify-content-center"),

    # Graphs Layout (2 Columns)
    dbc.Row([
        # Left Column
        dbc.Col([
            dbc.Card([dbc.CardHeader(html.H5("📈 Age Distribution")), dbc.CardBody(dcc.Graph(id="age-dist"))], className="mb-4"),
            dbc.Card([dbc.CardHeader(html.H5("⚧ Gender Distribution")), dbc.CardBody(dcc.Graph(id="gender-dist"))], className="mb-4"),
            dbc.Card([dbc.CardHeader(html.H5("🌎 Racial Distribution")), dbc.CardBody(dcc.Graph(id="race-dist"))], className="mb-4"),
            dbc.Card([dbc.CardHeader(html.H5("🤝 New Neighbors Met")), dbc.CardBody(dcc.Graph(id="new-neighbors"))], className="mb-4"),
        ], width=6),

        # Right Column
        dbc.Col([
            dbc.Card([dbc.CardHeader(html.H5("⭐ Event Rating")), dbc.CardBody(dcc.Graph(id="event-rating"))], className="mb-4"),
            dbc.Card([dbc.CardHeader(html.H5("🌎 Marketing Channels")), dbc.CardBody(dcc.Graph(id="mark-dist"))], className="mb-4"),
            dbc.Card([dbc.CardHeader(html.H5("📈 Net Promoter Score")), dbc.CardBody(dcc.Graph(id="net-prom"))], className="mb-4"),
            dbc.Card([dbc.CardHeader(html.H5("🧑‍🤝‍🧑 Better Knowledge of Neighbors")), dbc.CardBody(dcc.Graph(id="better-know"))], className="mb-4"),
        ], width=6),
    ]),

    # Heatmap
    dbc.Row([dbc.Col(html.H3("📍 Singapore Participant Heatmap", className="text-center"), width=12)], className="mb-2"),
    dbc.Row([dbc.Col(dl.Map(id="heatmap", style={'width': '100%', 'height': '500px'}, center=[1.3521, 103.8198], zoom=12), width=12)], className="mb-4"),
    
    # Data Table
    dbc.Row([dbc.Col(dcc.Dropdown(id="data-filter", multi=True, placeholder="Select Columns"), width=12)], className="mb-2"),
    dbc.Row([dbc.Col(html.Div(id="data-table"), width=12)], className="mb-4"),
], fluid=True)

# Callback to update dashboard
@app.callback(
    [
        Output("attendance-row", "children"),
        Output("averages-row", "children"),
        Output("age-dist", "figure"),
        Output("gender-dist", "figure"),
        Output("race-dist", "figure"),
        Output("new-neighbors", "figure"),
        Output("better-know", "figure"),
        Output("event-rating", "figure"),
        Output("net-prom", "figure"),
        Output("mark-dist", "figure"),
        Output("heatmap", "children"),
        Output("data-filter", "options"),
        Output("data-table", "children")
    ],
    [
        Input("sheet-selector", "value"),
        Input("date-filter", "start_date"),
        Input("date-filter", "end_date"),
    ]
)
def update_dashboard(sheet_name, start_date, end_date):
    df = load_sheet(sheet_name)  

    # Convert dates
    if start_date and end_date:
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
        df["Event Date"] = pd.to_datetime(df["Event Date"], errors='coerce') 
        df = df[(df["Event Date"] >= start_date) & (df["Event Date"] <= end_date)]

    averages = df.mean(numeric_only=True)

    # Attendance Stats
    total_attendance = df["Attendance"].sum()
    total_registrants = df["Attendance"].count()
    attrition_rate = ((total_registrants - total_attendance) / total_registrants) * 100 if total_registrants > 0 else 0

    attendance_row = dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("🎟️ Total Attendance"), dbc.CardBody(html.H4(f"{total_attendance}", className="text-center"))]), width=3),
        dbc.Col(dbc.Card([dbc.CardHeader("📉 Attrition Rate"), dbc.CardBody(html.H4(f"{attrition_rate:.2f}%", className="text-center"))]), width=3),
    ], className="justify-content-center")

    # Generate averages row dynamically
    averages_row = dbc.Row([
        dbc.Col(dbc.Card([dbc.CardHeader("📊 Avg Age"), dbc.CardBody(html.H4(f"{averages['Age']:.2f}"))]), width=2),
        dbc.Col(dbc.Card([dbc.CardHeader("🤝 Avg Neighbors Met"), dbc.CardBody(html.H4(f"{averages['How many new neighbours met?']:.2f}"))]), width=2),
        dbc.Col(dbc.Card([dbc.CardHeader("🧑‍🤝‍🧑 Avg Knowledge"), dbc.CardBody(html.H4(f"{averages['How much better do you know your neighbours?']:.2f}"))]), width=2),
        dbc.Col(dbc.Card([dbc.CardHeader("⭐ Avg Rating"), dbc.CardBody(html.H4(f"{averages['Rating of the whole event.']:.2f}"))]), width=2),
        dbc.Col(dbc.Card([dbc.CardHeader("📈 Net Promoter"), dbc.CardBody(html.H4(f"{averages['How likely are you to promote this event to your friend?']:.2f}"))]), width=2),
    ], className="justify-content-center"),

    # Update graphs
    age_fig = px.histogram(df, x="Age", nbins=10, title="Age Distribution")
    gender_fig = px.pie(df, names="Gender", title="Gender Distribution")
    race_fig = px.pie(df, names="Race", title="Racial Distribution")
    neighbors_fig = px.histogram(df, x="How many new neighbours met?", nbins=10, title="New Neighbors Met")
    better_know_fig = px.histogram(df, x="How much better do you know your neighbours?", nbins=5, title="Better Knowledge of Neighbors")
    event_rating_fig = px.histogram(df, x="Rating of the whole event.", nbins=10, title="Event Rating")
    net_prom_fig = px.histogram(df, x="How likely are you to promote this event to your friend?", nbins=10, title="Net Promoter Score")
    marketing_fig = px.pie(df, names="Marketing", title="How did you find out about our event?")

    # Update heatmap
    heatmap_layer = dl.Map([
        dl.TileLayer(),
        dl.LayerGroup([
            dl.GeoJSON(data={ 
                "type": "FeatureCollection", 
                "features": [{"type": "Feature", "geometry": {"type": "Point", "coordinates": [lon, lat]}} 
                             for lat, lon in zip(df["Latitude"], df["Longitude"])] 
            }, pointToLayer=geojson_heatmap)
        ])
    ], style={'width': '100%', 'height': '500px'}, center=[1.3521, 103.8198], zoom=12)

    # Update table dropdown
    column_options = [{"label": col, "value": col} for col in df.columns]

    return attendance_row, averages_row, age_fig, gender_fig, race_fig, neighbors_fig, better_know_fig, event_rating_fig, net_prom_fig, marketing_fig, heatmap_layer, column_options, ""


# Run the app
if __name__ == "__main__":
    app.run(debug=True, host="192.168.1.122", port=8050)
