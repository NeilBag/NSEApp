import os
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output
import pandas as pd
from nsepython import nsefetch
import io
import base64
import time
import functools
from flask_caching import Cache

# Initialize the Dash app
app = Dash(__name__)
server = app.server  # Expose the server variable for Gunicorn

# Set up Flask-Caching
cache = Cache(app.server, config={'CACHE_TYPE': 'SimpleCache'})

# Cache timeout (in seconds)
CACHE_TIMEOUT = 5 * 5  # Cache for 5 minutes

# Function to get stock data from all indices
@cache.memoize(timeout=CACHE_TIMEOUT)
def get_all_indices_data():
    indices = [
        "NIFTY 50", "NIFTY NEXT 50", "NIFTY MIDCAP 50", "NIFTY SMALLCAP 100", 
        "NIFTY 100", "NIFTY 200", "NIFTY 500", "NIFTY MIDCAP 100",
        "NIFTY MIDCAP 150", "NIFTY SMALLCAP 250", "NIFTY MIDSML 400"
    ]
    
    all_data = []
    
    for index in indices:
        url = f"https://www.nseindia.com/api/equity-stockIndices?index={index.replace(' ', '%20')}"
        data = nsefetch(url)
        df = pd.DataFrame(data['data'])
        df['index'] = index  # Add a column to identify the index
        all_data.append(df)
    
    # Combine all data into a single DataFrame
    combined_df = pd.concat(all_data, ignore_index=True)
    
    return combined_df

# Function to process the stock data and sort by 52-Week High
@functools.lru_cache(maxsize=1)
def get_stock_data():
    df = get_all_indices_data()
    
    # Convert relevant columns to numeric, handling missing columns
    df['totalTradedVolume'] = pd.to_numeric(df['totalTradedVolume'], errors='coerce')
    df['lastPrice'] = pd.to_numeric(df['lastPrice'], errors='coerce')
    df['previousClose'] = pd.to_numeric(df['previousClose'], errors='coerce')
    df['open'] = pd.to_numeric(df['open'], errors='coerce')
    df['dayHigh'] = pd.to_numeric(df['dayHigh'], errors='coerce')
    df['dayLow'] = pd.to_numeric(df['dayLow'], errors='coerce')
    df['yearHigh'] = pd.to_numeric(df['yearHigh'], errors='coerce')
    df['yearLow'] = pd.to_numeric(df['yearLow'], errors='coerce')

    # Handle missing columns
    if 'totalMktCap' in df.columns:
        df['totalMktCap'] = pd.to_numeric(df['totalMktCap'], errors='coerce')
    else:
        df['totalMktCap'] = None

    if 'pe' in df.columns:
        df['pe'] = pd.to_numeric(df['pe'], errors='coerce')
    else:
        df['pe'] = None

    if 'vwap' in df.columns:
        df['vwap'] = pd.to_numeric(df['vwap'], errors='coerce')
    else:
        df['vwap'] = None

    # Calculate additional metrics
    df['priceChange'] = df['lastPrice'] - df['previousClose']
    df['percentageChange'] = (df['priceChange'] / df['previousClose']) * 100

    # Sort the DataFrame by 'yearHigh' (52-Week High) in descending order
    sorted_df = df.sort_values(by='yearHigh', ascending=False)
    
    # Select relevant columns
    sorted_df = sorted_df[['symbol', 'index', 'lastPrice', 'previousClose', 'open', 'dayHigh', 
                           'dayLow', 'priceChange', 'percentageChange', 'yearHigh',
                           'yearLow', 'totalMktCap', 'pe', 'vwap']]
    
    return sorted_df

# Function to generate Excel download link
def generate_excel_download_link(df):
    # Write DataFrame to a BytesIO object
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Stocks')
    writer.save()
    output.seek(0)

    # Encode the output as a base64 string
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode('utf-8')
    
    return f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"

# Layout of the Dash app
app.layout = html.Div([
    html.H1("NSE Stocks Sorted by 52-Week High"),
    
    # Dropdown to filter by stock symbol
    dcc.Dropdown(
        id='symbol-dropdown',
        options=[{'label': symbol, 'value': symbol} for symbol in get_stock_data()['symbol'].unique()],
        placeholder="Select a stock symbol",
        multi=False,
        searchable=True,
        clearable=True
    ),
    
    # Button to download the data as Excel
    html.Button("Download as Excel", id="btn-download"),
    dcc.Download(id="download-dataframe-xlsx"),

    # Interval component for live refresh
    dcc.Interval(
        id='interval-component',
        interval=60*1000,  # Refresh every minute
        n_intervals=0
    ),
    
    # Data table for displaying stock data with pagination
    dash_table.DataTable(
        id='stock-table',
        columns=[
            {'name': 'Symbol', 'id': 'symbol'},
            {'name': 'Index', 'id': 'index'},
            {'name': 'Last Traded Price (LTP)', 'id': 'lastPrice'},
            {'name': 'Previous Close', 'id': 'previousClose'},
            {'name': 'Open Price', 'id': 'open'},
            {'name': "Day's High", 'id': 'dayHigh'},
            {'name': "Day's Low", 'id': 'dayLow'},
            {'name': 'Price Change', 'id': 'priceChange'},
            {'name': 'Percentage Change (%)', 'id': 'percentageChange'},
            {'name': '52-Week High', 'id': 'yearHigh'},
            {'name': '52-Week Low', 'id': 'yearLow'},
            {'name': 'Market Capitalization', 'id': 'totalMktCap'},
            {'name': 'P/E Ratio', 'id': 'pe'},
            {'name': 'VWAP', 'id': 'vwap'}
        ],
        data=get_stock_data().to_dict('records'),
        page_size=20,  # Number of rows per page
        page_action='native',  # Handle pagination natively
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left'},
    )
])

# Callback to update the data table based on dropdown selection
@app.callback(
    Output('stock-table', 'data'),
    Input('symbol-dropdown', 'value'),
    Input('interval-component', 'n_intervals')
)
def update_table(selected_symbol, n):
    df = get_stock_data()
    if selected_symbol:
        df = df[df['symbol'] == selected_symbol]
    return df.to_dict('records')

# Callback to export data to Excel
@app.callback(
    Output("download-dataframe-xlsx", "data"),
    Input("btn-download", "n_clicks"),
    prevent_initial_call=True,
)
def download_as_excel(n_clicks):
    df = get_stock_data()
    return dcc.send_data_frame(df.to_excel, "NSE_Stocks_Sorted_by_52_Week_High.xlsx", index=False)

# Add a simple route for testing
@server.route('/health')
def health_check():
    return "Server is running!", 200
# Run the app
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    app.run_server(debug=False, host='0.0.0.0', port=port)
