import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import pandas as pd
import base64
import io
import random
import string

app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server
app.title = "System Thinking - Excel"

# Function to generate a random ID for DataTable
def generate_random_id(size=8):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=size))

app.layout = html.Div(
    style={
        "backgroundColor": "#f0f8ff",  # Light Alice Blue background
        "padding": "20px",
        "borderRadius": "10px",
        "maxWidth": "800px",
        "margin": "auto",
        "boxShadow": "0 4px 8px rgba(0,0,0,0.2)",
        "overflow": "hidden"  # Ensure no overflow from the parent div
    },
    children=[
        html.H1(
            "System Thinking - Excel",
            style={"textAlign": "center", "color": "#0074D9", "fontFamily": "Arial, sans-serif"},
        ),
        dcc.Upload(
            id="upload-data",
            children=html.Div(["Drag and Drop or ", html.A("Select Files")]),
            style={
                "width": "100%",
                "height": "60px",
                "lineHeight": "60px",
                "borderWidth": "1px",
                "borderStyle": "dashed",
                "borderRadius": "5px",
                "textAlign": "center",
                "margin": "10px",
            },
            multiple=False,
        ),
        dcc.Dropdown(
            id='file-source-dropdown',
            options=[
                {'label': html.Div([
                    html.Img(src='/assets/google_drive.png', style={'height': '20px', 'verticalAlign': 'middle'}),
                    ' Google Drive'
                ]), 'value': 'google'},
                {'label': html.Div([
                    html.Img(src='/assets/dropbox.png', style={'height': '20px', 'verticalAlign': 'middle'}),
                    ' Dropbox'
                ]), 'value': 'dropbox'},
                {'label': html.Div([
                    html.Img(src='/assets/computer.png', style={'height': '20px', 'verticalAlign': 'middle'}),
                    ' Local Computer'
                ]), 'value': 'local'},
            ],
            value='local',
            placeholder="Select file source",
        ),
        html.Div(id='redirect-link'),
        dcc.Input(id="sheet-name", type="text", placeholder="Sheet Name"),
        dcc.Input(id="header-row", type="number", placeholder="Header Row Number"),
        html.Button("Load Data", id="load-data-button", n_clicks=0),
        html.Div(id="output-data-upload"),

        # Container for the DataTable with scrolling
        html.Div(
            id="data-table-container",
            style={
                "overflowY": "auto",  # Enable vertical scrolling
                "maxHeight": "400px",  # Set a maximum height for the table container
                "margin": "10px 0"  # Add some margin
            }
        ),
        dcc.Dropdown(
            id="column-dropdown", multi=True, placeholder="Select columns to sort"
        ),
        dcc.RadioItems(
            id="sort-order",
            options=[
                {"label": "Ascending", "value": "asc"},
                {"label": "Descending", "value": "desc"},
            ],
            value="asc",
        ),
        html.Button("Sort Data", id="sort-button", n_clicks=0),
        html.Div(id="sorted-data-table-container"),
        dcc.Dropdown(id="group-dropdown", placeholder="Select Group"),
        dcc.Input(
            id="classification-value", type="text", placeholder="Classification Value"
        ),
        html.Button("Update Group", id="update-group-button", n_clicks=0),
        html.Div(id="updated-group-table-container"),
        dcc.Input(
            id="first-blank-row-input", type="text", placeholder="First Blank Row"
        ),
        html.Button("Update Asset Codes", id="update-asset-codes-button", n_clicks=0),
        html.Div(id="updated-asset-table-container"),
        html.Button("Download Data as Excel", id="download-excel-button", n_clicks=0),
        dcc.Download(id="download-data-link"),
    ]
)

# Global variable to store the latest DataTable ID and DataFrame
latest_table_id = None
latest_table_data = None

def parse_contents(contents, filename, sheet_name, header_row):
    content_type, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    try:
        if "csv" in filename:
            df = pd.read_csv(io.StringIO(decoded.decode("utf-8")))
        elif "xls" in filename or "xlsx" in filename:
            df = pd.read_excel(
                io.BytesIO(decoded), sheet_name=sheet_name, header=header_row
            )
        else:
            return html.Div(["Unsupported file format: ", filename])
    except Exception as e:
        return html.Div(["There was an error processing this file. Error: ", str(e)])
    return df

@app.callback(
    Output('redirect-link', 'children'),
    Input('file-source-dropdown', 'value')
)
def redirect_user(file_source):
    if file_source == 'google':
        return dcc.Location(href='https://drive.google.com', id='go-to-google', refresh=True)
    elif file_source == 'dropbox':
        return dcc.Location(href='https://www.dropbox.com', id='go-to-dropbox', refresh=True)
    return ""

@app.callback(
    [
        Output("output-data-upload", "children"),
        Output("data-table-container", "children"),
        Output("column-dropdown", "options"),
        Output("group-dropdown", "options"),
    ],
    [Input("load-data-button", "n_clicks")],
    [
        State("upload-data", "contents"),
        State("upload-data", "filename"),
        State("sheet-name", "value"),
        State("header-row", "value"),
    ],
)
def update_output(n_clicks, contents, filename, sheet_name, header_row):
    if contents is not None:
        df = parse_contents(contents, filename, sheet_name, header_row)
        if isinstance(df, pd.DataFrame):
            global latest_table_data  # Access global variable to store the DataFrame
            latest_table_data = df  # Store the DataFrame for later use
            column_options = [{"label": col, "value": col} for col in df.columns]
            group_options = [{"label": grp, "value": grp} for grp in df["Site"].unique()]
            return (
                html.Div(
                    [
                        html.H5(filename),
                        dash_table.DataTable(
                            data=df.to_dict("records"),
                            columns=[{"name": i, "id": i} for i in df.columns],
                            id="data-table",
                            page_current=0,
                            page_size=10,
                            page_action="native",
                            style_cell={
                                'fontFamily': 'Arial, sans-serif',
                                'fontSize': '14px',
                                'textAlign': 'left',
                                'padding': '10px',
                            },
                            style_header={
                                'backgroundColor': '#0074D9',
                                'color': 'white',
                                'fontWeight': 'bold',
                                'fontSize': '16px'
                            },
                            style_table={
                                "overflowY": "auto",  # Enable vertical scrolling
                                "maxHeight": "400px",  # Set a maximum height for the table container
                                "margin": "10px 0"
                            },  # Add some margin
                        ),
                    ]
                ),
                dash_table.DataTable(
                    data=df.to_dict("records"),
                    columns=[{"name": i, "id": i} for i in df.columns],
                    id="data-table",
                    page_current=0,
                    page_size=10,
                    page_action="native",
                    style_cell={
                        'fontFamily': 'Arial, sans-serif',
                        'fontSize': '14px',
                        'textAlign': 'left',
                        'padding': '10px',
                    },
                    style_header={
                        'backgroundColor': '#0074D9',
                        'color': 'white',
                        'fontWeight': 'bold',
                        'fontSize': '16px'
                    },
                ),
                column_options,
                group_options,
            )
    return html.Div(), html.Div(), [], []

@app.callback(
    Output("sorted-data-table-container", "children"),
    [Input("sort-button", "n_clicks")],
    [
        State("data-table", "data"),
        State("column-dropdown", "value"),
        State("sort-order", "value"),
    ],
)
def sort_data(n_clicks, data, selected_columns, order):
    if data is not None and selected_columns is not None:
        df = pd.DataFrame(data)
        df[selected_columns] = df[selected_columns].apply(
            pd.to_numeric, errors="ignore"
        )
        df = df.sort_values(by=selected_columns, ascending=(order == "asc"))

        grouped = df.groupby("Site")
        new_rows = []
        for name, group in grouped:
            new_rows.append(group)
            if group["Site"].duplicated().any():
                empty_row = {col: "" for col in df.columns}
                empty_row["Quantity"] = group["Quantity"].sum()
                new_row = pd.DataFrame([empty_row])
                for col in group.columns:
                    if col not in ["Asset Code", "Quantity"]:
                        new_row[col] = group.iloc[0][col]
                new_rows.append(new_row)
        df = pd.concat(new_rows).reset_index(drop=True)
        return dash_table.DataTable(
            data=df.to_dict("records"),
            columns=[{"name": i, "id": i} for i in df.columns],
            id="sorted-data-table",
            page_current=0,
            page_size=30,
            page_action="native",
            style_cell={
                'fontFamily': 'Arial, sans-serif',
                'fontSize': '14px',
                'textAlign': 'left',
                'padding': '10px',
            },
            style_header={
                'backgroundColor': '#0074D9',
                'color': 'white',
                'fontWeight': 'bold',
                'fontSize': '16px'
            },
            style_table={ 
                "overflowY": "auto",  # Enable vertical scrolling
                "maxHeight": "400px",  # Set a maximum height for the table container
                "margin": "10px 0"
            },  # Add some margin
        )
    return dash_table.DataTable()

@app.callback(
    Output("updated-group-table-container", "children"),
    [Input("update-group-button", "n_clicks")],
    [
        State("sorted-data-table", "data"),
        State("group-dropdown", "value"),
        State("classification-value", "value"),
        State("updated-group-table-container", "children")
    ],
)
def update_group(n_clicks, data, group, classification_value, existing_table):
    global latest_table_id  # Use the global variable to save the latest DataTable ID
    global latest_table_data  # Access the stored DataFrame
    if data is not None and group and classification_value:
        df = pd.DataFrame(data)
        random_id = generate_random_id()  # Generate a random ID for the DataTable
        df.loc[df["Site"] == group, "Group.1"] = classification_value
        df["Group.1"] = df["Group.1"].fillna("")

        # Save the latest table ID and data for use in subsequent updates
        latest_table_id = random_id
        latest_table_data = df  # Update latest_table_data with the modified DataFrame

        return [
            dash_table.DataTable(
                data=df.to_dict("records"),
                columns=[{"name": i, "id": i} for i in df.columns],
                id=random_id,
                page_current=0,
                page_size=10,
                page_action="native",
                style_cell={
                    'fontFamily': 'Arial, sans-serif',
                    'fontSize': '14px',
                    'textAlign': 'left',
                    'padding': '10px',
                },
                style_header={
                    'backgroundColor': '#0074D9',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'fontSize': '16px'
                },
                style_table={
                    "overflowY": "auto",  # Enable vertical scrolling
                    "maxHeight": "400px",  # Set a maximum height for the table container
                    "margin": "10px 0"
                },  # Add some margin
            ),
            html.Div([
                html.Button('Yes', id='confirm-yes', n_clicks=0),
                html.Button('No', id='confirm-no', n_clicks=0),
                html.Div(id='confirmation-message')
            ])
        ]
    return dash_table.DataTable()

@app.callback(
    Output('confirmation-message', 'children'),
    [Input('confirm-yes', 'n_clicks'),
     Input('confirm-no', 'n_clicks')],
    [State("updated-group-table-container", "children")]
)
def confirm_selection(confirm_yes, confirm_no, updated_table):
    ctx = dash.callback_context
    if not ctx.triggered:
        return ""
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'confirm-yes':
        # Indicate that the update is completed
        return "Update completed and saved as 'group-data-table'."
    elif button_id == 'confirm-no':
        # Allow for further updates
        return "You can continue selecting items for updating."
    return ""

@app.callback(
    Output("updated-asset-table-container", "children"),
    [Input("update-asset-codes-button", "n_clicks")],
    [
        State("first-blank-row-input", "value"),
    ],
)
def update_asset_codes(n_clicks, first_blank_row):
    if n_clicks > 0 and latest_table_data is not None:
        df = latest_table_data  # Use the latest table data
        base_code = first_blank_row[:-3]  # Extract the base part of the code
        start_number = int(first_blank_row[-3:])  # Extract the starting number
        blank_rows = df[df["Asset Code"] == ""].index

        if not blank_rows.empty:
            for i, idx in enumerate(blank_rows):
                df.loc[idx, "Asset Code"] = f"{base_code}{str(start_number + i).zfill(3)}"

        grouped = df.groupby('Site')
        for name, group in grouped:
            if group['Asset Code'].str.startswith(base_code).any():
                asset_code_value = group.loc[group['Asset Code'].str.startswith(base_code), 'Asset Code'].iloc[0]
                df.loc[df['Site'] == name, 'Group Lead?'] = asset_code_value

        return dash_table.DataTable(
            data=df.to_dict("records"),
            columns=[{"name": i, "id": i} for i in df.columns],
            id="asset-data-table",
            page_current=0,
            page_size=10,
            page_action="native",
            style_cell={
                'fontFamily': 'Arial, sans-serif',
                'fontSize': '14px',
                'textAlign': 'left',
                'padding': '10px',
            },
            style_header={
                'backgroundColor': '#0074D9',
                'color': 'white',
                'fontWeight': 'bold',
                'fontSize': '16px'
            },
            style_table={
                "overflowY": "auto",  # Enable vertical scrolling
                "maxHeight": "400px",  # Set a maximum height for the table container
                "margin": "10px 0"
            },  # Add some margin
        )
    return dash_table.DataTable()

@app.callback(
    Output("download-data-link", "data"),
    Input("download-excel-button", "n_clicks"),
    [State("asset-data-table", "data")]
)

def download_data(download_excel, asset_table):
    ctx = dash.callback_context
    if not ctx.triggered:
        return None

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    if asset_table and button_id == 'download-excel-button':
        df = pd.DataFrame(asset_table)
        return dcc.send_data_frame(df.to_excel, "updated_data.xlsx", index=False)
    return None

if __name__ == "__main__":
    app.run_server(debug=False)