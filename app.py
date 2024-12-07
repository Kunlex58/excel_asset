from turtle import color
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import pandas as pd
import base64
import os
import psycopg2
from flask_session import Session
from flask import Flask, session
from config import DB_CONFIG, SECRET_KEY

# Create a Flask application
server = Flask(__name__)
server.secret_key = SECRET_KEY  # Set your secret key here

# Create a Dash application
app = dash.Dash(__name__, server=server, suppress_callback_exceptions=True)
app.title = "System Thinking - Excel"

# Function to create and return a new database connection
def get_db():
    conn = psycopg2.connect(**DB_CONFIG)
    return conn

# Function to close the database connection
def close_db(conn):
    conn.close()

# Directory to save uploaded files
UPLOAD_DIRECTORY = "uploads"
os.makedirs(UPLOAD_DIRECTORY, exist_ok=True)  # Create the uploads directory if it doesn't exist

# Function to authenticate user
def authenticate_user(username, password):
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username=%s AND password=%s", (username, password))
        user = cur.fetchone()
        close_db(conn)
        return user is not None
    except Exception as e:
        print(f"Error authenticating user: {e}")
        return False

# Application layout
app.layout = html.Div(id='app-content',
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
        html.Div(
            [
                html.H1(
                    "System Thinking - Excel",
                    style={"textAlign": "center", "color": "#0074D9", "fontFamily": "Arial, sans-serif"},
                ),
                html.Div(
                    [
                        html.Div("Username: "),
                        dcc.Input(id="username", type="text", placeholder="Username", style={"borderStyle": "solid", 
                                  "borderWidth": "1px", "borderRadius": "5px", "borderColor": "grey"}),
                        html.Div("Password: "),
                        dcc.Input(id="password", type="password", placeholder="Password", style={"borderStyle": "solid", 
                                  "borderWidth": "1px", "borderRadius": "5px", "borderColor": "grey"}),
                        html.Button("Login", id="login-button", n_clicks=0),
                        html.Button("Logout", id="logout-button", n_clicks=0, style={"display": "none"}),
                    ],
                    style={"display": "flex", "justifyContent": "flex-end"}
                ),
            ],
            style={"position": "relative"}
        ),
        html.Div(id="message"),
        dcc.Upload(
            id="upload-data",
            children=html.Div(["Drag and Drop or ", html.A("Select Files")]),
            style={
                "backgroundColor": "#273861",
                "color": "#f5f5f5",
                "font-weight": "bold",
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
        html.Div(id="data-table-container"),
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
        html.Button("Add to Update List", id="add-to-update-button", n_clicks=0),
        html.Button("Update Group", id="update-group-button", n_clicks=0),
        html.Div(id="update-list-container"),
        html.Div(id="updated-group-table-container"),
        dcc.Input(
            id="first-blank-row-input", type="text", placeholder="First Blank Row"
        ),
        html.Button("Update Asset Codes", id="update-asset-codes-button", n_clicks=0),
        html.Div(id="updated-asset-table-container"),
        html.Button("Download Data as Excel", id="download-excel-button", n_clicks=0),
        dcc.Download(id="download-data-link"),

        # Copyright Notice
        html.Div(
            "© System Thinking - Inc (2018)",
            style={
                'textAlign': 'right',
                'color': 'grey',
                'fontSize': '14px',
                'marginTop': '20px',
                'position': 'fluid',
                'bottom': '10px',
                'right': '20px'
                }
            )
    ]
)

# Global variable to store the latest DataFrame
latest_table_data = None  # Store the latest DataFrame
latest_uploaded_filename = None  # Store the filename of the uploaded file
update_list = []  # List to store updates as dictionaries

def save_uploaded_file(contents, filename):
    """ Save the uploaded file to the server """
    content_type, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    file_path = os.path.join(UPLOAD_DIRECTORY, filename)
    with open(file_path, 'wb') as f:
        f.write(decoded)
    return file_path

def parse_file(file_path, sheet_name, header_row):
    """ Parse the file into a DataFrame """
    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
        elif file_path.endswith(".xls") or file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        else:
            return None, f"Unsupported file format: {file_path}"
    except Exception as e:
        return None, f"There was an error processing this file. Error: {str(e)}"
    return df, None

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
    Output("output-data-upload", "children"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    State("username", "value"),
    State("password", "value")
)
def upload_file(contents, filename, username, password):
    """ Handle file upload and save it on the server """
    if session.get('logged_in'):
        if contents is not None:
            global latest_uploaded_filename
            latest_uploaded_filename = save_uploaded_file(contents, filename)  # Save the file to the server
            return f"Uploaded: {filename}"
    return ""

@app.callback(
    [
        Output("data-table-container", "children"),
        Output("column-dropdown", "options"),
        Output("group-dropdown", "options"),
    ],
    [Input("load-data-button", "n_clicks")],
    [State("sheet-name", "value"),
     State("header-row", "value")]
)
def load_data(n_clicks, sheet_name, header_row):
    """ Load data from the saved file on the server """
    if session.get('logged_in'):
        if n_clicks > 0 and latest_uploaded_filename:
            df, error = parse_file(latest_uploaded_filename, sheet_name, header_row)
            if error:
                return html.Div(error), [], []  # Show error message

            global latest_table_data  # Access global variable to store the DataFrame
            latest_table_data = df  # Store the DataFrame for later use
            column_options = [{"label": col, "value": col} for col in df.columns]
            group_options = [{"label": grp, "value": grp} for grp in df["Site"].unique()]
            
            return (
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
                column_options,
                group_options,
            )
        return html.Div(), [], []
    return html.Div(), [], []

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
    if session.get('logged_in'):
        if n_clicks > 0 and data is not None and selected_columns is not None:
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
                    # Convert "Quantity" column to numeric
                    group["Quantity"] = pd.to_numeric(group["Quantity"], errors="coerce")
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
    return html.Div()

@app.callback(
    Output("update-list-container", "children"),
    Input("add-to-update-button", "n_clicks"),
    [State("group-dropdown", "value"),
     State("classification-value", "value")]
)
def add_to_update_list(n_clicks, group, classification_value):
    """ Add selected group and classification value to the update list """
    global update_list
    if session.get('logged_in'):
        if n_clicks > 0 and group and classification_value:
            update_list.append({"Group": group, "Classification": classification_value})
            return f"Added: {group} - {classification_value} to the update list."
        return ""
    return html.Div()

@app.callback(
    [Output("login-button", "style"), Output("logout-button", "style"), Output("message", "children")],
    [Input("login-button", "n_clicks"), Input("logout-button", "n_clicks")],
    [State("username", "value"), State("password", "value")]
)
def handle_login_logout(login_n_clicks, logout_n_clicks, username, password):
    ctx = dash.callback_context
    if not ctx.triggered:
        return {}, {}, "Please log in to upload files!"

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if button_id == "login-button" and login_n_clicks > 0:
        if authenticate_user(username, password):
            session['logged_in'] = True
            session['username'] = username
            return {"display": "none"}, {"display": "inline"}, f"Welcome {username}! 😎"
        else:
            return {"display": "flex"}, {"display": "none"}, "Invalid username or password"
    elif button_id == "logout-button" and logout_n_clicks > 0:
        session.pop('logged_in', None)
        session.pop('username', None)
        # Clear tasks and delete uploaded files
        if os.path.exists(UPLOAD_DIRECTORY):
            for file in os.listdir(UPLOAD_DIRECTORY):
                os.remove(os.path.join(UPLOAD_DIRECTORY, file))
        # Refresh the page
        return {"display": "flex"}, {"display": "none"}, dcc.Location(href="/", id="refresh-page", refresh=True)
    return {}, {}, ""

@app.callback(
    Output("updated-group-table-container", "children"),
    Input("update-group-button", "n_clicks"),
    State("sorted-data-table", "data")
)
def update_group(n_clicks, data):
    """ Update all entries in the update list at once """
    global latest_table_data
    if session.get('logged_in'):
        if data is not None and n_clicks > 0 and update_list:
            df = pd.DataFrame(data)  # Make a copy of the latest data
            for update in update_list:
                group = update["Group"]
                classification = update["Classification"]
                df.loc[df["Site"] == group, "Group.1"] = classification
                df["Group.1"] = df["Group.1"].fillna("")  # Fill NaNs with empty strings

                # Save the latest table ID and data for use in subsequent updates
                latest_table_data = df  # Update latest_table_data with the modified DataFrame

            # Clear the update list after processing
            update_list.clear()

            return [
                dash_table.DataTable(
                    data=df.to_dict("records"),
                    columns=[{"name": i, "id": i} for i in df.columns],
                    id="updated-group-table",
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
                html.Div(f"Updated {len(update_list)} entries.")
            ]
        return html.Div()
    return html.Div() 

@app.callback(
    Output("updated-asset-table-container", "children"),
    [Input("update-asset-codes-button", "n_clicks")],
    [State("first-blank-row-input", "value")]
)

def update_asset_codes(n_clicks, first_blank_row):
    if session.get('logged_in'):
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
    return html.Div() 

@app.callback(
    Output("download-data-link", "data"),
    Input("download-excel-button", "n_clicks"),
    [State("asset-data-table", "data")]
)

def download_data(download_excel, asset_table):
    if session.get('logged_in'):
        ctx = dash.callback_context
        if not ctx.triggered:
            return None

        button_id = ctx.triggered[0]['prop_id'].split('.')[0]
        if asset_table and button_id == 'download-excel-button':
            df = pd.DataFrame(asset_table)
            return dcc.send_data_frame(df.to_excel, "updated_data.xlsx", index=False)
    return None

# Callback to delete uploaded files when the app is closed or refreshed
@app.callback(
    Output("upload-data", "contents"),
    Input("app-content", "style")
)

def delete_uploaded_files(style):
    if style['overflow'] == 'hidden':
        # Delete all files in the uploads directory
        for filename in os.listdir(UPLOAD_DIRECTORY):
            file_path = os.path.join(UPLOAD_DIRECTORY, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
    return None

if __name__ == '__main__':
    app.run_server(debug=True)