import dash
from dash import html

app = dash.Dash(__name__)
server = app.server

app.layout = html.Div("Hello from Azure!")

if __name__ == "__main__":
    app.run()
