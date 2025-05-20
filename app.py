import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import dash
from dash import html, dcc, dash_table
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objs as go


base_dir = os.getcwd() 

excel_file_path = os.path.join(os.getcwd(), "raw_data", "su-d-01.08.01.03.xlsx")

xlsx = pd.ExcelFile(excel_file_path)

dfs = {}

for sheet_name in xlsx.sheet_names:
    dfs[sheet_name] = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    # data clean-up
    dfs[sheet_name] = dfs[sheet_name].drop([0,1,2,3])    # drop empty rows at top
    dfs[sheet_name] = dfs[sheet_name].drop(index=range(16,25))  # drop empty rows at bottom
    dfs[sheet_name] = dfs[sheet_name].drop(dfs[sheet_name].columns[0], axis=1)    # delete empty or irrelevant columns
    dfs[sheet_name] = dfs[sheet_name].drop(dfs[sheet_name].columns[2], axis=1)
    dfs[sheet_name] = dfs[sheet_name].drop(dfs[sheet_name].columns[3], axis=1)
    dfs[sheet_name].columns = ["Language", "Number of speakers", "Percentage of workforce"]   # rename columns
    dfs[sheet_name] = dfs[sheet_name].reset_index(drop=True)


# Combine all dataframes into one and reshape to facilitate plotting
combined = []

for year, df in dfs.items():
    temp = df.copy()
    temp['Year'] = year
    combined.append(temp)

complete_df = pd.concat(combined, ignore_index=True)

pivoted_df = complete_df.pivot(index="Year", columns="Language", values="Number of speakers")

# Find the column name that starts with "Auskunft"
col_to_drop = [col for col in pivoted_df.columns if col.startswith("Auskunft")]

# Drop it if found
pivoted_df = pivoted_df.drop(columns=col_to_drop)

# Rename columns for convenience
pivoted_df.columns = ['ALBANIAN','OTHER','ENGLISH','FRENCH','GERMAN','ITALIAN','PORTUGUESE','RAETO','SWISS GERMAN','SERBIAN','SPANISH','TICINO']

clean_data = pivoted_df

# Pivot and rename for percentage
pivoted_df_percentages = complete_df.pivot(index="Year", columns="Language", values="Percentage of workforce")
pivoted_df_percentages = pivoted_df_percentages.drop(columns=col_to_drop)
pivoted_df_percentages.columns = ['ALBANIAN','OTHER','ENGLISH','FRENCH','GERMAN','ITALIAN','PORTUGUESE','RAETO','SWISS GERMAN','SERBIAN','SPANISH','TICINO']
clean_data_percentages = pivoted_df_percentages
clean_data_percentages.index = clean_data_percentages.index.astype(str)
latest_English_percentage = int(clean_data_percentages['ENGLISH'].iloc[-1])
english_2010_percentage = clean_data_percentages.loc['2010', 'ENGLISH']
percent_increase_english_percentage = int(((latest_English_percentage - english_2010_percentage) / english_2010_percentage) * 100)



# Ensure the index is of string type for clean labeling
clean_data.index = clean_data.index.astype(str)

english_2020 = clean_data.loc['2020', 'ENGLISH']
english_latest = clean_data['ENGLISH'].iloc[-1] 
english_latest_pretty = round(english_latest / 1_000_000, 2)
percent_increase_english = ((english_latest - english_2020) / english_2020) * 100

selected_columns = ['ENGLISH','FRENCH','GERMAN','ITALIAN','PORTUGUESE','SWISS GERMAN']
# Sort selected_columns by latest (2023) values in descending order

sorted_columns = clean_data.loc['2023', selected_columns].sort_values(ascending=False).index.tolist()


plt.figure(figsize=(14, 7))

# Plot each language as a line
for column in selected_columns:
    if column == 'ENGLISH':
        plt.plot(clean_data.index, clean_data[column], label="ENGLISH", color='#00FFFF', linewidth=2.5)  # Bright cyan
    else:
        plt.plot(clean_data.index, clean_data[column], label=column)

# Rotate x-axis labels
plt.xticks(rotation=45)

# Set y-axis tick interval to 0.5 million
plt.gca().yaxis.set_major_locator(ticker.MultipleLocator(500_000))

# Format ticks to show in millions (e.g., 0.5M, 1.0M)
plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x / 1_000_000:.1f}M'))

# Add labels and title
plt.xlabel("Year")
plt.ylabel("Number of Speakers (in millions)")
plt.legend(sorted_columns, title="Language", bbox_to_anchor=(1.05, 1), loc='upper left')

plt.tight_layout()

def make_highlighted_figure(df, selected_columns, yaxis_title, title, percent=False):
    fig = go.Figure()
    for col in selected_columns:
        if col == "ENGLISH":
            fig.add_trace(go.Scatter(
                x=df.index,
                y=df[col],
                mode='lines',
                name='ENGLISH',
                line=dict(color="#FF0095", width=6, dash='solid'),  # Bright cyan, thick
                opacity=1
            ))
        else:
            fig.add_trace(go.Scatter(
                x=df.index,
                y=df[col],
                mode='lines',
                name=col,
                line=dict(width=2)
            ))
    fig.update_layout(
    title={
        'text': title,
        'font': {
            'size': 22,
        },
    },
    xaxis_title="Year",
    yaxis_title=yaxis_title,
    legend_title="Language"
    )

    if percent:
        fig.update_yaxes(ticksuffix="%")
    return fig

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

app.title = "Languages in the Swiss Workplace"
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("Languages in the Swiss Workplace"), className="text-center m-4")
    ]),
    dbc.Row([
        dbc.Col(
            html.P([
                "Source: Swiss Federal Statistical Office – ",
                html.A(
                    "Die üblicherweise bei der Arbeit gesprochenen Sprachen",
                    href="https://opendata.swiss/en/dataset/die-ublicherweise-bei-der-arbeit-gesprochenen-sprachen1",
                    target="_blank"  
                )
            ]),
            className="text-center"
        )
    ]),
    dbc.Row([
        dbc.Col(
            dcc.Graph(
                figure=make_highlighted_figure(
                    clean_data, selected_columns,
                    "Number of speakers in millions",
                    "Regularly spoken languages in the Swiss workplace"
                )
            ),
            width={"size": 9}, className="m-4"
        ),
        dbc.Col(dbc.Card([
            dbc.CardHeader(html.H4("English")),
            dbc.CardBody([
                html.H2(f"{english_latest_pretty}M"), 
                html.P("English speakers in the Swiss workplace"),
                html.H2(f"{percent_increase_english:.2f}%"),
                html.P(f"increase since 2020"),
            ])
        ]), width={"size": 2}, className="m-2 align-self-center")
    ]),
    dbc.Row([
        dbc.Col(
            dcc.Graph(
                figure=make_highlighted_figure(
                    clean_data_percentages, selected_columns,
                    "Percentage of workforce (%)",
                    "Percentage of Swiss workforce that regularly speaks the language",
                    percent=True
                )
            ),
            width={"size": 9}, className="m-4"
        ),
        dbc.Col(dbc.Card([
            dbc.CardHeader(html.H4("English")),
            dbc.CardBody([
                html.H2(f"{latest_English_percentage}%"),
                html.P("of the Swiss workforce speaks English regularly"),
                html.H2(f"{percent_increase_english_percentage}%"),
                html.P("increase since 2010"),
            ])
        ]), width={"size": 2}, className="m-2 align-self-center")
    ])
], fluid=True)





if __name__ == '__main__':
    app.run(debug=True)