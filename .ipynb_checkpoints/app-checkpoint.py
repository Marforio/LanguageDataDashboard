import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import dash
from dash import html, dcc, dash_table
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px


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
pivoted_df.columns = ['ALBANIAN','AOTHER','ENGLISH','FRENCH','HIGH GERMAN','ITALIAN','PORTUGUESE','RAETO','SWISS GERMAN','SERBIAN','SPANISH','TICINO']

clean_data = pivoted_df

# Ensure the index is of string type for clean labeling
clean_data.index = clean_data.index.astype(str)

selected_columns = ['ENGLISH','FRENCH','HIGH GERMAN','ITALIAN','PORTUGUESE','SWISS GERMAN']

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
plt.title("Regularly Spoken Languages in the Swiss Workplace")
plt.legend(title="Language", bbox_to_anchor=(1.05, 1), loc='upper left')

plt.tight_layout()

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Swiss Languages in the Workplace"
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("Swiss Languages in the Workplace"), className="text-center mb-4")
    ]),
    dbc.Row([
        dbc.Col(dcc.Graph(figure=px.line(clean_data, x=clean_data.index, y=selected_columns, title="Regularly Spoken Languages in the Swiss Workplace")), className="mb-4")
    ]),
    
], fluid=True) 





if __name__ == '__main__':
    app.run(debug=True)