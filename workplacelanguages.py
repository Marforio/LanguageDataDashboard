import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker


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
pivoted_df_percentages = complete_df.pivot(index="Year", columns="Language", values="Percentage of workforce")

# Find the column name that starts with "Auskunft"
col_to_drop = [col for col in pivoted_df.columns if col.startswith("Auskunft")]

# Drop it if found
pivoted_df = pivoted_df.drop(columns=col_to_drop)
pivoted_df_percentages = pivoted_df_percentages.drop(columns=col_to_drop)

# Rename columns for convenience
pivoted_df.columns = ['ALBANIAN','OTHER','ENGLISH','FRENCH','GERMAN','ITALIAN','PORTUGUESE','RAETO','SWISS GERMAN','SERBIAN','SPANISH','TICINO']
pivoted_df_percentages.columns = ['ALBANIAN','OTHER','ENGLISH','FRENCH','GERMAN','ITALIAN','PORTUGUESE','RAETO','SWISS GERMAN','SERBIAN','SPANISH','TICINO']

clean_data = pivoted_df
clean_data_percentages = pivoted_df_percentages

# Ensure the index is of string type for clean labeling
clean_data.index = clean_data.index.astype(str)

selected_columns = ['SWISS GERMAN','GERMAN','FRENCH','ENGLISH','ITALIAN','PORTUGUESE']

fig, axes = plt.subplots(1, 2, figsize=(20, 7), sharex=True)

# --- First plot: Number of speakers ---
for column in selected_columns:
    if column == 'ENGLISH':
        axes[0].plot(clean_data.index, clean_data[column], label="ENGLISH", color='#00FFFF', linewidth=3.5)
    else:
        axes[0].plot(clean_data.index, clean_data[column], label=column)
axes[0].set_xticklabels(clean_data.index, rotation=45)
axes[0].yaxis.set_major_locator(ticker.MultipleLocator(500_000))
axes[0].yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x / 1_000_000:.1f}M'))
axes[0].set_xlabel("Year")
axes[0].set_ylabel("Number of Speakers (in millions)")
axes[0].set_title("Regularly Spoken Languages in the Swiss Workplace")
axes[0].legend(title="Language", bbox_to_anchor=(1.05, 1), loc='upper left')

# --- Second plot: Percentage of workforce ---
for column in selected_columns:
    if column == 'ENGLISH':
        axes[1].plot(clean_data_percentages.index, clean_data_percentages[column], label="ENGLISH", color='#00FFFF', linewidth=3.5)
    else:
        axes[1].plot(clean_data_percentages.index, clean_data_percentages[column], label=column)
axes[1].set_xticklabels(clean_data_percentages.index, rotation=45)
axes[1].yaxis.set_major_locator(ticker.MultipleLocator(5))
axes[1].yaxis.set_major_formatter(ticker.PercentFormatter())
axes[1].set_xlabel("Year")
axes[1].set_ylabel("Percentage of Workforce (%)")
axes[1].set_title("Languages as Percentage of Swiss Workforce")
axes[1].legend(title="Language", bbox_to_anchor=(1.05, 1), loc='upper left')

plt.tight_layout()
plt.show()


