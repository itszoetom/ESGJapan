import pandas as pd
import numpy as np

# Google Sheets document URL and sheet name
sheet_id = '1uYTDO_zU8P-tvVnDOnI_cHCseCNhbqk6xDrmj_sERD0'
financial_sheet = 'Financial'
social_sheet = 'Social'
environmental_sheet = 'Environmental'
governance_sheet = 'Governance'

# Read the Google Sheets document into a DataFrame
url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
financial_data = pd.read_excel(url, sheet_name=financial_sheet, header=0)

url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
social_data = pd.read_excel(url, sheet_name=social_sheet, header=0)

url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
environmental_data = pd.read_excel(url, sheet_name=environmental_sheet, header=0)

url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx'
governance_data = pd.read_excel(url, sheet_name=governance_sheet, header=0)

data = governance_data
data

%##

# Cleaning
# Convert 'Value' column to numeric
data['Value'] = pd.to_numeric(data['Value'], errors='coerce')

# Define a function to convert billion yen to million yen
def billion_to_million(value, unit):
    if unit == 'billion yen':
        return value * 1000  # Convert billion yen to million yen
    else:
        return value

# Apply the function to the DataFrame
data['Value'] = data.apply(lambda row: billion_to_million(row['Value'], row['Unit']), axis=1)

# Update 'Unit' column to reflect the conversion
data.loc[data['Unit'] == 'billion yen', 'Unit'] = 'million yen'

data

%##

df = data.drop(['Company', 'Unit', 'Year'], axis=1)

df = df.pivot(columns = 'KPI')

# Iterate over each column
for column in df.columns:
    # Iterate over each row in the column
    for i in range(len(df)):
        # If the value is NaN
        if pd.isna(df.loc[i, column]):
            # Find the first non-NaN value in the column below the NaN value
            first_valid_index = df[column].iloc[i:].first_valid_index()
            # If a non-NaN value exists below the NaN value
            if first_valid_index is not None:
                # Move down the value
                df.loc[i, column] = df.loc[first_valid_index, column]
                # Set the first valid index to NaN
                df.loc[first_valid_index, column] = pd.NA  # Use np.nan instead of pd.nan

df.head(15)

%##

#Sampling data to normalize number of rows for each KPI

def difference(column):
    return np.abs(column.min() - column.max())

difference_table = df.apply(difference)

df.dropna(axis=1, thresh=10, inplace=True)
df.head(15)

%##

# Compute correlation matrix
correlation_df = df.corr()

correlation_df

print(f'Significant Correlations \n')
for j in range(0, len(correlation_df)):
    for i in range(0, len(correlation_df)):
        if correlation_df.iloc[j, i] > 0.6 or correlation_df.iloc[j, i] < -0.6:
            print(f'{correlation_df.index.levels[1][j]} x {correlation_df.columns.levels[1][i]} : {correlation_df.iloc[j,i]}')
