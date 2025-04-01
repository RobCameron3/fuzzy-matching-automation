# Demo Version of Fuzzy Account Matcher
# This script uses mock data and replicates the logic used in a real-world professional tool.
# Proprietary data has been removed to protect client privacy.

# %%
import pandas as pd
import numpy as np
import os
import datetime
import multiprocessing
from fuzzywuzzy import fuzz, process
from joblib import Parallel, delayed
import win32com.client as win32

# %%
accounts_df = pd.read_excel(accounts_df = pd.read_excel("data/sample_clients.xlsx"))

# %%
accounts_df['Address'] = accounts_df['Address 1: Street 1'].astype(str) + " " + accounts_df['Address 1: Street 2'].fillna("")
accounts_df['Name'] = accounts_df['Account Name'].copy()
accounts_df['ZIP'] = accounts_df['Address 1: ZIP/Postal Code'].copy()
accounts_df['State'] = accounts_df['Address 1: State/Province'].copy()
accounts_df['State'] = accounts_df['State'].str.strip()

# %%
accounts_df

# %%

accounts_df['ZIP'] = accounts_df['ZIP'].str[:5].str.zfill(5)

# %%
accounts_df['Name'] = accounts_df['Name'].str.replace(' ', '')

# %%
retail_df = pd.read_excel(retail_df = pd.read_excel("data/unmapped_clients_sample.xlsx"))

# %%
retail_df

# %%
retail_df['Zip'] = retail_df['Zip'].astype(str)
retail_df['Zip'] = retail_df['Zip'].str[:5].str.zfill(5)
retail_df.rename(columns={'Zip': 'ZIP'}, inplace=True)

# %%

retail_df = retail_df.rename(columns={'Standard_Address': 'Address'})
retail_df

# %%
def find_best_match(row, restriction_column, match_column, choices_df, limit=5):
    value = row[restriction_column]
    choices_same_value = choices_df[choices_df[restriction_column] == value][match_column]
    if choices_same_value.empty:
        # Return a default value or handle it accordingly
        return pd.Series([row[match_column], value, None, None], index=[f'Merge {match_column}', restriction_column, 'Best Match', 'Score'])
    else:
        matches = process.extract(row[match_column], choices_same_value, limit=limit)
        return pd.Series([row[match_column], value, matches], index=[f'Merge {match_column}', restriction_column, 'Best Matches'])

# %%
# Define a function to be applied in parallel
def parallel_find_best_match(chunk, restriction_column, match_column, choices_df, limit):
    return chunk.apply(find_best_match, restriction_column=restriction_column, match_column=match_column, choices_df=choices_df, limit=limit, axis=1)

# %%
def run_parallel_find_best_match(retail_df, restriction_column='State', match_column='Address', choices_df=accounts_df, limit=2):                               
    # Define the number of CPU cores to use
    num_cores = multiprocessing.cpu_count()
    
    # Split the exercise DataFrame into chunks for parallel processing
    
    # If there aren't many observations, you only need one chunk (will break if more than cores)
    if len(retail_df) < num_cores:
        chunk_size = 1
    else:
        chunk_size = len(retail_df) // num_cores
    chunks = [retail_df.iloc[i:i+chunk_size] for i in range(0, len(retail_df), chunk_size)]
    
    # Run the parallel processing using joblib
    results = Parallel(n_jobs=num_cores)(delayed(parallel_find_best_match)(chunk, restriction_column, match_column, choices_df, limit) for chunk in chunks)
    
    # Concatenate the results from all processes
    best_matches_df = pd.concat(results, ignore_index=True)
    
    # Reset the index of the DataFrame
    best_matches_df.reset_index(drop=True, inplace=True)


    # Function to split tuples into separate columns
    def split_tuples_to_columns(row):
        tuples = row['Best Matches']
        num_tuples = limit
        for i, t in enumerate(tuples):
            row[f'Match {i+1}'] = t[0]
            row[f'Score {i+1}'] = t[1]
            row[f'Drop {i+1}'] = t[2]
            
        return row

    # Allow for errors (this is likely because the restriction column categories don't match the CRM, like the retail file may have full state vs the CRM's abbreviation)
    def split_tuples_to_columns2(row):
        try:
            row = split_tuples_to_columns(row)
            return row
        except:
            pass

    
    # Apply the function to split tuples into separate columns
    best_matches_df = best_matches_df.apply(split_tuples_to_columns2, axis=1)

    # Drop the original 'Best Matches' column
    best_matches_df.drop(columns=['Best Matches'], inplace=True)

    # Filter out columns containing "Drop" in their name and drop them
    best_matches_df = best_matches_df.filter(regex='^(?!.*Drop).*')

    
    return best_matches_df

# %%
%%time
# Find best matches
retail_matches_df = run_parallel_find_best_match(retail_df, restriction_column='ZIP', match_column='Address', choices_df=accounts_df, limit=2)

# %%
# Merge on account info from CRM
accounts_df2 = accounts_df.drop_duplicates(['Address'], keep='first')
retail_matches_df_with_accounts = retail_matches_df.merge(accounts_df2[['Address', 'Account Number', 'Account Name']], left_on=['Match 1'], right_on=['Address'], how='left')
retail_matches_df_with_accounts = retail_matches_df_with_accounts.merge(accounts_df2[['Address', 'Account Number', 'Account Name']], left_on=['Match 2'], right_on=['Address'], how='left', suffixes=(" 1", " 2"))
retail_matches_df_with_accounts = retail_matches_df_with_accounts[['Merge Address', 'Match 1', 'Score 1', 'Account Number 1','Account Name 1', 'Match 2', 'Score 2', 'Account Number 2', 'Account Name 2']]

retail_matches_df_with_accounts['Practice Name'] = retail_df['Practice Name']

retail_matches_df_with_accounts = retail_matches_df_with_accounts[['Practice Name', 'Merge Address', 'Match 1', 'Score 1', 'Account Number 1',
       'Account Name 1', 'Match 2', 'Score 2', 'Account Number 2',
       'Account Name 2']]

# %%
file_path = "data/client_report_by_zip.xlsx"

# Save the DataFrame to an Excel file
retail_matches_df_with_accounts.to_excel(file_path, index=False)

# %%



