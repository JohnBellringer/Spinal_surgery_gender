import pandas as pd

# Read the main data file
main_data = pd.read_excel('merged_excels_unfiltered_no_duplicati.xlsx')

# Read the list of journal names
journal_names = pd.read_excel('Final journal list.xlsx')

main_data['Journal'] = main_data['Journal'].str.lower()
journal_names['Journal name'] = journal_names['Journal name'].str.lower()

# Extract list of journal names from UfficialJournalNames
journal_list = journal_names['Journal name'].tolist()

# Filter main data based on journal list from OncologyResultsClean
filtered_data = main_data[main_data['Journal'].isin(journal_list)]

# Write filtered data to a new Excel file
filtered_data.to_excel('Excel_Spinal_Filtered.xlsx', index=False)