import pandas as pd

#lista con i nomi degli excel da unire
files = ["String 1 - No duplicati.xlsx", "String 2 - No duplicati.xlsx", "String 3 - No duplicati.xlsx", "String 4 - No duplicati.xlsx", "String 5 - No duplicati.xlsx",\
            "String 6 - No duplicati.xlsx", "String 7 - No duplicati.xlsx", "String 8 - No duplicati.xlsx"]

#inizializzare una lista vuota per memorizzare i dati di ciascun dataframe
dfs = []

#Legge ogni excel nella lista files e lo aggiunge alla lista "dfs"
for file in files:
    df = pd.read_excel(file)
    dfs.append(df)

#merge excels in the list dfs in one only excel.
merged_excels = pd.concat(dfs, ignore_index=True)

#salvare il risultato in un nuovo file excel.
merged_excels.to_excel("merged_excels_unfiltered.xlsx", index=False)

