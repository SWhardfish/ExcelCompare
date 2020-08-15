import pandas as pd


pd.set_option('display.max_rows', 10000)
pd.set_option('display.max_columns', 100)
pd.set_option('display.width', 1000)

Prod=pd.read_excel('Book1.xlsx')
PreProd=pd.read_excel('Book2.xlsx')
Prod['System'] = "Prod"
PreProd['System'] = "PreProd"

Prod_Unique_all = set(Prod['Unique'])
PreProd_Unique_all = set(PreProd['Unique'])

dropped_Unique = Prod_Unique_all - PreProd_Unique_all
added_Unique = PreProd_Unique_all - Prod_Unique_all

all_data = pd.concat([Prod, PreProd], ignore_index=True)
changes = all_data.drop_duplicates(subset=["Unique", "Version"], keep='last')

dupe_Unique = changes[changes['Unique'].duplicated() == True]['Unique'].tolist()
dupes = changes[changes["Unique"].isin(dupe_Unique)]
print(dropped_Unique)
print(added_Unique)

# Pull out the old and new data into separate dataframes
change_PreProd = dupes[(dupes["System"] == "PreProd")]
change_Prod = dupes[(dupes["System"] == "Prod")]

# Drop the temp columns - we don't need them now
change_PreProd = change_PreProd.drop(['System'], axis=1)
change_Prod = change_Prod.drop(['System'], axis=1)

# Index on the Unique numbers
change_PreProd.set_index('Unique', inplace=True)
change_Prod.set_index('Unique', inplace=True)

# Combine all the changes together
df_all_changes = pd.concat([change_Prod, change_PreProd],
                            axis='columns',
                            keys=['Prod', 'PreProd'],
                            join='outer')

df_all_changes.columns = ['ProdV', 'PreProdV']

df_all_changes = df_all_changes[df_all_changes.ProdV > df_all_changes.PreProdV]

print(df_all_changes)

df_all_changes.to_excel('./Diffv1.1.xlsx', index=True, header=True)