import pandas as pd

# Load the Rental Report DataFrame
df  = pd.read_excel(r"C:\Users\sadia.upama_ext.EDOTCO\Desktop\GL Project\Bangladesh Rental Report all Versions.xlsx")
df1 = pd.read_excel(r"C:\Users\sadia.upama_ext.EDOTCO\Desktop\GL Project\Project Ajax- Finance-Site Information_BD.xlsx")
df2 = pd.read_excel(r"C:\Users\sadia.upama_ext.EDOTCO\Desktop\GL Project\Bangladesh_Rental_Report.xlsx")
df3 = pd.read_excel(r"C:\Users\sadia.upama_ext.EDOTCO\Desktop\GL Project\Bangladesh Payee Details Report.xlsx")
df8 = pd.read_excel(r"C:\Users\sadia.upama_ext.EDOTCO\Desktop\GL Project\Total_Entry.xlsx")
df9 = pd.read_excel(r"C:\Users\sadia.upama_ext.EDOTCO\Desktop\GL Project\Hold_Unhold_Report.xlsx")
# Filter the DataFrame for 'Draft' Ground lease Status
df = df[df['Ground lease Status'] == 'Draft']   # STEP 1
# Merge the specified columns from the Rental Report and Site Information DataFrames
df['Merged'] = df['Site Reference'].astype(str) + " - " + df['Customer Name']

d10 =df [['Site Reference', 'Customer Code', 'Customer Name']].copy()

df1['Merged'] = df1['SiteRef'].astype(str) + " - " + df1['Anchor Tenant']  # STEP 2
df['Merged'] = df['Merged'].astype(str).fillna('')
df1['Merged'] = df1['Merged'].astype(str).fillna('')

df = df[~df['Merged'].str.contains('|'.join(df1['Merged']), case=False)]  # STEP 3

# Filter df2 based on Site Reference that matches with df
df = df[df['Site Reference'].isin(df2['Site Reference'])]

# Filter the DataFrame for 'Landlord Lease Executed' GL Status
df2 = df2[df2['GL Status'] == 'Landlord Lease Executed']  # STEP 4

# Extract the numeric portion from "ProjectRef" by removing 'GL-'
df2['ProjectRef_Num'] = df2['ProjectRef'].str.replace('GL-', '').astype(int)

# Sort the DataFrame by "ProjectRef_Num" and "VersionNumber" in descending order
df2.sort_values(['ProjectRef_Num', 'VersionNumber'], ascending=[True, False], inplace=True)  # STEP 5

# Keep only the rows with the latest version for each "ProjectRef_Num"
latest_versions = df2.groupby('ProjectRef_Num')['VersionNumber'].idxmax()
df2_latest = df2.loc[latest_versions]  # STEP 6

# Discard the rows with "Discontinued" or "Terminated" entities in the "Landlord Status" column
df2 = df2[~df2['Landlord Status'].str.contains('Discontinued|Terminated', na=False)]  # STEP 7

# Keep only rows where "Type of Agreement" is "Tenancy Agreement"
df2 = df2[df2['Type Of Agreement'] == 'Tenancy Agreement']  # STEP 8

# Keep the latest Expiry date for the same Site Ref but different GL
# Convert the 'Expiry Date' column to datetime
df2['Expiry Date'] = pd.to_datetime(df2['Expiry Date'])

# Sort the DataFrame by 'Site Reference' and 'Expiry Date' in descending order
df2.sort_values(['Site Reference', 'Expiry Date'], ascending=[True, False], inplace=True)

# Keep only the rows with the latest date for each unique 'Site Reference'
df2_latest_date = df2.groupby('Site Reference').first()

# Reset the index to retain the 'ProjectRef' column
df2_latest_date.reset_index(inplace=True)

# Filter the original DataFrame based on the 'ProjectRef' values in df2_latest_date
df2 = df2[df2['ProjectRef'].isin(df2_latest_date['ProjectRef'])]  # STEP 9

# Drop the column "Title No" from df2
df2.drop('Title No', axis=1, inplace=True)

# Create DataFrame d4 with NTC Site Code and Customer Site Ref
d4 = df1[['NTC Site Code', 'Customer Site Ref', 'SiteRef']].copy()

# Update the NTC Site Code where it is blank with Customer Site Ref
d4['NTC Site Code'].fillna(d4['Customer Site Ref'], inplace=True)

# Filter d4 based on the Site Reference values in df
d4 = d4[d4['NTC Site Code'].notnull() | d4['Customer Site Ref'].notnull()]
d4 = d4[d4['NTC Site Code'].str.strip() != '']

df2.drop(['Customer Name', 'Customer Code'],axis =1, inplace=True)

df2 = df2.merge(d4, left_on='Site Reference', right_on='SiteRef', how='left')
df2['Title No'] = df2['NTC Site Code'].astype(str) 

df2 = df2.merge(d10, left_on='Site Reference', right_on='Site Reference', how='left')
# Define the replacements
replacements = {
    'Grameenphone Ltd.': 'GP',
    'Banglalink Digital Communication Limited': 'BL',
    'Robi Axiata Limited': 'Robi',
    'Fibre@Home': 'F@H',
    'Teletalk Bangladesh Limited':'Tele'
}
# Replace the values in the "Customer Name" column
df2['Customer Name'] = df2['Customer Name'].replace(replacements)
df2['Title No'] = df2['NTC Site Code'].astype(str)+'_' + df2['Customer Name'].astype(str) + '_COLO_' + df2['Customer Code'].astype(str)

# Selecting the desired columns from df
df_selected = df[['Site Reference', 'ProjectRef']]

# Selecting the desired columns from df2
df2_selected = df2[['Site Reference', 'Title No','Lot/Pt No', 'Land use Type', 'Type of Land', 'Postcode', 'City Corporation']]
df2_selected.loc[df2_selected["City Corporation"].isnull(), "City Corporation"] = "Dhaka"
# Merging the selected columns based on Site Reference
df6 = pd.merge(df_selected, df2_selected, on='Site Reference', how='inner')

# Extract the ProjectRef values from df9
existing= df8["ProjectRef"].tolist()

# Filter df6 to discard entries with ProjectRef already present in df8
df6 = df6[~df6["ProjectRef"].isin(existing)]

# Drop duplicate rows based on 'ProjectRef'
df6.drop_duplicates(subset='ProjectRef', inplace=True)

#Rename the Project Ref to ProjectRef
df9.rename(columns={'Project Ref': 'ProjectRef'}, inplace=True)

#keep the ProjectRef that's Unhold Date is blank
df9 = df9[df9['Unhold Date'].isnull()]

#df9.to_excel("see.xlsx",index=None)

# Extract the ProjectRef values from df9
existing1= df9["ProjectRef"].tolist()

# Filter df6 to discard entries with ProjectRef already present in df9
df6 = df6[~df6["ProjectRef"].isin(existing1)]

df6.to_excel("GL DRAFT.xlsx",index=None)


df6_selected=df6[['Site Reference', 'ProjectRef']]

# Selecting the desired columns from df2
df2_selected1 = df2[['ProjectRef','Site Reference', 'Landlord Name','Father Name','Mother Name', 'Spouse Name', 'Date of Birth', 'Landlord Address', 'Landlord Status', 'Phone Number 1','Phone Number 2','Phone Number 3']]

# Fill blank cells with "N/A" in the specified column using .loc method
df2_selected1.loc[df2_selected1["Phone Number 1"].isnull(), "Phone Number 1"] = "N/A"
df2_selected1.loc[df2_selected1["Father Name"].isnull(), "Father Name"] = "N/A"
df2_selected1.loc[df2_selected1["Mother Name"].isnull(), "Mother Name"] = "N/A"
df2_selected1.loc[df2_selected1["Spouse Name"].isnull(), "Spouse Name"] = "N/A"


# Selecting the desired columns from df3
df3_selected = df3[['Parent Project Ref','Site Ref','Payee Name','Payee Address','IC / CO Registration No','Rent Payment Mode','Bank Account No','Bank Account Type','Bank Name','Bank Branch Name','Bank Routing Number','Distribution %']]

# Fill blank cells with "N/A" in the specified columns using .loc method
df3_selected.loc[df3_selected["IC / CO Registration No"].isnull(), "IC / CO Registration No"] = "N/A"
df3_selected.loc[df3_selected["Bank Routing Number"].isnull(), "Bank Routing Number"] = "N/A"
df3_selected.loc[df3_selected["Bank Account No"].isnull(), "Bank Account No"] = "N/A"
df3_selected.loc[df3_selected["Bank Account Type"].isnull(), "Bank Account Type"] = "N/A"
df3_selected.loc[df3_selected["Bank Name"].isnull(), "Bank Name"] = "Not Applicable (for Check/Cash Payment)"


# Merging the selected columns based on Site Reference
df7= df2_selected1.merge(df3_selected, left_on='ProjectRef', right_on='Parent Project Ref')
# Drop 'ProjectRef' and 'Parent Project Ref' columns from df7
df7.drop(['ProjectRef', 'Parent Project Ref','Site Ref'], axis=1, inplace=True)

# Merging the selected columns based on Site Reference
merged_df = pd.merge(df6_selected, df7, on='Site Reference', how='inner')

merged_df.drop(['Site Reference'], axis=1, inplace=True)

merged_df1=merged_df.copy()
merged_df2=merged_df.copy()
merged_df1.drop(['Payee Name','Payee Address','IC / CO Registration No','Rent Payment Mode','Bank Account No','Bank Account Type','Bank Name','Bank Branch Name','Bank Routing Number','Distribution %'],axis=1,inplace=True)
merged_df2.drop(['Landlord Name','Father Name',	'Mother Name','Spouse Name','Date of Birth','Landlord Address',	'Landlord Status','Phone Number 1','Phone Number 2','Phone Number 3'],axis=1,inplace=True)

import os
# Create the subfolder if it doesn't exist
subfolder_name1 = 'Subfolder1'
os.makedirs(subfolder_name1, exist_ok=True)

# Group the merged dataframe by 'ProjectRef'
groups = merged_df1.groupby('ProjectRef')

# Iterate over the groups and save each group to a separate Excel file
for name, group in groups:
    output_filename = os.path.join(subfolder_name1, f"{name}.xlsx")
    group.to_excel(output_filename, index=False)

# Create the subfolder if it doesn't exist
subfolder_name2 = 'Subfolder2'
os.makedirs(subfolder_name2, exist_ok=True)

# Group the merged dataframe by 'ProjectRef'
groups = merged_df2.groupby('ProjectRef')

# Iterate over the groups and save each group to a separate Excel file
for name, group in groups:
    output_filename = os.path.join(subfolder_name2, f"{name}.xlsx")
    group.to_excel(output_filename, index=False)

# Confirm the operation is complete
print("DONE")




