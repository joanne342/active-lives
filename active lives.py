import pyreadstat
import pandas as pd
import os
from datetime import datetime
import math

#==========================
#from first program section
#==========================

# Input filename
input_filename = "Nov 20-21 Full-Year Master.sav"  # You can change this to your actual input file name

# Determine Excel output filename based on input filename
if "Master" in input_filename:
    output_excel_file = input_filename.replace("Master", "Reference").replace(".sav", "") + '.xlsx'
else:
    output_excel_file = os.path.splitext(input_filename)[0] + ' Reference.xlsx'

# Load the .sav file along with its metadata
df, meta = pyreadstat.read_sav(input_filename)

# Initialize lists to store variable details and flattened value labels
variable_data = []
flattened_value_data = []

# Loop over metadata variables to extract information and value labels
for index, var_name in enumerate(meta.column_names):
    var_label = meta.column_labels[index]  # Descriptive label
    value_labels = meta.variable_value_labels.get(var_name, {})  # Value labels if available
    num_value_labels = len(value_labels)  # Count the number of value labels

    # Append data for variable details (for Columns sheet)
    variable_data.append([index + 1, var_name, var_label, num_value_labels])
    
    # Flatten value labels (for Labels sheet)
    for value, label in value_labels.items():
        flattened_value_data.append([index + 1, var_name, label, value])

# Create DataFrame for variable details (Columns sheet)
df_variables = pd.DataFrame(variable_data, columns=["ID", "Column", "Label", "Number of labels"])

# Create DataFrame for flattened value labels (Labels sheet)
df_flattened = pd.DataFrame(flattened_value_data, columns=["Column ID", "Column", "Label", "Value"])

# Create a DataFrame for the Info sheet
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
info_data = {
    'A': [
        "",
        "",
        "This reference workbook has been generated for KKP from a Python script created by Joanne O'Malley (KKP)",
        "",
        "All data is from Sport England - notes provided from the downloaded SPSS file",
        "",
        f"Generated on {timestamp}"
    ]
}
df_info = pd.DataFrame(info_data)

# Write DataFrames to an Excel file with four sheets
with pd.ExcelWriter(output_excel_file) as writer:
    df_info.to_excel(writer, sheet_name='Info', index=False, header=False)

    # Create and write the Contents sheet
    contents_data = {
        'A': ['Contents', '', 'Info', 'SPSS File', '', ''],
        'B': ['', '', '', '', 'Columns', 'Labels']
    }
    df_contents = pd.DataFrame(contents_data)
    df_contents.to_excel(writer, sheet_name='Contents', index=False, header=False)

    df_variables.to_excel(writer, sheet_name='Columns', index=False)
    df_flattened.to_excel(writer, sheet_name='Labels', index=False)

print(f"Excel file successfully saved as {output_excel_file}")

#-----------------------------------------------------------------------------------------------------
# respondents section

# Define the input and output file names
if "Master" in input_filename:
    respondents_sav_file = input_filename.replace("Master", "Respondents")
else:
    respondents_sav_file = os.path.splitext(input_filename)[0] + ' Respondents.sav'

# List of columns to include
include_columns = [
    'serial', 'mode', 'month', 'Month_Gr6', 'Quarter', 'group', 'wt_final', 
    'wt_final_online', 'wt_final_B', 'wt_final_C', 'wt_final_AB', 'wt_final_AC', 
    'wt_time', 'wt_online_time', 'xStrata', 'Overall', 'Age16plus', 'Age19plus', 
    'Filter_male16', 'Filter_female16', 'male16plus', 'female16plus', 'Filter_Act', 
    'Filter_Act_M', 'Filter_Act_F', 'Filter_InsAct', 'Filter_InsAct_M', 'Filter_InsAct_F', 
    'Filter_Inact', 'Filter_Inact_M', 'Filter_Inact_F', 'Filter_AnyVol', 'Filter_NotVol', 
    'Filter_VolYr', 'Filter_VolOcc', 'Filter_VolMth', 'Filter_VolWk', 'Filter_ActAnyVol', 
    'Filter_ActNotVol', 'Filter_NotActAnyVol', 'Filter_NotActNotVol', 'age', 'preg', 
    'NSSEC4', 'NSSEC5', 'NSSEC7', 'NSSEC8', 'EducStud2', 'EducStud4', 'Educ6', 
    'WorkStat5', 'WorkStat7', 'WorkStat10', 'Child4', 'HHLiv5', 'HHLiv12', 
    'HHLiv23', 'Age3', 'Age4', 'Age5', 'Age9', 'Age2_75', 'Age2_55', 'Age2_50', 
    'Age2_46', 'AgeTGC', 'Age1640', 'Age1660', 'Age17', 'Gend3', 'GendAge1640', 
    'GendAge1660', 'GendAge2', 'GendAge4', 'GendAge9', 'GendAgeTGC', 'GendDisab2_POP', 
    'GendImpair4', 'GendEth2', 'GendEth7', 'GendRelig7', 'GendOrient4', 'GendNSSEC5', 
    'GendNSSEC7', 'GendNSSEC8', 'Relig7', 'Orient4', 'Eth2', 'Eth7', 'EthAge2', 
    'EthAge4', 'Disab2_POP', 'Disab3', 'Impair4', 'disty1_POP', 'disty2_POP', 
    'disty3_POP', 'disty4_POP', 'disty5_POP', 'disty6_POP', 'disty7_POP', 
    'disty8_POP', 'disty9_POP', 'disty10_POP', 'disty11_POP', 'disty12_POP', 
    'disty13_POP', 'LA_2021', 'LA_2020', 'LA_2019', 'LA_2015', 'LondInOut', 'CSP', 
    'CountyCounc', 'Reg9', 'RegACE', 'AuthACE', 'CoastComm', 'LDP_Donc_Combined', 
    'CCGApr18', 'STPApr18', 'UrbRur10', 'UrbRur6_LA', 'UrbRur2', 'UrbRur6', 
    'IMD10', 'IMD4', 'IMD_Domain_Employ_rank', 'IMD_Domain_Employ_decile', 
    'IMD_Domain_Edu_rank', 'IMD_Domain_Edu_decile', 'IMD_Domain_Health_rank', 
    'IMD_Domain_Health_decile', 'IMD_Domain_Crime_rank', 'IMD_Domain_Crime_decile', 
    'IMD_Domain_Hous_rank', 'IMD_Domain_Hous_decile', 'IMD_Domain_Enviro_rank', 
    'IMD_Domain_Enviro_decile', 'IMD_Domain_Income_rank', 'IMD_Domain_Income_decile', 
    'ONS_SuperGroup', 'ONS_Group', 'DVHeightM', 'DVWeightKG', 'DVBMI', 'BMIG', 
    'indevtry', 'lifesat', 'happy', 'anxious', 'worthw', 'indev', 'comm1', 
    'health', 'lone', 'WellB_Lifesat_GR', 'WellB_Happy_GR', 'WellB_Anxious_GR', 
    'WellB_Worthw_GR', 'InDev_GR', 'Comm1_GR', 'InDevtry_GR', 'FVPor', 'FVPorG', 
    'FVPorG2', 'READYAB1_POP', 'READYOP_COMB2_POP', 'READYOP_CV_3_POP', 'Motiva_POP', 
    'motivb_POP', 'motivc_POP', 'motivd_POP', 'motive_POP', 'volint1', 'volint2', 
    'volint3', 'volint4', 'volint5', 'volint6', 'volint7', 'volint8', 'volint1_vol', 
    'volint2_vol', 'volint3_vol', 'volint4_vol', 'volint5_vol', 'volint6_vol', 
    'volint7_vol', 'volint8_vol', 'VolCnt', 'VolCnt_GR2', 'VolAny', 'VolFrqB_Pop', 
    'VolDur', 'VolDur_GR2', 'VolLong', 'VolLong_GR3'
]

def get_columns_from_user():
    # Commenting out user input for default behavior
    # user_input = input("Enter the columns you want to include, separated by commas, or press Enter to use defaults: ")
    # if user_input.strip():
    #     return [col.strip() for col in user_input.split(',')]
    # else:
    #     return include_columns

    # Always return default columns
    return include_columns

# Get the list of columns to include
include_columns = get_columns_from_user()

# Check for missing columns
missing_columns = [col for col in include_columns if col not in df.columns]
if missing_columns:
    print(f"Warning: The following columns are missing and will be skipped: {missing_columns}")

# Use only the columns that are present in df
available_columns = [col for col in include_columns if col in df.columns]
respondents_df = df[available_columns]

# Save the DataFrame to a .sav file using pyreadstat
pyreadstat.write_sav(respondents_df, respondents_sav_file)

#-----------------------------------------------------------------------------------------------------
# sports section

# Extract columns starting with MONTHS_12_
months_12_columns = [col for col in df.columns if col.startswith("MONTHS_12_")]

# Remove "MONTHS_12_" prefix from each column name
months_12_removed = [col.replace("MONTHS_12_", "") for col in months_12_columns]

# Create a DataFrame with cleaned column names as rows
df_months_12_removed = pd.DataFrame(months_12_removed, columns=["Column Names"])

# Define output CSV filename based on input filename
sports_csv_filename = input_filename.replace("Master.sav", "Sports.csv")

# Save the cleaned DataFrame to a CSV file with only column names
df_months_12_removed.to_csv(sports_csv_filename, index=False, header=False)
print(f"CSV file with cleaned column names saved as {sports_csv_filename}")

#-----------------------------------------------------------------------------------------------------
# categories section

search_string = months_12_removed[0] # Get the first cleaned column name

# Extract columns starting with months_12_removed[0]
categories_columns = [col for col in df.columns if search_string in col]

# Remove months_12_removed[0] prefix from each column name
filtered_columns = [col.replace(search_string, "") for col in categories_columns]

# Create a DataFrame with filtered column names as rows
df_filtered_columns = pd.DataFrame(filtered_columns, columns=["Column Names"])

# Define output CSV filename based on input filename
categories_csv_filename = input_filename.replace("Master.sav", "Categories.csv")

# Save the filtered DataFrame to a CSV file with only column names
df_filtered_columns.to_csv(categories_csv_filename, index=False, header=False)
print(f"CSV file with filtered column names saved as {categories_csv_filename}")

#-----------------------------------------------------------------------------------------------------
# population section

# Define the input filename for the Excel file
population_workbook = "ukpopestimatesmid2020on2021geography.xls"

# Load the data from the "MYE2 - Persons" tab
population_page = pd.read_excel(population_workbook, sheet_name="MYE2 - Persons", header=7)

# Remove unnecessary rows or columns
population_page = population_page.dropna(how='all').reset_index(drop=True)
population_page = population_page.loc[:, ~population_page.columns.str.contains('^Unnamed')]

# Define the age groups for Youth (ages 5–16) and Adults (ages 16+)
# For Youth: columns corresponding to ages 5 to 16 (index range 5 to 16 inclusive)
youth_columns = population_page.columns[9:21]  # Corresponding to ages 5–16 (0-indexed)
# For Adults: columns corresponding to ages 16 and above (index range 16 to 90+)
adult_columns = population_page.columns[20:]  # Corresponding to ages 16 and above (0-indexed)

# Create youth and adult dataframes as copies to avoid SettingWithCopyWarning
youth = population_page[['Code', 'Name', 'Geography'] + list(youth_columns)].copy()
adult = population_page[['Code', 'Name', 'Geography'] + list(adult_columns)].copy()

# Add a new column 'Total' that sums all age columns (excluding 'Code', 'Name', 'Geography')
youth['Total'] = youth.iloc[:, 3:].sum(axis=1)
adult['Total'] = adult.iloc[:, 3:].sum(axis=1)

# Create youth and adult *REGION* dataframes
regions_youth = youth[youth['Geography'] == 'Region']
regions_adult = adult[adult['Geography'] == 'Region']

# Create youth and adult *AUTHORITY* dataframes
authorities_youth = youth[youth['Geography'].isin(['London Borough', 'Metropolitan District', 'Non-metropolitan District', 'Unitary Authority'])]
authorities_adult = adult[adult['Geography'].isin(['London Borough', 'Metropolitan District', 'Non-metropolitan District', 'Unitary Authority'])]

# Create youth and adult *ENGLAND* dataframes
england_youth = youth[youth['Name'] == 'ENGLAND']
england_adult = adult[adult['Name'] == 'ENGLAND']

#-----------------------------------------------------------------------------------------------------
# local authorities section

# Hard-coded for debugging
local_authority = "LA_2021"

# Prompt the user to enter the local authority name
# local_authority = input("Enter the name of the local authority column: ")

# Filter the flattened value labels for the specific year
la_year_labels = df_flattened[df_flattened["Column"] == local_authority]

# Keep only the 'Value' and 'Label' columns
la_year_labels = la_year_labels[["Value", "Label"]]

# Remove rows with negative values in the 'Value' column
la_year_labels = la_year_labels[la_year_labels["Value"] >= 0]

# Split the 'Label' column into 'Code' and 'Area'
la_year_labels[["Code", "Area"]] = la_year_labels["Label"].str.split(" ", n=1, expand=True)

# Drop the original 'Label' column
la_year_labels = la_year_labels.drop(columns=["Label"])

# Merge with adult and youth population data to fill in population columns
la_year_labels = la_year_labels.merge(
    authorities_adult[['Code', 'Total']], on='Code', how='left', suffixes=('', '_Adult')
)
la_year_labels = la_year_labels.merge(
    authorities_youth[['Code', 'Total']], on='Code', how='left', suffixes=('', '_Youth')
)

# Rename the columns to match desired output
la_year_labels.rename(columns={'Total': 'Adult Population', 'Total_Youth': 'YP Population'}, inplace=True)

# Create the filename using the current year
la_filename = input_filename.replace("Master.sav", "Local Authorities.csv")

# Save the filtered DataFrame with the new columns to a CSV file
la_year_labels.to_csv(la_filename, index=False)

print(f"Value labels for {local_authority} extracted and saved to {la_filename}")

#-----------------------------------------------------------------------------------------------------
# regions section

#hard-coded for debugging
region = "Reg9"

# Prompt the user to enter the local authority name
#region = input("Enter the name of the region column: ")

# Filter the flattened value labels for the specific year
region_year_labels = df_flattened[df_flattened["Column"] == region]

# Keep only the 'Value' and 'Label' columns
region_year_labels = region_year_labels[["Value", "Label"]]

# Rename the 'Label' column to 'Area'
region_year_labels = region_year_labels.rename(columns={"Label": "Area"})

# Remove rows with negative values in the 'Value' column
region_year_labels = region_year_labels[region_year_labels["Value"] >= 0]

# Ensure that the 'Area' column in region_year_labels is in title case for comparison
region_year_labels["Area"] = region_year_labels["Area"].str.title()

# Initialize the "Adult Population" and "YP Population" columns
region_year_labels["Adult Population"] = None
region_year_labels["YP Population"] = None

# Loop through each row in region_year_labels to populate the 'Adult Population' and 'YP Population'
for index, row in region_year_labels.iterrows():
    area_name = row["Area"]
    
    # Match with regions_youth (case insensitive)
    youth_match = regions_youth[regions_youth["Name"].str.upper() == area_name.upper()]
    if not youth_match.empty:
        # Assign Total value from regions_youth to 'YP Population'
        region_year_labels.at[index, "YP Population"] = youth_match["Total"].values[0]
    
    # Match with regions_adult (case insensitive)
    adult_match = regions_adult[regions_adult["Name"].str.upper() == area_name.upper()]
    if not adult_match.empty:
        # Assign Total value from regions_adult to 'Adult Population'
        region_year_labels.at[index, "Adult Population"] = adult_match["Total"].values[0]

# Create the filename using the current year
region_filename = input_filename.replace("Master.sav", "Regions.csv")

# Save the filtered DataFrame to a CSV file
region_year_labels.to_csv(region_filename, index=False)

print(f"Value labels for {region} extracted and saved to {region_filename}")

#------------------------------------------------
#make england csv

# Get the Youth and Adult Population from the "Total" column in england_youth and england_adult
youth_population = england_youth['Total'].values[0]  # Total youth population for England
adult_population = england_adult['Total'].values[0]  # Total adult population for England

# Generate the output CSV filename based on the input filename
england_year_labels = input_filename.replace("Master.sav", "England.csv")

# Create a DataFrame for the blank CSV with the specified structure
data = {
    "Area": ["England"],
    "Adult Population": [adult_population],
    "YP Population": [youth_population]
}

df_blank = pd.DataFrame(data)

# Save the DataFrame to a CSV file
df_blank.to_csv(england_year_labels, index=False)

#===========================
#from second program section
#===========================

import pandas as pd
import pyreadstat
import math

sports = {
    "Combined": "SPORTCOUNT_A01",
    "Archery": "ARCHERY_J01",
    "Athletics": "ATHLETICS_D03",
    "Badminton": "BADMINTON_G02",
    "Basketball": "BASKETBALL_F07",
    "BMX": "CYCBMX_N03",
    "Bowls": "BOWLS_L10",
    "Canoeing": "CANOEING_H09",
    "Cricket": "CRICKET_F02",
    "Cycling": "CYCALL_C02",
    "Driving Range": "DRIVINGRANGE_U10",
    "Fitness": "FITNESS_B06",
    "Football": "FOOTBALL_F01",
    "Golf": "GOLF_L08",
    "Gymnastics": "GYMNASTICS_L13",
    "Hockey": "HOCKEY_F09",
    "Judo": "JUDO_J04",
    "Martial Arts": "MARTIAL_J07",
    "Netball": "NETBALL_F06",
    "Parkour": "PARKOUR_H05",
    "Rollerskating": "ROLLERSKATING_L16",
    "Rounders": "ROUNDERS_F11",
    "Rugby League": "RUGBYLEAGUE_F04",
    "Rugby Union": "RUGBYUNION_F03",
    "Skateboarding": "SKATEBOARDING_L17",
    "Squash": "SQUASH_G04",
    "Swimming": "SWIM_L01",
    "Table Tennis": "TABLETENNIS_G03",
    "Tennis": "TENNIS_G01",
    "Triathlon": "TRIATHLON_L05",
    "Volleyball": "VOLLEYBALL_F10",
    "Weightlifting": "WEIGHTLIFTING_P19",
    "Wrestling": "WRESTLING_J09",
}

default_activities = [
    "MONTHS_12", "ACTYRA", "ACTYRB", "ACTYRC", "ACTYR_7", "ACTYR_4", "ACTYR_3", "FREQUENCY", 
    "DURATION", "DUR_LHT", "DUR_MOD", "DUR_HVY", "DAYS10P", "MEMS7", "WHOWITHA", "WHOWITHB", 
    "WHOWITHC", "WHOWITHD", "CLUB", "DAYS10P60", "MINS_SESS", "FREQUENCYGR", "DURATIONGR", 
    "DURATION1PL", "DAYS10P60GR", "MEMS7GR", "ACT7GR", "MEMS7GR30MIN", "Mins_Sess_GR5min", 
    "Mins_Sess_GR4", "MEMS7_IN", "MEMS7_OUT", "MEMS7_IN_HOME", "MEMS7_IN_LEISURE", 
    "MEMS7_IN_COMMUNITY", "MEMS7_IN_SPECIALIST", "MEMS7_IN_OTHER", "MEMS7_OUT_BUILT", 
    "MEMS7_OUT_LOCAL", "MEMS7_OUT_COUNTRYCOAST", "MEMS7_OUT_OTHER", "MEMS7_OUT_LOCAL_HOME", 
    "MEMS7_OUT_LOCAL_PARK", "MEMS7_OUT_LOCAL_ROAD", "MEMS7_OUT_COUNTRYCOAST_WATER", 
    "MEMS7_OUT_COUNTRYCOAST_LAND", "MEMS7_OUT_BUILT_LEISURE", "MEMS7_OUT_BUILT_SPECIALIST", 
    "MEMS7GR_IN", "MEMS7GR_OUT", "MEMS7GR_IN_HOME", "MEMS7GR_IN_LEISURE", "MEMS7GR_IN_COMMUNITY", 
    "MEMS7GR_IN_SPECIALIST", "MEMS7GR_IN_OTHER", "MEMS7GR_OUT_BUILT", "MEMS7GR_OUT_LOCAL", 
    "MEMS7GR_OUT_COUNTRYCOAST", "MEMS7GR_OUT_OTHER", "MEMS7GR_OUT_LOCAL_HOME", 
    "MEMS7GR_OUT_LOCAL_PARK", "MEMS7GR_OUT_LOCAL_ROAD", "MEMS7GR_OUT_COUNTRYCOAST_WATER", 
    "MEMS7GR_OUT_COUNTRYCOAST_LAND", "MEMS7GR_OUT_BUILT_LEISURE", "MEMS7GR_OUT_BUILT_SPECIALIST", 
    "MEMS7_INOUT_HOME", "MEMS7_INOUT_LEISURE", "MEMS7_INOUT_SPECIALIST", "MEMS7_INOUT_OTHER", 
    "MEMS7_INOUT_BUILT", "MEMS7GR_INOUT_HOME", "MEMS7GR_INOUT_LEISURE", "MEMS7GR_INOUT_SPECIALIST", 
    "MEMS7GR_INOUT_OTHER", "MEMS7GR_INOUT_BUILT", "MEMS7GR2", "MUSCLE7GR"
]

for sport_or_sport_group_name, sport_codes_input in sports.items():
    print(f"Processing: {sport_or_sport_group_name} ({sport_codes_input})")

    # Split the input into a list of sport codes, removing leading/trailing whitespace and underscores
    sport_codes = [code.strip('_').strip() for code in sport_codes_input.split(",")]

    # Always use the default list
    activities = [activity.strip('_').strip() for activity in default_activities]

    # Generate concatenated pairs
    concatenated_pairs = [f"{activity}_{sport_code}" for sport_code in sport_codes for activity in activities]

    # Added "Age4" to the list of columns to load
    columns_to_load = ["wt_final", "serial", "Reg9", "LA_2021", "Age4", "Age5", "Gend3", "Eth7"] + concatenated_pairs

    # Load the .sav file and filter the data by selected columns
    sav_file_path = "Nov 20-21 Full-Year Master.sav"

    # Load the .sav file with only the relevant columns, including "serial"
    df_sav, meta = pyreadstat.read_sav(sav_file_path, usecols=columns_to_load)

    # Add new columns MONTHS_12 and Days10P
    df_sav["MONTHS_12"] = 0  # Start with 0
    df_sav["Days10P"] = 0  # Start with 0
    df_sav["MEMS7"] = 0  # Start with 0
    df_sav["FREQUENCYGR"] = 0  # Start with 0

    # Update MONTHS_12 by summing columns starting with "MONTHS_12_"
    for col in df_sav.columns:
        if col.startswith("MONTHS_12_"):
            df_sav["MONTHS_12"] += df_sav[col].apply(lambda x: max(x, 0))  # Replace negative values with 0

    # Update Days10P by summing columns starting with "DAYS10P60GR_"
    for col in df_sav.columns:
        if col.startswith("DAYS10P60GR_"):
            df_sav["Days10P"] += df_sav[col].apply(lambda x: max(x, 0))  # Replace negative values with 0

    # Update MEMS7 by summing columns starting with "MEMS7_"
    for col in df_sav.columns:
        if col.startswith("MEMS7_"):
            df_sav["MEMS7"] += df_sav[col].apply(lambda x: max(x, 0))  # Replace negative values with 0

    # Update FREQUENCYGR by summing columns starting with "FREQUENCYGR_"
    for col in df_sav.columns:
        if col.startswith("FREQUENCYGR_"):
            df_sav["FREQUENCYGR"] += df_sav[col].apply(lambda x: max(x, 0))  # Replace negative values with 0

    #=========================================================================================================
    #CALCULATIONS SECTION

    #Participated: If MONTHS_12 >= 1 then wt_final else 0
    df_sav["Participated"] = df_sav.apply(lambda row: row["wt_final"] if row["MONTHS_12"] >= 1 else 0, axis=1)

    #Not_Participated: If MONTHS_12 < 1 then wt_final else 0
    df_sav["Not_Participated"] = df_sav.apply(lambda row: row["wt_final"] if row["MONTHS_12"] < 1 else 0, axis=1)

    #Total: Participated + Not Participated
    df_sav["Total"] = df_sav["Participated"] + df_sav["Not_Participated"]

    #Regularly_Participated: If Days10P > 1 then wt_final else 0
    df_sav["Regularly_Participated"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 else 0, axis=1)

    #Inactive: If MEMS7 < 30 then wt_final else 0
    df_sav["Inactive"] = df_sav.apply(lambda row: row["wt_final"] if row["MEMS7"] < 30 else 0, axis=1)

    #Fairly_Active: If MEMS7 > 29 and MEMS7 < 150 then wt_final else 0
    df_sav["Fairly_Active"] = df_sav.apply(lambda row: row["wt_final"] if 29 < row["MEMS7"] < 150 else 0, axis=1)

    #Active: If MEMS7 > 149 then wt_final else 0
    df_sav["Active"] = df_sav.apply(lambda row: row["wt_final"] if row["MEMS7"] > 149 else 0, axis=1)

    #Regularly_Participated_Inactive: If Days10P > 1 and MEMS7 < 30 then wt_final else 0
    df_sav["Regularly_Participated_Inactive"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["MEMS7"] < 30 else 0, axis=1)

    #Regularly_Participated_Fairly_Active: If Days10P > 1 and MEMS7 > 29 and MEMS7 < 150 then wt_final else 0
    df_sav["Regularly_Participated_Fairly_Active"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and 29 < row["MEMS7"] < 150 else 0, axis=1)

    #Regularly_Participated_Active: If Days10P > 1 and MEMS7 > 149 then wt_final else 0
    df_sav["Regularly_Participated_Active"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["MEMS7"] > 149 else 0, axis=1)

    #Regularly_Participated_Male: If Days10P > 1 and Gend3 == 1 then wt_final else 0
    df_sav["Regularly_Participated_Male"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Gend3"] == 1 else 0, axis=1)

    #Regularly_Participated_Female: If Days10P > 1 and Gend3 == 2 then wt_final else 0
    df_sav["Regularly_Participated_Female"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Gend3"] == 2 else 0, axis=1)

    #Regularly_Participated_16_to_24: If Days10P > 1 and Age5 == 2 then wt_final else 0
    df_sav["Regularly_Participated_16_to_24"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Age5"] == 2 else 0, axis=1)

    #Regularly_Participated_25_to_44: If Days10P > 1 and Age5 == 3 then wt_final else 0
    df_sav["Regularly_Participated_25_to_44"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Age5"] == 3 else 0, axis=1)

    #Regularly_Participated_45_to_64: If Days10P > 1 and Age5 == 4 then wt_final else 0
    df_sav["Regularly_Participated_45_to_64"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Age5"] == 4 else 0, axis=1)

    #Regularly_Participated_65_plus: If Days10P > 1 and Age5 == 5 then wt_final else 0
    df_sav["Regularly_Participated_65_plus"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Age5"] == 5 else 0, axis=1)

    #Regularly_Participated_White_British: If Days10P > 1 and Eth7 == 1 then wt_final else 0
    df_sav["Regularly_Participated_White_British"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 1 else 0, axis=1)

    #Regularly_Participated_White_Other: If Days10P > 1 and Eth7 == 2 then wt_final else 0
    df_sav["Regularly_Participated_White_Other"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 2 else 0, axis=1)

    #Regularly_Participated_South_Asian: If Days10P > 1 and Eth7 == 3 then wt_final else 0
    df_sav["Regularly_Participated_South_Asian"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 3 else 0, axis=1)

    #Regularly_Participated_Black: If Days10P > 1 and Eth7 == 4 then wt_final else 0
    df_sav["Regularly_Participated_Black"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 4 else 0, axis=1)

    #Regularly_Participated_Chinese: If Days10P > 1 and Eth7 == 5 then wt_final else 0
    df_sav["Regularly_Participated_Chinese"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 5 else 0, axis=1)

    #Regularly_Participated_Mixed: If Days10P > 1 and Eth7 == 6 then wt_final else 0
    df_sav["Regularly_Participated_Mixed"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 6 else 0, axis=1)

    #Regularly_Participated_Other: If Days10P > 1 and Eth7 == 7 then wt_final else 0
    df_sav["Regularly_Participated_Other"] = df_sav.apply(lambda row: row["wt_final"] if row["Days10P"] > 1 and row["Eth7"] == 7 else 0, axis=1)

    #Participated_Male: if MONTHS_12 >= 1 and Gend3 == 1 then wt_final else 0
    df_sav["Participated_Male"] = df_sav.apply(lambda row: row["wt_final"] if row["MONTHS_12"] >= 1 and row["Gend3"] == 1 else 0, axis=1)

    #Participated_Female: if MONTHS_12 >= 1 and Gend3 == 2 then wt_final else 0
    df_sav["Participated_Female"] = df_sav.apply(lambda row: row["wt_final"] if row["MONTHS_12"] >= 1 and row["Gend3"] == 2 else 0, axis=1)

    #END CALCULATIONS SECTION
    #==================================================

    # Replace numeric values with labels
    value_labels = meta.variable_value_labels
    if "Reg9" in value_labels:
        reg9_labels = value_labels["Reg9"]
        df_sav["Reg9"] = df_sav["Reg9"].map(reg9_labels)

    if "LA_2021" in value_labels:
        la_labels = value_labels["LA_2021"]
        df_sav["LA_2021"] = df_sav["LA_2021"].map(la_labels)

        # Trim the first 10 characters from each value in LA_2021 if it corresponds to Authorities
        df_sav["LA_2021"] = df_sav["LA_2021"].apply(lambda x: x[10:] if isinstance(x, str) else x)

    if "Age4" in value_labels:
        age4_labels = value_labels["Age4"]
        df_sav["Age4"] = df_sav["Age4"].map(age4_labels)

    if "Age5" in value_labels:
        age5_labels = value_labels["Age5"]
        df_sav["Age5"] = df_sav["Age5"].map(age5_labels)

    if "Gend3" in value_labels:
        gend3_labels = value_labels["Gend3"]
        df_sav["Gend3"] = df_sav["Gend3"].map(gend3_labels)

    if "Eth7" in value_labels:
        eth7_labels = value_labels["Eth7"]
        df_sav["Eth7"] = df_sav["Eth7"].map(eth7_labels)

    # Read the CSV files
    england_df = pd.read_csv('Nov 20-21 Full-Year England.csv')
    local_authorities_df = pd.read_csv('Nov 20-21 Full-Year Local Authorities.csv')
    regions_df = pd.read_csv('Nov 20-21 Full-Year Regions.csv')

    # Extract relevant columns from each dataframe
    england_data = england_df[['Area', 'Adult Population']]
    local_authorities_data = local_authorities_df[['Area', 'Adult Population']]
    regions_data = regions_df[['Area', 'Adult Population']]

    # Concatenate the dataframes
    combined_population_df = pd.concat([england_data, local_authorities_data, regions_data], ignore_index=True)

    #==================================================================================    

    # Save the filtered DataFrame to a .sav file
    output_sav_path = f"Nov 20-21 Full-Year {sport_or_sport_group_name}.sav"
    pyreadstat.write_sav(df_sav, output_sav_path)
    print(f"The filtered data from the .sav file has been saved to {output_sav_path}")

    # Save the filtered DataFrame to a .csv file
    output_csv_path = f"Nov 20-21 Full-Year {sport_or_sport_group_name}.csv"
    df_sav.to_csv(output_csv_path, index=False)
    print(f"The filtered data has also been saved to {output_csv_path}")

    #==================================================================================

    # Calculate totals for "England"
    totals = {
        "Area": "England",
        "Level": "Country",
        **df_sav[[
            "Participated", "Not_Participated", "Total",
            "Regularly_Participated", "Inactive", "Fairly_Active", "Active",
            "Regularly_Participated_Inactive", "Regularly_Participated_Fairly_Active",
            "Regularly_Participated_Active",
            "Regularly_Participated_Male", "Regularly_Participated_Female",  
            "Regularly_Participated_16_to_24", "Regularly_Participated_25_to_44", 
            "Regularly_Participated_45_to_64", "Regularly_Participated_65_plus",
            "Regularly_Participated_White_British", "Regularly_Participated_White_Other",
            "Regularly_Participated_South_Asian", "Regularly_Participated_Black",
            "Regularly_Participated_Chinese", "Regularly_Participated_Mixed",
            "Regularly_Participated_Other",
            "Participated_Male", "Participated_Female"
        ]].sum()
    }

    # Calculate totals for each Region
    regions = df_sav["Reg9"].unique()
    region_totals = [
        {
            "Area": region,
            "Level": "Region",
            **df_sav[df_sav["Reg9"] == region][[ 
            "Participated", "Not_Participated", "Total",
            "Regularly_Participated", "Inactive", "Fairly_Active", "Active",
            "Regularly_Participated_Inactive", "Regularly_Participated_Fairly_Active",
            "Regularly_Participated_Active",
            "Regularly_Participated_Male", "Regularly_Participated_Female",  
            "Regularly_Participated_16_to_24", "Regularly_Participated_25_to_44", 
            "Regularly_Participated_45_to_64", "Regularly_Participated_65_plus",
            "Regularly_Participated_White_British", "Regularly_Participated_White_Other",
            "Regularly_Participated_South_Asian", "Regularly_Participated_Black",
            "Regularly_Participated_Chinese", "Regularly_Participated_Mixed",
            "Regularly_Participated_Other",
            "Participated_Male", "Participated_Female"
            ]].sum()
        }
        for region in regions
    ]

    # Calculate totals for each Authority
    authorities = df_sav["LA_2021"].unique()
    authority_totals = [
        {
            "Area": authority,
            "Level": "Authority",
            **df_sav[df_sav["LA_2021"] == authority][[ 
            "Participated", "Not_Participated", "Total",
            "Regularly_Participated", "Inactive", "Fairly_Active", "Active",
            "Regularly_Participated_Inactive", "Regularly_Participated_Fairly_Active",
            "Regularly_Participated_Active",
            "Regularly_Participated_Male", "Regularly_Participated_Female",  
            "Regularly_Participated_16_to_24", "Regularly_Participated_25_to_44", 
            "Regularly_Participated_45_to_64", "Regularly_Participated_65_plus",
            "Regularly_Participated_White_British", "Regularly_Participated_White_Other",
            "Regularly_Participated_South_Asian", "Regularly_Participated_Black",
            "Regularly_Participated_Chinese", "Regularly_Participated_Mixed",
            "Regularly_Participated_Other",
            "Participated_Male", "Participated_Female"
            ]].sum()
        }
        for authority in authorities
    ]

    # Combine totals into a single DataFrame
    combined_data = [totals] + region_totals + authority_totals
    combined_df = pd.DataFrame(combined_data)

    # Sort the DataFrame by the "Area" column alphabetically
    combined_df = combined_df.sort_values(by="Area", ascending=True)

    # Drop the "Level" column before saving the DataFrame
    combined_df = combined_df.drop(columns=["Level"])

    # Convert the 'Area' column values to title case in both dataframes before merging
    combined_df['Area'] = combined_df['Area'].str.title()
    combined_population_df['Area'] = combined_population_df['Area'].str.title()

    # Merge population data into combined_df based on the 'Area' column
    combined_df = combined_df.merge(combined_population_df.rename(columns={"Adult Population": "Pop"}), on="Area", how="left")

    #===========================================================================================================
    #CALCULATIONS SECTION 2

    #Percent_Participated: Participated / Total
    combined_df["Percent_Participated"] = combined_df["Participated"] / combined_df["Total"]

    #Percent_Not_Participated: Not Participated / Total
    combined_df["Percent_Not_Participated"] = combined_df["Not_Participated"] / combined_df["Total"]

    #Percent_Regularly_Participated: Regularly Participated / Total
    combined_df["Percent_Regularly_Participated"] = combined_df["Regularly_Participated"] / combined_df["Total"]

    #Percent_Inactive: Inactive / Total
    combined_df["Percent_Inactive"] = combined_df["Inactive"] / combined_df["Total"]

    #Percent_Fairly_Active: Fairly Active / Total
    combined_df["Percent_Fairly_Active"] = combined_df["Fairly_Active"] / combined_df["Total"]

    #Percent_Active: Active / Total
    combined_df["Percent_Active"] = combined_df["Active"] / combined_df["Total"]

    #Percent_Regularly_Participated_Inactive: If Regularly_Participated > 0 then Regularly_Participated_Inactive / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Inactive"] = combined_df.apply(lambda row: row["Regularly_Participated_Inactive"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Fairly_Active: Percent_Regularly_Participated_Fairly_Active - If Regularly_Participated > 0 then Regularly_Participated_Fairly_Active / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Fairly_Active"] = combined_df.apply(lambda row: row["Regularly_Participated_Fairly_Active"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Fairly_Active: Percent_Regularly_Participated_Active - If Regularly_Participated > 0 then Regularly_Participated_Active / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Active"] = combined_df.apply(lambda row: row["Regularly_Participated_Active"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Male: Percent_Regularly_Participated_Male - If Regularly_Participated > 0 then Regularly_Participated_Male / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Male"] = combined_df.apply(lambda row: row["Regularly_Participated_Male"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Female: Percent_Regularly_Participated_Female - If Regularly_Participated > 0 then Regularly_Participated_Female / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Female"] = combined_df.apply(lambda row: row["Regularly_Participated_Female"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_16_to_24: Percent_Regularly_Participated_16_to_24 - If Regularly_Participated > 0 then Regularly_Participated_16_to_24 / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_16_to_24"] = combined_df.apply(lambda row: row["Regularly_Participated_16_to_24"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_25_to_44: Percent_Regularly_Participated_25_to_44 - If Regularly_Participated > 0 then Regularly_Participated_25_to_44 / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_25_to_44"] = combined_df.apply(lambda row: row["Regularly_Participated_25_to_44"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_45_to_64: Percent_Regularly_Participated_45_to_64 - If Regularly_Participated > 0 then Regularly_Participated_45_to_64 / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_45_to_64"] = combined_df.apply(lambda row: row["Regularly_Participated_45_to_64"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_65_plus: Percent_Regularly_Participated_65_plus - If Regularly_Participated > 0 then Regularly_Participated_65_plus / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_65_plus"] = combined_df.apply(lambda row: row["Regularly_Participated_65_plus"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_White_British: Percent_Regularly_Participated_White_British - If Regularly_Participated > 0 then Regularly_Participated_White_British / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_White_British"] = combined_df.apply(lambda row: row["Regularly_Participated_White_British"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_White_Other: Percent_Regularly_Participated_White_Other - If Regularly_Participated > 0 then Regularly_Participated_White_Other / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_White_Other"] = combined_df.apply(lambda row: row["Regularly_Participated_White_Other"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_South_Asian: Percent_Regularly_Participated_South_Asian - If Regularly_Participated > 0 then Regularly_Participated_South_Asian / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_South_Asian"] = combined_df.apply(lambda row: row["Regularly_Participated_South_Asian"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Black: Percent_Regularly_Participated_Black - If Regularly_Participated > 0 then Regularly_Participated_Black / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Black"] = combined_df.apply(lambda row: row["Regularly_Participated_Black"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Chinese: Percent_Regularly_Participated_Chinese - If Regularly_Participated > 0 then Regularly_Participated_Chinese / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Chinese"] = combined_df.apply(lambda row: row["Regularly_Participated_Chinese"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Mixed: Percent_Regularly_Participated_Mixed - If Regularly_Participated > 0 then Regularly_Participated_Mixed / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Mixed"] = combined_df.apply(lambda row: row["Regularly_Participated_Mixed"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Regularly_Participated_Other: Percent_Regularly_Participated_Other - If Regularly_Participated > 0 then Regularly_Participated_Other / Regularly_Participated else 0
    combined_df["Percent_Regularly_Participated_Other"] = combined_df.apply(lambda row: row["Regularly_Participated_Other"] / row["Regularly_Participated"] if row["Regularly_Participated"] > 0 else 0, axis=1)

    #Percent_Participated_Male: Percent_Participated_Male - If Total > 0 then Participated_Male / Total else 0
    combined_df["Percent_Participated_Male"] = combined_df.apply(lambda row: row["Participated_Male"] / row["Total"] if row["Total"] > 0 else 0, axis=1)

    #Percent_Participated_Female: Percent_Participated_Female - If Total > 0 then Participated_Feale / Total else 0
    combined_df["Percent_Participated_Female"] = combined_df.apply(lambda row: row["Participated_Female"] / row["Total"] if row["Total"] > 0 else 0, axis=1)

    #Pop_Participated: Pop * Percent_Participated
    combined_df["Pop_Participated"] = (combined_df["Pop"] * combined_df["Percent_Participated"]).apply(math.floor)

    #Pop_Not_Participated: Pop * Percent_Not_Participated
    combined_df["Pop_Not_Participated"] = (combined_df["Pop"] * combined_df["Percent_Not_Participated"]).apply(math.floor)

    #Pop_Inactive: Pop * Percent_Inactive
    combined_df["Pop_Inactive"] = (combined_df["Pop"] * combined_df["Percent_Inactive"]).apply(math.floor)

    #Pop_Fairly_Active: Pop * Percent_Fairly_Active
    combined_df["Pop_Fairly_Active"] = (combined_df["Pop"] * combined_df["Percent_Fairly_Active"]).apply(math.floor)

    #Pop_Active: Pop * Percent_Active
    combined_df["Pop_Active"] = (combined_df["Pop"] * combined_df["Percent_Active"]).apply(math.floor)

    #Pop_Regularly_Participated: Pop * Percent_Regularly_Participated
    combined_df["Pop_Regularly_Participated"] = (combined_df["Pop"] * combined_df["Percent_Regularly_Participated"]).apply(math.floor)

    #Pop_Regularly_Participated_Inactive: Pop_Regularly_Participated * Percent_Regularly_Participated_Inactive
    combined_df["Pop_Regularly_Participated_Inactive"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Inactive"]).apply(math.floor)

    #Pop_Regularly_Participated_Fairly_Active: Pop_Regularly_Participated * Percent_Regularly_Participated_Fairly_Active
    combined_df["Pop_Regularly_Participated_Fairly_Active"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Fairly_Active"]).apply(math.floor)

    #Pop_Regularly_Participated_Active: Pop_Regularly_Participated * Percent_Regularly_Participated_Active
    combined_df["Pop_Regularly_Participated_Active"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Active"]).apply(math.floor)

    #Pop_Regularly_Participated_Male: Pop_Regularly_Participated * Percent_Regularly_Participated_Male
    combined_df["Pop_Regularly_Participated_Male"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Male"]).apply(math.floor)

    #Pop_Regularly_Participated_Female: Pop_Regularly_Participated * Percent_Regularly_Participated_Female
    combined_df["Pop_Regularly_Participated_Female"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Female"]).apply(math.floor)

    #Pop_Regularly_Participated_16_to_24: Pop_Regularly_Participated * Percent_Regularly_Participated_16_to_24
    combined_df["Pop_Regularly_Participated_16_to_24"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_16_to_24"]).apply(math.floor)

    #Pop_Regularly_Participated_25_to_44: Pop_Regularly_Participated * Percent_Regularly_Participated_25_to_44
    combined_df["Pop_Regularly_Participated_25_to_44"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_25_to_44"]).apply(math.floor)

    #Pop_Regularly_Participated_45_to_64: Pop_Regularly_Participated * Percent_Regularly_Participated_45_to_64
    combined_df["Pop_Regularly_Participated_45_to_64"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_45_to_64"]).apply(math.floor)

    #Pop_Regularly_Participated_65_plus: Pop_Regularly_Participated * Percent_Regularly_Participated_65_plus
    combined_df["Pop_Regularly_Participated_65_plus"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_65_plus"]).apply(math.floor)

    #Pop_Regularly_Participated_White_British: Pop_Regularly_Participated * Percent_Regularly_Participated_White_British
    combined_df["Pop_Regularly_Participated_White_British"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_White_British"]).apply(math.floor)

    #Pop_Regularly_Participated_White_Other: Pop_Regularly_Participated * Percent_Regularly_Participated_White_Other
    combined_df["Pop_Regularly_Participated_White_Other"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_White_Other"]).apply(math.floor)

    #Pop_Regularly_Participated_South_Asian: Pop_Regularly_Participated * Percent_Regularly_Participated_South_Asian
    combined_df["Pop_Regularly_Participated_South_Asian"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_South_Asian"]).apply(math.floor)

    #Pop_Regularly_Participated_Black: Pop_Regularly_Participated * Percent_Regularly_Participated_Black
    combined_df["Pop_Regularly_Participated_Black"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Black"]).apply(math.floor)

    #Pop_Regularly_Participated_Chinese: Pop_Regularly_Participated * Percent_Regularly_Participated_Chinese
    combined_df["Pop_Regularly_Participated_Chinese"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Chinese"]).apply(math.floor)

    #Pop_Regularly_Participated_Mixed: Pop_Regularly_Participated * Percent_Regularly_Participated_Mixed
    combined_df["Pop_Regularly_Participated_Mixed"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Mixed"]).apply(math.floor)

    #Pop_Regularly_Participated_Other: Pop_Regularly_Participated * Percent_Regularly_Participated_Other
    combined_df["Pop_Regularly_Participated_Other"] = (combined_df["Pop_Regularly_Participated"] * combined_df["Percent_Regularly_Participated_Other"]).apply(math.floor)

    #Pop_Participated_Male: Pop * Percent_Participated_Male
    combined_df["Pop_Participated_Male"] = (combined_df["Pop"] * combined_df["Percent_Participated_Male"]).apply(math.floor)

    #Pop_Participated_Female: Pop * Percent_Participated_Female
    combined_df["Pop_Participated_Female"] = (combined_df["Pop"] * combined_df["Percent_Participated_Female"]).apply(math.floor)

    #Male_Latent_Demand: Pop_Participated_Male - Pop_Regularly_Participated_Male
    combined_df["Male_Latent_Demand"] = (combined_df["Pop_Participated_Male"] - combined_df["Pop_Regularly_Participated_Male"]).apply(math.floor)

    #Female_Latent_Demand: Pop_Participated_Female - Pop_Regularly_Participated_Female
    combined_df["Female_Latent_Demand"] = (combined_df["Pop_Participated_Female"] - combined_df["Pop_Regularly_Participated_Female"]).apply(math.floor)

    #END CALCULATIONS SECTION 2
    #==================================================================================================================================================

    # List of columns to round to the nearest 100
    columns_to_round_100 = [
        "Pop_Participated",
        "Pop_Not_Participated",
        "Pop_Regularly_Participated",
        "Pop_Inactive",
        "Pop_Fairly_Active",
        "Pop_Active",
        "Pop_Regularly_Participated_Inactive",
        "Pop_Regularly_Participated_Fairly_Active",
        "Pop_Regularly_Participated_Active",
        "Pop_Regularly_Participated_Male",
        "Pop_Regularly_Participated_Female",
        "Pop_Regularly_Participated_16_to_24",
        "Pop_Regularly_Participated_25_to_44",
        "Pop_Regularly_Participated_45_to_64",
        "Pop_Regularly_Participated_65_plus",
        "Pop_Regularly_Participated_White_British",
        "Pop_Regularly_Participated_White_Other",
        "Pop_Regularly_Participated_South_Asian",
        "Pop_Regularly_Participated_Black",
        "Pop_Regularly_Participated_Chinese",
        "Pop_Regularly_Participated_Mixed",
        "Pop_Regularly_Participated_Other",
        "Pop_Participated_Male",
        "Pop_Participated_Female",
        "Male_Latent_Demand",
        "Female_Latent_Demand"
    ]

    # Round the specified columns to the nearest 100
    combined_df[columns_to_round_100] = combined_df[columns_to_round_100].apply(lambda col: col.map(lambda x: round(x, -2)))

    columns_to_round_2dp = [
        "Participated",
        "Not_Participated",
        "Total",
        "Regularly_Participated",
        "Inactive",
        "Fairly_Active",
        "Active",
        "Regularly_Participated_Inactive",
        "Regularly_Participated_Fairly_Active",
        "Regularly_Participated_Active",
        "Regularly_Participated_Male",
        "Regularly_Participated_Female",
        "Regularly_Participated_16_to_24",
        "Regularly_Participated_25_to_44",
        "Regularly_Participated_45_to_64",
        "Regularly_Participated_65_plus",
        "Regularly_Participated_White_British",
        "Regularly_Participated_White_Other",
        "Regularly_Participated_South_Asian",
        "Regularly_Participated_Black",
        "Regularly_Participated_Chinese",
        "Regularly_Participated_Mixed",
        "Regularly_Participated_Other",
        "Participated_Male",
        "Participated_Female"
    ]

    combined_df[columns_to_round_2dp] = combined_df[columns_to_round_2dp].round(2)

    # Desired column order
    desired_order = [
        "Area", "Participated", "Not_Participated", "Total", "Regularly_Participated",
        "Inactive", "Fairly_Active", "Active", "Regularly_Participated_Inactive",
        "Regularly_Participated_Fairly_Active", "Regularly_Participated_Active",
        "Regularly_Participated_Male", "Regularly_Participated_Female",
        "Regularly_Participated_16_to_24", "Regularly_Participated_25_to_44",
        "Regularly_Participated_45_to_64", "Regularly_Participated_65_plus",
        "Regularly_Participated_White_British", "Regularly_Participated_White_Other",
        "Regularly_Participated_South_Asian", "Regularly_Participated_Black",
        "Regularly_Participated_Chinese", "Regularly_Participated_Mixed",
        "Regularly_Participated_Other", "Percent_Participated",
        "Percent_Not_Participated", "Percent_Regularly_Participated", "Percent_Inactive",
        "Percent_Fairly_Active", "Percent_Active", "Percent_Regularly_Participated_Inactive",
        "Percent_Regularly_Participated_Fairly_Active", "Percent_Regularly_Participated_Active",
        "Percent_Regularly_Participated_Male", "Percent_Regularly_Participated_Female",
        "Percent_Regularly_Participated_16_to_24", "Percent_Regularly_Participated_25_to_44",
        "Percent_Regularly_Participated_45_to_64", "Percent_Regularly_Participated_65_plus",
        "Percent_Regularly_Participated_White_British", "Percent_Regularly_Participated_White_Other",
        "Percent_Regularly_Participated_South_Asian", "Percent_Regularly_Participated_Black",
        "Percent_Regularly_Participated_Chinese", "Percent_Regularly_Participated_Mixed",
        "Percent_Regularly_Participated_Other", "Pop", "Pop_Participated", "Pop_Not_Participated",
        "Pop_Inactive", "Pop_Fairly_Active", "Pop_Active", "Pop_Regularly_Participated",
        "Pop_Regularly_Participated_Inactive", "Pop_Regularly_Participated_Fairly_Active",
        "Pop_Regularly_Participated_Active", "Pop_Regularly_Participated_Male",
        "Pop_Regularly_Participated_Female", "Pop_Regularly_Participated_16_to_24",
        "Pop_Regularly_Participated_25_to_44", "Pop_Regularly_Participated_45_to_64",
        "Pop_Regularly_Participated_65_plus", "Pop_Regularly_Participated_White_British",
        "Pop_Regularly_Participated_White_Other", "Pop_Regularly_Participated_South_Asian",
        "Pop_Regularly_Participated_Black", "Pop_Regularly_Participated_Chinese",
        "Pop_Regularly_Participated_Mixed", "Pop_Regularly_Participated_Other",
        "Participated_Male", "Participated_Female", "Percent_Participated_Male",
        "Percent_Participated_Female", "Pop_Participated_Male", "Pop_Participated_Female",
        "Male_Latent_Demand", "Female_Latent_Demand"
    ]

    # Reorder the DataFrame
    combined_df = combined_df[desired_order]

    # Save the reordered DataFrame to CSV
    output_csv_path = f"export{sport_or_sport_group_name} Age16+.csv"
    combined_df.to_csv(output_csv_path, index=False)
    print(f"Data has been saved to {output_csv_path}")






