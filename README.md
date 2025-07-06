# active-lives
Active Lives SPSS to CSVS

Walkthrough including code: https://docs.google.com/document/d/1k9r3TWJMUNgx0z5Zg_t7FR1HXuF32SPQ/

The script is a data processing pipeline for:<br>
- Importing and cleaning SPSS data on sports participation.<br>
- Extracting metadata and respondent demographics.<br>
- Integrating external population data.<br>
- Computing detailed participation metrics and demographic breakdowns.<br>
- Aggregating data by geography (national, regional, local).<br>
- Generating multiple output files for reporting and further analysis.

________________________________________

More specifically:

1. Input and Metadata Handling<br>
- Reads an SPSS data file named "Nov 20-21 Full-Year Master.sav".<br>
- Extracts both the dataset and metadata (variable names, labels, coded values).<br>
- Creates an Excel file with multiple sheets containing:<br>
	- Info: Summary and notes.<br>
	- Contents: Overview.<br>
	- Columns: Variable details.<br>
	- Labels: Value labels (mapping codes to descriptions).

2. Respondents Data<br>
- Generates a filename for respondents data by replacing "Master" with "Respondents".<br>
- Filters relevant demographic/behavioral columns.<br>
- Saves this filtered subset as a new SPSS file.

3. Sports Participation Data<br>
- Extracts columns related to sports participation (prefix "MONTHS_12_").<br>
- Cleans column names and saves the list of sports participation variables to Sports.csv.

4. Categorization<br>
- Extracts variables by category prefixes.<br>
- Saves cleaned category variable names to Categories.csv.

5. Population Estimates<br>
- Loads population data from an Excel file (ukpopestimatesmid2020on2021geography.xls).<br>
- Breaks down population by youth (5-16 years) and adults (16+ years).<br>
- Aggregates population by:<br>
	- Regions<br>
	- Local Authorities<br>
	- National level (England)

6. Local Authorities and Regions Data<br>
- Extracts local authority and region labels from metadata.<br>
- Merges these with population estimates.<br>
- Saves results to Local Authorities.csv and Regions.csv.

7. England-Level Summary<br>
- Extracts and saves overall England population summary to England.csv.

________________________________________

8. Sports and Activity Codes Definition<br>
- Defines a dictionary mapping sports names to SPSS variable codes.<br>
- Generates activity-based column names dynamically.

9. Load and Filter Data<br>
- Loads the original SPSS file again.<br>
- Filters required columns: demographics + sports participation variables.

10. Data Cleaning and Aggregation<br>
- Initializes zero-value columns.<br>
- Sums participation metrics over 12 months, frequency, minutes played, etc.

11. Calculations on Participation<br>
- Creates metrics like:<br>
	- Participated vs Not Participated<br>
	- Regular participation<br>
	- Activity levels (Inactive, Fairly Active, Active)<br>
- Assigns weighted values (wt_final) depending on participation and demographics.

12. Demographic Breakdown<br>
- Breaks down participation by:<br>
	- Gender<br>
	- Age groups<br>
	- Ethnicities

13. Label Mapping<br>
- Replaces numeric codes with descriptive labels for regions, local authorities, age, gender, ethnicity.

14. Population Data Integration<br>
- Merges participation data with population estimates (England, regions, authorities).
________________________________________

15. Output Generation for Sports<br>
- Saves processed sports participation data to SPSS and CSV files per sport.

16. Aggregate Data by Geography<br>
- Aggregates participation metrics for:<br>
	- England (national)<br>
	- Regions<br>
	- Local Authorities

17. Sorting and Cleaning Aggregated Data<br>
- Sorts by area name.<br>
- Cleans up column names and removes unnecessary columns.

18. Calculate Participation Percentages and Estimates<br>
- Calculates percentage participation metrics.<br>
- Calculates population estimates for participation categories by multiplying percentages with population counts.<br>
- Breaks down participation by demographics for detailed analysis.<br>
- Calculates latent demand (potential for increased participation).<br>

19. Final Output<br>
- Rounds and sorts columns.<br>
- Saves the final aggregated and processed data to CSV files named by sport or group.

