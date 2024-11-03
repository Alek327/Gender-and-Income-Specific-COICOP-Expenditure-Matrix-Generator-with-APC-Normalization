# COICOP Expenditure Matrix Generator with APC Adjustment
This R code generates gender- and income-specific COICOP expenditure matrices with controlled APC variability, ensuring cumulative sums match original regional totals, and exports the results to Excel.


This R script generates gender- and income-specific COICOP expenditure matrices with controlled Average Propensity to Consume (APC) adjustments, ensuring cumulative sums match original regional totals. The script processes data on income quintiles, COICOP expenditures by region, APC averages, and gender distribution to create detailed expenditure matrices suitable for demographic analysis.

## Features

- **Controlled APC Variation**: APC values are varied around each region’s average APC, scaled by 10 to ensure values fall within the specified range of 0.7 to 0.9.
- **Gender and Income Distribution**: Applies income quintile shares and gender proportions to COICOP expenditure data, providing matrices segmented by both demographic dimensions.
- **Cumulative Normalization**: Ensures that the sum across all generated COICOP matrices (across quintiles and genders) matches the original column sums, preserving consistency with the input data.
- **Excel Export**: Outputs all generated matrices to an Excel file, with each sheet representing a specific combination of income quintile and gender, and each column sum aligned with original regional totals.

## Getting Started

### Prerequisites

Ensure you have the following R packages installed:
```R
install.packages(c("readxl", "dplyr", "openxlsx", "tibble"))

## Input Files

The script requires four input Excel files:

Household income quintile.xlsx: Income quintile data.
adjusted_coicop_expenditures_by_region.xlsx: Original COICOP expenditure data, used to determine regional spending patterns.
APC.xlsx: Average APC values by region.
carcom10.xlsx: Gender distribution data indicating primary income earners.

## Running the Script

Open the R script and run it interactively.
When prompted, select each of the input files in sequence.
The script will process the data and export the output to APC_and_COICOP_Quintile_Vectors_Refined_Gender_Distribution_with_Cumulative_Normalization.xlsx.

## Output

The generated Excel file contains:

Sheets for each combination of income quintile and gender (e.g., COICOP_first_Male, COICOP_first_Female).
Columns normalized to match the original COICOP expenditure totals, with APC adjustments that vary across COICOP categories and regions.

### LICENSE 

In the root of your GitHub repository, add a `LICENSE` file with the GNU General Public License content. You can find the full text of the GPLv3 at [https://www.gnu.org/licenses/gpl-3.0.txt](https://www.gnu.org/licenses/gpl-3.0.txt).

## Author

Developed by Oleksandr Galychyn


This `README.md` provides an overview of the functionality, setup instructions, file requirements, and usage details for GitHub. Let me know if you'd like further adjustments!

