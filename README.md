Cryptocurrency Analysis & Internship Analytics Project

Overview

This project is an extension of an existing Cryptocurrency Analysis training project. The internship tasks were implemented using the same project workflow, cryptocurrency dataset, Excel workbook, and source-code approach rather than creating an unrelated project.

The solution combines Python-based data collection and preparation with Microsoft Excel dashboards, formulas, PivotTables, PivotCharts, slicers, KPI cards, and VBA automation.

Objectives

Collect and prepare cryptocurrency market data.

Analyze price, market capitalization, trading volume, circulating supply, and percentage changes.

Build interactive dashboards and reports in Excel.

Add slicer-based filtering and KPI-driven insights.

Compare current and historical/estimated prices.

Provide investment-oriented decision support.

Automate selected dashboard behavior with VBA.

Technologies Used

Python

Python

Requests

BeautifulSoup

Pandas

OpenPyXL

Python is used for cryptocurrency data collection, preparation, and export.

Microsoft Excel

Excel Tables

Excel formulas

XLOOKUP

PivotTables

PivotCharts

Slicers

KPI cards

Dashboard formatting

VBA

VBA is used for the time-based visibility requirement in Task 1 and for automatic dashboard checking/refresh behavior.

Existing Training Project

The internship solution builds on the original cryptocurrency analytics project.

The training project already included cryptocurrency data collection, cleaning, market analysis, Excel dashboards, PivotTables/PivotCharts, KPI cards, and slicers.

The six internship tasks were implemented as additional analytics features and dashboards within the same overall project.

Internship Tasks

Task 1 — Time-Based Visibility & Sentiment Analysis

Objective

Display the Top 10 cryptocurrencies by 24-hour trading volume from eligible coins and control chart visibility according to working hours.

Eligibility Rule

Only cryptocurrencies whose names start with:

A

E

I

O

U

B

C

D

are considered.

The eligible coins are ranked by 24-hour volume, and the Top 10 are displayed.

Time-Based Visibility

9:00 AM–5:00 PM: Top 10 chart is visible.

Outside working hours: The chart is hidden and the message is displayed:

Please open in working hours (9 am to 5 pm).

VBA Automation

The workbook uses VBA to:

Check the current time.

Show or hide the Task 1 chart.

Display the working-hours message.

Run the visibility check when the workbook opens.

Repeat the check periodically.

Task 2 — Crypto Coin Comparison Dashboard

Objective

Compare two cryptocurrencies using user-entered coin names.

Information Displayed

For each selected coin:

Symbol

Price

Volume (24h)

Market Cap

Circulating Supply

KPI Comparisons

The dashboard calculates differences including:

Volume difference

Market capitalization difference

Circulating supply difference

Validation

The coin-name inputs are validated to support names between 3 and 10 characters and prevent numeric input.

Task 3 — Top 5 Crypto Liquidity Analysis

Objective

Identify the most liquid cryptocurrencies using 24-hour trading volume as the liquidity indicator.

Price Categories

$0–$50

Above $50

Analysis

For the selected price category:

Rank cryptocurrencies by 24-hour volume.

Identify the Top 5.

Group all remaining cryptocurrencies into Others.

Visualize the liquidity distribution using a pie chart.

Interactivity

A price-category slicer dynamically changes the analysis.

Task 4 — Price Change Comparison Analysis

Objective

Compare the current price of cryptocurrencies with their estimated previous 1-hour price, identify the Top 10 coins by 1-hour price change, and allow filtering by price range.

Previous 1-Hour Price

Because the 1-hour percentage is treated as a positive increase:

Previous 1h Price = Current Price / (1 + 1h Change / 100)

Price Ranges

Up to $10

$10 and above

Analysis

For the selected price range:

Rank cryptocurrencies by 1-hour percentage change.

Display the Top 10.

Compare current price and previous 1-hour price.

Display a linked table containing:

Coin Name

Symbol

1h Price Change (%)

Visualization

A clustered column chart compares:

Current Price

Previous 1h Price

The Top 10 and chart change according to the selected price range.

Task 5 — Historical Price Change Insights

Objective

Analyze cryptocurrency price changes using:

1-hour change

24-hour change

7-day change

The analysis focuses on cryptocurrencies priced between $0 and $5.

Assumptions

1-hour change represents an increase.

7-day change represents an increase.

24-hour change represents a decrease.

Historical Price Calculations

Estimated 7-day price:

7-Day Price = Current Price / (1 + 7d Change / 100)

Estimated 24-hour price:

24-Hour Price = Current Price / (1 - 24h Change / 100)

Analysis

Filter coins priced between $0 and $5.

Rank them by 1-hour percentage change.

Identify the Top 10.

Display:

Coin Name

Symbol

Current Price

7-Day Price

24-Hour Price

Compare the estimated historical prices.

Visualization

A clustered column chart compares 7-Day Price and 24-Hour Price for the Top 10 selected cryptocurrencies.

Task 6 — Low Budget Investment Insights

Objective

Create a KPI dashboard that identifies the cryptocurrency with the lowest average downfall percentage within a selected price range.

Assumption

The 1-hour, 24-hour, and 7-day percentage changes are treated as price declines.

Average Downfall

The average downfall is calculated using the absolute values of the three percentage changes:

Average Downfall =
AVERAGE(
    ABS(1h Change),
    ABS(24h Change),
    ABS(7d Change)
)

Price Ranges

$0–$0.05

$0.05–$0.5

$0.5–$5

$5–$50

>$50

KPI Dashboard

The dashboard displays five KPIs:

Coin Name

Symbol

Current Price

Average Downfall (%)

Total Coins Considered

Interactivity

The price-range slicer dynamically changes the five KPI results.

The selected range identifies the coin with the lowest average downfall, along with its symbol, current price, average downfall percentage, and the total number of coins considered.

Data Workflow

Cryptocurrency Market Data
          ↓
Python Data Collection
          ↓
Data Preparation
          ↓
CSV / Excel Dataset
          ↓
Excel Analysis
          ↓
Formulas + PivotTables + Slicers
          ↓
Charts + KPI Dashboards
          ↓
Investment & Trend Insights

Key Excel Features

XLOOKUP

XLOOKUP is used to dynamically retrieve related cryptocurrency information such as symbols and price changes.

Example:

=XLOOKUP(CoinName,LookupRange,ReturnRange,"Not Found")

PivotTables

PivotTables are used to summarize and filter analytical results where appropriate, especially for Top 10, liquidity, price-change, and supporting KPI analysis.

Slicers

Slicers provide interactive filtering by price categories and task-specific analysis ranges.

KPI Cards

KPI cards present the main decision-making indicators in a compact dashboard format.

VBA Automation

The workbook is saved as .xlsm so that VBA functionality is preserved.

The Task 1 VBA logic follows:

9:00 AM–5:00 PM
       ↓
Chart visible

Outside working hours
       ↓
Chart hidden
       ↓
"Please open in working hours (9 am to 5 pm)."

The workbook also checks the visibility condition when the workbook opens and continues periodic checking.

Workbook Contents

The final workbook contains the original training-project dashboard plus the six internship task implementations:

Task 1 — Time-Based Visibility & Sentiment Analysis

Task 2 — Crypto Coin Comparison

Task 3 — Top 5 Crypto Liquidity Analysis

Task 4 — Price Change Comparison

Task 5 — Historical Price Change Insights

Task 6 — Low Budget Investment Insights

The workbook also contains the required source data, calculations, dashboards, slicers, charts, and supporting analysis.

How to Use

Open the final .xlsm workbook in Microsoft Excel.

Enable macros when prompted if you want Task 1 automation.

Navigate through the task sheets.

Use the input fields and slicers.

Review the KPI cards, tables, and charts.

Change the available price ranges to explore the analysis dynamically.

Submission Package

The complete project should include:

Final Excel .xlsm workbook

Cryptocurrency CSV dataset

Python/source code

Task reports

Dashboard screenshots

README.md

The complete project is intended to be packaged into one ZIP file for submission.

The ZIP should then be uploaded to Google Drive and shared using a publicly accessible link.

The project/dashboard should also be published through a public GitHub repository as required by the internship instructions.
