# Stock Return CAPM Analysis

This repository analyzes **stock and industry returns** using the **Capital Asset Pricing Model (CAPM)**.  
It estimates each sector’s **systematic risk (β)** and **abnormal return (α)**, visualizing the risk–return tradeoff.

---
## Project Overview

### Objective
Use CAPM to measure the relationship between each industry’s stock returns and market returns.  
We estimate:
- **β (Beta)** – systematic risk relative to the market  
- **α (Alpha)** – abnormal or excess return unexplained by the market



### Model
$R_i = \alpha + \beta R_m$

Where:  
- $\ R_i \$: industry return, from Taiwan OTC (櫃買中心)
- $\ R_m \$: market return, from Taiwan OTC (櫃買中心)


## Methodology
1. **Data Retrieval**  
   - Download market and industry stock return data from the OTC (櫃買中心) website.  

2. **Data Conversion**  
   - Convert downloaded Excel files to CSV format using custom Python functions.

3. **Data Cleaning**  
   - Select relevant rows (3rd–24th) containing market and industry return data.  
   - Save as `data_ok.csv`.

4. **Regression Analysis**  
   - For each industry:  
     - Use market return as independent variable $\ R_m \$.  
     - Use industry return as dependent variable $\ R_i \$.  
     - Run OLS regression to estimate α and β.  

5. **Visualization**  
   - Plot α (Y-axis) vs β (X-axis) scatter chart with industry labels.  

## Interpretation
- **β > 0** → Industry moves with the market
- **β > 1** → higher volatility than the market
- **α > 0** → Industry earns excess returns beyond market influence
- Biotech and shipping sectors showed the strongest positive α in 2021

Example Findings:
- Biotech sector (ex. 高端疫苗 6547) showed the highest α in May 2021.  
- Shipping sector (ex. 萬海, 長榮, 陽明) had high β and α during market upswings.


---


## Requirements
- Python 3.9+
- Required libraries:
  - `pandas`
  - `numpy`
  - `matplotlib`
  - `scikit-learn`
  - `openpyxl`



