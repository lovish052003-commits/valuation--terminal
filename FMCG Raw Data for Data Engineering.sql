-- Step 1: Create Database in T-SQL
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'fmcg_valuation_db')
BEGIN
    CREATE DATABASE fmcg_valuation_db;
END
GO

USE fmcg_valuation_db;
GO

-- Step 2: Create Companies Master Table
IF OBJECT_ID('dbo.companies', 'U') IS NOT NULL DROP TABLE dbo.companies;
CREATE TABLE dbo.companies (
    company_id INT PRIMARY KEY,
    company_name VARCHAR(100) NOT NULL,
    ticker VARCHAR(20) NOT NULL,
    sector VARCHAR(50) DEFAULT 'FMCG'
);

-- Step 3: Create Valuation Outputs Table
IF OBJECT_ID('dbo.valuation_outputs', 'U') IS NOT NULL DROP TABLE dbo.valuation_outputs;
CREATE TABLE dbo.valuation_outputs (
    valuation_id INT IDENTITY(1,1) PRIMARY KEY,
    company_id INT FOREIGN KEY REFERENCES dbo.companies(company_id),
    current_price DECIMAL(10,2),
    dcf_price DECIMAL(10,2),
    ev_ebitda_price DECIMAL(10,2),
    pe_price DECIMAL(10,2),
    wacc_pct DECIMAL(5,2),
    pv_fcff_cr DECIMAL(12,2),
    pv_tv_cr DECIMAL(12,2),
    enterprise_val_cr DECIMAL(12,2),
    cash_cr DECIMAL(12,2),
    debt_cr DECIMAL(12,2),
    equity_val_cr DECIMAL(12,2),
    shares_cr DECIMAL(10,4)
);
GO

-- Step 4: Populate Data
INSERT INTO dbo.companies (company_id, company_name, ticker) 
VALUES
    (1, 'Hindustan Unilever', 'HINDUNILVR'),
    (2, 'ITC Ltd.', 'ITC'),
    (3, 'Nestle India', 'NESTLEIND'),
    (4, 'Varun Beverages', 'VBL'),
    (5, 'Britannia Industries', 'BRITANNIA');

INSERT INTO dbo.valuation_outputs 
(company_id, current_price, dcf_price, ev_ebitda_price, pe_price, wacc_pct, pv_fcff_cr, pv_tv_cr, enterprise_val_cr, cash_cr, debt_cr, equity_val_cr, shares_cr) 
VALUES
    (1, 2086.60, 694.03, 2225.87, 3264.23, 11.64, 55644.40, 105654.95, 161299.35, 3248.00, 1478.00, 163069.35, 234.9591),
    (2, 276.50, 176.76, 782.08, 852.36, 11.52, 70746.76, 150113.23, 220859.99, 3008.79, 2399.06, 221469.72, 1252.9468),
    (3, 1498.80, 1805.14, 915.07, 932.34, 11.32, 84650.94, 262540.70, 347191.64, 1340.87, 444.16, 348088.35, 192.8314),
    (4, 436.25, 92.12, 588.41, 504.18, 11.33, 873.49, 30789.79, 31663.28, 1998.49, 2508.18, 31153.59, 338.1989),
    (5, 5618.00, 1007.33, 5009.91, 5318.00, 11.33, 2282.84, 23008.40, 25291.25, 352.44, 1380.28, 24263.41, 24.0868);
GO