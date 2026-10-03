SELECT 
    c.company_name,
    c.ticker,
    v.current_price,
    v.dcf_price,
    ROUND(((v.dcf_price - v.current_price) / v.current_price) * 100, 2) AS dcf_upside_pct,
    CASE 
        WHEN ((v.dcf_price - v.current_price) / v.current_price) > 0.10 THEN 'BUY'
        WHEN ((v.dcf_price - v.current_price) / v.current_price) < -0.10 THEN 'OVERVALUED / SELL'
        ELSE 'HOLD'
    END AS recommendation
FROM dbo.companies c
INNER JOIN dbo.valuation_outputs v ON c.company_id = v.company_id
ORDER BY dcf_upside_pct DESC;

SELECT 
    c.company_name,
    v.enterprise_val_cr,
    v.pv_fcff_cr,
    v.pv_tv_cr,
    ROUND((v.pv_fcff_cr / NULLIF(v.enterprise_val_cr, 0)) * 100, 2) AS explicit_fcff_contrib_pct,
    ROUND((v.pv_tv_cr / NULLIF(v.enterprise_val_cr, 0)) * 100, 2) AS terminal_value_contrib_pct
FROM dbo.companies c
INNER JOIN dbo.valuation_outputs v ON c.company_id = v.company_id
ORDER BY terminal_value_contrib_pct DESC;

SELECT 
    c.company_name,
    v.cash_cr,
    v.debt_cr,
    ROUND(v.cash_cr - v.debt_cr, 2) AS net_cash_position_cr,
    CASE 
        WHEN (v.cash_cr - v.debt_cr) > 0 THEN 'Net Cash (Strong Liquidity)'
        ELSE 'Net Debt (Leveraged)'
    END AS liquidity_status
FROM dbo.companies c
INNER JOIN dbo.valuation_outputs v ON c.company_id = v.company_id;

-- optional
SELECT 
    c.company_name,
    v.current_price,
    v.dcf_price,
    v.wacc_pct,
    ROUND(AVG(v.dcf_price) OVER(), 2) AS sector_avg_dcf_price,
    ROUND(v.dcf_price - AVG(v.dcf_price) OVER(), 2) AS variance_from_sector_avg
FROM dbo.companies c
INNER JOIN dbo.valuation_outputs v ON c.company_id = v.company_id;

SELECT 
    c.company_name,
    v.pe_price,
    DENSE_RANK() OVER (ORDER BY v.pe_price DESC) AS pe_valuation_rank,
    v.ev_ebitda_price,
    DENSE_RANK() OVER (ORDER BY v.ev_ebitda_price DESC) AS ev_ebitda_valuation_rank
FROM dbo.companies c
INNER JOIN dbo.valuation_outputs v ON c.company_id = v.company_id;

GO
CREATE VIEW dbo.vw_fmcg_executive_summary AS
SELECT 
    c.company_name,
    c.ticker,
    v.current_price,
    v.dcf_price,
    v.wacc_pct,
    v.enterprise_val_cr,
    v.equity_val_cr,
    (v.cash_cr - v.debt_cr) AS net_cash_cr,
    ROUND(((v.dcf_price - v.current_price) / NULLIF(v.current_price, 0)) * 100, 2) AS dcf_upside_pct
FROM dbo.companies c
INNER JOIN dbo.valuation_outputs v ON c.company_id = v.company_id;
GO

-- Querying the View directly:
SELECT * FROM dbo.vw_fmcg_executive_summary WHERE dcf_upside_pct > 0;