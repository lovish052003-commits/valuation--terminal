"""
Universal Institutional Equity Valuation Engine
===============================================
A modular, universal equity valuation terminal that automatically identifies,
normalizes, forecasts, and values ANY publicly listed company across sectors
without company-specific hardcoding.
"""

from .config import VALUATION_CONFIG
from .company_master import (
    CompanyMaster,
    CanonicalMetric,
    ShareCountEngine,
    MarketCapEngine,
    ReportingBasisController
)
from .company_classifier import CompanyClassificationEngine
from .method_selector import ValuationMethodSelector
from .financial_normalizer import FinancialNormalizationEngine
from .forecasting_engine import ForecastEngine
from .roic_engine import SustainableROICEngine
from .cost_of_capital import CostOfCapitalEngine
from .peer_selection_engine import PeerSelectionEngine, PeerSelectionError
from .sotp_engine import SOTPEngine
from .dcf_engine import DCFEngine
from .reconciliation_engine import ValuationReconciliationEngine, DataSheetReconciliationEngine
from .confidence_engine import ValuationConfidenceEngine
from .quality_engine import ModelQualityEngine

from .canonical_financials import (
    CanonicalFinancialStatementLayer,
    CanonicalPeriodStatements,
    FinancialField,
    clean_fiscal_year_label
)
from .company_data import CompanyData, CompanyIdentity, MarketData, ValidationIssue
from .workbook_map import WorkbookMap, DEFAULT_DATA_SHEET_ROWS
from .canonical_data import (
    CanonicalMarketData,
    CanonicalFinancialData,
    CanonicalValuationInputs,
    CanonicalValuationOutput,
    build_canonical_data_objects
)

__all__ = [
    'VALUATION_CONFIG',
    'CompanyMaster',
    'CanonicalMetric',
    'ShareCountEngine',
    'MarketCapEngine',
    'ReportingBasisController',
    'CompanyClassificationEngine',
    'ValuationMethodSelector',
    'FinancialNormalizationEngine',
    'ForecastEngine',
    'SustainableROICEngine',
    'CostOfCapitalEngine',
    'PeerSelectionEngine',
    'PeerSelectionError',
    'SOTPEngine',
    'DCFEngine',
    'ValuationReconciliationEngine',
    'DataSheetReconciliationEngine',
    'ValuationConfidenceEngine',
    'ModelQualityEngine',
    'CanonicalFinancialStatementLayer',
    'CanonicalPeriodStatements',
    'FinancialField',
    'clean_fiscal_year_label',
    'UniversalFieldResolver',
    'CompanyData',
    'CompanyIdentity',
    'MarketData',
    'ValidationIssue',
    'WorkbookMap',
    'DEFAULT_DATA_SHEET_ROWS',
    'CanonicalMarketData',
    'CanonicalFinancialData',
    'CanonicalValuationInputs',
    'CanonicalValuationOutput',
    'build_canonical_data_objects'
]


