# Official Source Files

This directory contains downloaded official source files from Oregon state agencies for budget analysis and verification.

## Directory Structure

- **marijuana-tax-receipts/**: Monthly marijuana tax receipt files (PDF/XLSX) from Oregon Department of Revenue
- **revenue-distributions/**: Quarterly revenue distribution reports (PDF/XLSX) from Oregon Department of Revenue
- **statistical-reports/**: Annual statistical reports on marijuana tax characteristics
- **other-sources/**: Other official budget documents and reports

## Source Files

### Oregon Department of Revenue - Marijuana Tax Statistics

**Source URL**: https://www.oregon.gov/dor/gov-research/pages/default.aspx

**Downloaded Files:**

- `marijuana-tax-receipts-most-recent.pdf` (273KB) - Monthly tax receipts (PDF)
- `marijuana-tax-receipts-most-recent.xlsx` (48KB) - Monthly tax receipts (Excel)
- `revenue-distributions-most-recent.pdf` (177KB) - Quarterly distributions (PDF)
- `revenue-distributions-most-recent.xlsx` (26KB) - Quarterly distributions (Excel)

**Download Date**: October 31, 2025

**Analysis Files:**

- `marijuana-tax-analysis.md` - Comprehensive analysis of revenue data
- `analysis_results_detailed.json` - Raw analysis data extracted from Excel files
- `analyze_marijuana_tax_v2.py` - Python script for analyzing Excel files

**Contact**: dor.research@dor.oregon.gov | 503-945-8383

## File Naming Convention

- Monthly receipts: `marijuana-tax-receipts-YYYY-MM.pdf` or `.xlsx`
- Quarterly distributions: `revenue-distribution-YYYY-QX.pdf` or `.xlsx`
- Statistical reports: `marijuana-statistical-report-YYYY.pdf`

## Analysis Results

### Key Findings (Based on Official Data)

**2024 Calendar Year Revenue:**

- Total State Tax Distributed: $312.6 million
- Q1: $79.3M, Q2: $77.3M, Q3: $78.8M, Q4: $77.2M

**2025 Calendar Year (First 3 Quarters):**

- Q1: $73.0M, Q2: $70.8M, Q3: $73.4M
- Annualized: ~$289.6 million

**Critical Finding**: Actual revenue is ~$290-310M annually, significantly higher than previous estimates of $169.5M annually.

**Distribution Formula (Verified):**

- State School Fund: 40%
- Mental Health: 20%
- State Police: 15%
- Health Authority: 5%
- Cities/Counties: 20%
- Drug Treatment Fund: Variable (Measure 110)

See `marijuana-tax-analysis.md` for complete analysis.

## Analysis Notes

Files in this directory are used to:

1. ✅ Verify revenue figures in budget analysis documents - **COMPLETE**
2. ✅ Calculate year-over-year trends - **COMPLETE**
3. ⏳ Analyze monthly patterns and seasonality - **IN PROGRESS**
4. ✅ Cross-reference distribution formula percentages - **COMPLETE**
5. ✅ Ensure ballot measure proposals use accurate revenue data - **COMPLETE**

## Update Schedule

- **Monthly**: Download latest tax receipt files
- **Quarterly**: Download latest revenue distribution reports
- **Annually**: Download latest statistical report (when available)

## Next Steps

- Extract detailed monthly data for seasonal analysis
- Update all budget documents with verified figures
- Monitor for quarterly updates from DOR

---

**Last Updated**: October 31, 2025  
**Data Source**: Oregon Department of Revenue - Marijuana Tax Statistics
