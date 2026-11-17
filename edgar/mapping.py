"""
MAPPING: standardized line item -> ordered list of XBRL tags: us-gaap first, then IFRS (ifrs-full) for
foreign filers. A company's own extra tags go in overrides/<TICKER>.json, not here (see company_maps).
The first tag with data for a given period wins (coalesce per period).
Edit this to fix gaps for a specific company; it is also written to the
Mapping tab so every number is traceable.

(label, kind, unit, tags)   kind: "flow" (duration) or "instant" (point in time)
A tag entry may be a tuple of tags, meaning "sum these"; it is used only for periods
where every component has a value (so a missing piece never silently understates the total).

Rows that are computed rather than pulled (plugs, subtotal fallbacks) live in layout.py.
"""
IS_MAP = [
    ("Revenue", "flow", "USD", ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
                                "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet",
                                "SalesRevenueGoodsNet", "Revenue", "RevenueFromContractsWithCustomers"]),
    ("Cost of Revenue", "flow", "USD", ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold", "CostOfSales",
                                            "CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization"]),
    ("Gross Profit", "flow", "USD", ["GrossProfit"]),
    ("R&D Expense", "flow", "USD", ["ResearchAndDevelopmentExpense"]),
    ("SG&A Expense", "flow", "USD", ["SellingGeneralAndAdministrativeExpense", ("SellingAndMarketingExpense", "GeneralAndAdministrativeExpense"),
                                         ("MarketingExpense", "GeneralAndAdministrativeExpense"),
                                         ("SalesAndMarketingExpense", "GeneralAndAdministrativeExpense"),
                                         ("SellingExpense", "GeneralAndAdministrativeExpense"),
                                         ("DistributionCosts", "AdministrativeExpense")]),
    ("Operating Income", "flow", "USD", ["OperatingIncomeLoss", "ProfitLossFromOperatingActivities"]),
    ("Interest Expense", "flow", "USD", ["InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt", "FinanceCosts"]),
    ("Interest Income", "flow", "USD", ["InvestmentIncomeInterest", "InvestmentIncomeInterestAndDividend", "FinanceIncome"]),
    ("Pre-Tax Income", "flow", "USD", [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "ProfitLossBeforeTax"]),
    ("Income Tax Expense", "flow", "USD", ["IncomeTaxExpenseBenefit", "IncomeTaxExpenseContinuingOperations"]),
    ("Net Income to NCI", "flow", "USD", ["NetIncomeLossAttributableToNoncontrollingInterest",
                                              "IncomeLossFromContinuingOperationsAttributableToNoncontrollingEntity",
                                              "ProfitLossAttributableToNoncontrollingInterests"]),
    ("Net Income", "flow", "USD", ["NetIncomeLoss", "ProfitLossAttributableToOwnersOfParent", "ProfitLoss"]),
    ("Diluted Shares", "flow", "shares", ["WeightedAverageNumberOfDilutedSharesOutstanding", "AdjustedWeightedAverageShares"]),
    ("Diluted EPS", "flow", "USD/shares", ["EarningsPerShareDiluted", "DilutedEarningsLossPerShare"]),
]
BS_MAP = [
    ("Cash & Equivalents", "instant", "USD", ["CashAndCashEquivalentsAtCarryingValue",
                                              "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
                                              "CashAndCashEquivalents"]),
    ("Short-Term Investments", "instant", "USD", ["MarketableSecuritiesCurrent", "ShortTermInvestments",
                                                  "AvailableForSaleSecuritiesDebtSecuritiesCurrent", "CurrentInvestments",
                                                  "OtherCurrentFinancialAssets"]),
    ("Accounts Receivable", "instant", "USD", ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent",
                                               "TradeAndOtherCurrentReceivables", "CurrentTradeReceivables"]),
    ("Inventory", "instant", "USD", ["InventoryNet", "Inventories"]),
    ("Total Current Assets", "instant", "USD", ["AssetsCurrent", "CurrentAssets"]),
    ("PP&E, Net", "instant", "USD", ["PropertyPlantAndEquipmentNet",
        "PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
        "PropertyPlantAndEquipment"]),
    ("Goodwill", "instant", "USD", ["Goodwill"]),
    # FiniteLived... excludes indefinite-lived intangibles; used only when the total isn't tagged.
    ("Intangibles, Net", "instant", "USD", ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet",
                                            "IntangibleAssetsOtherThanGoodwill", "IntangibleAssetsAndGoodwill"]),
    ("Long-Term Investments", "instant", "USD", ["MarketableSecuritiesNoncurrent", "LongTermInvestments",
                                                 "AvailableForSaleSecuritiesDebtSecuritiesNoncurrent", "NoncurrentInvestments",
                                                 "OtherNoncurrentFinancialAssets"]),
    ("Total Assets", "instant", "USD", ["Assets"]),
    ("Accounts Payable", "instant", "USD", ["AccountsPayableCurrent", "AccountsPayableTradeCurrent",
                                            "AccountsPayableAndAccruedLiabilitiesCurrent",
                                            "TradeAndOtherCurrentPayablesToTradeSuppliers", "TradeAndOtherCurrentPayables"]),
    ("Accrued Liabilities", "instant", "USD", ["AccruedLiabilitiesCurrent"]),
    ("Deferred Revenue (Current)", "instant", "USD", ["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent",
                                                      "CurrentContractLiabilities"]),
    # Debt is split into three non-overlapping pieces so Net debt can sum them without double counting.
    # (Avoid DebtCurrent / LongTermDebt here: both usually include amounts reported on another line.)
    ("Short-Term Borrowings", "instant", "USD", ["CommercialPaper", "ShortTermBorrowings", "NotesAndLoansPayable",
                                                 "ShorttermBorrowings"]),
    ("Current Portion of LT Debt", "instant", "USD", ["LongTermDebtCurrent", "LongTermDebtAndCapitalLeaseObligationsCurrent",
                                                      ("CurrentBondsIssuedAndCurrentPortionOfNoncurrentBondsIssued",
                                                       "CurrentPortionOfLongtermBorrowings"),
                                                      "CurrentPortionOfLongtermBorrowings",
                                                      "CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings"]),
    ("Total Current Liabilities", "instant", "USD", ["LiabilitiesCurrent", "CurrentLiabilities"]),
    ("Long-Term Debt", "instant", "USD", ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations",
                                          ("NoncurrentPortionOfNoncurrentBondsIssued", "LongtermBorrowings"),
                                          "LongtermBorrowings", "NoncurrentPortionOfNoncurrentBorrowings"]),
    ("Operating Lease Liabilities", "instant", "USD", ["OperatingLeaseLiabilityNoncurrent", "NoncurrentLeaseLiabilities"]),
    ("Total Liabilities", "instant", "USD", ["Liabilities"]),
    ("Shareholders' Equity (Parent)", "instant", "USD", ["StockholdersEquity", "EquityAttributableToOwnersOfParent"]),
    ("Noncontrolling Interest", "instant", "USD", ["MinorityInterest", "NonredeemableNoncontrollingInterest", "NoncontrollingInterests"]),
    # Falls back to Parent + NCI (layout.py) when the consolidated total isn't tagged.
    ("Total Equity", "instant", "USD", ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "Equity"]),
    ("Total Liabilities & Equity", "instant", "USD", ["LiabilitiesAndStockholdersEquity", "EquityAndLiabilities"]),
]
CF_MAP = [
    # The cash flow statement starts from consolidated net income (incl. NCI), so ProfitLoss comes first.
    ("Net Income (CF)", "flow", "USD", ["ProfitLoss", "NetIncomeLoss"]),
    ("D&A", "flow", "USD", ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
                            "DepreciationAmortizationAndAccretionNet",
                            ("Depreciation", "AmortizationOfIntangibleAssets"),
                            "DepreciationAndAmortisationExpense", "AdjustmentsForDepreciationAndAmortisationExpense",
                            ("DepreciationExpense", "AmortisationExpense"),
                            ("AdjustmentsForDepreciationExpense", "AdjustmentsForAmortisationExpense"),
                            "Depreciation"]),
    ("Amortization of Intangibles", "flow", "USD", ["AmortizationOfIntangibleAssets", "AmortisationExpense",
                                                    "AdjustmentsForAmortisationExpense"]),
    ("Stock-Based Comp", "flow", "USD", ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense",
                                         "AdjustmentsForSharebasedPayments"]),
    ("Cash from Operations", "flow", "USD", ["NetCashProvidedByUsedInOperatingActivities",
                                             "CashFlowsFromUsedInOperatingActivities"]),
    ("Capex", "flow", "USD", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets",
                              "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"]),
    ("Acquisitions", "flow", "USD", ["PaymentsToAcquireBusinessesNetOfCashAcquired",
        "CashFlowsUsedInObtainingControlOfSubsidiariesOrOtherBusinessesClassifiedAsInvestingActivities"]),
    ("Cash from Investing", "flow", "USD", ["NetCashProvidedByUsedInInvestingActivities",
                                            "CashFlowsFromUsedInInvestingActivities"]),
    ("Debt Issued", "flow", "USD", ["ProceedsFromIssuanceOfLongTermDebt", "ProceedsFromIssuanceOfDebt",
                                    "ProceedsFromBorrowingsClassifiedAsFinancingActivities",
                                    ("ProceedsFromIssueOfBondsNotesAndDebentures", "ProceedsFromNoncurrentBorrowings")]),
    ("Debt Repaid", "flow", "USD", ["RepaymentsOfLongTermDebt", "RepaymentsOfDebt",
                                    "RepaymentsOfBorrowingsClassifiedAsFinancingActivities",
                                    ("RepaymentsOfBondsNotesAndDebentures", "RepaymentsOfNoncurrentBorrowings")]),
    ("Dividends Paid", "flow", "USD", ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock", "PaymentsOfOrdinaryDividends",
                                       "DividendsPaidClassifiedAsFinancingActivities",
                                       "DividendsPaidToEquityHoldersOfParentClassifiedAsFinancingActivities"]),
    ("Share Repurchases", "flow", "USD", ["PaymentsForRepurchaseOfCommonStock", "PaymentsToAcquireOrRedeemEntitysShares",
                                          "PurchaseOfTreasuryShares"]),
    ("Cash from Financing", "flow", "USD", ["NetCashProvidedByUsedInFinancingActivities",
                                            "CashFlowsFromUsedInFinancingActivities"]),
]
# Extra lines used only by the financials template (banks, insurers, asset managers): see xl_fin.py.
FIN_MAP = [
    ("Net Interest Income", "flow", "USD", ["InterestIncomeExpenseNet", "NetInterestIncome"]),
    ("Provision for Credit Losses", "flow", "USD", ["ProvisionForLoanLeaseAndOtherLosses", "ProvisionForCreditLosses",
                                                    "ProvisionForLoanAndLeaseLosses", "CreditLossExpenseReversal"]),
    ("Deposits", "instant", "USD", ["Deposits", "DepositsFromCustomers"]),
    ("Loans, Net", "instant", "USD", ["LoansAndLeasesReceivableNetReportedAmount", "NotesReceivableNet",
                                      "FinancingReceivableExcludingAccruedInterestAfterAllowanceForCreditLoss",
                                      "LoansAndAdvancesToCustomers"]),
]
ALL_MAPS = IS_MAP + BS_MAP + CF_MAP + FIN_MAP
MAP_BY_LABEL = {m[0]: m for m in ALL_MAPS}

ANNUAL_FORMS = {"10-K", "10-K/A", "10-KT", "20-F", "20-F/A", "40-F", "40-F/A"}   # 20-F / 40-F: foreign filers
FACT_NAMESPACES = ("us-gaap", "ifrs-full")
QUARTERLY_FORMS = {"10-Q", "10-Q/A"}


def fmt_tag(entry):
    return f"({' + '.join(entry)})" if isinstance(entry, tuple) else entry


def company_maps(overrides=None):
    """(IS, BS, CF, financials-only) maps for one company. `overrides` is the content of overrides/<TICKER>.json:
    {"Interest Income": ["SomeTag", ["TagA", "TagB"]]} - tags tried before the standard ones for that line;
    an inner list means "sum these". Unknown line names are reported by load_overrides, not here."""
    overrides = overrides or {}

    def merged(mapping):
        out = []
        for label, kind, unit, tags in mapping:
            extra = [tuple(t) if isinstance(t, list) else t for t in overrides.get(label, [])]
            out.append((label, kind, unit, extra + [t for t in tags if t not in extra]))
        return out
    return merged(IS_MAP), merged(BS_MAP), merged(CF_MAP), merged(FIN_MAP)
