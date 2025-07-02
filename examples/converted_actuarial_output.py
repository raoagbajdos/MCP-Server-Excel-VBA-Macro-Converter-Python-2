import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from typing import Any, Optional, Union
from datetime import datetime
from decimal import Decimal
import os
import sys

# Converted from VBA module: ActuarialCalculations

# Constants
MORTALITY_TABLE_BASE AS DOUBLE = 0.001
INTEREST_RATE AS DOUBLE = 0.05
EXPENSE_RATIO AS DOUBLE = 0.15
PROFIT_MARGIN AS DOUBLE = 0.08

def calculatelifeinsurancepremium() -> None:
    """
    Converted from VBA SUB: CalculateLifeInsurancePremium
    """
    ws: Any = None
    lastrow: int = None
    i: int = None
    age: int = None
    suminsured: float = None
    term: int = None
    premium: float = None
    reserves: float = None
    ws = ActiveSheet
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    # Process each policy
    for i in range(2, lastRow + 1, 1):
    age = ws.Cells(i, 1).Value        ' Column A: Age
    suminsured = ws.Cells(i, 2).Value ' Column B: Sum Insured
    term = ws.Cells(i, 3).Value       ' Column C: Term
    # Calculate premium using NPV method
    premium = CalculateNetPremium(age, sumInsured, term)
    premium = premium * (1 + EXPENSE_RATIO + PROFIT_MARGIN)
    # Calculate policy reserves
    reserves = CalculatePolicyReserve(age, sumInsured, term, premium)
    # Output results
    ws.cells(i, 4).value = premium    ' Column D: Premium
    ws.cells(i, 5).value = reserves   ' Column E: Reserves
    ws.cells(i, 6).value = CalculateCommissionRate(premium) ' Column F: Commission
    # Risk assessment
    ws.cells(i, 7).value = AssessRiskCategory(age, sumInsured)
    # Calculate portfolio statistics
    Call CalculatePortfolioMetrics
    MsgBox "Premium calculations completed for " + (lastRow - 1) + " policies."

def calculatenetpremium(age as integer: Any, suminsured as double: Any, term as integer: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculateNetPremium
    
    Args:
        age as integer: Variant
        suminsured as double: Variant
        term as integer: Variant
    """
    i: int = None
    qx: float = None
    lx: float = None
    dx: float = None
    vx: float = None
    commutationm: float = None
    commutationn: float = None
    netpremium: float = None
    # Initialize life table values
    lx = 100000 ' Radix
    commutationm = 0
    commutationn = 0
    # Calculate commutation functions using standard actuarial methods
    for i in range(0, term + 1, 1):
    qx = GetMortalityRate(age + i)
    dx = lx * qx
    vx = (1 + INTEREST_RATE) ^ -(i + 1)
    commutationm = commutationM + (dx * vx)
    commutationn = commutationN + (lx * vx)
    lx = lx - dx
    # Net single premium calculation
    netpremium = sumInsured * commutationM / commutationN
    calculatenetpremium = netPremium

def getmortalityrate(age as integer: Any) -> Any:
    """
    Converted from VBA FUNCTION: GetMortalityRate
    
    Args:
        age as integer: Variant
    """
    # Simplified Gompertz mortality model
    baserate: float = None
    growthfactor: float = None
    baserate = MORTALITY_TABLE_BASE
    growthfactor = 1.08 ' Annual mortality increase factor
    getmortalityrate = baseRate * (growthFactor ^ (age - 20))
    # Cap mortality rate at 1.0
    if GetMortalityRate > 1 Then GetMortalityRate = 1:

def calculatepolicyreserve(age as integer: Any, suminsured as double: Any, term as integer: Any, premium as double: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculatePolicyReserve
    
    Args:
        age as integer: Variant
        suminsured as double: Variant
        term as integer: Variant
        premium as double: Variant
    """
    currentage: int = None
    yearselapsed: int = None
    remainingterm: int = None
    futurenetpremium: float = None
    futurebenefits: float = None
    reserve: float = None
    # Assume policy is 1 year old for reserve calculation
    yearselapsed = 1
    currentage = age + yearsElapsed
    remainingterm = term - yearsElapsed
    if remainingTerm <= 0 Then:
    calculatepolicyreserve = 0
    Exit Function
    # Calculate prospective reserve
    futurebenefits = CalculateNetPremium(currentAge, sumInsured, remainingTerm)
    futurenetpremium = premium / (1 + EXPENSE_RATIO + PROFIT_MARGIN)
    reserve = futureBenefits - (futureNetPremium * CalculateAnnuityValue(currentAge, remainingTerm))
    calculatepolicyreserve = WorksheetFunction.Max(0, reserve)

def calculateannuityvalue(age as integer: Any, term as integer: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculateAnnuityValue
    
    Args:
        age as integer: Variant
        term as integer: Variant
    """
    i: int = None
    survivalprob: float = None
    discountfactor: float = None
    annuityvalue: float = None
    annuityvalue = 0
    survivalprob = 1
    for i in range(1, term + 1, 1):
    survivalprob = survivalProb * (1 - GetMortalityRate(age + i - 1))
    discountfactor = (1 + INTEREST_RATE) ^ -i
    annuityvalue = annuityValue + (survivalProb * discountFactor)
    calculateannuityvalue = annuityValue

def calculatecommissionrate(premium as double: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculateCommissionRate
    
    Args:
        premium as double: Variant
    """
    # Commission structure based on premium amount
    if premium < 1000 Then:
    calculatecommissionrate = 0.15
elif premium < 5000 Then:
    calculatecommissionrate = 0.12
elif premium < 10000 Then:
    calculatecommissionrate = 0.10
else:
    calculatecommissionrate = 0.08

def assessriskcategory(age as integer: Any, suminsured as double: Any) -> Any:
    """
    Converted from VBA FUNCTION: AssessRiskCategory
    
    Args:
        age as integer: Variant
        suminsured as double: Variant
    """
    riskscore: float = None
    # Risk assessment algorithm
    riskscore = (age / 100) + (sumInsured / 1000000) * 0.1
    if riskScore < 0.3 Then:
    assessriskcategory = "Low Risk"
elif riskScore < 0.6 Then:
    assessriskcategory = "Medium Risk"
elif riskScore < 0.8 Then:
    assessriskcategory = "High Risk"
else:
    assessriskcategory = "Declined"

def calculateportfoliometrics() -> None:
    """
    Converted from VBA SUB: CalculatePortfolioMetrics
    """
    ws: Any = None
    lastrow: int = None
    totalpremium: float = None
    totalreserves: float = None
    totalsuminsured: float = None
    avgage: float = None
    policycount: int = None
    i: int = None
    ws = ActiveSheet
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    policycount = lastRow - 1
    # Calculate portfolio totals
    for i in range(2, lastRow + 1, 1):
    totalpremium = totalPremium + ws.Cells(i, 4).Value
    totalreserves = totalReserves + ws.Cells(i, 5).Value
    totalsuminsured = totalSumInsured + ws.Cells(i, 2).Value
    avgage = avgAge + ws.Cells(i, 1).Value
    avgage = avgAge / policyCount
    # Output portfolio metrics
    ws.cells(lastrow + 3, 1).value = "Portfolio Metrics:"
    ws.cells(lastrow + 4, 1).value = "Total Premium:"
    ws.cells(lastrow + 4, 2).value = totalPremium
    ws.cells(lastrow + 5, 1).value = "Total Reserves:"
    ws.cells(lastrow + 5, 2).value = totalReserves
    ws.cells(lastrow + 6, 1).value = "Total Sum Insured:"
    ws.cells(lastrow + 6, 2).value = totalSumInsured
    ws.cells(lastrow + 7, 1).value = "Average Age:"
    ws.cells(lastrow + 7, 2).value = Round(avgAge, 1)
    ws.cells(lastrow + 8, 1).value = "Loss Ratio:"
    ws.cells(lastrow + 8, 2).value = Round((totalReserves / totalPremium) * 100, 2) + "%"

def calculatesurrendervalue(age as integer: Any, suminsured as double: Any, term as integer: Any, yearselapsed as integer: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculateSurrenderValue
    
    Args:
        age as integer: Variant
        suminsured as double: Variant
        term as integer: Variant
        yearselapsed as integer: Variant
    """
    policyreserve: float = None
    surrendercharges: float = None
    surrendervalue: float = None
    # Calculate current policy reserve
    policyreserve = CalculatePolicyReserve(age, sumInsured, term, 0)
    # Apply surrender charges (decreasing over time)
    if yearsElapsed <= 5 Then:
    surrendercharges = policyReserve * (0.05 * (6 - yearsElapsed))
else:
    surrendercharges = 0
    surrendervalue = policyReserve - surrenderCharges
    calculatesurrendervalue = WorksheetFunction.Max(0, surrenderValue)

def calculatedaly(age as integer: Any, gender as string: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculateDALY
    
    Args:
        age as integer: Variant
        gender as string: Variant
    """
    # Disability-Adjusted Life Years calculation
    lifeexpectancy: float = None
    disabilityweight: float = None
    # Simplified life expectancy calculation
    if gender = "M" Then:
    lifeexpectancy = 78.5 - age
else:
    lifeexpectancy = 82.3 - age
    # Age-based disability weighting
    if age < 65 Then:
    disabilityweight = 0.05
else:
    disabilityweight = 0.15
    calculatedaly = lifeExpectancy * (1 - disabilityWeight)

def runstochasticprojections() -> None:
    """
    Converted from VBA SUB: RunStochasticProjections
    """
    scenarios: int = None
    i: int = None
    j: int = None
    ws: Any = None
    projectionyears: int = None
    randomshock: float = None
    stressedmortality: float = None
    ws = Worksheets.Add
    ws.name = "Stochastic_Projections"
    scenarios = 1000
    projectionyears = 30
    # Monte Carlo simulation for mortality scenarios
    ws.cells(1, 1).value = "Scenario"
    ws.cells(1, 2).value = "Final Reserve"
    ws.cells(1, 3).value = "Profit/Loss"
    for i in range(1, scenarios + 1, 1):
    # Generate random mortality shock
    randomshock = WorksheetFunction.NormInv(Rnd(), 0, 0.2)
    stressedmortality = MORTALITY_TABLE_BASE * (1 + randomShock)
    # Project reserves under stressed scenario
    projectedreserve: float = None
    projectedreserve = ProjectReservesWithShock(stressedMortality, projectionYears)
    ws.cells(i + 1, 1).value = i
    ws.cells(i + 1, 2).value = projectedReserve
    ws.cells(i + 1, 3).value = CalculateProfitLoss(projectedReserve)
    # Calculate percentiles
    Call CalculateRiskMetrics(ws, scenarios)

def projectreserveswithshock(mortalityshock as double: Any, years as integer: Any) -> Any:
    """
    Converted from VBA FUNCTION: ProjectReservesWithShock
    
    Args:
        mortalityshock as double: Variant
        years as integer: Variant
    """
    # Simplified reserve projection with mortality shock
    basereserve: float = None
    shockimpact: float = None
    basereserve = 50000 ' Simplified base reserve
    shockimpact = mortalityShock * years * 1000
    projectreserveswithshock = baseReserve + shockImpact

def calculateprofitloss(finalreserve as double: Any) -> Any:
    """
    Converted from VBA FUNCTION: CalculateProfitLoss
    
    Args:
        finalreserve as double: Variant
    """
    targetreserve: float = None
    targetreserve = 55000 ' Target reserve level
    calculateprofitloss = targetReserve - finalReserve

def calculateriskmetrics(ws as worksheet: Any, scenarios as integer: Any) -> None:
    """
    Converted from VBA SUB: CalculateRiskMetrics
    
    Args:
        ws as worksheet: Variant
        scenarios as integer: Variant
    """
    profitlossrange: Any = None
    var95: float = None
    var99: float = None
    expectedshortfall: float = None
    profitlossrange = ws.Range(ws.worksheet.cell(2, 3), ws.Cells(scenarios + 1, 3))
    # Calculate Value at Risk
    var95 = WorksheetFunction.Percentile(profitLossRange, 0.05)
    var99 = WorksheetFunction.Percentile(profitLossRange, 0.01)
    # Output risk metrics
    ws.cells(scenarios + 3, 1).value = "Risk Metrics:"
    ws.cells(scenarios + 4, 1).value = "VaR 95%:"
    ws.cells(scenarios + 4, 2).value = var95
    ws.cells(scenarios + 5, 1).value = "VaR 99%:"
    ws.cells(scenarios + 5, 2).value = var99
    ws.cells(scenarios + 6, 1).value = "Mean P + L:"
    ws.cells(scenarios + 6, 2).value = WorksheetFunction.Average(profitLossRange)

