Attribute VB_Name = "K2_MonteCarlo_Enhanced"
Option Explicit

Sub RunEnhancedPortfolioMonteCarlo()
    ' Define all variables
    Dim wsPortfolio As Worksheet, wsCor As Worksheet, wsPortfolioMC As Worksheet
    Dim pnlResults As Variant, numScenarios As Long, results As Variant
    Dim targetRiskOfRuin As Double, tolerance As Double, currentRiskOfRuin As Double
    Dim yearsToConsider As Double, ruinedCount As Long, count As Long
    Dim startdate As Date, endDate As Date, tradeAdjustment As Double, solveRisk As String
    Dim startingEquity As Double, requiredMargin As Double, averageTradesPerYear As Long
    Dim clusterSize As Long, optimalBlockSize As Long
    Dim correlationMatrix() As Double, choleskyMatrix() As Double
    Dim numStrategies As Long, currentdate As Date
    Dim AverageTrade() As Double, adjustedTradeFactor() As Double
    Dim dailyEquityTracking() As Double
    
    ' Simulation parameters now customizable via input
    Dim factorModelPercentage As Double
    Dim enableCrisisMode As Boolean
    Dim crisisCorrelationIncrease As Double
    Dim crisisThreshold As Double
    Dim crisisFrequencyMultiplier As Double
    Dim numFactors As Long
    Dim blockSizeMultiplier As Double
    Dim CrisisReturns As Double
    
    ' License check
    If Not IsLicenseValid() Then
        MsgBox "Invalid or missing license.", vbCritical
        Exit Sub
    End If

    ' Initialize column constants
    Call InitializeColumnConstantsManually

    ' Assign worksheets
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsCor = ThisWorkbook.Sheets("Correlations")

    ' Create or replace the PortfolioMC tab
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PortfolioMC").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsPortfolioMC = ThisWorkbook.Sheets.Add(After:=wsPortfolio)
    wsPortfolioMC.name = "PortfolioMC"

    ' Read input parameters
    requiredMargin = GetNamedRangeValue("PortfolioCeaseTrading")
    startingEquity = GetNamedRangeValue("PortfolioStartingEquity")
    numScenarios = GetNamedRangeValue("PortfolioSimulations")
    tradeAdjustment = GetNamedRangeValue("PortfolioMCTradeAdjustment")
    targetRiskOfRuin = GetNamedRangeValue("PortfolioRiskRuinTarget")
    tolerance = GetNamedRangeValue("PortfolioRiskRuinTolerance")
    solveRisk = GetNamedRangeValue("Solve_Risk_Ruin")
    yearsToConsider = GetNamedRangeValue("PortfolioPeriod")
    
    ' Get customizable simulation parameters via input dialog
    If Not GetSimulationParameters(factorModelPercentage, enableCrisisMode, _
                                  crisisCorrelationIncrease, crisisThreshold, _
                                  crisisFrequencyMultiplier, numFactors, _
                                  blockSizeMultiplier, CrisisReturns) Then
        Exit Sub ' User cancelled
    End If

    ' Define date range
    currentdate = wsPortfolio.Cells(2, COL_PORT_LAST_DATE_ON_FILE).value
    startdate = DateAdd("yyyy", -Int(yearsToConsider), currentdate)
    endDate = currentdate

    ' Determine trading frequency
    averageTradesPerYear = 252 ' Daily trades

    ' Get PnL data using improved data cleaning
    pnlResults = CleanPortfolioDailyPnL(startdate, endDate)

    ' Exit if no PnL data
    If IsEmpty(pnlResults) Then
        MsgBox "No valid PnL data found.", vbExclamation
        Exit Sub
    End If

    ' Initialize arrays for trade adjustment per strategy
    numStrategies = UBound(pnlResults, 2)
    ReDim AverageTrade(1 To numStrategies)
    ReDim adjustedTradeFactor(1 To numStrategies)

    ' Compute the Average Trade per Strategy
    Dim j As Long, i As Long
    For j = 1 To numStrategies
        AverageTrade(j) = 0
        Dim TradeCount As Long
        TradeCount = 0
        
        For i = 1 To UBound(pnlResults, 1)
            If pnlResults(i, j) <> 0 Then ' Avoid counting zero trades
                AverageTrade(j) = AverageTrade(j) + pnlResults(i, j)
                TradeCount = TradeCount + 1
            End If
        Next i
        
        ' Prevent division by zero
        If TradeCount > 0 Then
            AverageTrade(j) = AverageTrade(j) / TradeCount
        Else
            AverageTrade(j) = 0
        End If

        ' Compute adjusted trade factor per strategy
        adjustedTradeFactor(j) = AverageTrade(j) * (1 - tradeAdjustment)
    Next j

    ' Read correlation matrix from "Correlations" sheet
    correlationMatrix = ReadCorrelationMatrix(numStrategies)
    
    ' Ensure correlation matrix is positive definite
    correlationMatrix = EnsurePositiveDefiniteMatrix(correlationMatrix, numStrategies)

    ' Prepare crisis correlation matrix if needed using improved method
    Dim crisisCorrelationMatrix() As Double
    If enableCrisisMode Then
        crisisCorrelationMatrix = CreateCrisisCorrelationMatrix(correlationMatrix, crisisCorrelationIncrease, numStrategies)
        ' Ensure crisis correlation matrix is also positive definite
        crisisCorrelationMatrix = EnsurePositiveDefiniteMatrix(crisisCorrelationMatrix, numStrategies)
    End If

    ' Perform Cholesky decomposition for normal correlation using improved method
    choleskyMatrix = CholeskyDecomposition(correlationMatrix, numStrategies)
    
    ' Perform Cholesky decomposition for crisis correlation if needed
    Dim crisisCholeskyMatrix() As Double
    If enableCrisisMode Then
        crisisCholeskyMatrix = CholeskyDecomposition(crisisCorrelationMatrix, numStrategies)
    End If

    ' Calculate optimal block size with improved method
    optimalBlockSize = Int(Sqr(UBound(pnlResults, 1)) * blockSizeMultiplier)
    optimalBlockSize = Application.WorksheetFunction.Max(5, Application.WorksheetFunction.Min(30, optimalBlockSize))
    
    ' Extract factors using improved PCA approach
    Dim factorLoadings As Variant
    Dim factorData As Variant
    factorLoadings = ExtractFactorLoadings(pnlResults, numStrategies, numFactors)
    factorData = ExtractFactorData(pnlResults, factorLoadings, numStrategies, numFactors)

    ' Monte Carlo loop with improved unified approach
    count = 0
    Do
        ' Run the unified Monte Carlo simulation WITH TRACKING
        results = RunUnifiedMonteCarlo(pnlResults, choleskyMatrix, factorLoadings, factorData, _
                                     averageTradesPerYear, startingEquity, numScenarios, _
                                     adjustedTradeFactor, requiredMargin, optimalBlockSize, _
                                     factorModelPercentage, dailyEquityTracking, enableCrisisMode, _
                                     IIf(enableCrisisMode, crisisCholeskyMatrix, Null), _
                                     crisisThreshold, CrisisReturns, crisisCorrelationIncrease, _
                                     crisisFrequencyMultiplier)
        
        ' Evaluate risk of ruin
        ruinedCount = 0
        For i = LBound(results, 1) To UBound(results, 1)
            If results(i, 6) = 1 Then
                ruinedCount = ruinedCount + 1
            End If
        Next i
    
        currentRiskOfRuin = ruinedCount / numScenarios

        ' Output interim results if solving for risk of ruin
        If solveRisk = "Yes" And count > 0 And count Mod 5 = 0 Then
            Debug.Print "Iteration " & count & ": Equity = " & Format(startingEquity, "$#,##0") & _
                      ", Risk of Ruin = " & Format(currentRiskOfRuin, "0.00%")
        End If

        ' Adjust starting equity if solving for risk of ruin with improved algorithm
        If solveRisk = "Yes" Then
            ' More sophisticated adjustment algorithm based on how far we are from target
            Dim adjustmentFactor As Double
            Dim distanceFromTarget As Double
            
            distanceFromTarget = currentRiskOfRuin - targetRiskOfRuin
            
            ' Larger adjustments when far from target, smaller when close
            If Abs(distanceFromTarget) > 0.1 Then
                adjustmentFactor = 1.1 ' Large adjustment
            ElseIf Abs(distanceFromTarget) > 0.05 Then
                adjustmentFactor = 1.05 ' Medium adjustment
            ElseIf Abs(distanceFromTarget) > 0.01 Then
                adjustmentFactor = 1.02 ' Small adjustment
            Else
                adjustmentFactor = 1.01 ' Tiny adjustment
            End If
            
            ' Apply adjustment in the appropriate direction
            If distanceFromTarget > 0 Then
                ' Risk too high, increase equity
                startingEquity = startingEquity * adjustmentFactor
            ElseIf distanceFromTarget < 0 Then
                ' Risk too low, decrease equity
                startingEquity = startingEquity / adjustmentFactor
            End If
        End If

        count = count + 1

    Loop While Abs(currentRiskOfRuin - targetRiskOfRuin) > tolerance And count < 500 And solveRisk = "Yes"

    ' If solving for risk of ruin, show the final equity needed
    If solveRisk = "Yes" Then
        MsgBox "Required starting equity: " & Format(startingEquity, "$#,##0") & vbCrLf & _
               "Risk of ruin: " & Format(currentRiskOfRuin, "0.00%") & vbCrLf & _
               "Target: " & Format(targetRiskOfRuin, "0.00%"), vbInformation, "Risk of Ruin Solution"
    End If

    
    Dim wsCharts As Worksheet
    On Error Resume Next
    Set wsCharts = ThisWorkbook.Sheets("MC_Charts")
    If Err.Number <> 0 Then
        Set wsCharts = ThisWorkbook.Sheets.Add(After:=wsPortfolioMC)
        wsCharts.name = "MC_Charts"
    End If
    On Error GoTo 0
    
    ' Create the Monte Carlo charts and statistics
    CreateMainMonteCarloResults wsCharts, dailyEquityTracking, results, startingEquity
    
    ' Generate and output Monte Carlo summary with improved visualization
    GenerateEnhancedMonteCarloSummary wsPortfolioMC, results, startingEquity, requiredMargin, _
                            numScenarios, currentRiskOfRuin, factorModelPercentage, _
                            enableCrisisMode, crisisCorrelationIncrease, crisisThreshold, _
                            optimalBlockSize, numFactors

    MsgBox "Enhanced Monte Carlo Simulation Completed!", vbInformation
End Sub



Function GetSimulationParameters(factorModelPercentage As Double, enableCrisisMode As Boolean, _
                                crisisCorrelationIncrease As Double, crisisThreshold As Double, _
                                crisisFrequencyMultiplier As Double, numFactors As Long, _
                                blockSizeMultiplier As Double, CrisisReturns As Double) As Boolean
    ' This function handles the input dialog for simulation parameters
    ' Returns True if user confirmed, False if cancelled
    
    ' Default values
    factorModelPercentage = 0.5 ' 60% factor model, 40% block bootstrap
    enableCrisisMode = True
    crisisCorrelationIncrease = 0.3 ' Increase correlations by 0.3 during crisis
    crisisThreshold = -0.05 ' -1% portfolio return triggers crisis
    crisisFrequencyMultiplier = 1#  ' Normal crisis frequency
    numFactors = 4 ' Number of factors to extract
    blockSizeMultiplier = 0.5   ' Multiplier for optimal block size
    CrisisReturns = 0.25
    ' Create UserForm programmatically
    'Dim frmParams As Object
    'Set frmParams = CreateSimulationParamsForm(factorModelPercentage, enableCrisisMode, _
                                            crisisCorrelationIncrease, crisisThreshold, _
                                            crisisFrequencyMultiplier, numFactors, _
                                            blockSizeMultiplier)
    
    ' Show the form and get results
    'frmParams.Show
    
    ' Check if user confirmed or cancelled
    'If frmParams.confirmed Then
        ' Get values from form
        factorModelPercentage = factorModelPercentage 'frmParams.factorModelPercentage
        enableCrisisMode = enableCrisisMode 'frmParams.enableCrisisMode
        crisisCorrelationIncrease = crisisCorrelationIncrease 'frmParams.crisisCorrelationIncrease
        crisisThreshold = crisisThreshold 'frmParams.crisisThreshold
        crisisFrequencyMultiplier = crisisFrequencyMultiplier 'frmParams.crisisFrequencyMultiplier
        numFactors = numFactors 'frmParams.numFactors
        blockSizeMultiplier = blockSizeMultiplier 'frmParams.blockSizeMultiplier
        CrisisReturns = CrisisReturns
        GetSimulationParameters = True
  '  Else
  '      GetSimulationParameters = False
  '  End If
End Function

Function CreateSimulationParamsForm(factorModelPercentage As Double, enableCrisisMode As Boolean, _
                                  crisisCorrelationIncrease As Double, crisisThreshold As Double, _
                                  crisisFrequencyMultiplier As Double, numFactors As Long, _
                                  blockSizeMultiplier As Double) As Object
    ' Create a simple input form using InputBox instead of UserForm for simplicity
    ' In a full implementation, you would create a proper UserForm
    
    ' For now, we'll use a series of InputBox calls
    Dim confirmed As Boolean
    confirmed = True
    
    ' Get factor model percentage
    On Error Resume Next
    Dim tempValue As Variant
    tempValue = InputBox("Enter Factor Model Percentage (0.0 to 1.0):" & vbCrLf & _
                         "0.6 = 60% factor model, 40% bootstrap", _
                         "Simulation Parameters", factorModelPercentage)
    
    If tempValue = "" Then
        confirmed = False
    Else
        factorModelPercentage = CDbl(tempValue)
    End If
    
    ' Get number of factors
    If confirmed Then
        tempValue = InputBox("Enter Number of Factors to Extract (1-3):" & vbCrLf & _
                             "2 is recommended for most portfolios", _
                             "Simulation Parameters", numFactors)
        
        If tempValue = "" Then
            confirmed = False
        Else
            numFactors = CLng(tempValue)
        End If
    End If
    
    ' Get block size multiplier
    If confirmed Then
        tempValue = InputBox("Enter Block Size Multiplier (0.5-2.0):" & vbCrLf & _
                             "1.0 = sqrt(days) (optimal), 0.5 = smaller blocks, 2.0 = larger blocks", _
                             "Simulation Parameters", blockSizeMultiplier)
        
        If tempValue = "" Then
            confirmed = False
        Else
            blockSizeMultiplier = CDbl(tempValue)
        End If
    End If
    
    ' Enable crisis mode?
    If confirmed Then
        tempValue = MsgBox("Enable Crisis Correlation Mode?" & vbCrLf & _
                          "This increases correlations during negative periods", _
                          vbYesNo + vbQuestion, "Simulation Parameters")
        
        enableCrisisMode = (tempValue = vbYes)
    End If
    
    ' If crisis mode enabled, get crisis parameters
    If confirmed And enableCrisisMode Then
        ' Crisis correlation increase
        tempValue = InputBox("Enter Crisis Correlation Increase (0.0-0.7):" & vbCrLf & _
                             "How much correlations increase during crisis (e.g., 0.3 = +0.3)", _
                             "Crisis Parameters", crisisCorrelationIncrease)
        
        If tempValue = "" Then
            confirmed = False
        Else
            crisisCorrelationIncrease = CDbl(tempValue)
        End If
        
        ' Crisis threshold
        If confirmed Then
            tempValue = InputBox("Enter Crisis Threshold (-0.05 to 0.0):" & vbCrLf & _
                                 "Portfolio return below this triggers crisis (e.g., -0.01 = -1%)", _
                                 "Crisis Parameters", crisisThreshold)
            
            If tempValue = "" Then
                confirmed = False
            Else
                crisisThreshold = CDbl(tempValue)
            End If
        End If
        
        ' Crisis frequency multiplier
        If confirmed Then
            tempValue = InputBox("Enter Crisis Frequency Multiplier (0.5-2.0):" & vbCrLf & _
                                 "1.0 = historical frequency, 2.0 = twice as frequent", _
                                 "Crisis Parameters", crisisFrequencyMultiplier)
            
            If tempValue = "" Then
                confirmed = False
            Else
                crisisFrequencyMultiplier = CDbl(tempValue)
            End If
        End If
    End If
    
    ' Create a simple object to return results
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Store parameters and confirmed status
    result.Add "FactorModelPercentage", factorModelPercentage
    result.Add "EnableCrisisMode", enableCrisisMode
    result.Add "CrisisCorrelationIncrease", crisisCorrelationIncrease
    result.Add "CrisisThreshold", crisisThreshold
    result.Add "CrisisFrequencyMultiplier", crisisFrequencyMultiplier
    result.Add "NumFactors", numFactors
    result.Add "BlockSizeMultiplier", blockSizeMultiplier
    result.Add "Confirmed", confirmed
    
    Set CreateSimulationParamsForm = result
End Function





Function CreateCrisisCorrelationMatrix(baseCorrelationMatrix As Variant, correlationIncrease As Double, _
                                      numStrategies As Long) As Variant
    ' Creates a correlation matrix for crisis periods using a more mathematically sound approach
    Dim crisisMatrix() As Double
    Dim i As Long, j As Long
    Dim targetCorr As Double
    Dim alpha As Double
    
    ReDim crisisMatrix(1 To numStrategies, 1 To numStrategies)
    
    ' Use a blending parameter approach (0 <= alpha <= 1)
    ' Higher alpha means more influence from the "extreme correlation" matrix
    alpha = Application.WorksheetFunction.Min(0.8, correlationIncrease * 2)
    
    ' Create the crisis correlation matrix
    For i = 1 To numStrategies
        ' Diagonal elements remain 1
        crisisMatrix(i, i) = 1
        
        For j = i + 1 To numStrategies
            ' Get base correlation
            Dim baseCorr As Double
            baseCorr = baseCorrelationMatrix(i, j)
            
            ' Target correlation is more extreme (closer to 1 for positive, closer to -1 for negative)
            If baseCorr >= 0 Then
                targetCorr = baseCorr + (1 - baseCorr) * alpha
            Else
                targetCorr = baseCorr + (-1 - baseCorr) * alpha
            End If
            
            crisisMatrix(i, j) = targetCorr
            crisisMatrix(j, i) = targetCorr  ' Maintain symmetry
        Next j
    Next i
    
    ' Ensure the matrix is valid (positive definite)
    crisisMatrix = EnsurePositiveDefiniteMatrix(crisisMatrix, numStrategies)
    
    CreateCrisisCorrelationMatrix = crisisMatrix
End Function


' ---------------------------------------
' Function: Cholesky Decomposition
' ---------------------------------------
Function CholeskyDecomposition(correlationMatrix As Variant, numStrategies As Long) As Variant
    Dim i As Integer, j As Integer, K As Integer
    Dim L() As Double
    Dim sum As Double
    
    ' First ensure matrix is positive definite using proper method
    correlationMatrix = EnsurePositiveDefiniteMatrix(correlationMatrix, numStrategies)
    
    ' Initialize L matrix
    ReDim L(1 To numStrategies, 1 To numStrategies)
    
    ' Perform standard Cholesky decomposition (now safer with positive definite matrix)
    On Error Resume Next
    For i = 1 To numStrategies
        For j = 1 To i
            sum = correlationMatrix(i, j)
            For K = 1 To j - 1
                sum = sum - L(i, K) * L(j, K)
            Next K

            If i = j Then
                L(i, j) = Sqr(sum)
            Else
                L(i, j) = sum / L(j, j)
            End If
        Next j
    Next i
    On Error GoTo 0
    
    CholeskyDecomposition = L
End Function



Function GenerateCorrelatedPnL(pnlResults As Variant, choleskyMatrix As Variant, tradeIndex As Long, numStrategies As Long) As Variant
    Dim correlatedPnL() As Double
    Dim i As Integer, j As Integer
    Dim rawPnL() As Double
    
    ReDim correlatedPnL(1 To numStrategies)
    ReDim rawPnL(1 To numStrategies)
    
    For i = 1 To numStrategies
        rawPnL(i) = pnlResults(tradeIndex, i)
    Next i
    
    For i = 1 To numStrategies
        correlatedPnL(i) = 0
        For j = 1 To numStrategies
            correlatedPnL(i) = correlatedPnL(i) + choleskyMatrix(i, j) * rawPnL(j)
        Next j
    Next i
    
    GenerateCorrelatedPnL = correlatedPnL
End Function


Function ComputeDrawdownClusters(pnlResults As Variant, clusterSize As Long) As Variant
    Dim numDays As Long, numStrategies As Long
    Dim i As Long, j As Long
    Dim drawdownClusters() As Variant
    
    numDays = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    
    ReDim drawdownClusters(1 To numDays, 1 To numStrategies)
    
    For j = 1 To numStrategies
        Dim drawdownStreak As Long
        drawdownStreak = 0
        
        For i = 1 To numDays
            If pnlResults(i, j) < 0 Then
                drawdownStreak = drawdownStreak + 1
            Else
                drawdownStreak = 0
            End If
            
            If drawdownStreak >= clusterSize Then
                drawdownClusters(i, j) = 1 ' Mark cluster point
            Else
                drawdownClusters(i, j) = 0
            End If
        Next i
    Next j
    
    ComputeDrawdownClusters = drawdownClusters
End Function


Function ReadCorrelationMatrix(numStrategies As Long) As Variant
    Dim wsCor As Worksheet
    Dim correlationMatrix() As Double
    Dim i As Integer, j As Integer
    
    Set wsCor = ThisWorkbook.Sheets("Correlations")
    ReDim correlationMatrix(1 To numStrategies, 1 To numStrategies)
    
    For i = 1 To numStrategies
        For j = 1 To numStrategies
            correlationMatrix(i, j) = wsCor.Cells(56 + i - 1, 4 + j - 1).value
        Next j
    Next i
    
    ReadCorrelationMatrix = correlationMatrix
End Function


Function ComputeStrategyPnLStdDev(pnlResults As Variant) As Variant
    Dim numDays As Long, numStrategies As Long
    Dim i As Long, j As Long
    Dim mean() As Double, variance() As Double, stdDev() As Double
    Dim tempValue As Double

    ' Get dimensions
    numDays = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)

    ' Initialize arrays
    ReDim mean(1 To numStrategies)
    ReDim variance(1 To numStrategies)
    ReDim stdDev(1 To numStrategies)

    ' Compute mean per strategy
    For j = 1 To numStrategies
        mean(j) = 0
        For i = 1 To numDays
            mean(j) = mean(j) + pnlResults(i, j)
        Next i
        mean(j) = mean(j) / numDays
    Next j

    ' Compute variance per strategy
    For j = 1 To numStrategies
        variance(j) = 0
        For i = 1 To numDays
            tempValue = pnlResults(i, j) - mean(j)
            variance(j) = variance(j) + tempValue * tempValue
        Next i
        variance(j) = variance(j) / (numDays - 1)
        stdDev(j) = Sqr(variance(j)) ' Standard deviation is sqrt(variance)
    Next j

    ComputeStrategyPnLStdDev = stdDev
End Function





Function ComputeSkewness(data() As Double) As Double
    Dim i As Long, meanVal As Double, stdDev As Double, n As Long
    Dim sumSkew As Double
    n = UBound(data)
    
    meanVal = Application.WorksheetFunction.Average(data)
    stdDev = Application.WorksheetFunction.stdev(data)

    If stdDev = 0 Then
        ComputeSkewness = 0
        Exit Function
    End If

    sumSkew = 0
    For i = 1 To n
        sumSkew = sumSkew + ((data(i) - meanVal) / stdDev) ^ 3
    Next i

    ComputeSkewness = (n / ((n - 1) * (n - 2))) * sumSkew
End Function

Function ComputeKurtosis(data As Variant) As Double
    Dim n As Double
    Dim meanVal As Double, stdDev As Double
    Dim sumKurt As Double, x As Double
    Dim i As Long
    Dim numerator As Double, denominator As Double, correction As Double

    n = UBound(data) - LBound(data) + 1
    
    ' Prevent division errors
    If n < 4 Then
        ComputeKurtosis = 0
        Exit Function
    End If

    ' Compute mean
    meanVal = Application.WorksheetFunction.Average(data)
    
    ' Compute standard deviation
    stdDev = Application.WorksheetFunction.stdev(data)
    If stdDev = 0 Then
        ComputeKurtosis = 0
        Exit Function
    End If

    ' Compute sumKurt
    sumKurt = 0
    For i = LBound(data) To UBound(data)
        x = (data(i) - meanVal) / stdDev
        sumKurt = sumKurt + x ^ 4
    Next i

    ' Convert n values to Double to prevent overflow
    Dim nDouble As Double
    nDouble = CDbl(n)

    numerator = nDouble * (nDouble + 1) * sumKurt
    denominator = CDbl(n - 1) * CDbl(n - 2) * CDbl(n - 3)  ' Convert values to Double before multiplication
    correction = (3 * (nDouble - 1) ^ 2) / ((nDouble - 2) * (nDouble - 3))

    ' Check for division by zero
    If denominator = 0 Then
        ComputeKurtosis = 0
        Exit Function
    End If

    ComputeKurtosis = (numerator / denominator) - correction
End Function

Function GenerateRandomFromDistribution(meanVal As Double, stdDev As Double, skew As Double, kurt As Double) As Double
    Dim normalRandom As Double
    Dim randValue As Double
    
    ' Get a random value between 0.001 and 0.999 to avoid numerical errors
    randValue = 0.001 + 0.998 * Rnd()
    
    ' Use bounded random value for inverse normal calculation
    normalRandom = Application.WorksheetFunction.Norm_S_Inv(randValue) * stdDev + meanVal
    
    ' Note: skew and kurt parameters are currently not used in this simplified implementation
    ' For a full implementation, you would use these to adjust the distribution
    
    GenerateRandomFromDistribution = normalRandom
End Function

Function ExtractFactorLoadings(pnlResults As Variant, numStrategies As Long, numFactors As Long) As Variant
    ' This function implements proper PCA for factor extraction
    Dim covMatrix As Variant
    Dim factorLoadings() As Double
    Dim eigenValues() As Double, eigenVectors() As Double
    Dim i As Long, j As Long
    Dim numDays As Long
    
    numDays = UBound(pnlResults, 1)
    
    ' Calculate covariance matrix of returns
    covMatrix = CalculateCovarianceMatrix(pnlResults, numDays, numStrategies)
    
    ' Extract eigenvalues and eigenvectors (using helper function)
    ReDim eigenValues(1 To numStrategies)
    ReDim eigenVectors(1 To numStrategies, 1 To numStrategies)
    
    Call ComputeEigenDecomposition(covMatrix, numStrategies, eigenValues, eigenVectors)
    
    ' Sort eigenvalues and eigenvectors by eigenvalue magnitude (descending)
    Call SortEigenSystem(eigenValues, eigenVectors, numStrategies)
    
    ' Take the top numFactors eigenvectors as factor loadings
    ReDim factorLoadings(1 To numStrategies, 1 To numFactors)
    
    ' Extract loadings from eigenvectors, scaled by sqrt of eigenvalues
    For i = 1 To numStrategies
        For j = 1 To numFactors
            ' Scale loading by sqrt of eigenvalue for proper variance explanation
            factorLoadings(i, j) = eigenVectors(i, j) * Sqr(eigenValues(j))
        Next j
    Next i
    
    ExtractFactorLoadings = factorLoadings
End Function

Function ComputeEigenDecomposition(matrix As Variant, n As Long, ByRef eigenValues() As Double, ByRef eigenVectors() As Double)
    ' Simplified Power Method for eigenvalue decomposition
    ' This is a basic implementation that works reasonably well for correlation matrices
    ' For production use, consider using an external library or more sophisticated algorithm
    
    Dim i As Long, j As Long, K As Long, iter As Long
    Dim maxIter As Long, prevVector() As Double, normFactor As Double
    Dim tempMatrix() As Double, tempVector() As Double, dotProduct As Double
    Dim converged As Boolean, tolerance As Double
    
    maxIter = 100  ' Maximum iterations for each eigenvalue
    tolerance = 0.000001  ' Convergence tolerance
    
    ' Create working copy of the matrix
    ReDim tempMatrix(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            tempMatrix(i, j) = matrix(i, j)
        Next j
    Next i
    
    ' Compute each eigenvalue/eigenvector pair
    For K = 1 To n
        ' Initialize with random vector
        ReDim eigenVectors(1 To n, K)
        ReDim prevVector(1 To n)
        
        Randomize
        For i = 1 To n
            eigenVectors(i, K) = Rnd()
            prevVector(i) = 0
        Next i
        
        ' Normalize initial vector
        normFactor = 0
        For i = 1 To n
            normFactor = normFactor + eigenVectors(i, K) ^ 2
        Next i
        normFactor = Sqr(normFactor)
        
        If normFactor > 0 Then
            For i = 1 To n
                eigenVectors(i, K) = eigenVectors(i, K) / normFactor
            Next i
        End If
        
        ' Power iteration
        converged = False
        For iter = 1 To maxIter
            ' Store previous vector
            For i = 1 To n
                prevVector(i) = eigenVectors(i, K)
            Next i
            
            ' Multiply matrix by vector
            ReDim tempVector(1 To n)
            For i = 1 To n
                tempVector(i) = 0
                For j = 1 To n
                    tempVector(i) = tempVector(i) + tempMatrix(i, j) * prevVector(j)
                Next j
            Next i
            
            ' Normalize result
            normFactor = 0
            For i = 1 To n
                normFactor = normFactor + tempVector(i) ^ 2
            Next i
            normFactor = Sqr(normFactor)
            
            If normFactor > 0 Then
                For i = 1 To n
                    eigenVectors(i, K) = tempVector(i) / normFactor
                Next i
            End If
            
            ' Check convergence - dot product of current and previous vectors
            dotProduct = 0
            For i = 1 To n
                dotProduct = dotProduct + eigenVectors(i, K) * prevVector(i)
            Next i
            
            If Abs(Abs(dotProduct) - 1) < tolerance Then
                converged = True
                Exit For
            End If
        Next iter
        
        ' Compute eigenvalue (Rayleigh quotient)
        eigenValues(K) = 0
        For i = 1 To n
            For j = 1 To n
                eigenValues(K) = eigenValues(K) + eigenVectors(i, K) * tempMatrix(i, j) * eigenVectors(j, K)
            Next j
        Next i
        
        ' Deflate the matrix for next eigenvalue
        For i = 1 To n
            For j = 1 To n
                tempMatrix(i, j) = tempMatrix(i, j) - eigenValues(K) * eigenVectors(i, K) * eigenVectors(j, K)
            Next j
        Next i
    Next K
End Function

Function SortEigenSystem(ByRef eigenValues() As Double, ByRef eigenVectors() As Double, n As Long)
    ' Sort eigenvalues and eigenvectors by eigenvalue magnitude (descending)
    Dim i As Long, j As Long, K As Long
    Dim tempVal As Double, tempVec() As Double
    
    ReDim tempVec(1 To n)
    
    ' Simple bubble sort
    For i = 1 To n - 1
        For j = i + 1 To n
            ' Sort in descending order of absolute value
            If Abs(eigenValues(i)) < Abs(eigenValues(j)) Then
                ' Swap eigenvalues
                tempVal = eigenValues(i)
                eigenValues(i) = eigenValues(j)
                eigenValues(j) = tempVal
                
                ' Swap corresponding eigenvectors
                For K = 1 To n
                    tempVal = eigenVectors(K, i)
                    eigenVectors(K, i) = eigenVectors(K, j)
                    eigenVectors(K, j) = tempVal
                Next K
            End If
        Next j
    Next i
End Function

Function CalculateCovarianceMatrix(data As Variant, numDays As Long, numStrategies As Long) As Variant
    ' Calculate covariance matrix from PnL data
    Dim covMatrix() As Double
    Dim means() As Double
    Dim i As Long, j As Long, K As Long
    
    ReDim covMatrix(1 To numStrategies, 1 To numStrategies)
    ReDim means(1 To numStrategies)
    
    ' Calculate means
    For j = 1 To numStrategies
        means(j) = 0
        For i = 1 To numDays
            means(j) = means(j) + data(i, j)
        Next i
        means(j) = means(j) / numDays
    Next j
    
    ' Calculate covariance matrix
    For i = 1 To numStrategies
        For j = 1 To numStrategies
            covMatrix(i, j) = 0
            For K = 1 To numDays
                covMatrix(i, j) = covMatrix(i, j) + (data(K, i) - means(i)) * (data(K, j) - means(j))
            Next K
            covMatrix(i, j) = covMatrix(i, j) / (numDays - 1)
        Next j
    Next i
    
    CalculateCovarianceMatrix = covMatrix
End Function

' More sophisticated version of EnsurePositiveDefiniteMatrix that uses eigenvalue decomposition
Function EnsurePositiveDefiniteMatrix(matrix As Variant, n As Long) As Variant
    ' Ensure a correlation/covariance matrix is positive definite
    ' by finding eigenvalues and replacing negative ones with small positive values
    
    Dim eigenValues() As Double, eigenVectors() As Double
    Dim result() As Double
    Dim i As Long, j As Long, K As Long
    Dim minEigenvalue As Double
    
    ReDim eigenValues(1 To n)
    ReDim eigenVectors(1 To n, 1 To n)
    ReDim result(1 To n, 1 To n)
    
    ' Compute eigendecomposition
    Call ComputeEigenDecomposition(matrix, n, eigenValues, eigenVectors)
    
    ' Fix non-positive eigenvalues
    minEigenvalue = 0.000001 ' Small positive threshold
    
    For i = 1 To n
        If eigenValues(i) < minEigenvalue Then
            eigenValues(i) = minEigenvalue
        End If
    Next i
    
    ' Reconstruct matrix: M = V * D * V'
    ' Where V is eigenvector matrix, D is diagonal eigenvalue matrix, V' is transpose of V
    
    ' Initialize result matrix
    For i = 1 To n
        For j = 1 To n
            result(i, j) = 0
        Next j
    Next i
    
    ' Matrix multiplication V * D * V'
    For i = 1 To n
        For j = 1 To n
            For K = 1 To n
                result(i, j) = result(i, j) + eigenVectors(i, K) * eigenValues(K) * eigenVectors(j, K)
            Next K
        Next j
    Next i
    
    ' Ensure diagonal is exactly 1 for correlation matrices
    For i = 1 To n
        ' Get the diagonal scaling factor
        Dim diagFactor As Double
        diagFactor = Sqr(result(i, i))
        
        If diagFactor > 0 Then
            ' Scale row and column to ensure diagonal is 1
            For j = 1 To n
                result(i, j) = result(i, j) / diagFactor
                result(j, i) = result(j, i) / diagFactor
            Next j
        End If
    Next i
    
    ' Final fix for diagonal elements
    For i = 1 To n
        result(i, i) = 1
    Next i
    
    EnsurePositiveDefiniteMatrix = result
End Function
Function CalculateCorrelationMatrix(data As Variant, numDays As Long, numStrategies As Long) As Variant
    ' Calculate correlation matrix from standardized data
    Dim corrMatrix() As Double
    Dim i As Long, j As Long, K As Long
    Dim sum As Double
    
    ReDim corrMatrix(1 To numStrategies, 1 To numStrategies)
    
    For i = 1 To numStrategies
        ' Diagonal elements are 1
        corrMatrix(i, i) = 1
        
        ' Calculate correlations
        For j = i + 1 To numStrategies
            sum = 0
            For K = 1 To numDays
                sum = sum + data(K, i) * data(K, j)
            Next K
            
            corrMatrix(i, j) = sum / (numDays - 1)
            corrMatrix(j, i) = corrMatrix(i, j) ' Symmetric matrix
        Next j
    Next i
    
    CalculateCorrelationMatrix = corrMatrix
End Function

Function ExtractFactorData(pnlResults As Variant, factorLoadings As Variant, numStrategies As Long, numFactors As Long) As Variant
    ' This function extracts factor time series from the PnL data using the loadings
    Dim factorData() As Double
    Dim i As Long, j As Long, f As Long
    Dim numDays As Long
    
    numDays = UBound(pnlResults, 1)
    ReDim factorData(1 To numDays, 1 To numFactors)
    
    ' Extract factor data for each day
    For i = 1 To numDays
        For f = 1 To numFactors
            factorData(i, f) = 0
            For j = 1 To numStrategies
                factorData(i, f) = factorData(i, f) + pnlResults(i, j) * factorLoadings(j, f)
            Next j
        Next f
    Next i
    
    ExtractFactorData = factorData
End Function


Function DetectCrisisMode(portfolioReturns As Variant, lookbackPeriod As Long, _
                        currentIndex As Long, inCrisisMode As Boolean, _
                        daysSinceCrisisStart As Long, crisisLength As Long) As Boolean
    ' More sophisticated crisis detection using volatility clustering and momentum
    Dim i As Long
    Dim recentReturns() As Double
    Dim recentVolatility As Double
    Dim historicalVolatility As Double
    Dim momentumIndicator As Double
    Dim threshold As Double
    
    ' If already in crisis mode, check for exit conditions
    If inCrisisMode Then
        daysSinceCrisisStart = daysSinceCrisisStart + 1
        
        ' Exit crisis after its duration is complete
        If daysSinceCrisisStart > crisisLength Then
            DetectCrisisMode = False
            Exit Function
        End If
        
        ' Check for strong recovery that would end crisis early
        If currentIndex > 5 Then
            momentumIndicator = 0
            For i = 0 To 4
                momentumIndicator = momentumIndicator + portfolioReturns(currentIndex - i)
            Next i
            
            ' Strong positive momentum can end crisis early
            If momentumIndicator > 0.03 Then ' 3% recovery over 5 days
                DetectCrisisMode = False
                Exit Function
            End If
        End If
        
        ' Still in crisis
        DetectCrisisMode = True
        Exit Function
    End If
    
    ' If not in crisis, check for entry conditions
    
    ' Calculate recent volatility (last 10 days)
    ReDim recentReturns(1 To Application.WorksheetFunction.Min(10, currentIndex))
    For i = 1 To UBound(recentReturns)
        recentReturns(i) = portfolioReturns(currentIndex - i + 1)
    Next i
    
    recentVolatility = Application.WorksheetFunction.stdev(recentReturns)
    
    ' Calculate longer-term historical volatility (30 days)
    If currentIndex >= 30 Then
        ReDim recentReturns(1 To 30)
        For i = 1 To 30
            recentReturns(i) = portfolioReturns(currentIndex - i + 1)
        Next i
        historicalVolatility = Application.WorksheetFunction.stdev(recentReturns)
    Else
        historicalVolatility = recentVolatility
    End If
    
    ' Get current return
    Dim currentReturn As Double
    currentReturn = portfolioReturns(currentIndex)
    
    ' Crisis conditions:
    ' 1. Sharp negative return (worse than -2%)
    ' 2. Volatility spike (recent vol > 1.5x historical)
    ' 3. Negative momentum (sum of last 3 days < -3%)
    
    Dim crisisConditions As Long
    crisisConditions = 0
    
    ' Check condition 1: Sharp negative return
    If currentReturn < -0.02 Then crisisConditions = crisisConditions + 1
    
    ' Check condition 2: Volatility spike
    If recentVolatility > 1.5 * historicalVolatility Then crisisConditions = crisisConditions + 1
    
    ' Check condition 3: Negative momentum
    If currentIndex >= 3 Then
        momentumIndicator = portfolioReturns(currentIndex) + _
                          portfolioReturns(currentIndex - 1) + _
                          portfolioReturns(currentIndex - 2)
        
        If momentumIndicator < -0.03 Then crisisConditions = crisisConditions + 1
    End If
    
    ' Enter crisis if at least 2 conditions are met
    If crisisConditions >= 2 Then
        DetectCrisisMode = True
        daysSinceCrisisStart = 1
        
        ' Crisis duration based on severity
        If crisisConditions = 3 Then
            ' Severe crisis lasts longer
            crisisLength = 10 + Int(Rnd() * 15) ' 10-25 days
        Else
            ' Moderate crisis
            crisisLength = 5 + Int(Rnd() * 10) ' 5-15 days
        End If
    Else
        DetectCrisisMode = False
    End If
End Function
Function RunUnifiedMonteCarlo(pnlResults As Variant, normalCholeskyMatrix() As Double, _
                             factorLoadings As Variant, factorData As Variant, _
                             averageTradesPerYear As Long, startingEquity As Double, _
                             numScenarios As Long, adjustedTradeFactor() As Double, _
                             requiredMargin As Double, blockSize As Long, _
                             factorModelPercentage As Double, _
                             ByRef dailyEquityTracking() As Double, _
                             Optional enableCrisisMode As Boolean = False, _
                             Optional crisisCholeskyMatrix As Variant = Null, _
                             Optional crisisThreshold As Double = -0.01, _
                             Optional CrisisReturns As Double = 0, _
                             Optional crisisCorrelationIncrease As Double = 0.2, _
                             Optional crisisFrequencyMultiplier As Double = 1) As Variant
    ' Modified version of RunUnifiedMonteCarlo that tracks daily equity values for all scenarios
    ' Parameters:
    '   pnlResults - Historical PnL data (days x strategies)
    '   normalCholeskyMatrix - Cholesky decomposition of normal correlation matrix
    '   factorLoadings - Factor loadings matrix (strategies x factors)
    '   factorData - Historical factor data (days x factors)
    '   averageTradesPerYear - Number of trading days to simulate per year
    '   startingEquity - Initial portfolio equity
    '   numScenarios - Number of Monte Carlo scenarios to run
    '   adjustedTradeFactor - Adjustment factor for each strategy's PnL
    '   requiredMargin - Margin requirement as fraction of starting equity
    '   blockSize - Size of blocks for bootstrap simulation
    '   factorModelPercentage - Proportion of simulations using factor model vs bootstrap
    '   dailyEquityTracking - Output array that will contain daily equity for all scenarios
    '   enableCrisisMode - Whether to enable crisis correlation mode
    '   crisisCholeskyMatrix - Cholesky matrix for crisis periods (required if enableCrisisMode=True)
    '   crisisThreshold - Drawdown threshold that triggers crisis mode (e.g., -0.10 for 10% drawdown)
    '   crisisFrequencyMultiplier - Multiplier for historical crisis frequency
    
    Dim results() As Variant
    Dim i As Long, j As Long, K As Long, f As Long
    Dim numStrategies As Long, numDays As Long, numFactors As Long
    Dim simEquity As Double, peakEquity As Double, drawdown As Double, maxDrawdown As Double
    Dim TradePNL As Double, dailyReturn As Double, marginThreshold As Double
    Dim factorScenarios As Long, bootstrapScenarios As Long
    Dim inCrisisMode As Boolean
    Dim daysSinceCrisisStart As Long, crisisLength As Long
    Dim factorValues() As Double, lastFactorValues() As Double
    Dim bootstrappedData As Variant
    Dim crisisModeCount As Long, crisisDayCount As Long
    
    ' Statistics for factors and residuals
    Dim factorMean() As Double, factorStdDev() As Double
    Dim factorSkew() As Double, factorKurt() As Double
    Dim residuals() As Double
    Dim residualMean() As Double, residualStdDev() As Double
    
    ' Initialize dimensions
    numDays = UBound(pnlResults, 1)
    numStrategies = UBound(pnlResults, 2)
    numFactors = UBound(factorLoadings, 2)
    
    ' Initialize the daily equity tracking array
    ReDim dailyEquityTracking(1 To numScenarios, 0 To averageTradesPerYear)
    
    ' Set initial equity for all scenarios (day 0)
    For i = 1 To numScenarios
        dailyEquityTracking(i, 0) = startingEquity
    Next i
    
    ' Determine results array size based on whether crisis mode is enabled
    If enableCrisisMode Then
        ReDim results(1 To numScenarios, 1 To 7)  ' Add column for crisis days
    Else
        ReDim results(1 To numScenarios, 1 To 6)
    End If
    
    ' Set margin threshold
    marginThreshold = requiredMargin * startingEquity
    
    ' Calculate how many scenarios to run with each method
    factorScenarios = Int(numScenarios * factorModelPercentage)
    bootstrapScenarios = numScenarios - factorScenarios
    
    ' Calculate historical crisis frequency based on drawdowns
    Dim crisisCount As Long
    crisisCount = 0
    
    If enableCrisisMode Then
        Dim historicalEquity As Double
        Dim historicalPeak As Double
        Dim historicalDrawdown As Double
        
        historicalEquity = startingEquity
        historicalPeak = startingEquity
        
        For i = 1 To numDays
            Dim dailyPortfolioPnL As Double
            dailyPortfolioPnL = 0
            
            ' Calculate portfolio return (equal weight)
            For j = 1 To numStrategies
                dailyPortfolioPnL = dailyPortfolioPnL + pnlResults(i, j)
            Next j
            
            ' Update historical equity and peak
            historicalEquity = historicalEquity + dailyPortfolioPnL
            
            If historicalEquity > historicalPeak Then
                historicalPeak = historicalEquity
            End If
            
            ' Calculate drawdown
            If historicalPeak > 0 Then
                historicalDrawdown = (historicalPeak - historicalEquity) / historicalPeak
            Else
                historicalDrawdown = 0
            End If
            
            ' Count crisis days based on drawdown threshold
            If historicalDrawdown > Abs(crisisThreshold) Then
                crisisCount = crisisCount + 1
            End If
        Next i
    End If
    
    ' Calculate factor statistics - cached for better performance
    ReDim factorMean(1 To numFactors)
    ReDim factorStdDev(1 To numFactors)
    ReDim factorSkew(1 To numFactors)
    ReDim factorKurt(1 To numFactors)
    
    For f = 1 To numFactors
        Dim factorArray() As Double
        ReDim factorArray(1 To numDays)
        
        For i = 1 To numDays
            factorArray(i) = factorData(i, f)
        Next i
        
        factorMean(f) = Application.WorksheetFunction.Average(factorArray)
        factorStdDev(f) = Application.WorksheetFunction.stdev(factorArray)
        factorSkew(f) = ComputeSkewness(factorArray)
        factorKurt(f) = ComputeKurtosis(factorArray)
    Next f
    
    ' Calculate residuals (idiosyncratic components)
    ReDim residuals(1 To numDays, 1 To numStrategies)
    
    For i = 1 To numDays
        For j = 1 To numStrategies
            residuals(i, j) = pnlResults(i, j)
            
            ' Subtract factor contributions
            For f = 1 To numFactors
                residuals(i, j) = residuals(i, j) - factorData(i, f) * factorLoadings(j, f)
            Next f
        Next j
    Next i
    
    ' Calculate residual statistics
    ReDim residualMean(1 To numStrategies)
    ReDim residualStdDev(1 To numStrategies)
    
    For j = 1 To numStrategies
        Dim residualArray() As Double
        ReDim residualArray(1 To numDays)
        
        For i = 1 To numDays
            residualArray(i) = residuals(i, j)
        Next i
        
        residualMean(j) = Application.WorksheetFunction.Average(residualArray)
        residualStdDev(j) = Application.WorksheetFunction.stdev(residualArray)
    Next j
    
    ' Pre-generate bootstrapped data for all bootstrap scenarios for efficiency
    Dim allBootstrappedData() As Variant
     
    If bootstrapScenarios = 0 Then bootstrapScenarios = 1
    
    ReDim allBootstrappedData(1 To bootstrapScenarios)
    
    For i = 1 To bootstrapScenarios
        allBootstrappedData(i) = SampleBlockBootstrap(pnlResults, blockSize, averageTradesPerYear, numStrategies)
    Next i
    
    Randomize
    
    ' Main simulation loop for all scenarios
    For i = 1 To numScenarios
        ' Determine if this is a factor model or bootstrap simulation
        Dim useFactorModel As Boolean
        useFactorModel = (i <= factorScenarios)
        
        ' Initialize scenario variables
        simEquity = startingEquity
        peakEquity = startingEquity
        maxDrawdown = 0
        drawdown = 0
        inCrisisMode = False
        daysSinceCrisisStart = 0
        crisisModeCount = 0
        crisisDayCount = 0
        
        ' Initialize factor values arrays for factor model
        If useFactorModel Then
            ReDim factorValues(1 To numFactors)
            ReDim lastFactorValues(1 To numFactors)
            
            ' Reset factor values
            For f = 1 To numFactors
                factorValues(f) = 0
                lastFactorValues(f) = 0
            Next f
        Else
            ' For bootstrap, get the pre-generated data
            Dim bootstrapIndex As Long
            bootstrapIndex = i - factorScenarios
            If bootstrapIndex <= 0 Then bootstrapIndex = 1
            bootstrappedData = allBootstrappedData(bootstrapIndex)
        End If
        
        ' Simulate trading days
        For j = 1 To averageTradesPerYear
            dailyReturn = 0
            
            ' Determine if we're in crisis mode (if enabled)
            If enableCrisisMode Then
                If inCrisisMode Then
                    daysSinceCrisisStart = daysSinceCrisisStart + 1
                    crisisDayCount = crisisDayCount + 1
                    
                    ' Exit crisis mode after its duration
                    If daysSinceCrisisStart > crisisLength Then
                        inCrisisMode = False
                        daysSinceCrisisStart = 0
                    End If
                Else
                    ' Random chance to enter crisis mode based on historical frequency
                    ' Only check if not already triggered by drawdown
                    If crisisCount > 0 And Rnd() < (crisisCount / numDays) * crisisFrequencyMultiplier Then
                        inCrisisMode = True
                        daysSinceCrisisStart = 1
                        crisisLength = 3 + Int(Rnd() * 8) ' Crisis lasts 3-10 days
                        crisisModeCount = crisisModeCount + 1
                        crisisDayCount = crisisDayCount + 1
                    End If
                End If
            End If
            
            ' Generate PnL based on simulation method
            If useFactorModel Then
                ' Choose appropriate Cholesky matrix based on market mode
                Dim currentCholeskyMatrix As Variant
                If inCrisisMode And enableCrisisMode Then
                    currentCholeskyMatrix = crisisCholeskyMatrix
                Else
                    currentCholeskyMatrix = normalCholeskyMatrix
                End If
                
                ' Save current factor values before generating new ones
                For f = 1 To numFactors
                    lastFactorValues(f) = factorValues(f)
                Next f
                
                ' Generate factor values with autocorrelation
                For f = 1 To numFactors
                    If j = 1 Then
                        ' First day: sample from factor distribution
                        factorValues(f) = GenerateRandomFromDistribution(factorMean(f), factorStdDev(f), factorSkew(f), factorKurt(f))
                    Else
                        ' Subsequent days: add autocorrelation
                        Dim autocorr As Double
                        autocorr = 0.2 ' Mild autocorrelation
                        factorValues(f) = factorMean(f) + autocorr * (lastFactorValues(f) - factorMean(f)) + _
                                         Sqr(1 - autocorr * autocorr) * GenerateRandomFromDistribution(0, factorStdDev(f), factorSkew(f), factorKurt(f))
                    End If
                Next f
                
                ' Generate PnL for each strategy
                TradePNL = 0
                For K = 1 To numStrategies
                    Dim strategyPnL As Double
                    strategyPnL = 0
                    
                    ' Factor component
                    For f = 1 To numFactors
                        strategyPnL = strategyPnL + factorValues(f) * factorLoadings(K, f)
                    Next f
                    
                    ' Residual component
                    strategyPnL = strategyPnL + GenerateRandomFromDistribution(residualMean(K), residualStdDev(K), 0, 3)
                    
                    ' During crisis, potentially amplify negative returns
                    If inCrisisMode And enableCrisisMode And strategyPnL < 0 Then
                        Dim crisisMultiplier As Double
                        crisisMultiplier = 1 + (Rnd() * CrisisReturns)
                        strategyPnL = strategyPnL * crisisMultiplier
                    End If
                    
                    ' Apply trade adjustment
                    strategyPnL = strategyPnL - adjustedTradeFactor(K)
                    
                    ' Add to total daily PnL
                    TradePNL = TradePNL + strategyPnL
                    
                    ' Keep track of daily return for calculations
                    dailyReturn = dailyReturn + strategyPnL / numStrategies
                Next K
            Else
                ' Bootstrap simulation - use pre-generated bootstrap data
                TradePNL = 0
                
                ' Sum PnL across strategies for this day
                For K = 1 To numStrategies
                    Dim dayPnL As Double
                    dayPnL = bootstrappedData(j, K)
                    
                    ' During crisis, amplify negative returns
                    If inCrisisMode And enableCrisisMode And dayPnL < 0 Then
                        Dim bootCrisisMultiplier As Double
                        bootCrisisMultiplier = 1 + (Rnd() * CrisisReturns)
                        dayPnL = dayPnL * bootCrisisMultiplier
                    End If
                    
                    ' Apply trade adjustment
                    dayPnL = dayPnL - adjustedTradeFactor(K)
                    
                    ' Add to total daily PnL
                    TradePNL = TradePNL + dayPnL
                    
                    ' Keep track of daily return for calculations
                    dailyReturn = dailyReturn + dayPnL / numStrategies
                Next K
                
                ' Apply correlated adjustments during crisis
                If inCrisisMode And enableCrisisMode Then
                    ' In crisis mode, correlations increase, so adjust the tradePnL
                    ' This simulates the increased correlation effect
                    Dim corrAdjustment As Double
                    corrAdjustment = 0
                    
                    ' If it's a negative day, make it more negative (more correlated down moves)
                    If TradePNL < 0 Then
                        corrAdjustment = TradePNL * crisisCorrelationIncrease
                    End If
                    
                    TradePNL = TradePNL + corrAdjustment
                End If
            End If
            
            ' Apply realistic limits to extreme values using smooth transition
            TradePNL = ApplyRealisticExtremeValueHandling(TradePNL, startingEquity)
            
            ' Update equity
            simEquity = simEquity + TradePNL
            
            ' Track daily equity value for this scenario and day
            dailyEquityTracking(i, j) = simEquity
            
            ' Track peak equity and drawdown
            If simEquity > peakEquity Then
                peakEquity = simEquity
            End If
            
            ' Calculate current drawdown
            If peakEquity > 0 Then
                drawdown = (peakEquity - simEquity) / peakEquity
                If drawdown > maxDrawdown Then maxDrawdown = drawdown
            Else
                drawdown = 0
            End If
            
            ' Check for ruin
            If simEquity < marginThreshold Then
                results(i, 6) = 1 ' Mark as ruined
                
                ' For ruined scenarios, fill remaining days with the ruin value
                For K = j + 1 To averageTradesPerYear
                    dailyEquityTracking(i, K) = simEquity
                Next K
                
                Exit For
            End If
            
            ' Potential crisis mode entry based on drawdown (NEW)
            If enableCrisisMode And Not inCrisisMode And drawdown > Abs(crisisThreshold) Then
                inCrisisMode = True
                daysSinceCrisisStart = 1
                crisisLength = 3 + Int(Rnd() * 8) ' Crisis lasts 3-10 days
                crisisModeCount = crisisModeCount + 1
                crisisDayCount = crisisDayCount + 1
            End If
        Next j
        
        ' Store results
        results(i, 1) = simEquity  ' Final equity
        results(i, 2) = simEquity - startingEquity  ' Net profit
        results(i, 3) = (simEquity - startingEquity) / startingEquity  ' Return %
        results(i, 4) = IIf(maxDrawdown > 0, results(i, 3) / maxDrawdown, 4)  ' Return/Drawdown
        results(i, 5) = maxDrawdown  ' Max drawdown
        If results(i, 6) <> 1 Then results(i, 6) = 0  ' Ensure ruin flag is set
        
        ' Add crisis days if applicable
        If enableCrisisMode Then
            results(i, 7) = crisisDayCount
        End If
    Next i
    
    RunUnifiedMonteCarlo = results
End Function

Function CalculateDailyEquityStatistics(dailyEquityTracking() As Double, ByRef dailyStats() As Double) As Boolean
    ' Calculate daily statistics from the equity tracking array
    '
    ' Parameters:
    '   dailyEquityTracking - Array containing daily equity values (scenarios x days)
    '   dailyStats - Output array that will contain statistics for each day:
    '                Column 1: Average
    '                Column 2: Median
    '                Column 3: 10th percentile
    '                Column 4: 25th percentile
    '                Column 5: 75th percentile
    '                Column 6: 90th percentile
    '                Column 7: 99th percentile (worst case)
    '                Column 8: Minimum value
    '                Column 9: Maximum value
    '
    ' Returns True if successful, False otherwise
    
    On Error GoTo ErrorHandler
    
    Dim numScenarios As Long, numDays As Long
    Dim i As Long, j As Long
    
    ' Get dimensions
    numScenarios = UBound(dailyEquityTracking, 1)
    numDays = UBound(dailyEquityTracking, 2)
    
    ' Initialize output array
    ReDim dailyStats(0 To numDays, 1 To 9)
    
    ' Create a temporary array for sorting values for each day
    Dim tempValues() As Double
    ReDim tempValues(1 To numScenarios)
    
    ' Calculate statistics for each day
    For j = 0 To numDays
        ' Extract values for this day across all scenarios
        For i = 1 To numScenarios
            tempValues(i) = dailyEquityTracking(i, j)
        Next i
        
        ' Calculate statistics
        dailyStats(j, 1) = Application.WorksheetFunction.Average(tempValues)  ' Average
        dailyStats(j, 2) = Application.WorksheetFunction.Median(tempValues)   ' Median
        
        ' Sort values for percentile calculations
        Call SortArray(tempValues, numScenarios)
        
        ' Calculate percentiles
        dailyStats(j, 3) = GetPercentile(tempValues, 0.1)   ' 10th percentile
        dailyStats(j, 4) = GetPercentile(tempValues, 0.25)  ' 25th percentile
        dailyStats(j, 5) = GetPercentile(tempValues, 0.75)  ' 75th percentile
        dailyStats(j, 6) = GetPercentile(tempValues, 0.9)   ' 90th percentile
        dailyStats(j, 7) = GetPercentile(tempValues, 0.99)  ' 99th percentile (worst case)
        
        ' Min and Max
        dailyStats(j, 8) = Application.WorksheetFunction.Min(tempValues)  ' Minimum
        dailyStats(j, 9) = Application.WorksheetFunction.Max(tempValues)  ' Maximum
    Next j
    
    CalculateDailyEquityStatistics = True
    Exit Function
    
ErrorHandler:
    CalculateDailyEquityStatistics = False
End Function

Sub SortArray(arr() As Double, n As Long)
    ' Simple bubble sort implementation
    Dim i As Long, j As Long
    Dim temp As Double
    
    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub

Function GetPercentile(sortedArr() As Double, percentile As Double) As Double
    ' Calculate percentile from a sorted array
    ' Note: Array must be already sorted in ascending order
    
    Dim n As Long
    Dim index As Double
    Dim lowerIndex As Long, upperIndex As Long
    Dim lowerValue As Double, upperValue As Double
    Dim fraction As Double
    
    n = UBound(sortedArr)
    
    ' Calculate the position
    index = 1 + (n - 1) * percentile
    
    ' Get the lower and upper indices
    lowerIndex = Int(index)
    upperIndex = lowerIndex + 1
    
    ' Ensure indices are within bounds
    If lowerIndex < 1 Then lowerIndex = 1
    If upperIndex > n Then upperIndex = n
    
    ' Get the values
    lowerValue = sortedArr(lowerIndex)
    upperValue = sortedArr(upperIndex)
    
    ' Calculate the fractional part
    fraction = index - lowerIndex
    
    ' Interpolate
    GetPercentile = lowerValue + fraction * (upperValue - lowerValue)
End Function

Sub GenerateEquityChartData(wsOutput As Worksheet, dailyStats() As Double)
    ' Generate a chart data table in the specified worksheet
    '
    ' Parameters:
    '   wsOutput - Worksheet to output the data
    '   dailyStats - Array containing daily statistics
    
    Dim numDays As Long
    Dim i As Long
    
    ' Get dimensions
    numDays = UBound(dailyStats, 1)
    
    ' Clear existing content
    wsOutput.Cells.Clear
    
    ' Add headers
    wsOutput.Cells(1, 1).value = "Day"
    wsOutput.Cells(1, 2).value = "Average"
    wsOutput.Cells(1, 3).value = "Median"
    wsOutput.Cells(1, 4).value = "10th Percentile"
    wsOutput.Cells(1, 5).value = "25th Percentile"
    wsOutput.Cells(1, 6).value = "75th Percentile"
    wsOutput.Cells(1, 7).value = "90th Percentile"
    wsOutput.Cells(1, 8).value = "99th Percentile"
    wsOutput.Cells(1, 9).value = "Minimum"
    wsOutput.Cells(1, 10).value = "Maximum"
    
    ' Format headers
    wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(1, 10)).Font.Bold = True
    
    ' Add data
    For i = 0 To numDays
        wsOutput.Cells(i + 2, 1).value = i
        wsOutput.Cells(i + 2, 2).value = dailyStats(i, 1)  ' Average
        wsOutput.Cells(i + 2, 3).value = dailyStats(i, 2)  ' Median
        wsOutput.Cells(i + 2, 4).value = dailyStats(i, 3)  ' 10th percentile
        wsOutput.Cells(i + 2, 5).value = dailyStats(i, 4)  ' 25th percentile
        wsOutput.Cells(i + 2, 6).value = dailyStats(i, 5)  ' 75th percentile
        wsOutput.Cells(i + 2, 7).value = dailyStats(i, 6)  ' 90th percentile
        wsOutput.Cells(i + 2, 8).value = dailyStats(i, 7)  ' 99th percentile
        wsOutput.Cells(i + 2, 9).value = dailyStats(i, 8)  ' Minimum
        wsOutput.Cells(i + 2, 10).value = dailyStats(i, 9) ' Maximum
    Next i
    
    ' Format numbers
    wsOutput.Range(wsOutput.Cells(2, 2), wsOutput.Cells(numDays + 2, 10)).NumberFormat = "$#,##0"
    
    ' Autofit columns
    wsOutput.Columns("A:J").AutoFit
End Sub

Sub CreateEquityChart(wsOutput As Worksheet, chartTitle As String, Optional includeBands As Boolean = True)
    ' Create a chart from the data in the specified worksheet
    '
    ' Parameters:
    '   wsOutput - Worksheet containing the data
    '   chartTitle - Title for the chart
    '   includeBands - Whether to include percentile bands (default: True)
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numRows As Long
    
    ' Determine number of data rows
    numRows = wsOutput.Cells(wsOutput.rows.count, 1).End(xlUp).row
    
    ' Delete any existing charts
    On Error Resume Next
    wsOutput.ChartObjects.Delete
    On Error GoTo 0
    
    ' Create chart object
    Set chartObj = wsOutput.ChartObjects.Add(left:=50, top:=50, Width:=800, Height:=400)
    Set cht = chartObj.chart
    
    ' Set chart type
    cht.ChartType = xlLine
    
    ' Add data series
    With cht
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Average"
        .SeriesCollection(1).values = wsOutput.Range(wsOutput.Cells(2, 2), wsOutput.Cells(numRows, 2))
        .SeriesCollection(1).XValues = wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(numRows, 1))
        .SeriesCollection(1).Format.line.Weight = 2.5
        .SeriesCollection(1).Format.line.ForeColor.RGB = RGB(0, 112, 192) ' Blue
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).name = "Median"
        .SeriesCollection(2).values = wsOutput.Range(wsOutput.Cells(2, 3), wsOutput.Cells(numRows, 3))
        .SeriesCollection(2).XValues = wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(numRows, 1))
        .SeriesCollection(2).Format.line.Weight = 2.5
        .SeriesCollection(2).Format.line.ForeColor.RGB = RGB(0, 176, 80) ' Green
        
        If includeBands Then
            ' Add percentile bands
            .SeriesCollection.NewSeries
            .SeriesCollection(3).name = "10th Percentile"
            .SeriesCollection(3).values = wsOutput.Range(wsOutput.Cells(2, 4), wsOutput.Cells(numRows, 4))
            .SeriesCollection(3).XValues = wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(numRows, 1))
            .SeriesCollection(3).Format.line.Weight = 1.5
            .SeriesCollection(3).Format.line.DashStyle = msoLineDash
            .SeriesCollection(3).Format.line.ForeColor.RGB = RGB(192, 0, 0) ' Dark Red
            
            .SeriesCollection.NewSeries
            .SeriesCollection(4).name = "90th Percentile"
            .SeriesCollection(4).values = wsOutput.Range(wsOutput.Cells(2, 7), wsOutput.Cells(numRows, 7))
            .SeriesCollection(4).XValues = wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(numRows, 1))
            .SeriesCollection(4).Format.line.Weight = 1.5
            .SeriesCollection(4).Format.line.DashStyle = msoLineDash
            .SeriesCollection(4).Format.line.ForeColor.RGB = RGB(0, 176, 240) ' Light Blue
            
            .SeriesCollection.NewSeries
            .SeriesCollection(5).name = "99th Percentile"
            .SeriesCollection(5).values = wsOutput.Range(wsOutput.Cells(2, 8), wsOutput.Cells(numRows, 8))
            .SeriesCollection(5).XValues = wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(numRows, 1))
            .SeriesCollection(5).Format.line.Weight = 1.5
            .SeriesCollection(5).Format.line.DashStyle = msoLineDashDot
            .SeriesCollection(5).Format.line.ForeColor.RGB = RGB(112, 48, 160) ' Purple
        End If
    End With
    
    ' Format chart
    With cht
        .HasTitle = True
        .chartTitle.text = chartTitle
        .chartTitle.Font.Size = 14
        .chartTitle.Font.Bold = True
        
        ' Add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Trading Day"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 12
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Portfolio Equity ($)"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 12
        
        ' Format axes
        .Axes(xlCategory).MajorGridlines.Format.line.Visible = msoFalse
        .Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = RGB(200, 200, 200)
        .Axes(xlValue).MajorGridlines.Format.line.Transparency = 0.5
        
        ' Format value axis to show currency
        .Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
        
        ' Add horizontal reference line at starting equity (first data point)
        Dim startingEquity As Double
        startingEquity = wsOutput.Cells(2, 2).value
        
        .SeriesCollection.NewSeries
        .SeriesCollection(.SeriesCollection.count).name = "Starting Equity"
        .SeriesCollection(.SeriesCollection.count).ChartType = xlLine
        .SeriesCollection(.SeriesCollection.count).values = Array(startingEquity, startingEquity)
        .SeriesCollection(.SeriesCollection.count).XValues = Array(0, numRows - 1)
        .SeriesCollection(.SeriesCollection.count).Format.line.ForeColor.RGB = RGB(192, 0, 0)
        .SeriesCollection(.SeriesCollection.count).Format.line.DashStyle = msoLineDashDotDot
        .SeriesCollection(.SeriesCollection.count).Format.line.Weight = 1.25
        
        ' Add legend
        .HasLegend = True
        .Legend.position = xlLegendPositionBottom
        .Legend.Font.Size = 10
        
        ' Add data labels to the last point of each main series
        .SeriesCollection(1).Points(.SeriesCollection(1).Points.count).HasDataLabel = True
        .SeriesCollection(1).Points(.SeriesCollection(1).Points.count).DataLabel.ShowValue = True
        .SeriesCollection(1).Points(.SeriesCollection(1).Points.count).DataLabel.Font.Size = 9
        .SeriesCollection(1).Points(.SeriesCollection(1).Points.count).DataLabel.Font.Bold = True
        .SeriesCollection(1).Points(.SeriesCollection(1).Points.count).DataLabel.Format.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(0, 112, 192)
        
        .SeriesCollection(2).Points(.SeriesCollection(2).Points.count).HasDataLabel = True
        .SeriesCollection(2).Points(.SeriesCollection(2).Points.count).DataLabel.ShowValue = True
        .SeriesCollection(2).Points(.SeriesCollection(2).Points.count).DataLabel.Font.Size = 9
        .SeriesCollection(2).Points(.SeriesCollection(2).Points.count).DataLabel.Font.Bold = True
        .SeriesCollection(2).Points(.SeriesCollection(2).Points.count).DataLabel.Format.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(0, 176, 80)
    End With
End Sub

Sub CreateEquityDistributionChart(wsOutput As Worksheet, dailyEquityTracking() As Double, tradingDay As Long, chartTitle As String)
    ' Create a histogram chart showing the distribution of equity values for a specific trading day
    '
    ' Parameters:
    '   wsOutput - Worksheet to output the chart
    '   dailyEquityTracking - Array containing daily equity values (scenarios x days)
    '   tradingDay - The specific trading day to show distribution for
    '   chartTitle - Title for the chart
    
    Dim chartObj As ChartObject
    Dim cht As chart
    Dim numScenarios As Long
    Dim i As Long, j As Long, bin As Long
    Dim minValue As Double, maxValue As Double, binWidth As Double
    Dim numBins As Long
    Dim counts() As Long
    Dim binEdges() As Double
    
    ' Get dimensions
    numScenarios = UBound(dailyEquityTracking, 1)
    
    ' Clear existing content in the area where we'll put the histogram data
    wsOutput.Range("M1:O100").Clear
    
    ' Find min and max values for the specified day
    minValue = dailyEquityTracking(1, tradingDay)
    maxValue = minValue
    
    For i = 1 To numScenarios
        If dailyEquityTracking(i, tradingDay) < minValue Then
            minValue = dailyEquityTracking(i, tradingDay)
        ElseIf dailyEquityTracking(i, tradingDay) > maxValue Then
            maxValue = dailyEquityTracking(i, tradingDay)
        End If
    Next i
    
    ' Create bins for histogram (20 bins)
    numBins = 20
    binWidth = (maxValue - minValue) / numBins
    
    If binWidth = 0 Then
        ' All values are the same, create a single bin
        numBins = 1
        binWidth = 1
    End If
    
    ReDim counts(1 To numBins)
    ReDim binEdges(0 To numBins)
    
    ' Initialize bin edges
    For i = 0 To numBins
        binEdges(i) = minValue + i * binWidth
    Next i
    
    ' Count values in each bin
    For i = 1 To numScenarios
        Dim value As Double
        value = dailyEquityTracking(i, tradingDay)
        
        ' Find the bin for this value
        For bin = 1 To numBins
            If value < binEdges(bin) Or bin = numBins Then
                counts(bin) = counts(bin) + 1
                Exit For
            End If
        Next bin
    Next i
    
    ' Output histogram data
    wsOutput.Cells(1, 13).value = "Bin"
    wsOutput.Cells(1, 14).value = "Value"
    wsOutput.Cells(1, 15).value = "Count"
    
    For i = 1 To numBins
        wsOutput.Cells(i + 1, 13).value = i
        wsOutput.Cells(i + 1, 14).value = binEdges(i - 1)
        wsOutput.Cells(i + 1, 15).value = counts(i)
    Next i
    
    ' Delete any existing charts
    On Error Resume Next
    For Each chartObj In wsOutput.ChartObjects
        If InStr(chartObj.chart.chartTitle.text, "Distribution") > 0 Then
            chartObj.Delete
        End If
    Next chartObj
    On Error GoTo 0
    
    ' Create chart object
    Set chartObj = wsOutput.ChartObjects.Add(left:=50, top:=475, Width:=800, Height:=350)
    Set cht = chartObj.chart
    
    ' Set chart type and data
    With cht
        .ChartType = xlColumnClustered
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Equity Distribution"
        .SeriesCollection(1).values = wsOutput.Range(wsOutput.Cells(2, 15), wsOutput.Cells(numBins + 1, 15))
        .SeriesCollection(1).XValues = wsOutput.Range(wsOutput.Cells(2, 14), wsOutput.Cells(numBins + 1, 14))
        .SeriesCollection(1).Format.fill.ForeColor.RGB = RGB(91, 155, 213)
        
        ' Format chart
        .HasTitle = True
        .chartTitle.text = chartTitle & " (Day " & tradingDay & ")"
        .chartTitle.Font.Size = 14
        .chartTitle.Font.Bold = True
        
        ' Add axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.text = "Equity Value ($)"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 12
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.text = "Number of Scenarios"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 12
        
        ' Format axes
        .Axes(xlCategory).TickLabels.NumberFormat = "$#,##0"
        
        ' Format plot area
        .PlotArea.Format.fill.ForeColor.RGB = RGB(240, 240, 240)
        
        ' No legend needed for histogram
        .HasLegend = False
    End With
End Sub

Sub CreateMainMonteCarloResults(wsOutput As Worksheet, dailyEquityTracking() As Double, results As Variant, startingEquity As Double)
    ' Create a summary sheet with key metrics and multiple charts
    '
    ' Parameters:
    '   wsOutput - Worksheet to output the results
    '   dailyEquityTracking - Array containing daily equity values (scenarios x days)
    '   results - Standard Monte Carlo results array
    '   startingEquity - Initial portfolio equity
    
    Dim numScenarios As Long, numDays As Long
    Dim dailyStats() As Double
    Dim avgFinalReturn As Double, medianFinalReturn As Double
    Dim avgMaxDrawdown As Double, medianMaxDrawdown As Double
    Dim riskOfRuin As Double
    Dim i As Long
    
    ' Get dimensions
    numScenarios = UBound(dailyEquityTracking, 1)
    numDays = UBound(dailyEquityTracking, 2)
    
    ' Clear the worksheet
    wsOutput.Cells.Clear
    
    ' Add title
    wsOutput.Cells(1, 1).value = "Monte Carlo Simulation Results"
    wsOutput.Cells(1, 1).Font.Size = 16
    wsOutput.Cells(1, 1).Font.Bold = True
    
    ' Calculate key metrics
    avgFinalReturn = 0
    medianFinalReturn = Application.WorksheetFunction.Median(Application.index(results, 0, 3))
    
    avgMaxDrawdown = Application.WorksheetFunction.Average(Application.index(results, 0, 5))
    medianMaxDrawdown = Application.WorksheetFunction.Median(Application.index(results, 0, 5))
    
    riskOfRuin = 0
    For i = 1 To numScenarios
        If results(i, 6) = 1 Then
            riskOfRuin = riskOfRuin + 1
        End If
    Next i
    riskOfRuin = riskOfRuin / numScenarios
    
    ' Output key metrics
    wsOutput.Cells(3, 1).value = "Key Metrics:"
    wsOutput.Cells(3, 1).Font.Bold = True
    
    wsOutput.Cells(4, 1).value = "Number of Scenarios:"
    wsOutput.Cells(4, 2).value = numScenarios
    
    wsOutput.Cells(5, 1).value = "Trading Days Simulated:"
    wsOutput.Cells(5, 2).value = numDays
    
    wsOutput.Cells(6, 1).value = "Starting Equity:"
    wsOutput.Cells(6, 2).value = startingEquity
    wsOutput.Cells(6, 2).NumberFormat = "$#,##0"
    
    wsOutput.Cells(7, 1).value = "Median Final Return:"
    wsOutput.Cells(7, 2).value = medianFinalReturn
    wsOutput.Cells(7, 2).NumberFormat = "0.00%"
    
    wsOutput.Cells(8, 1).value = "Median Maximum Drawdown:"
    wsOutput.Cells(8, 2).value = medianMaxDrawdown
    wsOutput.Cells(8, 2).NumberFormat = "0.00%"
    
    wsOutput.Cells(9, 1).value = "Risk of Ruin:"
    wsOutput.Cells(9, 2).value = riskOfRuin
    wsOutput.Cells(9, 2).NumberFormat = "0.00%"
    
    ' Format metrics area
    wsOutput.Range("A3:B9").Borders.LineStyle = xlContinuous
    wsOutput.Range("A3:A9").Font.Bold = True
    wsOutput.Range("A3:B3").Interior.Color = RGB(217, 217, 217)
    
    ' Calculate daily statistics
    Call CalculateDailyEquityStatistics(dailyEquityTracking, dailyStats)
    
    ' Output data for charts
    Call GenerateEquityChartData(wsOutput, dailyStats)
    
    ' Create equity progression chart
    Call CreateEquityChart(wsOutput, "Portfolio Equity Progression Over Time", True)
    
    ' Create distribution chart for final day
    Call CreateEquityDistributionChart(wsOutput, dailyEquityTracking, numDays, "Final Equity Distribution")
    
    ' Create distribution chart for half-way point
    Call CreateEquityDistributionChart(wsOutput, dailyEquityTracking, Int(numDays / 2), "Mid-Year Equity Distribution")
    
    ' Autofit columns
    wsOutput.Columns("A:O").AutoFit
End Sub
    

Function ApplyRealisticExtremeValueHandling(TradePNL As Double, startingEquity As Double) As Double
    ' More sophisticated extreme value handling using a smooth transition function
    Dim extremeThresholdMultiple As Double
    Dim maxLossMultiple As Double
    Dim maxGainMultiple As Double
    
    ' Set thresholds
    extremeThresholdMultiple = 5    ' 5x starting equity is extreme
    maxLossMultiple = 10           ' Maximum possible loss
    maxGainMultiple = 10           ' Maximum possible gain
    
    ' Check if we're in extreme territory
    If Abs(TradePNL) <= extremeThresholdMultiple * startingEquity Then
        ' Not extreme, return as is
        ApplyRealisticExtremeValueHandling = TradePNL
    Else
        ' For extreme values, use a dampening function
        If TradePNL > 0 Then
            ' Positive extreme: apply dampening toward max gain
            Dim excessPositive As Double
            excessPositive = TradePNL - extremeThresholdMultiple * startingEquity
            
            ' Logarithmic dampening to approach but not exceed the limit
            ApplyRealisticExtremeValueHandling = extremeThresholdMultiple * startingEquity + _
                (maxGainMultiple - extremeThresholdMultiple) * startingEquity * _
                (1 - Exp(-excessPositive / startingEquity))
        Else
            ' Negative extreme: apply dampening toward max loss
            Dim excessNegative As Double
            excessNegative = Abs(TradePNL) - extremeThresholdMultiple * startingEquity
            
            ' Logarithmic dampening to approach but not exceed the limit
            ApplyRealisticExtremeValueHandling = -extremeThresholdMultiple * startingEquity - _
                (maxLossMultiple - extremeThresholdMultiple) * startingEquity * _
                (1 - Exp(-excessNegative / startingEquity))
        End If
    End If
End Function

' Improved block bootstrap function
Function SampleBlockBootstrap(pnlResults As Variant, blockSize As Long, numDays As Long, numStrategies As Long) As Variant
    ' Improved block bootstrap with better overlap handling
    Dim bootstrappedData() As Variant
    Dim i As Long, j As Long, block As Long
    Dim blockStart As Long, originalDays As Long
    Dim transitionProb As Double
    Dim currentBlock As Long
    
    originalDays = UBound(pnlResults, 1)
    ReDim bootstrappedData(1 To numDays, 1 To numStrategies)
    
    ' Probability of starting a new block
    transitionProb = 1 / blockSize
    
    ' Choose initial block
    currentBlock = Int((originalDays - blockSize + 1) * Rnd()) + 1
    
    ' Generate bootstrap sample
    i = 1
    While i <= numDays
        ' Position in current block
        Dim posInBlock As Long
        posInBlock = 0
        
        ' Copy from current block until transition or end of required days
        Do While posInBlock < blockSize And i <= numDays
            ' Copy data from this position
            For j = 1 To numStrategies
                ' Wrap around if we reach the end of original data
                Dim sourceIdx As Long
                sourceIdx = ((currentBlock + posInBlock - 1) Mod originalDays) + 1
                bootstrappedData(i, j) = pnlResults(sourceIdx, j)
            Next j
            
            ' Move to next day
            posInBlock = posInBlock + 1
            i = i + 1
            
            ' Possibly start a new block with probability transitionProb
            If i <= numDays And Rnd() < transitionProb Then
                ' Choose a new block
                currentBlock = Int((originalDays - blockSize + 1) * Rnd()) + 1
                Exit Do
            End If
        Loop
        
        ' If we completed the block, choose next block
        If posInBlock >= blockSize And i <= numDays Then
            ' Choose a new block with overlap considerations
            ' Make transitions more likely to nearby blocks for smoother changes
            Dim prevBlock As Long
            prevBlock = currentBlock
            
            ' 50% chance of choosing a nearby block for better continuity
            If Rnd() < 0.5 Then
                ' Choose a nearby block (+/- 10 days from end of current block)
                Dim rangeStart As Long, rangeEnd As Long
                rangeStart = Application.WorksheetFunction.Max(1, prevBlock + blockSize - 10)
                rangeEnd = Application.WorksheetFunction.Min(originalDays - blockSize + 1, prevBlock + blockSize + 10)
                
                currentBlock = rangeStart + Int((rangeEnd - rangeStart + 1) * Rnd())
            Else
                ' Choose a completely random block
                currentBlock = Int((originalDays - blockSize + 1) * Rnd()) + 1
            End If
        End If
    Wend
    
    SampleBlockBootstrap = bootstrappedData
End Function


' Additional function to generate the enhanced summary report
Sub GenerateEnhancedMonteCarloSummary(wsPortfolioMC As Worksheet, results As Variant, _
                                startingEquity As Double, requiredMargin As Double, _
                                numScenarios As Long, currentRiskOfRuin As Double, _
                                factorModelPercentage As Double, enableCrisisMode As Boolean, _
                                crisisCorrelationIncrease As Double, crisisThreshold As Double, _
                                blockSize As Long, numFactors As Long)
    
    Dim avgProfit As Double, avgReturn As Double, avgMaxDrawdown As Double, avgReturnToDrawdown As Double
    Dim medianReturn As Double, medianProfit As Double, medianDrawdown As Double, medianReturnToDrawdown As Double
    Dim summaryRow As Long, maxProfit As Double, binWidth As Double
    Dim i As Long
    Dim returnArray() As Variant, profitArray() As Variant, drawdownArray() As Variant
    Dim avgCrisisDays As Double, medianCrisisDays As Double
    
    ' Calculate summary statistics
    medianReturn = WorksheetFunction.Median(Application.index(results, 0, 3))
    medianDrawdown = WorksheetFunction.Median(Application.index(results, 0, 5))
    medianProfit = WorksheetFunction.Median(Application.index(results, 0, 2))
    medianReturnToDrawdown = WorksheetFunction.Median(Application.index(results, 0, 4))
    
    avgReturn = WorksheetFunction.Average(Application.index(results, 0, 3))
    avgMaxDrawdown = WorksheetFunction.Average(Application.index(results, 0, 5))
    avgProfit = WorksheetFunction.Average(Application.index(results, 0, 2))
    avgReturnToDrawdown = WorksheetFunction.Average(Application.index(results, 0, 4))
    
    ' Crisis-specific statistics if enabled
    If enableCrisisMode Then
        avgCrisisDays = WorksheetFunction.Average(Application.index(results, 0, 7))
        medianCrisisDays = WorksheetFunction.Median(Application.index(results, 0, 7))
    End If
    
    ' Output summary metrics
    summaryRow = 3
    With wsPortfolioMC
        ' Title Formatting
        .Cells(1, 1).value = "Enhanced Monte Carlo Simulation Summary"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(0, 102, 204) ' Light blue background
        .Cells(1, 1).Font.Color = RGB(255, 255, 255) ' White font
    
        ' Header row formatting
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 16, 1)).Interior.Color = RGB(224, 224, 224)
        .Range(.Cells(summaryRow, 1), .Cells(summaryRow + 16, 1)).Font.Bold = True
    
        ' Populate summary table
        .Cells(summaryRow, 1).value = "Starting Capital"
        .Cells(summaryRow, 2).value = startingEquity
        .Cells(summaryRow, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 1, 1).value = "Minimum Portfolio Value"
        .Cells(summaryRow + 1, 2).value = requiredMargin * startingEquity
        .Cells(summaryRow + 1, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 2, 1).value = "Simulation Method"
        .Cells(summaryRow + 2, 2).value = "Combined (" & Format(factorModelPercentage * 100, "0") & "% Factor Model, " & _
                                           Format((1 - factorModelPercentage) * 100, "0") & "% Block Bootstrap)"
        
        .Cells(summaryRow + 3, 1).value = "Number of Factors"
        .Cells(summaryRow + 3, 2).value = numFactors
        
        .Cells(summaryRow + 4, 1).value = "Block Size"
        .Cells(summaryRow + 4, 2).value = blockSize & " days"
        
        If enableCrisisMode Then
            .Cells(summaryRow + 5, 1).value = "Crisis Mode"
            .Cells(summaryRow + 5, 2).value = "Enabled (+" & Format(crisisCorrelationIncrease, "0.0") & _
                                              " correlation increase, " & Format(crisisThreshold * 100, "0.0") & "% threshold)"
            
            .Cells(summaryRow + 6, 1).value = "Avg. Crisis Days"
            .Cells(summaryRow + 6, 2).value = avgCrisisDays
            .Cells(summaryRow + 6, 2).NumberFormat = "0.0"
            
            .Cells(summaryRow + 7, 1).value = "Median Crisis Days"
            .Cells(summaryRow + 7, 2).value = medianCrisisDays
            .Cells(summaryRow + 7, 2).NumberFormat = "0.0"
            
            ' Shift subsequent rows
            summaryRow = summaryRow + 2
        Else
            .Cells(summaryRow + 5, 1).value = "Crisis Mode"
            .Cells(summaryRow + 5, 2).value = "Disabled"
        End If
        
        .Cells(summaryRow + 6, 1).value = "Average Profit ($)"
        .Cells(summaryRow + 6, 2).value = avgProfit
        .Cells(summaryRow + 6, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 7, 1).value = "Median Profit ($)"
        .Cells(summaryRow + 7, 2).value = medianProfit
        .Cells(summaryRow + 7, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 8, 1).value = "Average Return (%)"
        .Cells(summaryRow + 8, 2).value = avgReturn
        .Cells(summaryRow + 8, 2).NumberFormat = "0.0%"
        
        .Cells(summaryRow + 9, 1).value = "Median Return (%)"
        .Cells(summaryRow + 9, 2).value = medianReturn
        .Cells(summaryRow + 9, 2).NumberFormat = "0.0%"
        
        .Cells(summaryRow + 10, 1).value = "Average Max Drawdown ($)"
        .Cells(summaryRow + 10, 2).value = avgMaxDrawdown * startingEquity
        .Cells(summaryRow + 10, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 11, 1).value = "Median Max Drawdown ($)"
        .Cells(summaryRow + 11, 2).value = medianDrawdown * startingEquity
        .Cells(summaryRow + 11, 2).NumberFormat = "$#,##0"
        
        .Cells(summaryRow + 12, 1).value = "Average Max Drawdown (%)"
        .Cells(summaryRow + 12, 2).value = avgMaxDrawdown
        .Cells(summaryRow + 12, 2).NumberFormat = "0.0%"
        
        .Cells(summaryRow + 13, 1).value = "Median Max Drawdown (%)"
        .Cells(summaryRow + 13, 2).value = medianDrawdown
        .Cells(summaryRow + 13, 2).NumberFormat = "0.0%"
        
        .Cells(summaryRow + 14, 1).value = "Average Return to Drawdown"
        .Cells(summaryRow + 14, 2).value = avgReturnToDrawdown
        .Cells(summaryRow + 14, 2).NumberFormat = "0.0"
        
        .Cells(summaryRow + 15, 1).value = "Median Return to Drawdown"
        .Cells(summaryRow + 15, 2).value = medianReturnToDrawdown
        .Cells(summaryRow + 15, 2).NumberFormat = "0.0"
        
        .Cells(summaryRow + 16, 1).value = "Risk of Ruin"
        .Cells(summaryRow + 16, 2).value = currentRiskOfRuin
        .Cells(summaryRow + 16, 2).NumberFormat = "0.0%"
    
        ' Apply borders
        With .Range(.Cells(summaryRow - 2, 1), .Cells(summaryRow + 16, 2)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    
        ' Autofit columns
        .Columns("A:Z").AutoFit
    End With

    ' Determine maximum profit value
    maxProfit = WorksheetFunction.Max(Application.index(results, 0, 2))

    ' Calculate bin width as one-tenth of maxprofit
    Dim magnitude As Double
    If maxProfit > 0 Then
        magnitude = 10 ^ Int(Log(maxProfit / 10) / Log(10)) ' Find nearest power of 10
        binWidth = Application.WorksheetFunction.Round(maxProfit / 10, -Int(Log(magnitude) / Log(10)))
    Else
        binWidth = 1
    End If
    
    ' Convert to 1D arrays
    ReDim returnArray(1 To UBound(results, 1))
    ReDim profitArray(1 To UBound(results, 1))
    ReDim drawdownArray(1 To UBound(results, 1))
    
    For i = 1 To UBound(results, 1)
        returnArray(i) = results(i, 3)
        drawdownArray(i) = results(i, 5)
        profitArray(i) = results(i, 2)
    Next i
    
    ' Generate histograms
    CreateHistogram wsPortfolioMC, returnArray, "Return Histogram", 10, 14, summaryRow, 2, 0.05
    CreateHistogram wsPortfolioMC, drawdownArray, "Max Drawdown Histogram", 10, 17, summaryRow, 16, 0.05
    CreateHistogram wsPortfolioMC, profitArray, "Profit Histogram", 10, 20, summaryRow, 30, binWidth
    
    ' Add crisis days histogram if crisis mode is enabled
    If enableCrisisMode Then
        Dim crisisDaysArray() As Variant
        ReDim crisisDaysArray(1 To UBound(results, 1))
        
        For i = 1 To UBound(results, 1)
            crisisDaysArray(i) = results(i, 7)
        Next i
        
        CreateHistogram wsPortfolioMC, crisisDaysArray, "Crisis Days Histogram", 10, 23, summaryRow, 44, 5
    End If
End Sub


