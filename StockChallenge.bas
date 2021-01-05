Attribute VB_Name = "StockChallenge"
Option Explicit
'Declare/Initialize constants
'Data columns
Const TICKER_COL = 1
Const DATE_COL = 2
Const OPEN_COL = 3
Const HIGH_COL = 4
Const LOW_COL = 5
Const CLOSE_COL = 6
Const VOL_COL = 7
'Colors
Const RED_COLOR = 3
Const GREEN_COLOR = 4
'Blank String
Const BLANK_STRING = ""
'Popup messages to show progress
Const NUMBER_POPUP_MESSAGES = 15
Const SECONDS_TO_WAIT = 1
'Results table columns
Const RESULTS_TABLE_TICKER_COL = 9
Const RESULTS_TABLE_YEARLY_CHANGE_COL = 10
Const RESULTS_TABLE_PERCENT_CHANGE_COL = 11
Const RESULTS_TABLE_TOTAL_STOCK_VALUE_COL = 12
'Bonus table columns
Const BONUS_TABLE_LABELS_COL = 15
Const BONUS_TABLE_TICKER_COL = 16
Const BONUS_TABLE_VALUE_COL = 17
    
Public Function ConvertYYYYMMDDtoDate(YYYYMMDDStr As String) As Date
    'We slice and dice the number, feeding it all into DateSerial
    ConvertYYYYMMDDtoDate = DateSerial(Left(YYYYMMDDStr, 4), Mid(YYYYMMDDStr, 5, 2), Right(YYYYMMDDStr, 2))
End Function

Sub stockChallenge()
    'Declare variables
    Dim numberSheetsInt As Integer
    Dim pctComplDbl As Double
    Dim partialRowsCountDbl As Double
    Dim totalRowsCountDbl As Double
    Dim blockSizeDbl As Double
    Dim currentSheetInt As Integer
    Dim numberRowsLng As Long
    Dim currentResultsTableRowLng As Long
    Dim currentRowLng As Long
    Dim currentTickerStr As String
    Dim currentDateStr As String
    Dim currentDateDat As Date
    Dim currentOpenDbl As Double
    Dim currentCloseDbl  As Double
    Dim currentVolDbl  As Double
    Dim previuosTickerStr As String
    Dim minDateDat As Date
    Dim maxDateDat As Date
    Dim yearlyOpenDbl As Double
    Dim yearlyCloseDbl As Double
    Dim yearlyChangeDbl As Double
    Dim percentChangeDbl As Double
    Dim totalVolDbl As Double
    Dim greatestPercentIncreaseDbl As Double
    Dim leastPercentDecreaseDbl As Double
    Dim greatestTotalVolDbl As Double
    Dim greatestPercentIncreaseTickerStr As String
    Dim leastPercentDecreaseTickerStr As String
    Dim greatestTotalVolTickerStr As String
    Dim startTimeDbl As Double
    Dim elapsedTimeStr As String
    Dim elapsedTimeDbl As Double
    Dim estimatedTimeStr As String
    Dim msgStr As String
    
    'Remember time when macro starts
    startTimeDbl = Timer

    'Hourglass cursor ON
    Application.Cursor = xlWait

    'Get number of worksheets in the active workbook.
    numberSheetsInt = ActiveWorkbook.Worksheets.Count
    
    pctComplDbl = 0
    partialRowsCountDbl = 0
    totalRowsCountDbl = 0
    'Determine the total number of rows to be processed
    For currentSheetInt = 1 To numberSheetsInt
        numberRowsLng = ActiveWorkbook.Worksheets(currentSheetInt).Cells(ActiveWorkbook.Worksheets(currentSheetInt).Rows.Count, 2).End(xlUp).Row
        totalRowsCountDbl = totalRowsCountDbl + numberRowsLng - 2 + 1
    Next currentSheetInt
    blockSizeDbl = Round(totalRowsCountDbl / NUMBER_POPUP_MESSAGES, 0)

    'For each WorkSheet
    For currentSheetInt = 1 To numberSheetsInt
        'Find the last non-blank cell in column A, current worksheet
        With ActiveWorkbook.Worksheets(currentSheetInt)
            numberRowsLng = .Cells(.Rows.Count, 2).End(xlUp).Row
            .Columns("I:Q").EntireColumn.Delete
            'Initialize variables for each WorkSheet
            currentResultsTableRowLng = 0
            previuosTickerStr = BLANK_STRING
            greatestPercentIncreaseDbl = 0
            greatestPercentIncreaseTickerStr = BLANK_STRING
            leastPercentDecreaseDbl = 0
            leastPercentDecreaseTickerStr = BLANK_STRING
            greatestTotalVolDbl = 0
            greatestTotalVolTickerStr = BLANK_STRING
            percentChangeDbl = 0
            If numberRowsLng > 1 Then
                'If data available then place Headers for results table
                currentResultsTableRowLng = currentResultsTableRowLng + 1
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_TICKER_COL) = "Ticker"
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL) = "Yearly Change"
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_PERCENT_CHANGE_COL) = "Percent Change"
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_TOTAL_STOCK_VALUE_COL) = "Total Stock Volume"
            End If
            'For each row within current WorkSheet
            For currentRowLng = 2 To numberRowsLng
                'Get values of current row
                currentTickerStr = .Cells(currentRowLng, TICKER_COL)
                currentDateStr = Trim(Str(.Cells(currentRowLng, DATE_COL)))
                currentDateDat = ConvertYYYYMMDDtoDate(currentDateStr)
                currentOpenDbl = .Cells(currentRowLng, OPEN_COL)
                currentCloseDbl = .Cells(currentRowLng, CLOSE_COL)
                currentVolDbl = .Cells(currentRowLng, VOL_COL)
                If previuosTickerStr <> currentTickerStr Then
                'New Ticker
                    If previuosTickerStr <> BLANK_STRING Then
                    'Not the first Ticker in WorkSheet then print previousTicker
                        .Cells(currentResultsTableRowLng, RESULTS_TABLE_TICKER_COL) = previuosTickerStr
                        .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL) = yearlyChangeDbl
                        If yearlyChangeDbl >= 0 Then
                            .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL).Interior.ColorIndex = GREEN_COLOR ' 4 indicates Green Color
                        Else
                            .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL).Interior.ColorIndex = RED_COLOR ' 3 indicates Red Color
                        End If
                        .Cells(currentResultsTableRowLng, RESULTS_TABLE_PERCENT_CHANGE_COL) = percentChangeDbl
                        .Cells(currentResultsTableRowLng, RESULTS_TABLE_PERCENT_CHANGE_COL).NumberFormat = "0.00%"
                        .Cells(currentResultsTableRowLng, RESULTS_TABLE_TOTAL_STOCK_VALUE_COL) = totalVolDbl
                        If percentChangeDbl > greatestPercentIncreaseDbl Then
                            greatestPercentIncreaseDbl = percentChangeDbl
                            greatestPercentIncreaseTickerStr = previuosTickerStr
                        End If
                        If percentChangeDbl < leastPercentDecreaseDbl Then
                            leastPercentDecreaseDbl = percentChangeDbl
                            leastPercentDecreaseTickerStr = previuosTickerStr
                        End If
                        If totalVolDbl > greatestTotalVolDbl Then
                            greatestTotalVolDbl = totalVolDbl
                            greatestTotalVolTickerStr = previuosTickerStr
                        End If
                    End If
                    currentResultsTableRowLng = currentResultsTableRowLng + 1
                    minDateDat = currentDateDat
                    yearlyOpenDbl = currentOpenDbl
                    maxDateDat = currentDateDat
                    yearlyCloseDbl = currentCloseDbl
                    yearlyChangeDbl = yearlyCloseDbl - yearlyOpenDbl
                    If yearlyOpenDbl = 0 Then
                        percentChangeDbl = 0
                    Else
                        percentChangeDbl = yearlyChangeDbl / yearlyOpenDbl
                    End If
                    totalVolDbl = currentVolDbl
                Else
                'Still same Ticker
                    If currentDateDat < minDateDat Then
                        minDateDat = currentDateDat
                        yearlyOpenDbl = currentOpenDbl
                    End If
                    If currentDateDat > maxDateDat Then
                        maxDateDat = currentDateDat
                        yearlyCloseDbl = currentCloseDbl
                    End If
                    yearlyChangeDbl = yearlyCloseDbl - yearlyOpenDbl
                    If yearlyOpenDbl = 0 Then
                        percentChangeDbl = 0
                    Else
                        percentChangeDbl = yearlyChangeDbl / yearlyOpenDbl
                    End If
                    totalVolDbl = totalVolDbl + currentVolDbl
                End If
                previuosTickerStr = currentTickerStr
                partialRowsCountDbl = partialRowsCountDbl + 1
                pctComplDbl = pctComplDbl + (1 / totalRowsCountDbl)
                If partialRowsCountDbl Mod blockSizeDbl = 0 Then
                'Every blockSizeDbl rows it will show progress
                    'Determine how long the code has taken to run so far
                    elapsedTimeDbl = Timer - startTimeDbl
                    elapsedTimeStr = Format(elapsedTimeDbl / 86400, "hh:mm:ss")
                    If pctComplDbl > 0 Then
                        estimatedTimeStr = Format((1 - pctComplDbl) * elapsedTimeDbl / (pctComplDbl * 86400), "hh:mm:ss")
                    Else
                        estimatedTimeStr = "UNKNOWN"
                    End If
                    msgStr = "Current Sheet: [" + CStr(currentSheetInt) + "] " + .Name + _
                        " / " + Format(pctComplDbl, "0.00%") + " complete." + vbCrLf _
                        + "Elapsed time: " + elapsedTimeStr + ", Estimated time remaining: " + estimatedTimeStr
                    Call MessageBoxTimer(SECONDS_TO_WAIT, msgStr, "Progress")
                End If
            Next currentRowLng
            If previuosTickerStr <> BLANK_STRING Then
                'If last row of WorkSheet then print Ticker
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_TICKER_COL) = previuosTickerStr
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL) = yearlyChangeDbl
                If yearlyChangeDbl >= 0 Then
                    .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL).Interior.ColorIndex = GREEN_COLOR ' 4 indicates Green Color
                Else
                    .Cells(currentResultsTableRowLng, RESULTS_TABLE_YEARLY_CHANGE_COL).Interior.ColorIndex = RED_COLOR ' 3 indicates Red Color
                End If
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_PERCENT_CHANGE_COL) = percentChangeDbl
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_PERCENT_CHANGE_COL).NumberFormat = "0.00%"
                .Cells(currentResultsTableRowLng, RESULTS_TABLE_TOTAL_STOCK_VALUE_COL) = totalVolDbl
                If percentChangeDbl > greatestPercentIncreaseDbl Then
                    greatestPercentIncreaseDbl = percentChangeDbl
                    greatestPercentIncreaseTickerStr = previuosTickerStr
                End If
                If percentChangeDbl < leastPercentDecreaseDbl Then
                    leastPercentDecreaseDbl = percentChangeDbl
                    leastPercentDecreaseTickerStr = previuosTickerStr
                End If
                If totalVolDbl > greatestTotalVolDbl Then
                    greatestTotalVolDbl = totalVolDbl
                    greatestTotalVolTickerStr = previuosTickerStr
                End If
            End If
            'Bonus table
            'Headers
            .Cells(1, BONUS_TABLE_TICKER_COL) = "Ticker"
            .Cells(1, BONUS_TABLE_VALUE_COL) = "Value"
            .Cells(2, BONUS_TABLE_LABELS_COL) = "Greatest % Increase"
            .Cells(3, BONUS_TABLE_LABELS_COL) = "Greatest % Decrease"
            .Cells(4, BONUS_TABLE_LABELS_COL) = "Greatest Total Volume"
            'Values
            .Cells(2, BONUS_TABLE_TICKER_COL) = greatestPercentIncreaseTickerStr
            .Cells(3, BONUS_TABLE_TICKER_COL) = leastPercentDecreaseTickerStr
            .Cells(4, BONUS_TABLE_TICKER_COL) = greatestTotalVolTickerStr
            .Cells(2, BONUS_TABLE_VALUE_COL) = greatestPercentIncreaseDbl
            .Cells(2, BONUS_TABLE_VALUE_COL).NumberFormat = "0.00%"
            .Cells(3, BONUS_TABLE_VALUE_COL) = leastPercentDecreaseDbl
            .Cells(3, BONUS_TABLE_VALUE_COL).NumberFormat = "0.00%"
            .Cells(4, BONUS_TABLE_VALUE_COL) = greatestTotalVolDbl
            .Cells(4, BONUS_TABLE_VALUE_COL).NumberFormat = "##.0000E+0"
            .Columns("A:Q").EntireColumn.AutoFit
        End With
    Next currentSheetInt
    Call MessageBoxTimer(SECONDS_TO_WAIT, "Total Sheets: " + CStr(numberSheetsInt) + _
        " / " + Format(pctComplDbl, "0.00%") + " complete", "Done")

    'Notify user that the application has finished
    Application.Speech.Speak "I've finished"
    
    'Hourglass cursor OFF
    Application.Cursor = xlDefault
    
    'Determine how long the code took to run
    elapsedTimeStr = Format((Timer - startTimeDbl) / 86400, "hh:mm:ss")
    
    'Notify user elapsed time
    MsgBox "This code ran successfully in " & elapsedTimeStr & " minutes", vbInformation
End Sub

Sub MessageBoxTimer(PauseTimeSecondsInt As Integer, MessageStr As String, TitleStr As String)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 10 seconds
    PauseTimeSecondsInt = 1
    Select Case InfoBox.Popup(MessageStr, PauseTimeSecondsInt, TitleStr, 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub

