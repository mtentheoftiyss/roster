Attribute VB_Name = "setHoliday"
Option Explicit

'休日の行の背景色を変更
Public Sub setHoliday(ByVal target As Range)
    '初期処理
    Call commonUtil.startProcess
    
    '年月の変更時のみ実行
    If Intersect(target, Range("年月")) Is Nothing Then
        GoTo endProcess
    End If
    
    '年月どちらも入力時のみ実行
    If WorksheetFunction.CountA(Range("年月")) < 2 Then
        GoTo endProcess
    End If
    
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    '最終列の特定
    Dim lastCol: lastCol = Range("備考")(Range("備考").Count).Column
    
    '最終行の特定
    Dim lastRow As Long: lastRow = Range("開始・終了時間リスト")(Range("開始・終了時間リスト").Count).Row
    
    '先頭行
    Dim firstRow As Long: firstRow = Range("基点").Row
    '最初に諸々クリアする
    With Range(Cells(firstRow, 1), Cells(lastRow, lastCol))
        '塗りつぶしの色
        .Interior.ColorIndex = xlNone
        With .Font
            'フォントの色
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            '太字
            .Bold = False
            '斜体
            .Italic = False
            '下線
            .Underline = xlUnderlineStyleNone
            '取り消し線
            .Strikethrough = False
        End With
    End With
    
    '先頭行から最終行までループ
    Dim i As Long
    Dim dayCnt As Integer: dayCnt = 0
    Dim dayFlg As Boolean: dayFlg = False
    Dim baseCol As Long: baseCol = Range("基点").Column + 1
    For i = firstRow To lastRow
        Dim cellFormula As String: cellFormula = Cells(i, baseCol).Formula
        Dim cellValue As String: cellValue = Cells(i, baseCol).Value
        Dim cellObject As Range: Set cellObject = Cells(i, baseCol)
        Dim holidayFlg As Boolean: holidayFlg = False
        Dim holidayFind As Range
        Dim holidayRow As Long
        Dim holidayNameCell As Range
        
        '計算式が空の場合、上行のものを使用する
        Dim j As Long: j = i
        Do While cellFormula = ""
            j = j - 1
            cellFormula = Cells(j, baseCol).Formula
            cellValue = Cells(j, baseCol).Value
            Set cellObject = Cells(j, baseCol)
            dayFlg = True
        Loop
        If cellValue <> "" Then
            holidayFlg = isHoliday(cellObject)
        Else
            '31日に満たない月の31日
            holidayFlg = True
        End If
        '休日は塗りつぶし
        If holidayFlg Then
            Range(Cells(i, 1), Cells(i, lastCol)).Interior.ColorIndex = 16
        Else
            If Not dayFlg Then
                dayCnt = dayCnt + 1
            End If
        End If
        '祝日を備考に入力
        If holidayFlg Then
            If Not dayFlg Then
                If cellValue <> "" Then
                    Set holidayFind = Range("祝日リスト").Find(What:=CDate(cellValue), LookIn:=xlValues, lookAt:=xlWhole)
                    If Not holidayFind Is Nothing Then
                        holidayRow = holidayFind.Row
                        Set holidayNameCell = Range("祝日リスト").Resize(1, 1).Offset(holidayRow - Range("祝日リスト").Row, -1)
                        Cells(cellObject.Row, Range("備考").Column).Value = holidayNameCell.Value
                    End If
                End If
            End If
        End If
        dayFlg = False
    Next
    
    '日数設定
    Range("日数").Value = dayCnt
    
    '選択位置を初期に戻す
    initSelection.setInitSelection
    
endProcess:
    '終了処理
    Call commonUtil.endProcess
End Sub

'休日判定
Private Function isHoliday(ByVal targetCell As Range)
    Dim returnValue As Boolean: returnValue = False
    
    '祝日
    If WorksheetFunction.CountIf(Range("祝日リスト"), targetCell) Then
        returnValue = True
    End If
    '日曜日
    If Weekday(targetCell.Value) = 1 Then
        returnValue = True
    End If
    '土曜日
    If Weekday(targetCell.Value) = 7 Then
        returnValue = True
    End If
    isHoliday = returnValue
End Function




