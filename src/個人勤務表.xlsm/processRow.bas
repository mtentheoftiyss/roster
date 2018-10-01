Attribute VB_Name = "processRow"
Option Explicit

'行追加
Public Sub addRow()
    '初期処理
    Call commonUtil.startProcess
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    Dim startCol As Long: startCol = Range("作業場所").Column
    Dim endCol As Long: endCol = Range("備考").Column
    Dim pasteRow As Long: pasteRow = targetCell.Row
    Dim copyRow As Long: copyRow = targetCell.Row + 1

    '行挿入
    Rows(targetCell.Row).Insert Shift:=xlDown
    '上に行挿入されるので対象行のデータをコピペ
    Range(Cells(copyRow, startCol), Cells(copyRow, endCol)).Copy
    With Range(Cells(pasteRow, startCol), Cells(pasteRow, endCol))
        '値貼り付け
        .PasteSpecial _
         Paste:=xlPasteValues, _
         Operation:=xlNone, _
         SkipBlanks:=False, _
         Transpose:=False
        '上下罫線
        .Borders(xlEdgeTop).LineStyle = xlDot
        .Borders(xlEdgeBottom).LineStyle = xlDot
    End With
    Application.CutCopyMode = False
    'クリア
    Range(Cells(copyRow, startCol), Cells(copyRow, endCol)).ClearContents

    '行追加ボタン追加
    With ActiveSheet.Buttons.Add(ActiveSheet.Shapes(Application.Caller).Left, _
                                 ActiveSheet.Shapes(Application.Caller).Top - targetCell.Height, _
                                 ActiveSheet.Shapes(Application.Caller).Width, _
                                 ActiveSheet.Shapes(Application.Caller).Height)
        .OnAction = "addRow"
        .Characters.Text = "＋"
        .Font.Name = "ＭＳ ゴシック"
        .Font.Size = 10
    End With

    '行削除ボタン追加
    With ActiveSheet.Buttons.Add(ActiveSheet.Shapes(Application.Caller).Left + targetCell.Width, _
                                 ActiveSheet.Shapes(Application.Caller).Top - targetCell.Height, _
                                 ActiveSheet.Shapes(Application.Caller).Width, _
                                 ActiveSheet.Shapes(Application.Caller).Height)
        .OnAction = "delRow"
        .Characters.Text = "－"
        .Font.Name = "ＭＳ ゴシック"
        .Font.Size = 10
    End With

    '選択位置を初期に戻す
    initSelection.setInitSelection
    '終了処理
    Call commonUtil.endProcess
End Sub

'行削除
Public Sub delRow()
    '初期処理
    Call commonUtil.startProcess
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    Dim startCol As Long: startCol = Range("作業場所").Column
    Dim endCol As Long: endCol = Range("備考").Column
    Dim copyRow As Long: copyRow = targetCell.Row
    Dim pasteRow As Long: pasteRow = targetCell.Row + 1
    
    'ボタンを押した行の入力がある、かつ、下の行の入力がない場合、下の行にコピー
    If WorksheetFunction.CountA(Range(Cells(copyRow, startCol), Cells(copyRow, endCol))) <> 0 Then
        If WorksheetFunction.CountA(Range(Cells(pasteRow, startCol), Cells(pasteRow, endCol))) = 0 Then
            '値貼り付け
            Range(Cells(copyRow, startCol), Cells(copyRow, endCol)).Copy
            Range(Cells(pasteRow, startCol), Cells(pasteRow, endCol)).PasteSpecial _
                Paste:=xlPasteValues, _
                Operation:=xlNone, _
                SkipBlanks:=False, _
                Transpose:=False
        End If
    End If
    
    
    'ボタン削除
    '行追加ボタン削除
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        'たまに謎のドロップダウンが作られるのであったら削除
        If shp.FormControlType = xlDropDown Then
            shp.Delete
        '行削除ボタンの左隣のセル
        ElseIf Not (Intersect(shp.TopLeftCell, targetCell.Offset(0, -1)) Is Nothing) Then
            shp.Delete
        End If
    Next
    
    '行削除ボタン削除
    ActiveSheet.Shapes(Application.Caller).Delete
    '行削除
    Rows(targetCell.Row).Delete Shift:=xlUp

    '選択位置を初期に戻す
    initSelection.setInitSelection
    '終了処理
    Call commonUtil.endProcess
End Sub
