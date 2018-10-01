VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "initSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private initSheet As Long
Private initFirstRow As Long
Private initFirstCol As Long
Private initLastRow As Long
Private initLastCol As Long
Private initActiveRow As Long
Private initActiveCol As Long

Public Property Get getInitSheet() As Long
    getInitSheet = initSheet
End Property

Public Property Get getInitFirstRow() As Long
    getInitFirstRow = initFirstRow
End Property

Public Property Get getInitFirstCol() As Long
    getInitFirstCol = initFirstCol
End Property

Public Property Get getInitLastRow() As Long
    getInitLastRow = initLastRow
End Property

Public Property Get getInitLastCol() As Long
    getInitLastCol = initLastCol
End Property

Public Property Get getInitActiveRow() As Long
    getInitActiveRow = initActiveRow
End Property

Public Property Get getInitActiveCol() As Long
    getInitActiveCol = initActiveCol
End Property

'初期選択位置を取得
Private Sub Class_Initialize()
    'シート
    initSheet = ActiveSheet.Index
    With Selection
        '開始位置
        initFirstRow = .Row
        initFirstCol = .Column
        '終了位置
        initLastRow = .Row + .Rows.Count - 1
        initLastCol = .Column + .Columns.Count - 1
    End With
    With ActiveCell
        'アクティブセル
        initActiveRow = .Row
        initActiveCol = .Column
    End With
End Sub

Private Sub Class_Terminate()

End Sub

'選択位置を初期に戻す
Public Sub setInitSelection()
    Worksheets(initSheet).Activate
    Range(Cells(initFirstRow, initFirstCol), Cells(initLastRow, initLastCol)).Select
    Cells(initActiveRow, initActiveCol).Activate
End Sub

