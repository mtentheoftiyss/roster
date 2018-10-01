Attribute VB_Name = "commonUtil"
Option Explicit

'初期処理
Public Sub startProcess()
    '自動計算を手動にする
    Application.Calculation = xlCalculationManual
    '画面の描画をOFFにする
    Application.ScreenUpdating = False
    'イベントの発生を抑止する
    Application.EnableEvents = False
    '確認メッセージを非表示にする
    Application.DisplayAlerts = False
End Sub

'終了処理
Public Sub endProcess()
    '確認メッセージを表示にする
    Application.DisplayAlerts = True
    'イベントの発生抑止を解除する
    Application.EnableEvents = True
    '画面の描画をONにする
    Application.ScreenUpdating = True
    '自動計算を自動にする
    Application.Calculation = xlCalculationAutomatic
End Sub

'行選択判定
Public Function isRowSelect(ByVal target As Range)
    Dim returnValue As Boolean: returnValue = False
    If target.Address = target.EntireRow.Address Then
        returnValue = True
    End If
    
    isRowSelect = returnValue
End Function

'列選択判定
Public Function isColumnSelect(ByVal target As Range)
    Dim returnValue As Boolean: returnValue = False
    If target.Address = target.EntireColumn.Address Then
        returnValue = True
    End If
    
    isColumnSelect = returnValue
End Function

'第何曜日の日付を取得
Public Function getWhatWeekDay(ByVal targetYear As Long, ByVal targetMonth As Long, ByVal targetWeek As Long, ByVal targetDay As Long)
    Dim targetDate As Date
    Dim i As Long
    Dim startDay As Long
    Dim endDay As Long
    startDay = (targetWeek - 1) * 7 + 1
    endDay = startDay + 7
    For i = startDay To endDay
        targetDate = DateSerial(targetYear, targetMonth, i)
        If Weekday(targetDate) = targetDay Then
            GoTo break
        End If
    Next
    
break:
    getWhatWeekDay = targetDate
    
End Function

'改行区切りでメッセージを作成
Public Function createMsg(ByVal msg As String, ByVal addMsg As String)
    If msg <> Empty Then
        msg = msg & vbCrLf
    End If
    
    createMsg = msg & addMsg
End Function

'取込ファイル指定
Public Sub setInputFile()
    Dim targetCell As Range
    Dim openFileName As String
    openFileName = Application.GetOpenFilename("テキスト,*.txt")
    If openFileName <> "False" Then
        Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(0, -1)
        If targetCell.MergeCells Then
            targetCell.MergeArea.Offset(0, 0).Value = openFileName
        Else
            targetCell.Value = openFileName
        End If
    End If
End Sub

'対象ファイル指定
Public Sub setTargetFile()
    Dim targetCell As Range
    Dim openFileName As String
    openFileName = Application.GetOpenFilename("ワークシート,*.xls")
    If openFileName <> "False" Then
        Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(0, -1)
        If targetCell.MergeCells Then
            targetCell.MergeArea.Offset(0, 0).Value = openFileName
        Else
            targetCell.Value = openFileName
        End If
    End If
End Sub

'対象フォルダ指定
Public Sub setTargetFolder()
    Dim targetCell As Range
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(0, -1)
            If targetCell.MergeCells Then
                targetCell.MergeArea.Offset(0, 0).Value = .SelectedItems(1)
            Else
                targetCell.Value = .SelectedItems(1)
            End If
        End If
    End With
End Sub

'時間を切り上げ／切り捨て
Public Function roundTime(ByVal targetTime As Double, ByVal timeDivide As Long, ByVal isRoundUp As Boolean)
    Dim targetMinute As Double
    Dim correctMinute As Double
    
    '分を取得
    targetMinute = Minute(targetTime)
    '単位時間で割り切れる場合は調整不要
    If targetMinute Mod timeDivide <> 0 Then
        '0～29分は0、30～59分は30に変換(30分単位の場合)
        correctMinute = (targetMinute \ timeDivide) * timeDivide
        '元の分を差し引いて、単位時間に変換した値を加える
        targetTime = DateAdd("n", correctMinute, DateAdd("n", targetMinute * -1, targetTime))
        '切り上げの場合、更に単位時間を加える
        If (isRoundUp) Then
            targetTime = DateAdd("n", timeDivide, targetTime)
        End If
    End If
    
    roundTime = targetTime
End Function

'URLエンコード
Public Function encodeURL(ByVal sWord As String) As String
    Dim d As Object
    Dim elm As Object
    
    sWord = Replace(sWord, "\", "\\")
    sWord = Replace(sWord, "'", "\'")
    Set d = CreateObject("htmlfile")
    Set elm = d.createElement("span")
    elm.setAttribute "id", "result"
    d.appendChild elm
    d.parentWindow.execScript "document.getElementById('result').innerText = encodeURIComponent('" & sWord & "');", "JScript"
    encodeURL = elm.innerText
 End Function

'URLデコード
Public Function decodeURL(ByVal sWord As String) As String
    Dim d As Object
    Dim elm As Object
    
    sWord = Replace(sWord, "\", "\\")
    sWord = Replace(sWord, "'", "\'")
    Set d = CreateObject("htmlfile")
    Set elm = d.createElement("span")
    elm.setAttribute "id", "result"
    d.appendChild elm
    d.parentWindow.execScript "document.getElementById('result').innerText = decodeURIComponent('" & sWord & "');", "JScript"
    decodeURL = elm.innerText
End Function

'IEを開く
Public Sub openIE(strUrl As String)
    Const navOpenInNewTab = &H800
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate strUrl, &H800
End Sub

'年月を今月にする
Public Sub setThisMonth()
    Call setTargetMonth(Date)
End Sub

'年月を来月にする
Public Sub setNextMonth()
    Call setTargetMonth(DateAdd("m", 1, Date))
End Sub

'年月を指定月に変更する
Private Sub setTargetMonth(ByVal target As Date)
    Range("年").Value = Format(target, "yyyy")
    Range("月").Value = Format(target, "m")
    
    Dim thisMonth As String: thisMonth = Format(target, "yyyymm")
    Dim beforeMonth As String: beforeMonth = Range("年月").Value
    Range("年月") = Replace(Range("年月"), beforeMonth, thisMonth)
    Range("対象ファイル") = Replace(Range("対象ファイル"), beforeMonth, thisMonth)
End Sub

'非表示の名前定義を表示
Public Sub showInvisibleNames()
    Dim oName As Object
    For Each oName In Names
        If oName.Visible = False Then
            oName.Visible = True
        End If
    Next
    MsgBox "非表示の名前定義を表示しました", vbOKOnly
End Sub
