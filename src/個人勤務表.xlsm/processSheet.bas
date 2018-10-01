Attribute VB_Name = "processSheet"
Option Explicit
    
Private Const requiredErrorMsg As String = "を入力してください。"
Private Const numericErrorMsg As String = "には数値を入力してください。"
Private Const intErrorMsg As String = "には整数を入力してください。"
Private Const yearRangeErrorMsg As String = "年には1900～9999の値を入力してください。"
Private Const monthRangeErrorMsg As String = "月には1～12の値を入力してください。"
Private Const yearMonthErrorMsg As String = "対象年月のシートが既に存在します。"
Private Const delConfirmMsg As String = "対象年月のシートを削除します。よろしいですか？"
Private Const noSheetMsg As String = "対象年月のシートが存在しません。"

'新規シートを作成
Public Sub createNewSheet()
    '画面の描画をOFFにします
    Call commonUtil.startProcess
    
    Dim yearVal As String: yearVal = Range("年").Value
    Dim monthVal As String: monthVal = Range("月").Value
    
    '年月どちらか入力チェックに引っかかった場合、エラー
    If inputValidate(yearVal, monthVal) Then
        GoTo endProcess
    End If
    
    '対象年月のシートが存在する場合、エラー
    Dim yearMonth As String: yearMonth = zeroAdd(yearVal, 4) & zeroAdd(monthVal, 2)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = yearMonth Then
            MsgBox yearMonthErrorMsg, vbCritical
            GoTo endProcess
        End If
    Next ws
    
    'ベースシートからコピー
    Sheets("base").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = yearMonth
'    ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).Properties("_CodeName") = yearMonth
'    ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).Name = yearMonth
'    ActiveSheet.["_CodeName"] = yearMonth
    Application.Goto Reference:=ActiveWindow.ActiveSheet.Range("A1"), Scroll:=True
    
    'イベントの発生抑止を解除する
    Application.EnableEvents = True
    
    Range("年").Value = yearVal
    Range("月").Value = monthVal
    
    '年月リスト作成
    createYMList
    
endProcess:
    '終了処理
    Call commonUtil.endProcess
End Sub

'シートを削除
Public Sub deleteSheet()
    '画面の描画をOFFにします
    Call commonUtil.startProcess
    
    Dim yearVal As String: yearVal = Range("年").Value
    Dim monthVal As String: monthVal = Range("月").Value
    
    '年月どちらか入力チェックに引っかかった場合、エラー
    If inputValidate(yearVal, monthVal) Then
        GoTo endProcess
    End If

    '対象年月のシートが存在する場合、削除する
    Dim yearMonth As String: yearMonth = zeroAdd(yearVal, 4) & zeroAdd(monthVal, 2)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = yearMonth Then
            Dim rc As Integer
            rc = MsgBox(delConfirmMsg, vbYesNo + vbExclamation + vbDefaultButton2)
            If rc = vbYes Then
                'シート削除
                Sheets(yearMonth).Delete
            End If
    
            '年月リスト作成
            createYMList
            
            GoTo endProcess
        End If
    Next ws
    
    '対象年月のシートが存在しない場合、エラー
    MsgBox noSheetMsg, vbCritical

endProcess:
    '終了処理
    Call commonUtil.endProcess
End Sub

'年月リスト作成
Private Sub createYMList()
    Dim ymList As String: ymList = ""
    Dim ws As Worksheet
    For Each ws In Worksheets
        If Not ws.Name Like "20[0-9][0-9][0-1][0-9]" Then
            GoTo continue
        End If
        If ymList = "" Then
            ymList = ws.Name
        Else
            ymList = ymList & "," & ws.Name
        End If
continue:
    Next

    With Worksheets("main").Range("年月").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ymList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = False
        .ShowError = False
    End With

End Sub

'入力チェック
Private Function inputValidate(ByVal yearVal As String, ByVal monthVal As String)
    Dim errorFlg As Boolean: errorFlg = False
    Dim monthFlg As Boolean: monthFlg = False
    Dim errorMsg As String: errorMsg = ""
    
    '年の未入力チェック
    If yearVal = Empty Then
        errorFlg = True
        errorMsg = createMsg(errorMsg, "年" + requiredErrorMsg)
    Else
        '年の数値チェック
        If Not IsNumeric(yearVal) Then
            errorFlg = True
            errorMsg = createMsg(errorMsg, "年" + numericErrorMsg)
        Else
            '年の整数チェック
            If Int(yearVal) <> yearVal Then
                errorFlg = True
                errorMsg = createMsg(errorMsg, "年" + intErrorMsg)
            Else
                '年の範囲チェック
                If yearVal < 1900 Or yearVal > 9999 Then
                    errorFlg = True
                    errorMsg = createMsg(errorMsg, yearRangeErrorMsg)
                End If
            End If
        End If
    End If
    
    '月の未入力チェック
    If monthVal = Empty Then
        errorFlg = True
        errorMsg = createMsg(errorMsg, "月" + requiredErrorMsg)
    Else
        '月の数値チェック
        If Not IsNumeric(monthVal) Then
            errorFlg = True
            errorMsg = createMsg(errorMsg, "月" + numericErrorMsg)
        Else
            '月の整数チェック
            If Int(monthVal) <> monthVal Then
                errorFlg = True
                errorMsg = createMsg(errorMsg, "月" + intErrorMsg)
            Else
                '月の範囲チェック
                If monthVal < 1 Or monthVal > 12 Then
                    errorFlg = True
                    errorMsg = createMsg(errorMsg, monthRangeErrorMsg)
                End If
            End If
        End If
    End If
    
    '年月どちらかチェックに引っかかった場合、エラー
    If errorFlg Then
        MsgBox errorMsg, vbCritical
    End If
    
    inputValidate = errorFlg
End Function

'改行区切りでメッセージを作成
Private Function createMsg(ByVal msg As String, ByVal addMsg As String)
    If msg <> Empty Then
        msg = msg & vbCrLf
    End If
    
    createMsg = msg & addMsg
End Function

'数値形式文字列のゼロ埋め
Private Function zeroAdd(ByVal str As String, ByVal length As Integer)
    Dim zeroStr As String: zeroStr = ""
    Dim i As Integer:
    
    For i = 0 To length
        zeroStr = zeroStr & "0"
    Next
    
    zeroAdd = Right(zeroStr & str, length)
End Function
