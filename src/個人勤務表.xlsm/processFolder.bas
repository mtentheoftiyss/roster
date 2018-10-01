Attribute VB_Name = "processFolder"
Option Explicit

'固定値フィールド
Private projectName As String
Private teamName As String
Private businessFunctionName As String
Private checkUser As String
Private checkDate As String

'セルフチェックリスト(RD)
Public Sub selfCheckRD()
    Call processFolder("selfCheckRD")
End Sub

'レビュー依頼書(RD)
Public Sub reviewRequestRD()
    Call processFolder("reviewRequestRD")
End Sub

'セルフチェックリスト(ED)
Public Sub selfCheckED()
    Call processFolder("selfCheckED")
End Sub

'レビュー依頼書(ED)
Public Sub reviewRequestED()
    Call processFolder("reviewRequestED")
End Sub

'DB設計書
Public Sub dbLayout()
    Call processFolder("dbLayout")
End Sub

'要件定義
Public Sub requirementDefinition()
    Call processFolder("requirementDefinition")
End Sub

'フォルダー操作
Public Sub processFolder(ByVal processMode As String)
    '初期処理
    Call commonUtil.startProcess
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim objFSO As FileSystemObject
    Dim targetFolder As String
    
    'ファイルシステムオブジェクト生成
    Set objFSO = New FileSystemObject
    
    'ルート指定
    targetFolder = Range("対象フォルダ")
    
    '探索開始
    Call searchSubFolder(objFSO.GetFolder(targetFolder), processMode)
    
    'オブジェクトを破棄
    Set objFSO = Nothing
    
    '完了メッセージを表示
    MsgBox "処理が終了しました。", vbInformation
    
    '選択位置を初期に戻す
    initSelection.setInitSelection
    '終了処理
    Call commonUtil.endProcess
End Sub

'サブフォルダー操作
Private Sub searchSubFolder(ByVal objFOLDER As Folder, ByVal processMode As String)
    Dim objSubFOLDER As Folder
    Dim objFILE As File
    
    '****************************************************************************************************
    'メイン処理ここから
    '****************************************************************************************************
    
    '固定値をフィールドに格納
    Call setPrefix
    
    For Each objFILE In objFOLDER.Files
    Select Case processMode
        Case "selfCheckRD"
            Call selfCheckRDMain(objFILE)
        Case "reviewRequestRD"
            Call reviewRequestRDMain(objFILE)
        Case "selfCheckED"
            Call selfCheckEDMain(objFILE)
        Case "reviewRequestED"
            Call reviewRequestEDMain(objFILE)
        Case "dbLayout"
            Call dbLayoutMain(objFILE)
        Case "requirementDefinition"
            Call requirementDefinitionMain(objFILE)
    End Select
    Next objFILE
    '****************************************************************************************************
    'メイン処理ここまで
    '****************************************************************************************************
    
    'サブフォルダーを探索
    For Each objSubFOLDER In objFOLDER.SubFolders
        Call searchSubFolder(objSubFOLDER, processMode)
    Next objSubFOLDER
    
    'オブジェクトを破棄
    Set objFOLDER = Nothing
End Sub

'固定値をフィールドに格納
Private Sub setPrefix()
projectName = Range("プロジェクト名").Value
teamName = Range("チーム名").Value
businessFunctionName = Range("業務機能名").Value
checkUser = Range("チェック実施者").Value
checkDate = Range("チェック実施日").Value

End Sub

'セルフチェックリスト(RD)
Private Sub selfCheckRDMain(ByVal objFILE As File)
    Dim targetSheet As String
    Dim targetCell As String
    Dim fixedValue As String
    
    'セルフチェックリストのみ対象とする
    Dim targetFileName As String
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "セルフチェックリスト") > 0 Then
        Exit Sub
    End If
    
    '画面かどうかの判定
    Dim functionKind As Integer
    functionKind = 0
    If targetFileName Like "D11-F02_2[1-5]_RD*" Then
        functionKind = 1
    ElseIf targetFileName Like "D11-F02_2[6-9]_RD*" Or targetFileName Like "D11-F02_3[0-3]_RD*" Then
        functionKind = 2
    Else
        functionKind = 3
    End If
    
    '書き込むシート、セル、値を指定
    targetSheet = "要件定義成果物 セルフチェックリスト"
    targetCell = "C6"
    Dim functionId As String
    Dim startPosition As Integer
    Dim endPosition As Integer
    Dim functionName As String
    
    Select Case functionKind
        '通常画面
        Case 1
            '機能ID
            functionId = Mid(targetFileName, 15, 5)
            '機能名開始位置
            startPosition = InStr(20, targetFileName, "_") + 1
            '機能名終了位置
            endPosition = InStrRev(targetFileName, "_")
            '機能名
            functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
            '設計書ファイル名
            fixedValue = "ES0303-F01-PTN2_" & functionId & "_機能設計書(" & functionName & ").xlsx"
        
        '補助画面
        Case 2
            '機能ID
            functionId = Mid(targetFileName, 15, 5)
            '機能名開始位置
            startPosition = InStr(20, targetFileName, "_") + 1
            '機能名終了位置
            endPosition = InStrRev(targetFileName, "_")
            '機能名
            functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
            '設計書ファイル名
            fixedValue = "ES0303-F01-PTN2_機能設計書_" & functionId & "_" & functionName & ".xlsx"
        
        'バッチ
        Case 3
            '機能ID
            functionId = Mid(targetFileName, 15, 8)
            '機能名開始位置
            startPosition = InStr(23, targetFileName, "_") + 1
            '機能名終了位置
            endPosition = InStrRev(targetFileName, "_")
            '機能名
            functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
            '設計書ファイル名
            fixedValue = "ES0303-F02-PTN2_" & functionId & "_機能設計書(" & functionName & ").xlsm"
    End Select
    
    'ファイルを開く
    Workbooks.Open (objFILE.Path)
    '書き込む
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell).Value = fixedValue
    'ファイルを保存して閉じる
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'レビュー依頼書(RD)
Private Sub reviewRequestRDMain(ByVal objFILE As File)
    Dim targetSheet As String
    Dim targetCell As String
    Dim targetCell7 As String
    Dim fixedValue As String
    Dim fixedValue7 As String
    Dim targetFileName As String
    
    'レビュー依頼書のみ対象とする
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "レビュー依頼書兼報告書") > 0 Then
        Exit Sub
    End If
    
    '画面かどうかの判定
    Dim fixedStartPosition As Integer
    Dim fixedIdLength As Integer
    If targetFileName Like "D11-F02_2[1-5]_RD*" Then
        fixedStartPosition = 20
        fixedIdLength = 5
    ElseIf targetFileName Like "D11-F04_2[1-9]_RD*" Or targetFileName Like "D11-F04_3[0-3]_RD*" Then
        fixedStartPosition = 20
        fixedIdLength = 5
    Else
        fixedStartPosition = 23
        fixedIdLength = 8
    End If
    
    '書き込むシート、セル、値を指定
    targetSheet = "レビュー依頼書兼報告書"
    targetCell = "機能名"
    targetCell7 = "レビュー管理番号"
    
    '開始位置
    Dim startPosition As Integer
    startPosition = InStr(fixedStartPosition, targetFileName, "_") + 1
    '終了位置
    Dim endPosition As Integer
    endPosition = InStrRev(targetFileName, "_")
    '機能名
    fixedValue = Mid(targetFileName, startPosition, endPosition - startPosition)
    'レビュー管理番号
    fixedValue7 = Mid(targetFileName, 12, 2) & "_" & Mid(targetFileName, 9, 2) & "_" & Mid(targetFileName, 15, fixedIdLength)
    
    'ファイルを開く
    Workbooks.Open (objFILE.Path)
    '書き込む
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell).Value = fixedValue
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell7).Value = fixedValue7
    'ファイルを保存して閉じる
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'セルフチェックリスト(ED)
Private Sub selfCheckEDMain(ByVal objFILE As File)
    Dim targetSheet As String
    
    Dim targetCell1 As String
    Dim targetCell2 As String
    Dim targetCell3 As String
    Dim targetCell4 As String
    Dim targetCell5 As String
    Dim targetCell6 As String
    Dim targetCell7 As String
    
    Dim fixedValue1 As String
    Dim fixedValue2 As String
    Dim fixedValue3 As String
    Dim fixedValue4 As String
    Dim fixedValue5 As String
    Dim fixedValue6 As String
    Dim fixedValue7 As String
    
    'セルフチェックリストのみ対象とする
    Dim targetFileName As String
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "セルフチェックリスト") > 0 Then
        Exit Sub
    End If
    
    '画面かどうかの判定
    Dim funcClass As String
    Dim funcPrefix As String
    If targetFileName Like "D11-F02_21_ED*" Then
        funcClass = "画面"
        funcPrefix = "ES0303-F01-PTN2_"
    Else
        funcClass = "メール"
        funcPrefix = "ES0302-F13_"
    End If
    
    '書き込むシート、セル、値を指定
    targetSheet = "基本設計(システム設計・外部設計・AP基盤セルフチェックリスト"
    targetCell1 = "H3"
    targetCell2 = "X3"
    targetCell3 = "AQ3"
    targetCell4 = "H4"
    targetCell5 = "X4"
    targetCell6 = "AQ4"
    targetCell7 = "C6"
    Dim functionId As String
    Dim startPosition As Integer
    Dim endPosition As Integer
    Dim functionName As String
    
    'プロジェクト名
    fixedValue1 = projectName
    'チーム名
    fixedValue2 = teamName
    'チェック実施者
    fixedValue3 = checkUser
    '業務名
    fixedValue4 = "機能設計書（" & funcClass & "）"
    '業務機能名
    fixedValue5 = businessFunctionName
    'チェック実施日
    fixedValue6 = checkDate
    
    '機能ID
    functionId = Mid(targetFileName, 15, 5)
    '機能名開始位置
    startPosition = InStr(20, targetFileName, "_") + 1
    '機能名終了位置
    endPosition = InStr(startPosition, targetFileName, "_")
    '機能名
    functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
    '設計書ファイル名
    fixedValue7 = funcPrefix & functionId & "_機能設計書(" & functionName & ").xlsx"
    
    'ファイルを開く
    Workbooks.Open (objFILE.Path)
    '書き込む
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell1).Value = fixedValue1   'プロジェクト名
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell2).Value = fixedValue2   'チーム名
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell3).Value = fixedValue3   'チェック実施者
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell4).Value = fixedValue4   '業務名
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell5).Value = fixedValue5   '業務機能名
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell6).Value = fixedValue6   'チェック実施日
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell7).Value = fixedValue7   '設計書ファイル名
    'ファイルを保存して閉じる
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'レビュー依頼書(ED)
Private Sub reviewRequestEDMain(ByVal objFILE As File)
    Dim targetFileName As String
    Dim targetSheet As String
    
    Dim targetCell1 As String
    Dim targetCell2 As String
    Dim targetCell3 As String
    Dim targetCell4 As String
    Dim targetCell5 As String
    Dim targetCell6 As String
    Dim targetCell7 As String
    Dim targetCell8 As String
    
    Dim fixedValue1 As String
    Dim fixedValue2 As String
    Dim fixedValue3 As String
    Dim fixedValue4 As String
    Dim fixedValue5 As String
    Dim fixedValue6 As String
    Dim fixedValue7 As String
    Dim fixedValue8 As String
    
    'レビュー依頼書のみ対象とする
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "レビュー依頼書兼報告書") > 0 Then
        Exit Sub
    End If
    
    
    '画面かどうかの判定
    Dim funcClass As String
    Dim fixedStartPosition As Integer
    Dim fixedIdLength As Integer
    If targetFileName Like "D11-F04_21_ED*" Then
        funcClass = "画面"
        fixedStartPosition = 20
        fixedIdLength = 5
    Else
        funcClass = "メール"
        fixedStartPosition = 20
        fixedIdLength = 5
    End If
    
    
    '書き込むシート、セル、値を指定
    targetSheet = "レビュー依頼書兼報告書"
    targetCell1 = "プロジェクト名"
    targetCell2 = "チーム名"
    targetCell3 = "対象構成管理名"
    targetCell4 = "業務機能名"
    targetCell5 = "対象成果物名"
    targetCell6 = "機能名"
    targetCell7 = "レビュー管理番号"
    targetCell8 = "ページ数"
    
    '開始位置
    Dim startPosition As Integer
    startPosition = InStr(fixedStartPosition, targetFileName, "_") + 1
    '終了位置
    Dim endPosition As Integer
    endPosition = InStr(startPosition, targetFileName, "_")
    
    'プロジェクト名
    fixedValue1 = projectName
    'チーム名
    fixedValue2 = teamName
    '対象構成管理名
    fixedValue3 = "機能設計書（" & funcClass & "）"
    '業務機能名
    fixedValue4 = businessFunctionName
    '対象成果物名
    fixedValue5 = funcClass & "定義書(" & Mid(targetFileName, 12, 2) & "_" & Mid(targetFileName, 9, 2) & "_" & Mid(targetFileName, 15, fixedIdLength) & ")"
    '機能名
    fixedValue6 = Mid(targetFileName, startPosition, endPosition - startPosition)
    'レビュー管理番号
    fixedValue7 = Mid(targetFileName, 12, 2) & "_" & Mid(targetFileName, 9, 2) & "_" & Mid(targetFileName, 15, fixedIdLength)
    'ページ数
    Dim wb As Workbook
    Set wb = Workbooks.Open("D:\zz_endo-work\コマンド\Windowsコマンド\copy\ページ数カウント.xlsx")
    fixedValue8 = Application.WorksheetFunction.VLookup(Mid(targetFileName, startPosition, endPosition - startPosition), Range("B1:C40"), 2, False)
    wb.Close
    
    'ファイルを開く
    Workbooks.Open (objFILE.Path)
    '書き込む
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell1).Value = fixedValue1   'プロジェクト名
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell2).Value = fixedValue2   'チーム名
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell3).Value = fixedValue3   '対象構成管理名
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell4).Value = fixedValue4   '業務機能名
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell5).Value = fixedValue5   '対象成果物名
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell6).Value = fixedValue6   '機能名
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell7).Value = fixedValue7   'レビュー管理番号
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell8).Value = fixedValue8   'ページ数
    'ファイルを保存して閉じる
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'要件定義
Public Sub requirementDefinitionMain(ByVal objFILE As File)
    Dim ws As Worksheet

    Dim targetSheet1 As String
    Dim targetSheet2 As String
    Dim targetSheet3 As String
    Dim targetSheet4 As String
    Dim targetSheet5 As String
    
    Dim targetCells1 As String
    Dim targetCells2 As String
    Dim targetCells3 As String
    Dim targetCells4 As String
    Dim targetCells5 As String
    Dim targetCells6 As String
    Dim targetCells7 As String
    Dim targetCells8 As String
    Dim targetCells9 As String
    
    Dim targetCell1 As String
    Dim targetCell2 As String
    Dim targetCell3 As String
    Dim targetCell4 As String
    Dim targetCell5 As String
    Dim targetCell6 As String
    Dim targetCell7 As String
    Dim targetCell8 As String
    Dim targetCell9 As String
    
    Dim fixedValue1 As String
    Dim fixedValue2 As String
    Dim fixedValue3 As String
    Dim fixedValue4 As String
    Dim fixedValue5 As String
    Dim fixedValue6 As String
    Dim fixedValue7 As String
    Dim fixedValue8 As String
    Dim fixedValue9 As String

    '表紙
    targetSheet1 = "表紙"
    targetCells1 = "H19:H20"
    targetCells2 = "AD19:AD20"
    '作成年月日
    targetCell1 = "H19"
    fixedValue1 = "2018/9/27"
    '作成者
    targetCell2 = "AD19"
    fixedValue2 = "SCSK"
    
    '改訂履歴
    targetSheet2 = "改訂履歴"
    targetCells3 = "C3:G30"
    '変更日
    targetCell3 = "C3"
    fixedValue3 = "2018/9/27"
    '変更者
    targetCell4 = "D3"
    fixedValue4 = "SCSK"
    '変更内容
    targetCell5 = "F3"
    fixedValue5 = "新規作成"
    
    'ファイルを開く
    Workbooks.Open (objFILE.Path)
    
    '書き換え
    For Each ws In Worksheets
        If ws.Visible = xlSheetVisible Then
            '表紙
            If ws.Name = targetSheet1 Then
                '最初にクリア
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCells1).Value = ""
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCells2).Value = ""
            
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCell1).Value = fixedValue1   '作成年月日
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCell2).Value = fixedValue2   '作成者
            End If
        
            '改訂履歴
            If ws.Name = targetSheet2 Then
                '最初にクリア
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCells3).Value = ""
            
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCell3).Value = fixedValue3   '変更日
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCell4).Value = fixedValue4   '変更者
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCell5).Value = fixedValue5   '変更内容
            End If
            
            '左上
            ActiveWorkbook.Worksheets(ws.Name).Activate
            ActiveWorkbook.Worksheets(ws.Name).Select
            Application.Goto Reference:=Range("A1"), Scroll:=True
        End If
    Next
    
    'ファイルを保存して閉じる
    ActiveWorkbook.Worksheets(1).Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'ページ数カウントメイン
Public Sub pageCountMain()
    '初期処理
    Call commonUtil.startProcess
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim objFSO As FileSystemObject
    Dim targetFolder As String
    Dim resultCollection As Collection
    Dim rowCnt As Integer
    rowCnt = 1
    
    'ファイルシステムオブジェクト生成
    Set objFSO = New FileSystemObject
    
    'ルート指定
    targetFolder = Range("対象フォルダ")
    
    '探索開始
    Set resultCollection = New Collection
    Set resultCollection = pageCountSub(objFSO.GetFolder(targetFolder), resultCollection)
    
    If resultCollection.Count > 0 Then
        '新規ブック作成
        Dim wb As Workbook
        Set wb = Workbooks.Add
        Dim i As Integer
        
        '見出し
        wb.Sheets(1).Cells(rowCnt, 1).Value = "フォルダ名"
        wb.Sheets(1).Cells(rowCnt, 2).Value = "ファイル名"
        wb.Sheets(1).Cells(rowCnt, 3).Value = "シート名"
        wb.Sheets(1).Cells(rowCnt, 4).Value = "ページ数"
        wb.Sheets(1).Cells(rowCnt, 5).Value = "非表示"
        
        For i = 1 To resultCollection.Count
            rowCnt = rowCnt + 1
            wb.Sheets(1).Cells(rowCnt, 1).Value = resultCollection(i)(0)
            wb.Sheets(1).Cells(rowCnt, 2).Value = resultCollection(i)(1)
            wb.Sheets(1).Cells(rowCnt, 3).Value = resultCollection(i)(2)
            wb.Sheets(1).Cells(rowCnt, 4).Value = resultCollection(i)(3)
            wb.Sheets(1).Cells(rowCnt, 5).Value = resultCollection(i)(4)
        Next i
        
        '列幅を調整
        wb.Sheets(1).Range("A:D").Columns.AutoFit
        
        'ブックを保存
        wb.SaveAs fileName:=targetFolder & "\【ページ数カウント】.xlsx"
        wb.Close
    End If
    
    'オブジェクトを破棄
    Set objFSO = Nothing
    
    '完了メッセージを表示
    MsgBox "処理が終了しました。", vbInformation
    
    '選択位置を初期に戻す
    initSelection.setInitSelection
    '終了処理
    Call commonUtil.endProcess
End Sub

'ページ数カウントサブ
Private Function pageCountSub(ByVal objFOLDER As Folder, ByRef objCollection As Collection)
    Dim objSubFOLDER As Folder
    Dim objFILE As File
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim collectionItem(4) As String
    Dim rowCnt As Integer
    rowCnt = 0
    Dim objFSO As FileSystemObject
    Set objFSO = New FileSystemObject
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\~\$.*$"
    Dim ext As String
    Dim fileAttr As Long
    
    For Each objFILE In objFOLDER.Files
        collectionItem(0) = objFILE.ParentFolder
        collectionItem(1) = objFILE.Name
        
        '隠しファイルと一時ファイルは除外する
        'カウントするのはEXCELファイルのみ対象とする
        ext = LCase(objFSO.GetExtensionName(objFILE.Path))
        fileAttr = GetAttr(objFILE.Path)
        If ((fileAttr And vbHidden) = False And re.test(objFILE.Name) = False) And (ext = "xls" Or ext = "xlsx" Or ext = "xlsm") Then
            'ファイルを開く
            Workbooks.Open (objFILE.Path)
            
            'ページカウント
            For Each ws In ActiveWorkbook.Worksheets
                collectionItem(4) = ""
                If ws.Visible <> xlSheetVisible Then
                    ws.Visible = xlSheetVisible
                    collectionItem(4) = "○"
                End If
                
                ws.Activate
                ActiveWindow.View = xlPageBreakPreview
                collectionItem(2) = ws.Name
'                collectionItem(3) = Application.ExecuteExcel4Macro("get.document(50)")
                collectionItem(3) = ws.PageSetup.Pages.Count
                objCollection.Add collectionItem
            Next ws
        
            'ファイルを閉じる
            ActiveWorkbook.Close
        Else
            collectionItem(2) = "-"
            collectionItem(3) = "-"
            collectionItem(4) = ""
            objCollection.Add collectionItem
        End If
    Next objFILE
    
    'サブフォルダーを探索
    For Each objSubFOLDER In objFOLDER.SubFolders
        Set objCollection = pageCountSub(objSubFOLDER, objCollection)
    Next objSubFOLDER
    
    'オブジェクトを破棄
    Set objFOLDER = Nothing
    Set objFSO = Nothing
    
    'コレクションを返却
    Set pageCountSub = objCollection
End Function

'DB設計書
Private Sub dbLayoutMain(ByVal objFILE As File)
    Const idCol As Integer = 3
    Const typeCol As Integer = 6
    Const startRow As Integer = 4
    Const entityCol As Integer = 4
    Const entityRow As Integer = 1
    
    Const strTypeName As String = "nvarchar2"
    Const numTypeName As String = "number"
    Const dateTypeNmae As String = "date"
    
    Const strPrefix As String = "S_"
    Const numPrefix As String = "N_"
    Const datePrefix As String = "D_"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim entityVal As String
    Dim idVal As String
    Dim beforeTypeVal As String
    Dim afterTypeVal As String
    Dim targetRow As Integer
    Dim lastRow As Integer
    
    'ファイルを開く
    Set wb = Workbooks.Open(objFILE.Path)
    
    'シートをループ
    For Each ws In wb.Worksheets
        '表示シートのみを対象とする
        If ws.Visible = xlSheetVisible Then
            'シート名を判定
            If ws.Name <> "表紙" And ws.Name <> "改訂履歴" And ws.Name <> "ER図" And ws.Name <> "論理ER図" And ws.Name <> "論理エンティティ一覧" And ws.Name <> "こだわり条件" And ws.Name <> "リハウスコメント" And ws.Name <> "インデックス定義" And ws.Name <> "データビュー一覧" And ws.Name <> "データビュー・エンティティ関連定義" And ws.Name <> "SACで追加した項目の履歴" Then
                'エンティティ名を書き換え
                entityVal = ws.Cells(entityRow, entityCol).Value
                ws.Cells(entityRow, entityCol).Value = Replace(entityVal, "_", "_MYP_", 1, 1)
                
                '最終行を取得
                lastRow = ws.Cells(startRow, typeCol).End(xlDown).Row
                
                For targetRow = startRow To lastRow
                    'データ型の置換
                    beforeTypeVal = ws.Cells(targetRow, typeCol).Value
                    Select Case beforeTypeVal
                        Case "char"
                            afterTypeVal = strTypeName
                        Case "nchar"
                            afterTypeVal = strTypeName
                        Case "date"
                            afterTypeVal = dateTypeNmae
                        Case "datetime"
                            afterTypeVal = dateTypeNmae
                        Case "decimal"
                            afterTypeVal = numTypeName
                        Case "int"
                            afterTypeVal = numTypeName
                        Case "varchar"
                            afterTypeVal = strTypeName
                        Case "nvarchar"
                            afterTypeVal = strTypeName
                        Case Else
                            afterTypeVal = beforeTypeVal
                    End Select
                    
                    'データ項目IDにプレフィックスを付加
                    idVal = ws.Cells(targetRow, idCol).Value
                    'FastAPP共通項目は対象外
                    Select Case idVal
                        Case "HRZ_ROLE_ID"
                            idVal = idVal
                        Case "VRT_ROLE_ID"
                            idVal = idVal
                        Case "DEL_FLG"
                            idVal = idVal
                        Case "CRE_DT"
                            idVal = idVal
                        Case "CRE_USR"
                            idVal = idVal
                        Case "UPD_DT"
                            idVal = idVal
                        Case "UPD_USR"
                            idVal = idVal
                        Case "UPD_CNT"
                            idVal = idVal
                        Case Else
                            Select Case afterTypeVal
                                Case strTypeName
                                    idVal = strPrefix & idVal
                                Case numTypeName
                                    idVal = numPrefix & idVal
                                Case dateTypeNmae
                                    idVal = datePrefix & idVal
                            End Select
                    End Select
                    
                    'セルの書き換え
                    ws.Cells(targetRow, typeCol).Value = afterTypeVal
                    ws.Cells(targetRow, idCol).Value = idVal
                Next targetRow
                
'                'セル選択位置を戻す
'                ws.Range("A1").Select
            End If
        End If
    Next ws
    
    'ファイルを保存して閉じる
'    wb.Worksheets(1).Range("A1").Select
    wb.Save
    wb.Close
End Sub


