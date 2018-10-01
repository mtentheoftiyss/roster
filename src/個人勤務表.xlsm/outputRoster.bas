Attribute VB_Name = "outputRoster"
Option Explicit

Private Const yearMonthErrorMsg As String = "年月を入力してください。"
Private Const yearMonthPatternErrorMsg As String = "年月はYYYYMM形式で入力してください。"
Private Const yearMonthSheetErrorMsg As String = "年月のシートが存在しません。"
Private Const placeErrorMsg As String = "勤務地を入力してください。"
Private Const cdErrorMsg As String = "社員コードを入力してください。"
Private Const userErrorMsg As String = "記入者を入力してください。"
Private Const targetFileErrorMsg As String = "対象ファイルを入力してください。"
Private Const targetFileExistErrorMsg As String = "対象ファイルが存在しません。"
Private Const targetFileNameErrorMsg As String = "対象ファイル名が「YYYYMMsacXXXXX.xls」ではありません。"
Private Const targetFileYMWarnMsg As String = "対象ファイルの年月が指定した年月と異なります。よろしいですか？"
Private Const workTimeEmptyInfoMsg As String = "勤務時間が未入力の場合、8を設定します。"
Private Const endNoDataMsg As String = "出力するデータがありません。"
Private Const endSuccessMsg As String = "ファイル出力が完了しました。"

'勤務表ファイルに出力
Public Sub outputRoster()
    '初期処理
    Call commonUtil.startProcess
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim inputPlace As String: inputPlace = Range("勤務地").Value
    Dim inputCd As String: inputCd = Range("社員コード").Value
    Dim inputUser As String: inputUser = Range("記入者").Value
    Dim inputWorkTime As String: inputWorkTime = Range("勤務時間").Value
    Dim inputYearMonth As String: inputYearMonth = Range("年月").Value
    Dim targetFile As String: targetFile = Range("対象ファイル").Value
    
    Dim errorFlg As Boolean: errorFlg = False
    Dim errorMsg As String: errorMsg = ""
    Dim warnFlg As Boolean: warnFlg = False
    Dim warnMsg As String: warnMsg = ""
    Dim infoFlg As Boolean: infoFlg = False
    Dim infoMsg As String: infoMsg = ""
    
    Dim inputYearMonthFlg: inputYearMonthFlg = True
    Dim targetFileFlg: targetFileFlg = True
    
    '入力チェック
    '勤務地
    If inputPlace = Empty Then
        '必須チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, placeErrorMsg)
    End If
    
    '社員コード
    If inputCd = Empty Then
        '必須チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, cdErrorMsg)
    End If
    
    '記入者
    If inputUser = Empty Then
        '必須チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, userErrorMsg)
    End If
    
    '年月
    If inputYearMonth = Empty Then
        '必須チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, yearMonthErrorMsg)
        inputYearMonthFlg = False
    ElseIf Not inputYearMonth Like "20[0-9][0-9][0-1][0-9]" Then
        '形式チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, yearMonthPatternErrorMsg)
        inputYearMonthFlg = False
    Else
        '対象シート存在チェック
        Dim ws As Worksheet
        Dim isExist As Boolean: isExist = False
        For Each ws In Worksheets
            If ws.Name = inputYearMonth Then
                isExist = True
            End If
        Next ws
        If Not isExist Then
            errorFlg = True
            errorMsg = commonUtil.createMsg(errorMsg, yearMonthSheetErrorMsg)
            inputYearMonthFlg = False
        End If
    End If
    
    '対象ファイル
    If targetFile = Empty Then
        '必須チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, targetFileErrorMsg)
        targetFileFlg = False
    ElseIf Dir(targetFile) = Empty Then
        'ファイル存在チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, targetFileExistErrorMsg)
        targetFileFlg = False
    ElseIf Not Dir(targetFile) Like "20[0-9][0-9][0-1][0-9]sac[X0-9][X0-9][X0-9][X0-9][X0-9].xls" Then
        'ファイル名形式チェック
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, targetFileNameErrorMsg)
        targetFileFlg = False
    End If
    
    If errorFlg Then
        'チェックエラー
        MsgBox errorMsg, vbCritical
        GoTo endProcess
    End If
    
    If inputYearMonthFlg And targetFileFlg Then
        '指定した年月と対象ファイルの年月が異なる場合、警告
        If inputYearMonth <> Left(Dir(targetFile), 6) Then
            warnFlg = True
            warnMsg = commonUtil.createMsg(warnMsg, targetFileYMWarnMsg)
        End If
    End If
    
    If warnFlg Then
        'チェックワーニング
        Dim result As Long
        result = MsgBox(warnMsg, vbOKCancel)
        If result <> vbOK Then
            GoTo endProcess
        End If
    End If
    
    '勤務時間
    If inputWorkTime = Empty Then
        '未入力の場合、8とする
        infoFlg = True
        infoMsg = commonUtil.createMsg(infoMsg, workTimeEmptyInfoMsg)
        inputWorkTime = "8"
        Range("勤務時間").Value = inputWorkTime
    End If
    
    If infoFlg Then
        'チェックインフォ
        MsgBox infoMsg, vbInformation
    End If
    
    'シートの移動
    Worksheets(inputYearMonth).Activate
    
    'データを読み込む
    Dim i As Long
    Dim j As Long
    Dim str As String
    Dim dataArray(31, 4) As String
    i = 1
    j = 0
    Do While i < Range("ファイル出力リスト").Rows.Count
        Dim place As String
        Dim places As String
        Dim c As Variant
        Dim project As String
        
        '作業場所をカンマ区切りで連結する(重複は排除)
        places = ""
        For Each c In Range("ファイル出力リスト")(i, Range("日").Column).MergeArea
            place = c.Offset(0, Range("作業場所").Column - Range("日").Column)
            If place = "" Then
                GoTo continue
'            ElseIf place = "その他" Then
'                GoTo continue
            Else
                If places = "" Then
                    places = place
                ElseIf InStr(places, place) = 0 Then
                    places = places & "," & place
                End If
            End If
continue:
        Next c
        
        '作業場所が入力されている場合、入力ありと判断する
        If places <> "" Then
            dataArray(j, 0) = Format(Range("ファイル出力リスト")(i, Range("日").Column).Value, "d")
            dataArray(j, 1) = Format(Range("ファイル出力リスト")(i, Range("開始").Column).Value, "hh:nn")
            dataArray(j, 2) = Format(Range("ファイル出力リスト")(i, Range("終了").Column).Value, "hh:nn")
            '休暇判定
            project = Range("ファイル出力リスト")(i, Range("案件").Column).Value
            If project = "休暇" Then
                dataArray(j, 3) = Range("ファイル出力リスト")(i, Range("作業内容").Column).Value
                dataArray(j, 4) = "有休"
            ElseIf project = "夏期休暇" Then
                dataArray(j, 3) = Range("ファイル出力リスト")(i, Range("作業内容").Column).Value
                dataArray(j, 4) = "特休"
            Else
                dataArray(j, 3) = places
            End If
            j = j + 1
        End If
        
        i = i + Range("ファイル出力リスト")(i, Range("日").Column).MergeArea.Rows.Count
    Loop
    
    'データなしの場合、終了
    If j = 0 Then
        MsgBox endNoDataMsg, vbInformation
        GoTo endProcess
    End If
    
    '勤務表に書き出す
    'ファイルを開く
    Workbooks.Open (targetFile)
    
    Const dayCol As Long = 1
    Const startCol As Long = 3
    Const endCol As Long = 4
    Const holidayCol As Long = 13
    Const placeCol As Long = 14
    Const startRow As Long = 8
    Const endRow As Long = 38
    
    '勤務地
    Const outputPlaceRow As Long = 3
    Const outputPlaceCol As Long = 3
    Cells(outputPlaceRow, outputPlaceCol).Value = inputPlace
    
    '社員コード
    Const outputCdRow As Long = 5
    Const outputCdCol As Long = 3
    Cells(outputCdRow, outputCdCol).Value = inputCd
    
    '記入者
    Const outputUserRow As Long = 61
    Const outputUserCol As Long = 15
    Cells(outputUserRow, outputUserCol).Value = inputUser
    
    '勤務時間
    Const outputWorkTimeRow As Long = 61
    Const outputWorkTimeCol As Long = 4
    Cells(outputWorkTimeRow, outputWorkTimeCol).Value = inputWorkTime
    
    '出社／退社時間をクリア
    Range(Cells(startRow, startCol), Cells(endRow, endCol)).Select
    Selection.ClearContents
    
    '休暇等／行先をクリア
    Range(Cells(startRow, holidayCol), Cells(endRow, placeCol)).Select
    Selection.ClearContents
    
    'データをループ
    Dim k As Long
    Dim targetRow As Long
    Dim targetTime As Variant
    Dim outputMinute As Double
    Dim correctOutputMinute
    For k = 0 To j - 1
        targetRow = startRow - 1 + CLng(dataArray(k, 0))
        
        '30分単位で出社時間は切り上げ／退社時間は切り捨て
        '出社時間
        If dataArray(k, 1) <> "" Then
            targetTime = TimeValue(dataArray(k, 1))
            '時間を規定時間(30分単位)で調整
            targetTime = commonUtil.roundTime(targetTime, commonConstants.timeDivide, True)
            Cells(targetRow, startCol).Value = targetTime
        End If
        
        '退社時間
        If dataArray(k, 2) <> "" Then
            targetTime = TimeValue(dataArray(k, 2))
            '時間を規定時間(30分単位)で調整
            targetTime = commonUtil.roundTime(targetTime, commonConstants.timeDivide, False)
            Cells(targetRow, endCol).Value = targetTime
        End If
        
        '行先
        Cells(targetRow, placeCol).Value = dataArray(k, 3)
        
        '休暇等
        Cells(targetRow, holidayCol).Value = dataArray(k, 4)
    Next

    '勤務表ファイルを保存して閉じる
    Range("C5").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close

    '完了メッセージを表示
    MsgBox endSuccessMsg, vbInformation
    
endProcess:
    '選択位置を初期に戻す
    initSelection.setInitSelection
    '終了処理
    Call commonUtil.endProcess
End Sub

