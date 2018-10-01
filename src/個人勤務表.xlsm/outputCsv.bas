Attribute VB_Name = "outputCsv"
'※※※※※※※
'※　未使用　※
'※※※※※※※


Option Explicit

Private Const endSuccessMsg As String = "ファイル出力が完了しました。"

'CSVファイルを出力
Public Sub outputCsv()
    '初期処理
    Call commonUtil.startProcess

    'CSVファイル
    Dim yearVal As String: yearVal = Right("0000" & Range("年").Value, 4)
    Dim monthVal As String: monthVal = Right("00" & Range("月").Value, 2)
    Dim csvFile As String
    csvFile = ThisWorkbook.Path & "\outputCsv" & yearVal & monthVal & ".txt"

'    'SJISで書き出す
'    '空いているファイル番号を取得
'    Dim fileNumber As Integer
'    fileNumber = FreeFile
'    'ファイルを出力モードで開く
'    Open csvFile For Output As #fileNumber
'
'    'データを書き込む
'    Dim i As Long
'    Dim str As String
'    i = 1
'    Do While i < Range("ファイル出力リスト").Rows.Count
'        Dim attendance As String
'        Dim dt As String
'        Dim tm As String
'        Dim place As String
'        Dim c As Variant
'        For Each c In Range("ファイル出力リスト")(i, Range("日").Column).MergeArea
'            dt = Format(Range("ファイル出力リスト")(i, Range("日").Column).Value, "mmmm dd, yyyy")
'            place = c.Offset(0, Range("作業場所").Column - Range("日").Column)
'            If place = "" Then
'                GoTo Continue
'            End If
'
'            '出勤
'            tm = Format(Range("ファイル出力リスト")(i, Range("開始").Column).Value, " at hh:nnAM/PM")
'            If tm <> "" Then
'                str = "entered,"
'                str = str & dt & tm
'                str = str & ","
'                str = str & place
'
'                Print #fileNumber, str
'            End If
'
'            '退勤
'            tm = Format(Range("ファイル出力リスト")(i, Range("終了").Column).Value, " at hh:nnAM/PM")
'            If tm <> "" Then
'                str = "exited,"
'                str = str & dt & tm
'                str = str & ","
'                str = str & place
'
'                Print #fileNumber, str
'            End If
'
'Continue:
'        Next c
'
'        i = i + Range("ファイル出力リスト")(i, Range("日").Column).MergeArea.Rows.Count
'    Loop
'
'    'ファイルを閉じる
'    Close #fileNumber
'
    'UTF-8で書き出す
    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 2
    outStream.Charset = "utf-8"
    outStream.LineSeparator = 10
    outStream.Open
    
    'データを書き込む
    Dim i As Long
    Dim str As String
    i = 1
    Do While i < Range("ファイル出力リスト").Rows.Count
        Dim attendance As String
        Dim dt As String
        Dim tm As String
        Dim place As String
        Dim c As Variant
        For Each c In Range("ファイル出力リスト")(i, Range("日").Column).MergeArea
            dt = Format(Range("ファイル出力リスト")(i, Range("日").Column).Value, "mmmm dd, yyyy")
            place = c.Offset(0, Range("作業場所").Column - Range("日").Column)
            If place = "" Then
                GoTo continue
            End If
            
            '出勤
            tm = Format(Range("ファイル出力リスト")(i, Range("開始").Column).Value, " at hh:nnAM/PM")
            If tm <> "" Then
                str = "entered,"
                str = str & dt & tm
                str = str & ","
                str = str & place
                
                outStream.WriteText str, 1
            End If
            
            '退勤
            tm = Format(Range("ファイル出力リスト")(i, Range("終了").Column).Value, " at hh:nnAM/PM")
            If tm <> "" Then
                str = "exited,"
                str = str & dt & tm
                str = str & ","
                str = str & place
                
                outStream.WriteText str, 1
            End If
            
continue:
        Next c
        
        i = i + Range("ファイル出力リスト")(i, Range("日").Column).MergeArea.Rows.Count
    Loop
    
    'ファイルを閉じる
    outStream.Position = 0
    outStream.Type = 1
    outStream.Position = 3
    Dim csvStream As Object
    Set csvStream = CreateObject("ADODB.Stream")
    csvStream.Type = 1
    csvStream.Open
    outStream.CopyTo csvStream
    
    csvStream.SaveToFile csvFile, 2
    
    outStream.Close
    csvStream.Close

    MsgBox endSuccessMsg, vbInformation
    
    '終了処理
    Call commonUtil.endProcess
End Sub
