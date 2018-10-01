Attribute VB_Name = "outputCsv"
'��������������
'���@���g�p�@��
'��������������


Option Explicit

Private Const endSuccessMsg As String = "�t�@�C���o�͂��������܂����B"

'CSV�t�@�C�����o��
Public Sub outputCsv()
    '��������
    Call commonUtil.startProcess

    'CSV�t�@�C��
    Dim yearVal As String: yearVal = Right("0000" & Range("�N").Value, 4)
    Dim monthVal As String: monthVal = Right("00" & Range("��").Value, 2)
    Dim csvFile As String
    csvFile = ThisWorkbook.Path & "\outputCsv" & yearVal & monthVal & ".txt"

'    'SJIS�ŏ����o��
'    '�󂢂Ă���t�@�C���ԍ����擾
'    Dim fileNumber As Integer
'    fileNumber = FreeFile
'    '�t�@�C�����o�̓��[�h�ŊJ��
'    Open csvFile For Output As #fileNumber
'
'    '�f�[�^����������
'    Dim i As Long
'    Dim str As String
'    i = 1
'    Do While i < Range("�t�@�C���o�̓��X�g").Rows.Count
'        Dim attendance As String
'        Dim dt As String
'        Dim tm As String
'        Dim place As String
'        Dim c As Variant
'        For Each c In Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).MergeArea
'            dt = Format(Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).Value, "mmmm dd, yyyy")
'            place = c.Offset(0, Range("��Əꏊ").Column - Range("��").Column)
'            If place = "" Then
'                GoTo Continue
'            End If
'
'            '�o��
'            tm = Format(Range("�t�@�C���o�̓��X�g")(i, Range("�J�n").Column).Value, " at hh:nnAM/PM")
'            If tm <> "" Then
'                str = "entered,"
'                str = str & dt & tm
'                str = str & ","
'                str = str & place
'
'                Print #fileNumber, str
'            End If
'
'            '�ދ�
'            tm = Format(Range("�t�@�C���o�̓��X�g")(i, Range("�I��").Column).Value, " at hh:nnAM/PM")
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
'        i = i + Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).MergeArea.Rows.Count
'    Loop
'
'    '�t�@�C�������
'    Close #fileNumber
'
    'UTF-8�ŏ����o��
    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 2
    outStream.Charset = "utf-8"
    outStream.LineSeparator = 10
    outStream.Open
    
    '�f�[�^����������
    Dim i As Long
    Dim str As String
    i = 1
    Do While i < Range("�t�@�C���o�̓��X�g").Rows.Count
        Dim attendance As String
        Dim dt As String
        Dim tm As String
        Dim place As String
        Dim c As Variant
        For Each c In Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).MergeArea
            dt = Format(Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).Value, "mmmm dd, yyyy")
            place = c.Offset(0, Range("��Əꏊ").Column - Range("��").Column)
            If place = "" Then
                GoTo continue
            End If
            
            '�o��
            tm = Format(Range("�t�@�C���o�̓��X�g")(i, Range("�J�n").Column).Value, " at hh:nnAM/PM")
            If tm <> "" Then
                str = "entered,"
                str = str & dt & tm
                str = str & ","
                str = str & place
                
                outStream.WriteText str, 1
            End If
            
            '�ދ�
            tm = Format(Range("�t�@�C���o�̓��X�g")(i, Range("�I��").Column).Value, " at hh:nnAM/PM")
            If tm <> "" Then
                str = "exited,"
                str = str & dt & tm
                str = str & ","
                str = str & place
                
                outStream.WriteText str, 1
            End If
            
continue:
        Next c
        
        i = i + Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).MergeArea.Rows.Count
    Loop
    
    '�t�@�C�������
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
    
    '�I������
    Call commonUtil.endProcess
End Sub
