Attribute VB_Name = "outputRoster"
Option Explicit

Private Const yearMonthErrorMsg As String = "�N������͂��Ă��������B"
Private Const yearMonthPatternErrorMsg As String = "�N����YYYYMM�`���œ��͂��Ă��������B"
Private Const yearMonthSheetErrorMsg As String = "�N���̃V�[�g�����݂��܂���B"
Private Const placeErrorMsg As String = "�Ζ��n����͂��Ă��������B"
Private Const cdErrorMsg As String = "�Ј��R�[�h����͂��Ă��������B"
Private Const userErrorMsg As String = "�L���҂���͂��Ă��������B"
Private Const targetFileErrorMsg As String = "�Ώۃt�@�C������͂��Ă��������B"
Private Const targetFileExistErrorMsg As String = "�Ώۃt�@�C�������݂��܂���B"
Private Const targetFileNameErrorMsg As String = "�Ώۃt�@�C�������uYYYYMMsacXXXXX.xls�v�ł͂���܂���B"
Private Const targetFileYMWarnMsg As String = "�Ώۃt�@�C���̔N�����w�肵���N���ƈقȂ�܂��B��낵���ł����H"
Private Const workTimeEmptyInfoMsg As String = "�Ζ����Ԃ������͂̏ꍇ�A8��ݒ肵�܂��B"
Private Const endNoDataMsg As String = "�o�͂���f�[�^������܂���B"
Private Const endSuccessMsg As String = "�t�@�C���o�͂��������܂����B"

'�Ζ��\�t�@�C���ɏo��
Public Sub outputRoster()
    '��������
    Call commonUtil.startProcess
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim inputPlace As String: inputPlace = Range("�Ζ��n").Value
    Dim inputCd As String: inputCd = Range("�Ј��R�[�h").Value
    Dim inputUser As String: inputUser = Range("�L����").Value
    Dim inputWorkTime As String: inputWorkTime = Range("�Ζ�����").Value
    Dim inputYearMonth As String: inputYearMonth = Range("�N��").Value
    Dim targetFile As String: targetFile = Range("�Ώۃt�@�C��").Value
    
    Dim errorFlg As Boolean: errorFlg = False
    Dim errorMsg As String: errorMsg = ""
    Dim warnFlg As Boolean: warnFlg = False
    Dim warnMsg As String: warnMsg = ""
    Dim infoFlg As Boolean: infoFlg = False
    Dim infoMsg As String: infoMsg = ""
    
    Dim inputYearMonthFlg: inputYearMonthFlg = True
    Dim targetFileFlg: targetFileFlg = True
    
    '���̓`�F�b�N
    '�Ζ��n
    If inputPlace = Empty Then
        '�K�{�`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, placeErrorMsg)
    End If
    
    '�Ј��R�[�h
    If inputCd = Empty Then
        '�K�{�`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, cdErrorMsg)
    End If
    
    '�L����
    If inputUser = Empty Then
        '�K�{�`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, userErrorMsg)
    End If
    
    '�N��
    If inputYearMonth = Empty Then
        '�K�{�`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, yearMonthErrorMsg)
        inputYearMonthFlg = False
    ElseIf Not inputYearMonth Like "20[0-9][0-9][0-1][0-9]" Then
        '�`���`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, yearMonthPatternErrorMsg)
        inputYearMonthFlg = False
    Else
        '�ΏۃV�[�g���݃`�F�b�N
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
    
    '�Ώۃt�@�C��
    If targetFile = Empty Then
        '�K�{�`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, targetFileErrorMsg)
        targetFileFlg = False
    ElseIf Dir(targetFile) = Empty Then
        '�t�@�C�����݃`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, targetFileExistErrorMsg)
        targetFileFlg = False
    ElseIf Not Dir(targetFile) Like "20[0-9][0-9][0-1][0-9]sac[X0-9][X0-9][X0-9][X0-9][X0-9].xls" Then
        '�t�@�C�����`���`�F�b�N
        errorFlg = True
        errorMsg = commonUtil.createMsg(errorMsg, targetFileNameErrorMsg)
        targetFileFlg = False
    End If
    
    If errorFlg Then
        '�`�F�b�N�G���[
        MsgBox errorMsg, vbCritical
        GoTo endProcess
    End If
    
    If inputYearMonthFlg And targetFileFlg Then
        '�w�肵���N���ƑΏۃt�@�C���̔N�����قȂ�ꍇ�A�x��
        If inputYearMonth <> Left(Dir(targetFile), 6) Then
            warnFlg = True
            warnMsg = commonUtil.createMsg(warnMsg, targetFileYMWarnMsg)
        End If
    End If
    
    If warnFlg Then
        '�`�F�b�N���[�j���O
        Dim result As Long
        result = MsgBox(warnMsg, vbOKCancel)
        If result <> vbOK Then
            GoTo endProcess
        End If
    End If
    
    '�Ζ�����
    If inputWorkTime = Empty Then
        '�����͂̏ꍇ�A8�Ƃ���
        infoFlg = True
        infoMsg = commonUtil.createMsg(infoMsg, workTimeEmptyInfoMsg)
        inputWorkTime = "8"
        Range("�Ζ�����").Value = inputWorkTime
    End If
    
    If infoFlg Then
        '�`�F�b�N�C���t�H
        MsgBox infoMsg, vbInformation
    End If
    
    '�V�[�g�̈ړ�
    Worksheets(inputYearMonth).Activate
    
    '�f�[�^��ǂݍ���
    Dim i As Long
    Dim j As Long
    Dim str As String
    Dim dataArray(31, 4) As String
    i = 1
    j = 0
    Do While i < Range("�t�@�C���o�̓��X�g").Rows.Count
        Dim place As String
        Dim places As String
        Dim c As Variant
        Dim project As String
        
        '��Əꏊ���J���}��؂�ŘA������(�d���͔r��)
        places = ""
        For Each c In Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).MergeArea
            place = c.Offset(0, Range("��Əꏊ").Column - Range("��").Column)
            If place = "" Then
                GoTo continue
'            ElseIf place = "���̑�" Then
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
        
        '��Əꏊ�����͂���Ă���ꍇ�A���͂���Ɣ��f����
        If places <> "" Then
            dataArray(j, 0) = Format(Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).Value, "d")
            dataArray(j, 1) = Format(Range("�t�@�C���o�̓��X�g")(i, Range("�J�n").Column).Value, "hh:nn")
            dataArray(j, 2) = Format(Range("�t�@�C���o�̓��X�g")(i, Range("�I��").Column).Value, "hh:nn")
            '�x�ɔ���
            project = Range("�t�@�C���o�̓��X�g")(i, Range("�Č�").Column).Value
            If project = "�x��" Then
                dataArray(j, 3) = Range("�t�@�C���o�̓��X�g")(i, Range("��Ɠ��e").Column).Value
                dataArray(j, 4) = "�L�x"
            ElseIf project = "�Ċ��x��" Then
                dataArray(j, 3) = Range("�t�@�C���o�̓��X�g")(i, Range("��Ɠ��e").Column).Value
                dataArray(j, 4) = "���x"
            Else
                dataArray(j, 3) = places
            End If
            j = j + 1
        End If
        
        i = i + Range("�t�@�C���o�̓��X�g")(i, Range("��").Column).MergeArea.Rows.Count
    Loop
    
    '�f�[�^�Ȃ��̏ꍇ�A�I��
    If j = 0 Then
        MsgBox endNoDataMsg, vbInformation
        GoTo endProcess
    End If
    
    '�Ζ��\�ɏ����o��
    '�t�@�C�����J��
    Workbooks.Open (targetFile)
    
    Const dayCol As Long = 1
    Const startCol As Long = 3
    Const endCol As Long = 4
    Const holidayCol As Long = 13
    Const placeCol As Long = 14
    Const startRow As Long = 8
    Const endRow As Long = 38
    
    '�Ζ��n
    Const outputPlaceRow As Long = 3
    Const outputPlaceCol As Long = 3
    Cells(outputPlaceRow, outputPlaceCol).Value = inputPlace
    
    '�Ј��R�[�h
    Const outputCdRow As Long = 5
    Const outputCdCol As Long = 3
    Cells(outputCdRow, outputCdCol).Value = inputCd
    
    '�L����
    Const outputUserRow As Long = 61
    Const outputUserCol As Long = 15
    Cells(outputUserRow, outputUserCol).Value = inputUser
    
    '�Ζ�����
    Const outputWorkTimeRow As Long = 61
    Const outputWorkTimeCol As Long = 4
    Cells(outputWorkTimeRow, outputWorkTimeCol).Value = inputWorkTime
    
    '�o�Ё^�ގЎ��Ԃ��N���A
    Range(Cells(startRow, startCol), Cells(endRow, endCol)).Select
    Selection.ClearContents
    
    '�x�ɓ��^�s����N���A
    Range(Cells(startRow, holidayCol), Cells(endRow, placeCol)).Select
    Selection.ClearContents
    
    '�f�[�^�����[�v
    Dim k As Long
    Dim targetRow As Long
    Dim targetTime As Variant
    Dim outputMinute As Double
    Dim correctOutputMinute
    For k = 0 To j - 1
        targetRow = startRow - 1 + CLng(dataArray(k, 0))
        
        '30���P�ʂŏo�Ў��Ԃ͐؂�グ�^�ގЎ��Ԃ͐؂�̂�
        '�o�Ў���
        If dataArray(k, 1) <> "" Then
            targetTime = TimeValue(dataArray(k, 1))
            '���Ԃ��K�莞��(30���P��)�Œ���
            targetTime = commonUtil.roundTime(targetTime, commonConstants.timeDivide, True)
            Cells(targetRow, startCol).Value = targetTime
        End If
        
        '�ގЎ���
        If dataArray(k, 2) <> "" Then
            targetTime = TimeValue(dataArray(k, 2))
            '���Ԃ��K�莞��(30���P��)�Œ���
            targetTime = commonUtil.roundTime(targetTime, commonConstants.timeDivide, False)
            Cells(targetRow, endCol).Value = targetTime
        End If
        
        '�s��
        Cells(targetRow, placeCol).Value = dataArray(k, 3)
        
        '�x�ɓ�
        Cells(targetRow, holidayCol).Value = dataArray(k, 4)
    Next

    '�Ζ��\�t�@�C����ۑ����ĕ���
    Range("C5").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close

    '�������b�Z�[�W��\��
    MsgBox endSuccessMsg, vbInformation
    
endProcess:
    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    '�I������
    Call commonUtil.endProcess
End Sub

