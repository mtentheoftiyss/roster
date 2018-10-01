Attribute VB_Name = "commonUtil"
Option Explicit

'��������
Public Sub startProcess()
    '�����v�Z���蓮�ɂ���
    Application.Calculation = xlCalculationManual
    '��ʂ̕`���OFF�ɂ���
    Application.ScreenUpdating = False
    '�C�x���g�̔�����}�~����
    Application.EnableEvents = False
    '�m�F���b�Z�[�W���\���ɂ���
    Application.DisplayAlerts = False
End Sub

'�I������
Public Sub endProcess()
    '�m�F���b�Z�[�W��\���ɂ���
    Application.DisplayAlerts = True
    '�C�x���g�̔����}�~����������
    Application.EnableEvents = True
    '��ʂ̕`���ON�ɂ���
    Application.ScreenUpdating = True
    '�����v�Z�������ɂ���
    Application.Calculation = xlCalculationAutomatic
End Sub

'�s�I�𔻒�
Public Function isRowSelect(ByVal target As Range)
    Dim returnValue As Boolean: returnValue = False
    If target.Address = target.EntireRow.Address Then
        returnValue = True
    End If
    
    isRowSelect = returnValue
End Function

'��I�𔻒�
Public Function isColumnSelect(ByVal target As Range)
    Dim returnValue As Boolean: returnValue = False
    If target.Address = target.EntireColumn.Address Then
        returnValue = True
    End If
    
    isColumnSelect = returnValue
End Function

'�扽�j���̓��t���擾
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

'���s��؂�Ń��b�Z�[�W���쐬
Public Function createMsg(ByVal msg As String, ByVal addMsg As String)
    If msg <> Empty Then
        msg = msg & vbCrLf
    End If
    
    createMsg = msg & addMsg
End Function

'�捞�t�@�C���w��
Public Sub setInputFile()
    Dim targetCell As Range
    Dim openFileName As String
    openFileName = Application.GetOpenFilename("�e�L�X�g,*.txt")
    If openFileName <> "False" Then
        Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(0, -1)
        If targetCell.MergeCells Then
            targetCell.MergeArea.Offset(0, 0).Value = openFileName
        Else
            targetCell.Value = openFileName
        End If
    End If
End Sub

'�Ώۃt�@�C���w��
Public Sub setTargetFile()
    Dim targetCell As Range
    Dim openFileName As String
    openFileName = Application.GetOpenFilename("���[�N�V�[�g,*.xls")
    If openFileName <> "False" Then
        Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(0, -1)
        If targetCell.MergeCells Then
            targetCell.MergeArea.Offset(0, 0).Value = openFileName
        Else
            targetCell.Value = openFileName
        End If
    End If
End Sub

'�Ώۃt�H���_�w��
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

'���Ԃ�؂�グ�^�؂�̂�
Public Function roundTime(ByVal targetTime As Double, ByVal timeDivide As Long, ByVal isRoundUp As Boolean)
    Dim targetMinute As Double
    Dim correctMinute As Double
    
    '�����擾
    targetMinute = Minute(targetTime)
    '�P�ʎ��ԂŊ���؂��ꍇ�͒����s�v
    If targetMinute Mod timeDivide <> 0 Then
        '0�`29����0�A30�`59����30�ɕϊ�(30���P�ʂ̏ꍇ)
        correctMinute = (targetMinute \ timeDivide) * timeDivide
        '���̕������������āA�P�ʎ��Ԃɕϊ������l��������
        targetTime = DateAdd("n", correctMinute, DateAdd("n", targetMinute * -1, targetTime))
        '�؂�グ�̏ꍇ�A�X�ɒP�ʎ��Ԃ�������
        If (isRoundUp) Then
            targetTime = DateAdd("n", timeDivide, targetTime)
        End If
    End If
    
    roundTime = targetTime
End Function

'URL�G���R�[�h
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

'URL�f�R�[�h
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

'IE���J��
Public Sub openIE(strUrl As String)
    Const navOpenInNewTab = &H800
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate strUrl, &H800
End Sub

'�N���������ɂ���
Public Sub setThisMonth()
    Call setTargetMonth(Date)
End Sub

'�N���𗈌��ɂ���
Public Sub setNextMonth()
    Call setTargetMonth(DateAdd("m", 1, Date))
End Sub

'�N�����w�茎�ɕύX����
Private Sub setTargetMonth(ByVal target As Date)
    Range("�N").Value = Format(target, "yyyy")
    Range("��").Value = Format(target, "m")
    
    Dim thisMonth As String: thisMonth = Format(target, "yyyymm")
    Dim beforeMonth As String: beforeMonth = Range("�N��").Value
    Range("�N��") = Replace(Range("�N��"), beforeMonth, thisMonth)
    Range("�Ώۃt�@�C��") = Replace(Range("�Ώۃt�@�C��"), beforeMonth, thisMonth)
End Sub

'��\���̖��O��`��\��
Public Sub showInvisibleNames()
    Dim oName As Object
    For Each oName In Names
        If oName.Visible = False Then
            oName.Visible = True
        End If
    Next
    MsgBox "��\���̖��O��`��\�����܂���", vbOKOnly
End Sub
