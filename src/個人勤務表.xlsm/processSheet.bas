Attribute VB_Name = "processSheet"
Option Explicit
    
Private Const requiredErrorMsg As String = "����͂��Ă��������B"
Private Const numericErrorMsg As String = "�ɂ͐��l����͂��Ă��������B"
Private Const intErrorMsg As String = "�ɂ͐�������͂��Ă��������B"
Private Const yearRangeErrorMsg As String = "�N�ɂ�1900�`9999�̒l����͂��Ă��������B"
Private Const monthRangeErrorMsg As String = "���ɂ�1�`12�̒l����͂��Ă��������B"
Private Const yearMonthErrorMsg As String = "�Ώ۔N���̃V�[�g�����ɑ��݂��܂��B"
Private Const delConfirmMsg As String = "�Ώ۔N���̃V�[�g���폜���܂��B��낵���ł����H"
Private Const noSheetMsg As String = "�Ώ۔N���̃V�[�g�����݂��܂���B"

'�V�K�V�[�g���쐬
Public Sub createNewSheet()
    '��ʂ̕`���OFF�ɂ��܂�
    Call commonUtil.startProcess
    
    Dim yearVal As String: yearVal = Range("�N").Value
    Dim monthVal As String: monthVal = Range("��").Value
    
    '�N���ǂ��炩���̓`�F�b�N�Ɉ������������ꍇ�A�G���[
    If inputValidate(yearVal, monthVal) Then
        GoTo endProcess
    End If
    
    '�Ώ۔N���̃V�[�g�����݂���ꍇ�A�G���[
    Dim yearMonth As String: yearMonth = zeroAdd(yearVal, 4) & zeroAdd(monthVal, 2)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = yearMonth Then
            MsgBox yearMonthErrorMsg, vbCritical
            GoTo endProcess
        End If
    Next ws
    
    '�x�[�X�V�[�g����R�s�[
    Sheets("base").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = yearMonth
'    ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).Properties("_CodeName") = yearMonth
'    ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).Name = yearMonth
'    ActiveSheet.["_CodeName"] = yearMonth
    Application.Goto Reference:=ActiveWindow.ActiveSheet.Range("A1"), Scroll:=True
    
    '�C�x���g�̔����}�~����������
    Application.EnableEvents = True
    
    Range("�N").Value = yearVal
    Range("��").Value = monthVal
    
    '�N�����X�g�쐬
    createYMList
    
endProcess:
    '�I������
    Call commonUtil.endProcess
End Sub

'�V�[�g���폜
Public Sub deleteSheet()
    '��ʂ̕`���OFF�ɂ��܂�
    Call commonUtil.startProcess
    
    Dim yearVal As String: yearVal = Range("�N").Value
    Dim monthVal As String: monthVal = Range("��").Value
    
    '�N���ǂ��炩���̓`�F�b�N�Ɉ������������ꍇ�A�G���[
    If inputValidate(yearVal, monthVal) Then
        GoTo endProcess
    End If

    '�Ώ۔N���̃V�[�g�����݂���ꍇ�A�폜����
    Dim yearMonth As String: yearMonth = zeroAdd(yearVal, 4) & zeroAdd(monthVal, 2)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = yearMonth Then
            Dim rc As Integer
            rc = MsgBox(delConfirmMsg, vbYesNo + vbExclamation + vbDefaultButton2)
            If rc = vbYes Then
                '�V�[�g�폜
                Sheets(yearMonth).Delete
            End If
    
            '�N�����X�g�쐬
            createYMList
            
            GoTo endProcess
        End If
    Next ws
    
    '�Ώ۔N���̃V�[�g�����݂��Ȃ��ꍇ�A�G���[
    MsgBox noSheetMsg, vbCritical

endProcess:
    '�I������
    Call commonUtil.endProcess
End Sub

'�N�����X�g�쐬
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

    With Worksheets("main").Range("�N��").Validation
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

'���̓`�F�b�N
Private Function inputValidate(ByVal yearVal As String, ByVal monthVal As String)
    Dim errorFlg As Boolean: errorFlg = False
    Dim monthFlg As Boolean: monthFlg = False
    Dim errorMsg As String: errorMsg = ""
    
    '�N�̖����̓`�F�b�N
    If yearVal = Empty Then
        errorFlg = True
        errorMsg = createMsg(errorMsg, "�N" + requiredErrorMsg)
    Else
        '�N�̐��l�`�F�b�N
        If Not IsNumeric(yearVal) Then
            errorFlg = True
            errorMsg = createMsg(errorMsg, "�N" + numericErrorMsg)
        Else
            '�N�̐����`�F�b�N
            If Int(yearVal) <> yearVal Then
                errorFlg = True
                errorMsg = createMsg(errorMsg, "�N" + intErrorMsg)
            Else
                '�N�͈̔̓`�F�b�N
                If yearVal < 1900 Or yearVal > 9999 Then
                    errorFlg = True
                    errorMsg = createMsg(errorMsg, yearRangeErrorMsg)
                End If
            End If
        End If
    End If
    
    '���̖����̓`�F�b�N
    If monthVal = Empty Then
        errorFlg = True
        errorMsg = createMsg(errorMsg, "��" + requiredErrorMsg)
    Else
        '���̐��l�`�F�b�N
        If Not IsNumeric(monthVal) Then
            errorFlg = True
            errorMsg = createMsg(errorMsg, "��" + numericErrorMsg)
        Else
            '���̐����`�F�b�N
            If Int(monthVal) <> monthVal Then
                errorFlg = True
                errorMsg = createMsg(errorMsg, "��" + intErrorMsg)
            Else
                '���͈̔̓`�F�b�N
                If monthVal < 1 Or monthVal > 12 Then
                    errorFlg = True
                    errorMsg = createMsg(errorMsg, monthRangeErrorMsg)
                End If
            End If
        End If
    End If
    
    '�N���ǂ��炩�`�F�b�N�Ɉ������������ꍇ�A�G���[
    If errorFlg Then
        MsgBox errorMsg, vbCritical
    End If
    
    inputValidate = errorFlg
End Function

'���s��؂�Ń��b�Z�[�W���쐬
Private Function createMsg(ByVal msg As String, ByVal addMsg As String)
    If msg <> Empty Then
        msg = msg & vbCrLf
    End If
    
    createMsg = msg & addMsg
End Function

'���l�`��������̃[������
Private Function zeroAdd(ByVal str As String, ByVal length As Integer)
    Dim zeroStr As String: zeroStr = ""
    Dim i As Integer:
    
    For i = 0 To length
        zeroStr = zeroStr & "0"
    Next
    
    zeroAdd = Right(zeroStr & str, length)
End Function
