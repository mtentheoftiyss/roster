Attribute VB_Name = "setHoliday"
Option Explicit

'�x���̍s�̔w�i�F��ύX
Public Sub setHoliday(ByVal target As Range)
    '��������
    Call commonUtil.startProcess
    
    '�N���̕ύX���̂ݎ��s
    If Intersect(target, Range("�N��")) Is Nothing Then
        GoTo endProcess
    End If
    
    '�N���ǂ�������͎��̂ݎ��s
    If WorksheetFunction.CountA(Range("�N��")) < 2 Then
        GoTo endProcess
    End If
    
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    '�ŏI��̓���
    Dim lastCol: lastCol = Range("���l")(Range("���l").Count).Column
    
    '�ŏI�s�̓���
    Dim lastRow As Long: lastRow = Range("�J�n�E�I�����ԃ��X�g")(Range("�J�n�E�I�����ԃ��X�g").Count).Row
    
    '�擪�s
    Dim firstRow As Long: firstRow = Range("��_").Row
    '�ŏ��ɏ��X�N���A����
    With Range(Cells(firstRow, 1), Cells(lastRow, lastCol))
        '�h��Ԃ��̐F
        .Interior.ColorIndex = xlNone
        With .Font
            '�t�H���g�̐F
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            '����
            .Bold = False
            '�Α�
            .Italic = False
            '����
            .Underline = xlUnderlineStyleNone
            '��������
            .Strikethrough = False
        End With
    End With
    
    '�擪�s����ŏI�s�܂Ń��[�v
    Dim i As Long
    Dim dayCnt As Integer: dayCnt = 0
    Dim dayFlg As Boolean: dayFlg = False
    Dim baseCol As Long: baseCol = Range("��_").Column + 1
    For i = firstRow To lastRow
        Dim cellFormula As String: cellFormula = Cells(i, baseCol).Formula
        Dim cellValue As String: cellValue = Cells(i, baseCol).Value
        Dim cellObject As Range: Set cellObject = Cells(i, baseCol)
        Dim holidayFlg As Boolean: holidayFlg = False
        Dim holidayFind As Range
        Dim holidayRow As Long
        Dim holidayNameCell As Range
        
        '�v�Z������̏ꍇ�A��s�̂��̂��g�p����
        Dim j As Long: j = i
        Do While cellFormula = ""
            j = j - 1
            cellFormula = Cells(j, baseCol).Formula
            cellValue = Cells(j, baseCol).Value
            Set cellObject = Cells(j, baseCol)
            dayFlg = True
        Loop
        If cellValue <> "" Then
            holidayFlg = isHoliday(cellObject)
        Else
            '31���ɖ����Ȃ�����31��
            holidayFlg = True
        End If
        '�x���͓h��Ԃ�
        If holidayFlg Then
            Range(Cells(i, 1), Cells(i, lastCol)).Interior.ColorIndex = 16
        Else
            If Not dayFlg Then
                dayCnt = dayCnt + 1
            End If
        End If
        '�j������l�ɓ���
        If holidayFlg Then
            If Not dayFlg Then
                If cellValue <> "" Then
                    Set holidayFind = Range("�j�����X�g").Find(What:=CDate(cellValue), LookIn:=xlValues, lookAt:=xlWhole)
                    If Not holidayFind Is Nothing Then
                        holidayRow = holidayFind.Row
                        Set holidayNameCell = Range("�j�����X�g").Resize(1, 1).Offset(holidayRow - Range("�j�����X�g").Row, -1)
                        Cells(cellObject.Row, Range("���l").Column).Value = holidayNameCell.Value
                    End If
                End If
            End If
        End If
        dayFlg = False
    Next
    
    '�����ݒ�
    Range("����").Value = dayCnt
    
    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    
endProcess:
    '�I������
    Call commonUtil.endProcess
End Sub

'�x������
Private Function isHoliday(ByVal targetCell As Range)
    Dim returnValue As Boolean: returnValue = False
    
    '�j��
    If WorksheetFunction.CountIf(Range("�j�����X�g"), targetCell) Then
        returnValue = True
    End If
    '���j��
    If Weekday(targetCell.Value) = 1 Then
        returnValue = True
    End If
    '�y�j��
    If Weekday(targetCell.Value) = 7 Then
        returnValue = True
    End If
    isHoliday = returnValue
End Function




