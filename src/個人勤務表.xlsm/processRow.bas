Attribute VB_Name = "processRow"
Option Explicit

'�s�ǉ�
Public Sub addRow()
    '��������
    Call commonUtil.startProcess
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    Dim startCol As Long: startCol = Range("��Əꏊ").Column
    Dim endCol As Long: endCol = Range("���l").Column
    Dim pasteRow As Long: pasteRow = targetCell.Row
    Dim copyRow As Long: copyRow = targetCell.Row + 1

    '�s�}��
    Rows(targetCell.Row).Insert Shift:=xlDown
    '��ɍs�}�������̂őΏۍs�̃f�[�^���R�s�y
    Range(Cells(copyRow, startCol), Cells(copyRow, endCol)).Copy
    With Range(Cells(pasteRow, startCol), Cells(pasteRow, endCol))
        '�l�\��t��
        .PasteSpecial _
         Paste:=xlPasteValues, _
         Operation:=xlNone, _
         SkipBlanks:=False, _
         Transpose:=False
        '�㉺�r��
        .Borders(xlEdgeTop).LineStyle = xlDot
        .Borders(xlEdgeBottom).LineStyle = xlDot
    End With
    Application.CutCopyMode = False
    '�N���A
    Range(Cells(copyRow, startCol), Cells(copyRow, endCol)).ClearContents

    '�s�ǉ��{�^���ǉ�
    With ActiveSheet.Buttons.Add(ActiveSheet.Shapes(Application.Caller).Left, _
                                 ActiveSheet.Shapes(Application.Caller).Top - targetCell.Height, _
                                 ActiveSheet.Shapes(Application.Caller).Width, _
                                 ActiveSheet.Shapes(Application.Caller).Height)
        .OnAction = "addRow"
        .Characters.Text = "�{"
        .Font.Name = "�l�r �S�V�b�N"
        .Font.Size = 10
    End With

    '�s�폜�{�^���ǉ�
    With ActiveSheet.Buttons.Add(ActiveSheet.Shapes(Application.Caller).Left + targetCell.Width, _
                                 ActiveSheet.Shapes(Application.Caller).Top - targetCell.Height, _
                                 ActiveSheet.Shapes(Application.Caller).Width, _
                                 ActiveSheet.Shapes(Application.Caller).Height)
        .OnAction = "delRow"
        .Characters.Text = "�|"
        .Font.Name = "�l�r �S�V�b�N"
        .Font.Size = 10
    End With

    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    '�I������
    Call commonUtil.endProcess
End Sub

'�s�폜
Public Sub delRow()
    '��������
    Call commonUtil.startProcess
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim targetCell As Range
    Set targetCell = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    Dim startCol As Long: startCol = Range("��Əꏊ").Column
    Dim endCol As Long: endCol = Range("���l").Column
    Dim copyRow As Long: copyRow = targetCell.Row
    Dim pasteRow As Long: pasteRow = targetCell.Row + 1
    
    '�{�^�����������s�̓��͂�����A���A���̍s�̓��͂��Ȃ��ꍇ�A���̍s�ɃR�s�[
    If WorksheetFunction.CountA(Range(Cells(copyRow, startCol), Cells(copyRow, endCol))) <> 0 Then
        If WorksheetFunction.CountA(Range(Cells(pasteRow, startCol), Cells(pasteRow, endCol))) = 0 Then
            '�l�\��t��
            Range(Cells(copyRow, startCol), Cells(copyRow, endCol)).Copy
            Range(Cells(pasteRow, startCol), Cells(pasteRow, endCol)).PasteSpecial _
                Paste:=xlPasteValues, _
                Operation:=xlNone, _
                SkipBlanks:=False, _
                Transpose:=False
        End If
    End If
    
    
    '�{�^���폜
    '�s�ǉ��{�^���폜
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        '���܂ɓ�̃h���b�v�_�E���������̂ł�������폜
        If shp.FormControlType = xlDropDown Then
            shp.Delete
        '�s�폜�{�^���̍��ׂ̃Z��
        ElseIf Not (Intersect(shp.TopLeftCell, targetCell.Offset(0, -1)) Is Nothing) Then
            shp.Delete
        End If
    Next
    
    '�s�폜�{�^���폜
    ActiveSheet.Shapes(Application.Caller).Delete
    '�s�폜
    Rows(targetCell.Row).Delete Shift:=xlUp

    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    '�I������
    Call commonUtil.endProcess
End Sub
