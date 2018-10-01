Attribute VB_Name = "processFolder"
Option Explicit

'�Œ�l�t�B�[���h
Private projectName As String
Private teamName As String
Private businessFunctionName As String
Private checkUser As String
Private checkDate As String

'�Z���t�`�F�b�N���X�g(RD)
Public Sub selfCheckRD()
    Call processFolder("selfCheckRD")
End Sub

'���r���[�˗���(RD)
Public Sub reviewRequestRD()
    Call processFolder("reviewRequestRD")
End Sub

'�Z���t�`�F�b�N���X�g(ED)
Public Sub selfCheckED()
    Call processFolder("selfCheckED")
End Sub

'���r���[�˗���(ED)
Public Sub reviewRequestED()
    Call processFolder("reviewRequestED")
End Sub

'DB�݌v��
Public Sub dbLayout()
    Call processFolder("dbLayout")
End Sub

'�v����`
Public Sub requirementDefinition()
    Call processFolder("requirementDefinition")
End Sub

'�t�H���_�[����
Public Sub processFolder(ByVal processMode As String)
    '��������
    Call commonUtil.startProcess
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim objFSO As FileSystemObject
    Dim targetFolder As String
    
    '�t�@�C���V�X�e���I�u�W�F�N�g����
    Set objFSO = New FileSystemObject
    
    '���[�g�w��
    targetFolder = Range("�Ώۃt�H���_")
    
    '�T���J�n
    Call searchSubFolder(objFSO.GetFolder(targetFolder), processMode)
    
    '�I�u�W�F�N�g��j��
    Set objFSO = Nothing
    
    '�������b�Z�[�W��\��
    MsgBox "�������I�����܂����B", vbInformation
    
    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    '�I������
    Call commonUtil.endProcess
End Sub

'�T�u�t�H���_�[����
Private Sub searchSubFolder(ByVal objFOLDER As Folder, ByVal processMode As String)
    Dim objSubFOLDER As Folder
    Dim objFILE As File
    
    '****************************************************************************************************
    '���C��������������
    '****************************************************************************************************
    
    '�Œ�l���t�B�[���h�Ɋi�[
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
    '���C�����������܂�
    '****************************************************************************************************
    
    '�T�u�t�H���_�[��T��
    For Each objSubFOLDER In objFOLDER.SubFolders
        Call searchSubFolder(objSubFOLDER, processMode)
    Next objSubFOLDER
    
    '�I�u�W�F�N�g��j��
    Set objFOLDER = Nothing
End Sub

'�Œ�l���t�B�[���h�Ɋi�[
Private Sub setPrefix()
projectName = Range("�v���W�F�N�g��").Value
teamName = Range("�`�[����").Value
businessFunctionName = Range("�Ɩ��@�\��").Value
checkUser = Range("�`�F�b�N���{��").Value
checkDate = Range("�`�F�b�N���{��").Value

End Sub

'�Z���t�`�F�b�N���X�g(RD)
Private Sub selfCheckRDMain(ByVal objFILE As File)
    Dim targetSheet As String
    Dim targetCell As String
    Dim fixedValue As String
    
    '�Z���t�`�F�b�N���X�g�̂ݑΏۂƂ���
    Dim targetFileName As String
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "�Z���t�`�F�b�N���X�g") > 0 Then
        Exit Sub
    End If
    
    '��ʂ��ǂ����̔���
    Dim functionKind As Integer
    functionKind = 0
    If targetFileName Like "D11-F02_2[1-5]_RD*" Then
        functionKind = 1
    ElseIf targetFileName Like "D11-F02_2[6-9]_RD*" Or targetFileName Like "D11-F02_3[0-3]_RD*" Then
        functionKind = 2
    Else
        functionKind = 3
    End If
    
    '�������ރV�[�g�A�Z���A�l���w��
    targetSheet = "�v����`���ʕ� �Z���t�`�F�b�N���X�g"
    targetCell = "C6"
    Dim functionId As String
    Dim startPosition As Integer
    Dim endPosition As Integer
    Dim functionName As String
    
    Select Case functionKind
        '�ʏ���
        Case 1
            '�@�\ID
            functionId = Mid(targetFileName, 15, 5)
            '�@�\���J�n�ʒu
            startPosition = InStr(20, targetFileName, "_") + 1
            '�@�\���I���ʒu
            endPosition = InStrRev(targetFileName, "_")
            '�@�\��
            functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
            '�݌v���t�@�C����
            fixedValue = "ES0303-F01-PTN2_" & functionId & "_�@�\�݌v��(" & functionName & ").xlsx"
        
        '�⏕���
        Case 2
            '�@�\ID
            functionId = Mid(targetFileName, 15, 5)
            '�@�\���J�n�ʒu
            startPosition = InStr(20, targetFileName, "_") + 1
            '�@�\���I���ʒu
            endPosition = InStrRev(targetFileName, "_")
            '�@�\��
            functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
            '�݌v���t�@�C����
            fixedValue = "ES0303-F01-PTN2_�@�\�݌v��_" & functionId & "_" & functionName & ".xlsx"
        
        '�o�b�`
        Case 3
            '�@�\ID
            functionId = Mid(targetFileName, 15, 8)
            '�@�\���J�n�ʒu
            startPosition = InStr(23, targetFileName, "_") + 1
            '�@�\���I���ʒu
            endPosition = InStrRev(targetFileName, "_")
            '�@�\��
            functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
            '�݌v���t�@�C����
            fixedValue = "ES0303-F02-PTN2_" & functionId & "_�@�\�݌v��(" & functionName & ").xlsm"
    End Select
    
    '�t�@�C�����J��
    Workbooks.Open (objFILE.Path)
    '��������
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell).Value = fixedValue
    '�t�@�C����ۑ����ĕ���
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'���r���[�˗���(RD)
Private Sub reviewRequestRDMain(ByVal objFILE As File)
    Dim targetSheet As String
    Dim targetCell As String
    Dim targetCell7 As String
    Dim fixedValue As String
    Dim fixedValue7 As String
    Dim targetFileName As String
    
    '���r���[�˗����̂ݑΏۂƂ���
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "���r���[�˗������񍐏�") > 0 Then
        Exit Sub
    End If
    
    '��ʂ��ǂ����̔���
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
    
    '�������ރV�[�g�A�Z���A�l���w��
    targetSheet = "���r���[�˗������񍐏�"
    targetCell = "�@�\��"
    targetCell7 = "���r���[�Ǘ��ԍ�"
    
    '�J�n�ʒu
    Dim startPosition As Integer
    startPosition = InStr(fixedStartPosition, targetFileName, "_") + 1
    '�I���ʒu
    Dim endPosition As Integer
    endPosition = InStrRev(targetFileName, "_")
    '�@�\��
    fixedValue = Mid(targetFileName, startPosition, endPosition - startPosition)
    '���r���[�Ǘ��ԍ�
    fixedValue7 = Mid(targetFileName, 12, 2) & "_" & Mid(targetFileName, 9, 2) & "_" & Mid(targetFileName, 15, fixedIdLength)
    
    '�t�@�C�����J��
    Workbooks.Open (objFILE.Path)
    '��������
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell).Value = fixedValue
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell7).Value = fixedValue7
    '�t�@�C����ۑ����ĕ���
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'�Z���t�`�F�b�N���X�g(ED)
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
    
    '�Z���t�`�F�b�N���X�g�̂ݑΏۂƂ���
    Dim targetFileName As String
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "�Z���t�`�F�b�N���X�g") > 0 Then
        Exit Sub
    End If
    
    '��ʂ��ǂ����̔���
    Dim funcClass As String
    Dim funcPrefix As String
    If targetFileName Like "D11-F02_21_ED*" Then
        funcClass = "���"
        funcPrefix = "ES0303-F01-PTN2_"
    Else
        funcClass = "���[��"
        funcPrefix = "ES0302-F13_"
    End If
    
    '�������ރV�[�g�A�Z���A�l���w��
    targetSheet = "��{�݌v(�V�X�e���݌v�E�O���݌v�EAP��ՃZ���t�`�F�b�N���X�g"
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
    
    '�v���W�F�N�g��
    fixedValue1 = projectName
    '�`�[����
    fixedValue2 = teamName
    '�`�F�b�N���{��
    fixedValue3 = checkUser
    '�Ɩ���
    fixedValue4 = "�@�\�݌v���i" & funcClass & "�j"
    '�Ɩ��@�\��
    fixedValue5 = businessFunctionName
    '�`�F�b�N���{��
    fixedValue6 = checkDate
    
    '�@�\ID
    functionId = Mid(targetFileName, 15, 5)
    '�@�\���J�n�ʒu
    startPosition = InStr(20, targetFileName, "_") + 1
    '�@�\���I���ʒu
    endPosition = InStr(startPosition, targetFileName, "_")
    '�@�\��
    functionName = Mid(targetFileName, startPosition, endPosition - startPosition)
    '�݌v���t�@�C����
    fixedValue7 = funcPrefix & functionId & "_�@�\�݌v��(" & functionName & ").xlsx"
    
    '�t�@�C�����J��
    Workbooks.Open (objFILE.Path)
    '��������
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell1).Value = fixedValue1   '�v���W�F�N�g��
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell2).Value = fixedValue2   '�`�[����
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell3).Value = fixedValue3   '�`�F�b�N���{��
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell4).Value = fixedValue4   '�Ɩ���
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell5).Value = fixedValue5   '�Ɩ��@�\��
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell6).Value = fixedValue6   '�`�F�b�N���{��
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell7).Value = fixedValue7   '�݌v���t�@�C����
    '�t�@�C����ۑ����ĕ���
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'���r���[�˗���(ED)
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
    
    '���r���[�˗����̂ݑΏۂƂ���
    targetFileName = objFILE.Name
    If Not InStr(targetFileName, "���r���[�˗������񍐏�") > 0 Then
        Exit Sub
    End If
    
    
    '��ʂ��ǂ����̔���
    Dim funcClass As String
    Dim fixedStartPosition As Integer
    Dim fixedIdLength As Integer
    If targetFileName Like "D11-F04_21_ED*" Then
        funcClass = "���"
        fixedStartPosition = 20
        fixedIdLength = 5
    Else
        funcClass = "���[��"
        fixedStartPosition = 20
        fixedIdLength = 5
    End If
    
    
    '�������ރV�[�g�A�Z���A�l���w��
    targetSheet = "���r���[�˗������񍐏�"
    targetCell1 = "�v���W�F�N�g��"
    targetCell2 = "�`�[����"
    targetCell3 = "�Ώۍ\���Ǘ���"
    targetCell4 = "�Ɩ��@�\��"
    targetCell5 = "�Ώې��ʕ���"
    targetCell6 = "�@�\��"
    targetCell7 = "���r���[�Ǘ��ԍ�"
    targetCell8 = "�y�[�W��"
    
    '�J�n�ʒu
    Dim startPosition As Integer
    startPosition = InStr(fixedStartPosition, targetFileName, "_") + 1
    '�I���ʒu
    Dim endPosition As Integer
    endPosition = InStr(startPosition, targetFileName, "_")
    
    '�v���W�F�N�g��
    fixedValue1 = projectName
    '�`�[����
    fixedValue2 = teamName
    '�Ώۍ\���Ǘ���
    fixedValue3 = "�@�\�݌v���i" & funcClass & "�j"
    '�Ɩ��@�\��
    fixedValue4 = businessFunctionName
    '�Ώې��ʕ���
    fixedValue5 = funcClass & "��`��(" & Mid(targetFileName, 12, 2) & "_" & Mid(targetFileName, 9, 2) & "_" & Mid(targetFileName, 15, fixedIdLength) & ")"
    '�@�\��
    fixedValue6 = Mid(targetFileName, startPosition, endPosition - startPosition)
    '���r���[�Ǘ��ԍ�
    fixedValue7 = Mid(targetFileName, 12, 2) & "_" & Mid(targetFileName, 9, 2) & "_" & Mid(targetFileName, 15, fixedIdLength)
    '�y�[�W��
    Dim wb As Workbook
    Set wb = Workbooks.Open("D:\zz_endo-work\�R�}���h\Windows�R�}���h\copy\�y�[�W���J�E���g.xlsx")
    fixedValue8 = Application.WorksheetFunction.VLookup(Mid(targetFileName, startPosition, endPosition - startPosition), Range("B1:C40"), 2, False)
    wb.Close
    
    '�t�@�C�����J��
    Workbooks.Open (objFILE.Path)
    '��������
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell1).Value = fixedValue1   '�v���W�F�N�g��
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell2).Value = fixedValue2   '�`�[����
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell3).Value = fixedValue3   '�Ώۍ\���Ǘ���
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell4).Value = fixedValue4   '�Ɩ��@�\��
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell5).Value = fixedValue5   '�Ώې��ʕ���
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell6).Value = fixedValue6   '�@�\��
'    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell7).Value = fixedValue7   '���r���[�Ǘ��ԍ�
    ActiveWorkbook.Worksheets(targetSheet).Range(targetCell8).Value = fixedValue8   '�y�[�W��
    '�t�@�C����ۑ����ĕ���
    ActiveWorkbook.Worksheets(targetSheet).Range("A1").Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'�v����`
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

    '�\��
    targetSheet1 = "�\��"
    targetCells1 = "H19:H20"
    targetCells2 = "AD19:AD20"
    '�쐬�N����
    targetCell1 = "H19"
    fixedValue1 = "2018/9/27"
    '�쐬��
    targetCell2 = "AD19"
    fixedValue2 = "SCSK"
    
    '��������
    targetSheet2 = "��������"
    targetCells3 = "C3:G30"
    '�ύX��
    targetCell3 = "C3"
    fixedValue3 = "2018/9/27"
    '�ύX��
    targetCell4 = "D3"
    fixedValue4 = "SCSK"
    '�ύX���e
    targetCell5 = "F3"
    fixedValue5 = "�V�K�쐬"
    
    '�t�@�C�����J��
    Workbooks.Open (objFILE.Path)
    
    '��������
    For Each ws In Worksheets
        If ws.Visible = xlSheetVisible Then
            '�\��
            If ws.Name = targetSheet1 Then
                '�ŏ��ɃN���A
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCells1).Value = ""
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCells2).Value = ""
            
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCell1).Value = fixedValue1   '�쐬�N����
                ActiveWorkbook.Worksheets(targetSheet1).Range(targetCell2).Value = fixedValue2   '�쐬��
            End If
        
            '��������
            If ws.Name = targetSheet2 Then
                '�ŏ��ɃN���A
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCells3).Value = ""
            
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCell3).Value = fixedValue3   '�ύX��
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCell4).Value = fixedValue4   '�ύX��
                ActiveWorkbook.Worksheets(targetSheet2).Range(targetCell5).Value = fixedValue5   '�ύX���e
            End If
            
            '����
            ActiveWorkbook.Worksheets(ws.Name).Activate
            ActiveWorkbook.Worksheets(ws.Name).Select
            Application.Goto Reference:=Range("A1"), Scroll:=True
        End If
    Next
    
    '�t�@�C����ۑ����ĕ���
    ActiveWorkbook.Worksheets(1).Select
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

'�y�[�W���J�E���g���C��
Public Sub pageCountMain()
    '��������
    Call commonUtil.startProcess
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    Dim objFSO As FileSystemObject
    Dim targetFolder As String
    Dim resultCollection As Collection
    Dim rowCnt As Integer
    rowCnt = 1
    
    '�t�@�C���V�X�e���I�u�W�F�N�g����
    Set objFSO = New FileSystemObject
    
    '���[�g�w��
    targetFolder = Range("�Ώۃt�H���_")
    
    '�T���J�n
    Set resultCollection = New Collection
    Set resultCollection = pageCountSub(objFSO.GetFolder(targetFolder), resultCollection)
    
    If resultCollection.Count > 0 Then
        '�V�K�u�b�N�쐬
        Dim wb As Workbook
        Set wb = Workbooks.Add
        Dim i As Integer
        
        '���o��
        wb.Sheets(1).Cells(rowCnt, 1).Value = "�t�H���_��"
        wb.Sheets(1).Cells(rowCnt, 2).Value = "�t�@�C����"
        wb.Sheets(1).Cells(rowCnt, 3).Value = "�V�[�g��"
        wb.Sheets(1).Cells(rowCnt, 4).Value = "�y�[�W��"
        wb.Sheets(1).Cells(rowCnt, 5).Value = "��\��"
        
        For i = 1 To resultCollection.Count
            rowCnt = rowCnt + 1
            wb.Sheets(1).Cells(rowCnt, 1).Value = resultCollection(i)(0)
            wb.Sheets(1).Cells(rowCnt, 2).Value = resultCollection(i)(1)
            wb.Sheets(1).Cells(rowCnt, 3).Value = resultCollection(i)(2)
            wb.Sheets(1).Cells(rowCnt, 4).Value = resultCollection(i)(3)
            wb.Sheets(1).Cells(rowCnt, 5).Value = resultCollection(i)(4)
        Next i
        
        '�񕝂𒲐�
        wb.Sheets(1).Range("A:D").Columns.AutoFit
        
        '�u�b�N��ۑ�
        wb.SaveAs fileName:=targetFolder & "\�y�y�[�W���J�E���g�z.xlsx"
        wb.Close
    End If
    
    '�I�u�W�F�N�g��j��
    Set objFSO = Nothing
    
    '�������b�Z�[�W��\��
    MsgBox "�������I�����܂����B", vbInformation
    
    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    '�I������
    Call commonUtil.endProcess
End Sub

'�y�[�W���J�E���g�T�u
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
        
        '�B���t�@�C���ƈꎞ�t�@�C���͏��O����
        '�J�E���g����̂�EXCEL�t�@�C���̂ݑΏۂƂ���
        ext = LCase(objFSO.GetExtensionName(objFILE.Path))
        fileAttr = GetAttr(objFILE.Path)
        If ((fileAttr And vbHidden) = False And re.test(objFILE.Name) = False) And (ext = "xls" Or ext = "xlsx" Or ext = "xlsm") Then
            '�t�@�C�����J��
            Workbooks.Open (objFILE.Path)
            
            '�y�[�W�J�E���g
            For Each ws In ActiveWorkbook.Worksheets
                collectionItem(4) = ""
                If ws.Visible <> xlSheetVisible Then
                    ws.Visible = xlSheetVisible
                    collectionItem(4) = "��"
                End If
                
                ws.Activate
                ActiveWindow.View = xlPageBreakPreview
                collectionItem(2) = ws.Name
'                collectionItem(3) = Application.ExecuteExcel4Macro("get.document(50)")
                collectionItem(3) = ws.PageSetup.Pages.Count
                objCollection.Add collectionItem
            Next ws
        
            '�t�@�C�������
            ActiveWorkbook.Close
        Else
            collectionItem(2) = "-"
            collectionItem(3) = "-"
            collectionItem(4) = ""
            objCollection.Add collectionItem
        End If
    Next objFILE
    
    '�T�u�t�H���_�[��T��
    For Each objSubFOLDER In objFOLDER.SubFolders
        Set objCollection = pageCountSub(objSubFOLDER, objCollection)
    Next objSubFOLDER
    
    '�I�u�W�F�N�g��j��
    Set objFOLDER = Nothing
    Set objFSO = Nothing
    
    '�R���N�V������ԋp
    Set pageCountSub = objCollection
End Function

'DB�݌v��
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
    
    '�t�@�C�����J��
    Set wb = Workbooks.Open(objFILE.Path)
    
    '�V�[�g�����[�v
    For Each ws In wb.Worksheets
        '�\���V�[�g�݂̂�ΏۂƂ���
        If ws.Visible = xlSheetVisible Then
            '�V�[�g���𔻒�
            If ws.Name <> "�\��" And ws.Name <> "��������" And ws.Name <> "ER�}" And ws.Name <> "�_��ER�}" And ws.Name <> "�_���G���e�B�e�B�ꗗ" And ws.Name <> "����������" And ws.Name <> "���n�E�X�R�����g" And ws.Name <> "�C���f�b�N�X��`" And ws.Name <> "�f�[�^�r���[�ꗗ" And ws.Name <> "�f�[�^�r���[�E�G���e�B�e�B�֘A��`" And ws.Name <> "SAC�Œǉ��������ڂ̗���" Then
                '�G���e�B�e�B������������
                entityVal = ws.Cells(entityRow, entityCol).Value
                ws.Cells(entityRow, entityCol).Value = Replace(entityVal, "_", "_MYP_", 1, 1)
                
                '�ŏI�s���擾
                lastRow = ws.Cells(startRow, typeCol).End(xlDown).Row
                
                For targetRow = startRow To lastRow
                    '�f�[�^�^�̒u��
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
                    
                    '�f�[�^����ID�Ƀv���t�B�b�N�X��t��
                    idVal = ws.Cells(targetRow, idCol).Value
                    'FastAPP���ʍ��ڂ͑ΏۊO
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
                    
                    '�Z���̏�������
                    ws.Cells(targetRow, typeCol).Value = afterTypeVal
                    ws.Cells(targetRow, idCol).Value = idVal
                Next targetRow
                
'                '�Z���I���ʒu��߂�
'                ws.Range("A1").Select
            End If
        End If
    Next ws
    
    '�t�@�C����ۑ����ĕ���
'    wb.Worksheets(1).Range("A1").Select
    wb.Save
    wb.Close
End Sub


