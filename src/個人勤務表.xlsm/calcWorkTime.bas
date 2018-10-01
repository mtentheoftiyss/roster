Attribute VB_Name = "calcWorkTime"
Option Explicit

'�Ζ����Ԃ��v�Z
Public Sub calcWorkTime(ByVal target As Range)
    '��ʂ̕`���OFF�ɂ��܂�
    Call commonUtil.startProcess
    
    Dim startCol As Long: startCol = Range("�J�n").Column
    Dim endCol As Long: endCol = Range("�I��").Column
    Dim sacWorkCol As Long: sacWorkCol = Range("SAC�Ζ�����").Column
    Dim siteWorkCol As Long: siteWorkCol = Range("����Ζ�����").Column
    
    '�I���ʒu���擾
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    '�J�n�A�I���̕ύX���̂ݎ��s
    If Intersect(target, Range("�J�n�E�I�����ԃ��X�g")) Is Nothing Then
        GoTo endProcess
    End If
    
    Dim targetCell As Range
    For Each targetCell In target
        '�J�n�A�I���̕ύX���̂ݎ��s
        If Intersect(target, Range("�J�n�E�I�����ԃ��X�g")) Is Nothing Then
            GoTo continue
        End If
        
        '�J�n�A�I�����Ȃ��Ƃ��ǂ��炩�������͂̏ꍇ�͋Ζ����Ԃ��폜
        If Cells(targetCell.Row, startCol).Value = "" Or Cells(targetCell.Row, endCol).Value = "" Then
            Cells(targetCell.Row, sacWorkCol).Value = ""
            Cells(targetCell.Row, siteWorkCol).Value = ""
            GoTo continue
        End If
        
        '�x�e���Ԃ̎Z�o
        Dim startTime As Double: startTime = Cells(targetCell.Row, startCol).Value
        Dim endTime As Double: endTime = Cells(targetCell.Row, endCol).Value
        '���Ԃ��K�莞��(30���P��)�Œ���
        startTime = commonUtil.roundTime(startTime, commonConstants.timeDivide, True)
        endTime = commonUtil.roundTime(endTime, commonConstants.timeDivide, False)
        Cells(targetCell.Row, startCol).Value = startTime
        Cells(targetCell.Row, endCol).Value = endTime
        '�J�n�ƏI�����t�]���Ă���ꍇ�A�I����24���ԉ��Z
        If startTime > endTime Then
            endTime = calc24(endTime)
            Cells(targetCell.Row, endCol).Value = endTime
        End If
        '�}�N����DAY�֐����g���ƁA���܂��l�����Ȃ��̂ň�U�Z���Ɏ��𖄂ߍ���
        Cells(targetCell.Row, sacWorkCol).Formula = "=DAY(" + Cells(targetCell.Row, endCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) + ")"
        Cells(targetCell.Row, siteWorkCol).Formula = "=DAY(" + Cells(targetCell.Row, endCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) + ")"
        Dim startCell As Range: Set startCell = Cells(targetCell.Row, startCol)
        Dim endCell As Range: Set endCell = Cells(targetCell.Row, endCol)
        Dim sacCell As Range: Set sacCell = Cells(targetCell.Row, sacWorkCol)
        Dim siteCell As Range: Set siteCell = Cells(targetCell.Row, siteWorkCol)
        '�x�e���Ԃ̎Z�o
        Dim sacRestTime As Double: sacRestTime = calcRestTime(startCell, endCell, sacCell, True)
        Dim siteRestTime As Double: siteRestTime = calcRestTime(startCell, endCell, siteCell, False)
        
        '�Ζ����Ԃ̎Z�o
        Worksheets(initSelection.getInitSheet).Activate
        Cells(targetCell.Row, sacWorkCol).Value = (endTime - startTime - sacRestTime)
        Cells(targetCell.Row, siteWorkCol).Value = (endTime - startTime - siteRestTime)
        
        '����Ζ����Ԃ��玩�Е���������
        Dim sacTime As Double: sacTime = 0
        Dim ma As Range
        For Each ma In targetCell.MergeArea
            Dim wp As Range: Set wp = Cells(ma.Row, Range("��Əꏊ").Column)
            Dim wpl As Range
            For Each wpl In Range("��Əꏊ���X�g")
                If wp.Value = wpl.Value Then
                    Dim jisha As Range: Set jisha = wpl.Offset(0, 1)
                    If jisha.Value = "����" Then
                        sacTime = sacTime + Cells(ma.Row, Range("�H��").Column).Value / 24
                    End If
                End If
            Next wpl
        Next ma
        Cells(targetCell.Row, siteWorkCol).Value = Cells(targetCell.Row, siteWorkCol).Value - sacTime

continue:
    Next targetCell
    
endProcess:
    '�I���ʒu�������ɖ߂�
    initSelection.setInitSelection
    '�I������
    Call commonUtil.endProcess
End Sub

'�x�e���Ԃ̎Z�o
Private Function calcRestTime(ByVal targetStart As Range, ByVal targetEnd As Range, ByVal targetWork As Range, ByVal isSac As Boolean)
    Worksheets("data").Activate
    
    Dim startCol As Long
    Dim endCol As Long
    
    Dim targetStartTime As Double
    Dim targetEndTime As Double
    Dim restStartTime As Double
    Dim restEndTime As Double
    Dim calcStartTime As Double
    Dim calcEndTime As Double
    Dim rt As Range
    Dim i As Long
    Dim j As Long
    Dim targetDay As Double
    Dim targetHour As Double
    Dim dayCnt As Double
    Dim restStartCol As Long
    Dim restEndCol As Long
    
    targetStartTime = targetStart.Value
    targetEndTime = targetEnd.Value
        '�J�n�ƏI�����t�]���Ă���ꍇ�A�I����24���ԉ��Z
    If targetStartTime > targetEndTime Then
        targetEndTime = calc24(targetEndTime)
    End If
    Dim restTime As Double: restTime = 0
    
    '�I�����Ԃ���ׂ��ł���������v�Z
    targetDay = targetWork.Value
    targetHour = targetDay * 24 + Hour(targetEndTime)
    dayCnt = targetHour \ 24
    
    'SAC�����ꂩ
    If isSac Then
        restStartCol = 1
        restEndCol = 2
    Else
        restStartCol = 3
        restEndCol = 4
    End If
    
    For Each rt In Range("�x�e���ԃ��X�g")
        i = rt.Row
        startCol = rt.Column + restStartCol
        endCol = rt.Column + restEndCol
        restStartTime = Cells(i, startCol).Value
        restEndTime = Cells(i, endCol).Value
        '�x�e���Ԃ��ݒ肳��Ă���ꍇ
        If restStartTime <> restEndTime Then
            '�J�n�ƏI�����t�]���Ă���ꍇ�A�I����24���ԉ��Z
            If restStartTime > restEndTime Then
                restEndTime = calc24(restEndTime)
            End If
            
            j = 0
            '�����ׂ��ł���ꍇ�A�ׂ��ł���������A�x�e���Ԃ̊J�n�E�I����24���ԑ����Čv�Z����
            Do
                If Not (targetEndTime < restStartTime Or targetStartTime > restEndTime) Then
                    If targetStartTime < restStartTime Then
                        calcStartTime = restStartTime
                    Else
                        calcStartTime = targetStartTime
                    End If
                    If targetEndTime < restEndTime Then
                        calcEndTime = targetEndTime
                    Else
                        calcEndTime = restEndTime
                    End If
                    restTime = restTime + (calcEndTime - calcStartTime)
                End If
                
                restStartTime = calc24(restStartTime)
                restEndTime = calc24(restEndTime)
                j = j + 1
            Loop While j <= dayCnt
        End If
    Next
    
    calcRestTime = restTime
End Function

'24���ȍ~���H
Private Function calc24(ByVal target As Double)
    Dim targetTime As Double: targetTime = target
    
    targetTime = DateAdd("h", 24, targetTime)
    
    calc24 = targetTime
End Function
