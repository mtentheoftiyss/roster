Attribute VB_Name = "calcWorkTime"
Option Explicit

'勤務時間を計算
Public Sub calcWorkTime(ByVal target As Range)
    '画面の描画をOFFにします
    Call commonUtil.startProcess
    
    Dim startCol As Long: startCol = Range("開始").Column
    Dim endCol As Long: endCol = Range("終了").Column
    Dim sacWorkCol As Long: sacWorkCol = Range("SAC勤務時間").Column
    Dim siteWorkCol As Long: siteWorkCol = Range("現場勤務時間").Column
    
    '選択位置を取得
    Dim initSelection As initSelection
    Set initSelection = New initSelection
    
    '開始、終了の変更時のみ実行
    If Intersect(target, Range("開始・終了時間リスト")) Is Nothing Then
        GoTo endProcess
    End If
    
    Dim targetCell As Range
    For Each targetCell In target
        '開始、終了の変更時のみ実行
        If Intersect(target, Range("開始・終了時間リスト")) Is Nothing Then
            GoTo continue
        End If
        
        '開始、終了少なくともどちらかが未入力の場合は勤務時間を削除
        If Cells(targetCell.Row, startCol).Value = "" Or Cells(targetCell.Row, endCol).Value = "" Then
            Cells(targetCell.Row, sacWorkCol).Value = ""
            Cells(targetCell.Row, siteWorkCol).Value = ""
            GoTo continue
        End If
        
        '休憩時間の算出
        Dim startTime As Double: startTime = Cells(targetCell.Row, startCol).Value
        Dim endTime As Double: endTime = Cells(targetCell.Row, endCol).Value
        '時間を規定時間(30分単位)で調整
        startTime = commonUtil.roundTime(startTime, commonConstants.timeDivide, True)
        endTime = commonUtil.roundTime(endTime, commonConstants.timeDivide, False)
        Cells(targetCell.Row, startCol).Value = startTime
        Cells(targetCell.Row, endCol).Value = endTime
        '開始と終了が逆転している場合、終了に24時間加算
        If startTime > endTime Then
            endTime = calc24(endTime)
            Cells(targetCell.Row, endCol).Value = endTime
        End If
        'マクロでDAY関数を使うと、うまく値が取れないので一旦セルに式を埋め込む
        Cells(targetCell.Row, sacWorkCol).Formula = "=DAY(" + Cells(targetCell.Row, endCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) + ")"
        Cells(targetCell.Row, siteWorkCol).Formula = "=DAY(" + Cells(targetCell.Row, endCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) + ")"
        Dim startCell As Range: Set startCell = Cells(targetCell.Row, startCol)
        Dim endCell As Range: Set endCell = Cells(targetCell.Row, endCol)
        Dim sacCell As Range: Set sacCell = Cells(targetCell.Row, sacWorkCol)
        Dim siteCell As Range: Set siteCell = Cells(targetCell.Row, siteWorkCol)
        '休憩時間の算出
        Dim sacRestTime As Double: sacRestTime = calcRestTime(startCell, endCell, sacCell, True)
        Dim siteRestTime As Double: siteRestTime = calcRestTime(startCell, endCell, siteCell, False)
        
        '勤務時間の算出
        Worksheets(initSelection.getInitSheet).Activate
        Cells(targetCell.Row, sacWorkCol).Value = (endTime - startTime - sacRestTime)
        Cells(targetCell.Row, siteWorkCol).Value = (endTime - startTime - siteRestTime)
        
        '現場勤務時間から自社分を差引く
        Dim sacTime As Double: sacTime = 0
        Dim ma As Range
        For Each ma In targetCell.MergeArea
            Dim wp As Range: Set wp = Cells(ma.Row, Range("作業場所").Column)
            Dim wpl As Range
            For Each wpl In Range("作業場所リスト")
                If wp.Value = wpl.Value Then
                    Dim jisha As Range: Set jisha = wpl.Offset(0, 1)
                    If jisha.Value = "自社" Then
                        sacTime = sacTime + Cells(ma.Row, Range("工数").Column).Value / 24
                    End If
                End If
            Next wpl
        Next ma
        Cells(targetCell.Row, siteWorkCol).Value = Cells(targetCell.Row, siteWorkCol).Value - sacTime

continue:
    Next targetCell
    
endProcess:
    '選択位置を初期に戻す
    initSelection.setInitSelection
    '終了処理
    Call commonUtil.endProcess
End Sub

'休憩時間の算出
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
        '開始と終了が逆転している場合、終了に24時間加算
    If targetStartTime > targetEndTime Then
        targetEndTime = calc24(targetEndTime)
    End If
    Dim restTime As Double: restTime = 0
    
    '終了時間から跨いでいる日数を計算
    targetDay = targetWork.Value
    targetHour = targetDay * 24 + Hour(targetEndTime)
    dayCnt = targetHour \ 24
    
    'SACか現場か
    If isSac Then
        restStartCol = 1
        restEndCol = 2
    Else
        restStartCol = 3
        restEndCol = 4
    End If
    
    For Each rt In Range("休憩時間リスト")
        i = rt.Row
        startCol = rt.Column + restStartCol
        endCol = rt.Column + restEndCol
        restStartTime = Cells(i, startCol).Value
        restEndTime = Cells(i, endCol).Value
        '休憩時間が設定されている場合
        If restStartTime <> restEndTime Then
            '開始と終了が逆転している場合、終了に24時間加算
            If restStartTime > restEndTime Then
                restEndTime = calc24(restEndTime)
            End If
            
            j = 0
            '日を跨いでいる場合、跨いでいる日数分、休憩時間の開始・終了に24時間足して計算する
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

'24時以降加工
Private Function calc24(ByVal target As Double)
    Dim targetTime As Double: targetTime = target
    
    targetTime = DateAdd("h", 24, targetTime)
    
    calc24 = targetTime
End Function
