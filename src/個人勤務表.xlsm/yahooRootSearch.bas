Attribute VB_Name = "yahooRootSearch"
Option Explicit

'Yahoo!�H�����Ō�ʔ����������
Public Sub yahooRootSearch()
    Dim strUrl As String
    Dim myWS As Worksheet
    Set myWS = ActiveSheet
    Dim intRow As Integer
    intRow = ActiveCell.Row
    
    If myWS.Range("�o��").Value <> "" And myWS.Range("����").Value <> "" Then
        'flatlon�F�s��
        'from   �F�o��
        'tlatlon�F�s��
        'to     �F����
        'via    �F�o�R1
        'via    �F�o�R2
        'via    �F�o�R3
        'y      �F�N
        'm      �F��
        'd      �F��
        'hh     �F��
        'm2     �F��(��̈�)
        'm1     �F��(�\�̈�)
        'type   �F�����w��(1:�o��,4:����,3:�n��,4:�I�d,5:�w��Ȃ�)
        'ticket �F�^�����(ic:IC�J�[�h�D��,normal:����(������)�D��)
        'al     �F��H
        'shin   �F�V����
        'ex     �F�L�����}
        'hb     �F�����o�X
        'lb     �F�H��/�A���o�X
        'sr     �F�t�F���[
        's      �F�\������(0:������������,2:��芷���񐔏�,1:������������)
        'expkind�F�Ȏw��(1:���R�ȗD��,2:�w��ȗD��,3:�O���[���ԗD��)
        'ws     �F�������x(1:�}����,2:�W��,3:�����������,4:�������)
        'kw     �F�����Ɠ���
        
        strUrl = "http://transit.yahoo.co.jp/search/"
        strUrl = strUrl & "result?"
        strUrl = strUrl & "flatlon="
        strUrl = strUrl & "&from=" & commonUtil.encodeURL(myWS.Range("�o��").Value)
        strUrl = strUrl & "&tlatlon="
        strUrl = strUrl & "&to=" & commonUtil.encodeURL(myWS.Range("����").Value)
        strUrl = strUrl & "&via=" & commonUtil.encodeURL(myWS.Range("�o�R").Value)
        strUrl = strUrl & "&via="
        strUrl = strUrl & "&via="
        strUrl = strUrl & "&y=" & Format(Date, "yyyy")
        strUrl = strUrl & "&m=" & Format(Date, "mm")
        strUrl = strUrl & "&d=" & Format(Date, "dd")
        strUrl = strUrl & "&hh=" & Format(Date, "hh")
        strUrl = strUrl & "&m2=0"
        strUrl = strUrl & "&m1=0"
        strUrl = strUrl & "&type=5"
        strUrl = strUrl & "&ticket=ic"
        strUrl = strUrl & "&al=1"
        strUrl = strUrl & "&shin=1"
        strUrl = strUrl & "&ex=1"
        strUrl = strUrl & "&hb=1"
        strUrl = strUrl & "&lb=1"
        strUrl = strUrl & "&sr=1"
        strUrl = strUrl & "&s=0"
        strUrl = strUrl & "&expkind=1"
        strUrl = strUrl & "&ws=2"
        strUrl = strUrl & "&kw=" & commonUtil.encodeURL(myWS.Range("����").Value)
        
        Call openIE(strUrl)
    Else
        MsgBox "�o���������͓��������͂���Ă��܂���B"
    End If

End Sub

