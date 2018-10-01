Attribute VB_Name = "yahooRootSearch"
Option Explicit

'Yahoo!路線情報で交通費を検索する
Public Sub yahooRootSearch()
    Dim strUrl As String
    Dim myWS As Worksheet
    Set myWS = ActiveSheet
    Dim intRow As Integer
    intRow = ActiveCell.Row
    
    If myWS.Range("出発").Value <> "" And myWS.Range("到着").Value <> "" Then
        'flatlon：不明
        'from   ：出発
        'tlatlon：不明
        'to     ：到着
        'via    ：経由1
        'via    ：経由2
        'via    ：経由3
        'y      ：年
        'm      ：月
        'd      ：日
        'hh     ：時
        'm2     ：分(一の位)
        'm1     ：分(十の位)
        'type   ：日時指定(1:出発,4:到着,3:始発,4:終電,5:指定なし)
        'ticket ：運賃種別(ic:ICカード優先,normal:現金(きっぷ)優先)
        'al     ：空路
        'shin   ：新幹線
        'ex     ：有料特急
        'hb     ：高速バス
        'lb     ：路線/連絡バス
        'sr     ：フェリー
        's      ：表示順序(0:到着が早い順,2:乗り換え回数順,1:料金が安い順)
        'expkind：席指定(1:自由席優先,2:指定席優先,3:グリーン車優先)
        'ws     ：歩く速度(1:急いで,2:標準,3:少しゆっくり,4:ゆっくり)
        'kw     ：到着と同一
        
        strUrl = "http://transit.yahoo.co.jp/search/"
        strUrl = strUrl & "result?"
        strUrl = strUrl & "flatlon="
        strUrl = strUrl & "&from=" & commonUtil.encodeURL(myWS.Range("出発").Value)
        strUrl = strUrl & "&tlatlon="
        strUrl = strUrl & "&to=" & commonUtil.encodeURL(myWS.Range("到着").Value)
        strUrl = strUrl & "&via=" & commonUtil.encodeURL(myWS.Range("経由").Value)
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
        strUrl = strUrl & "&kw=" & commonUtil.encodeURL(myWS.Range("到着").Value)
        
        Call openIE(strUrl)
    Else
        MsgBox "出発もしくは到着が入力されていません。"
    End If

End Sub

