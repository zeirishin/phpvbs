<%
'BASP21を使用しない場合
'http://www.ac.cyberhome.ne.jp/‾mattn/AcrobatASP/1.html
'http://www.ac.cyberhome.ne.jp/‾mattn/AcrobatASP/StrConv.inc
'=======================================================================
' カナを("全角かな"、"半角かな"等に)変換する
'=======================================================================
'【引数】
'  str      = string  文字列
'  option   = mixed   変換オプション
'【戻り値】
'  変換された文字列を返します。
'【処理】
'  ・文字列 str  に関して「半角」-「全角」変換を行います。
'=======================================================================
Function mb_convert_kana( str , option_name )

    Dim outstr
    Dim J
    Dim tmp
    Dim bobj

    If Len( str ) = 0 Then Exit Function

    Set bobj = Server.CreateObject("basp21")

    If Len( option_name ) = 0 Then
        option_name = "K"
    End If

    outstr = str

    For J = 1 To Len( option_name )

        tmp = mid(option_name,J,1)

        If tmp = "r" Then
            '全角文字 (2 バイト) を半角文字 (1 バイト) に変換
            outstr = bobj.StrConv(outstr,8)

        ElseIf tmp = "R" Then
            '半角文字 (1 バイト) を全角文字 (2 バイト) に変換
            outstr = bobj.StrConv(outstr,4)

        ElseIf tmp = "c" Then
            'カタカナをひらがなに変換
            outstr = bobj.StrConv(outstr,32)

        ElseIf tmp = "C" Then
            'ひらがなをカタカナに変換
            outstr = bobj.StrConv(outstr,16)

        ElseIf tmp = "K" or tmp = "V" THen
            '文字列の中の半角カナを全角カナに変換します。濁音にも対応。
            outstr = bobj.HAN2ZEN(outstr)

        End If
    Next

    mb_convert_kana = outstr

End Function
%>
