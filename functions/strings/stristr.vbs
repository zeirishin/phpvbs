<%
'=======================================================================
' 大文字小文字を区別しない strstr()
'=======================================================================
'【引数】
'  haystack     = string    検索を行う文字列。
'  needle       = string    needle は、 ひとつまたは複数の文字であることに注意しましょう。needle が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  before_needle= string    TRUE にすると (デフォルトは FALSE です)、stristr()  の返り値は、haystack  の中で最初に needle  があらわれる箇所より前の部分となります。
'【戻り値】
'  マッチした部分文字列を返します。needle  が見つからない場合は FALSE を返します。
'【処理】
'  ・haystack  において needle  が最初に見つかった位置から最後までを返します。
'  ・needle  および haystack  は大文字小文字を区別せずに評価されます。
'=======================================================================
Function stristr( haystack, needle, before_needle )

    Dim pos
    If varType(before_needle) <> 11 Then before_needle = false

    pos = Instr(1,haystack,needle,vbTextCompare)

    If pos <= 0 Then
        stristr = false
    Else
        If before_needle Then
            stristr = Mid(haystack,1,pos-1)
        Else
            stristr = Mid(haystack,pos)
        End If
    End If

End Function
%>
