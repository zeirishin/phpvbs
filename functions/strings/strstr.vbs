<%
'=======================================================================
' 文字列が最初に現れる位置を見つける
'=======================================================================
'【引数】
'  haystack     = string    入力文字列。
'  needle       = string    needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  before_needle= string    TRUE にすると (デフォルトは FALSE です)、strstr()  の返り値は、haystack  の中で最初に needle  があらわれる箇所より前の部分となります。
'【戻り値】
'  部分文字列を返します。 needle  が見つからない場合は FALSE を返します。
'【処理】
'  ・haystack  の中で needle  が最初に現れる場所から文字列の終わりまでを返します。
'=======================================================================
Function strstr( haystack, needle, before_needle )

    Dim pos
    If varType(before_needle) <> 11 Then before_needle = false

    pos = Instr(1,haystack,needle,vbBinaryCompare)

    If pos <= 0 Then
        strstr = false
    Else
        If before_needle Then
            strstr = Mid(haystack,1,pos-1)
        Else
            strstr = Mid(haystack,pos)
        End If
    End If

End Function
%>
