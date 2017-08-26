<%
'=======================================================================
' 文字列中で、特定の(大文字小文字を区別しない)文字列が最後に現れた位置を探す
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string   needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  needle  が最後に現れた位置を返します。
'  needle  が見つからない場合、FALSE を返します。
'【処理】
'  ・ 文字列の中で、 大文字小文字を区別しないある文字列が最後に現れた位置を返します。
'  ・ strrpos() と異なり、strripos()  は大文字小文字を区別しません。
'=======================================================================
Function strripos( haystack, needle, offset)

    Dim i
    strripos = false

    If len(offset) = 0 Then
        offset = len( haystack)
    End If

    If len(needle) > 1 Then needle = Left(needle,1)

    i = InStrRev(haystack,needle,offset,vbTextCompare)

    If i > 0 Then
        strripos = i
    End If

End Function
%>
