<%
'=======================================================================
' 文字列が最初に現れる場所を見つける
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string   needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  位置を表す整数値を返します。 needle  が見つからない場合、 strpos() は boolean FALSE を返します。
'【処理】
'  ・ 文字列 haystack  の中で、 needle  が最初に現れた位置を数字で返します。
'  ・ PHP 5 以前の strrpos() とは異なり、この関数は needle  パラメータとして文字列全体をとり、 その文字列全体が検索対象となります。
'=======================================================================
Function strpos( haystack, needle, offset)

    Dim i
    strpos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbBinaryCompare)

    If i > 0 Then
        strpos = i
    End If

End Function
%>
