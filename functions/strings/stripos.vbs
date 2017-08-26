<%
'=======================================================================
' 大文字小文字を区別せずに文字列が最初に現れる位置を探す
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string    needle は、 ひとつまたは複数の文字であることに注意しましょう。needle が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  needle  がみつからない場合、 stripos() は boolean FALSE  を返します。
'【処理】
'  ・ 文字列 haystack  の中で needle  が最初に現れる位置を数字で返します。
'  ・ strpos() と異なり、stripos() は大文字小文字を区別しません。 
'=======================================================================
Function stripos( haystack, needle, offset)

    Dim i
    stripos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbTextCompare)

    If i > 0 Then
        stripos = i
    End If

End Function
%>
