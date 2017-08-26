<%
'=======================================================================
' 副文字列の出現回数を数える
'=======================================================================
'【引数】
'  haystack     = string    検索対象の文字列
'  needle       = string    検索する副文字列
'  offset       = int       開始位置のオフセット
'  length       = int       指定したオフセット以降に副文字列で検索する最大長。
'【戻り値】
'  この関数は 整数 を返します。
'【処理】
'  ・substr_count() は、文字列 haystack  の中での副文字列 needle  の出現回数を返します。
'  ・needle  は英大小文字を区別することに注意してください。
'=======================================================================
Function substr_count( haystack, needle, offset, length )

    Dim pos,cnt : cnt = 0

    If not isNumeric(offset) Then offset = 1
    If not isNumeric(length) Then length = 0

    Do While inStr(offset+1,haystack,needle,vbBinaryCompare) > 0
        offset = inStr(offset+1,haystack,needle,vbBinaryCompare)
        If length > 0 and offset + len(needle) > length Then
            Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop

    substr_count = cnt

End Function
%>
