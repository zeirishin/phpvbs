<%
'=======================================================================
' 文字列を文字列により分割する
'=======================================================================
'【引数】
'  delimiter    = string    区切り文字列。
'  string       = string    入力文字列。
'  limit        = string    limit  が指定された場合、返される配列には 最大 limit  の要素が含まれ、その最後の要素には string  の残りの部分が全て含まれます。
'【戻り値】
'  空の文字列 ("") が delimiter  として使用された場合、 explode() は FALSE  を返します。
'  delimiter  に引数 string  に含まれていない値が含まれている場合、 explode() は、引数 string  を含む配列を返します。
'【処理】
'  ・文字列の配列を返します。この配列の各要素は、 string  を文字列 delimiter  で区切った部分文字列となります。
'=======================================================================
Function explode(delimiter,string,limit)

    explode = false
    If len(delimiter) = 0 Then Exit Function
    If len(limit) = 0 Then limit = 0

    If limit > 0 Then
        explode = Split(string,delimiter,limit)
    Else
        explode = Split(string,delimiter)
    End If

End Function
%>
