<%
'=======================================================================
' 文字列の一部を置換する
'=======================================================================
'【引数】
'  str          = string    入力文字列。
'  replacement  = string    置換する文字列。
'  start        = string    start が正の場合、置換は string で start 番目の文字から始まります。start が負の場合、置換は string の終端から start 番目の文字から始まります。
'  length       = string    正の値を指定した場合、 string 　の置換される部分の長さを表します。 負の場合、置換を停止する位置が string  の終端から何文字目であるかを表します。このパラメータが省略された場合、 デフォルト値は strlen(string )、すなわち、 string  の終端まで置換することになります。 当然、もし length  がゼロだったら、 この関数は string  の最初から start  の位置に replacement  を挿入するということになります。
'【戻り値】
'  結果の文字列を返します。もし、string  が配列の場合、配列が返されます。
'【処理】
'  ・substr_replace()は、文字列 string の start  および (オプションの) length  パラメータで区切られた部分を replacement  で指定した文字列に置換します。
'=======================================================================
Function substr_replace(str, replacement, start, length)

    Dim key

    If isArray(str) Then
        For key = 0 to uBound(str)
            substr_replace(key) = substr_replace(str(key), replacement, start, length)
        Next
        Exit Function
    ElseIf isObject(str) Then
        For Each key In str
            substr_replace(key) = substr_replace(str(key), replacement, start, length)
        Next
        Exit Function
    End If

    Dim result

    If start < 0 Then
        start = len(str) + start
    End If

    If start <> 0 Then
        result = Left(str,start)
    End If

    result = result & replacement

    If len(length) > 0 Then
        If length > 0 Then
            result = result & Mid(str,start + length)
        Else
            result = result & Right(str,abs(length))
        End If

    End If

    substr_replace = result
End Function
%>
