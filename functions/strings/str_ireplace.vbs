<%
'=======================================================================
' 大文字小文字を区別しない str_replace()
'=======================================================================
'【引数】
'  search    = mixed 検索文字列
'  strReplace= mixed 置換文字列
'  subject   = mixed subject  が配列の場合は、そのすべての要素に 対して検索と置換が行われ、返される結果も配列となります。
'  cnt       = mixed needles  の中で、マッチして置換を行った数を count  に返します。このパラメータは参照渡しとします。
'【戻り値】
'  置換した文字列あるいは配列を返します。
'【処理】
'  ・この関数は、subject  の中に現れるすべての search (大文字小文字を区別しない)を replace  に置き換えた文字列あるいは配列を 返します。
'=======================================================================
Function str_ireplace(search, strReplace, subject, ByRef cnt)

    If is_string(search) and isArray(strReplace) Then Exit Function

    If Not isArray(search) Then search = array(search)
    search = array_values(search)

    Dim replace_string,i
    If not isArray(strReplace) Then
        replace_string = strReplace

        strReplace = array()
        For i = 0 to uBound(search)
            [] strReplace, replace_string
        Next
    End if

    strReplace = array_values(strReplace)

    Dim length_replace,length_search
    length_replace = count(strReplace,"")
    length_search  = count(search,"")
    if length_replace < length_search Then
        For i = length_replace to length_search
            strReplace(i) = ""
        Next
    End If

    Dim was_array : was_array = false
    If isArray(subject) Then
        was_array = true
        subject = array( subject )
    End If

    For i = 0 to uBound( search )
        search(i) = "/" & preg_quote(search(i),"") & "/"
    Next

    For i = 0 to uBound( strReplace )
        strReplace(i) = str_replace( array(chr(92),"$"),array(chr(92) & chr(92), "¥$"),strReplace(i) )
    Next

    Dim result
    result = preg_replace(search,strReplace,subject,"",cnt)

    If was_array = true Then
        str_ireplace = result(0)
    Else
        str_ireplace = result
    End If

End Function
%>
