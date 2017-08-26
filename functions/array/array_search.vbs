<%
'=======================================================================
'指定した値を配列で検索し、見つかった場合に対応するキーを返す
'=======================================================================
'【引数】
'  needle   = mixed 探したい値。
'  haystack = array 配列。
'  strict   = mixed TRUE が指定された場合、array_search() は haystack  の中で needle  の型に一致するかどうかも確認します。
'【戻り値】
'  ・needle  が見つかった場合に配列のキー、 それ以外の場合に FALSE を返します。
'【処理】
'  ・haystack  において needle  を検索します。
'=======================================================================
Function array_search(needle,haystack,strict)

    array_search = false

    If IsArray(needle) or isObject(needle) Then Exit Function

    If VarType(strict) <> 11 Then strict = false

    Dim key
    If isObject( haystack ) Then
        For Each key In haystack
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                array_search = array_search(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    array_search = key
            End If
            If array_search <> false Then Exit For
        Next
    ElseIf isArray( haystack ) Then
        For key = 0 to uBound( haystack )
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                array_search = array_search(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    array_search = key
            End If
            If array_search <> false Then Exit For
        Next
    End If

End Function
%>
