<%
'=======================================================================
'配列に値があるかチェックする
'=======================================================================
'【引数】
'  needle     = mixed 探す値。
'  haystack   = Array  配列。
'  strict     = bool   三番目のパラメータ strict が TRUE に設定された場合、 haystack の中の needle の型も確認します。
'【戻り値】
'  配列で needle  が見つかった場合に TRUE、それ以外の場合は、FALSE を返します。
'【処理】
'  ・haystack配列内にneedleが含まれるかチェック
'=======================================================================
Function in_array(needle, haystack,strict)

    in_array = False

    If Not IsArray(needle) Then
        If Len( needle ) = 0 Then Exit Function
    End If

    If VarType(strict) <> 11 Then strict = false

    Dim key

    If isArray(needle) Then
        For Each key In needle
            in_array = in_array(key,haystack,strict)
            If in_array = true Then Exit For
        Next
        Exit Function
    End If

    If isObject( haystack ) Then
        For Each key In haystack
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                in_array = in_array(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    in_array = true
            End If
            If in_array = true Then Exit For
        Next
    ElseIf isArray( haystack ) Then
        For key = 0 to uBound( haystack )
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                in_array = in_array(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    in_array = true
            End If
            If in_array = true Then Exit For
        Next
    End If

End Function

%>
