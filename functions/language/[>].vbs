<%
'=======================================================================
'a が b より多い TRUE。
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  比較する値
'【戻り値】
'  aがbより多い場合にTRUE を、多くない場合に FALSE を返します。
'【処理】
'  ・左辺と右辺を比較します。
'=======================================================================
Function [>](a, b)

    [>] = false

    Dim tmp_a,tmp_b
    Dim key

    If (isArray(a) or isArray(b)) or (isObject(a) or isObject(b)) Then

        If isObject(a) and isObject(b) Then
            If a.count < b.count Then Exit Function

            tmp_a = a.keys : tmp_b = b.keys
            If Not [>](uBound(tmp_a),uBound(tmp_b)) Then Exit Function
            [>] = true
        End If

        If isObject(a) and not isObject(b) Then
            [>] = true
        End If

        If isArray(a) and isArray(b) Then
            If uBound(a) < uBound(b) Then Exit Function

            For key = 0 to uBound(a)
                If Not [>](a(key),b(key) ) Then Exit Function
            Next

            [>] = true
        End If

        If isArray(a) and not isArray(b) Then
            [>] = true
        End If
    Else
        [>] = eval(a > b)
    End If

End Function
%>
