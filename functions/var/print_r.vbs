<%
'=======================================================================
'指定した変数に関する情報を解りやすく出力する
'=======================================================================
'【引数】
'  expression   = mixed 表示したい式。
'  ret          = bool  print_r() はデフォルトでは結果を直接表示してしまいますが この引数が TRUE の場合には結果を戻します。
'【戻り値】
'  値が出力されます。
'【処理】
'  ・変数の値に関する情報を解り易い形式で表示します。
'=======================================================================
Function print_r(expression,ret)
    print_r = print_r_helper(expression,ret,0)
End Function

'*************************
Function print_r_helper(expression,ret,tab)

    If VarType(tab) <> 2 Then tab = 0
    If VarType(ret) <> 11 Then ret = false

    Dim strPrint

    If IsObject(expression) Then
        strPrint = strPrint & "Dictionary Object" & vbCrLf
    ElseIf IsArray(expression) Then
        strPrint = strPrint & "Array" & vbCrLf
    End If

    strPrint = strPrint & String(tab,vbTab) & "(" & vbCrLf

    Dim a,i
    i = 0
    If IsObject(expression) Then
        For Each a In expression
            strPrint = strPrint & String(tab,vbTab)
            If IsArray(a) or IsObject(a) Then
                strPrint = strPrint & vbTab & "[] => " & _
                           print_r_helper(a,true,tab + 1)
            ElseIf isArray(expression(a)) or isObject( expression(a) ) Then
                strPrint = strPrint & vbTab & "[" & a & "] => " & _
                           print_r_helper(expression(a),true,tab + 1)

            Else
               strPrint = strPrint & vbTab & ("[" & a & "]" & " => " & _
                          expression(a)) & vbCrLf
            End If
        Next
    ElseIf IsArray(expression) Then
        For Each a In expression
            strPrint = strPrint & String(tab,vbTab)
            If IsArray(a) or IsObject(a) Then
                strPrint = strPrint & vbTab & "[" & i & "] => " & _
                           print_r_helper(a,true,tab + 1)
            Else
                strPrint = strPrint & vbTab & ("[" & i & "] => " & a) & vbCrLf
            End If

            i =  i+1
        Next
    Else
        strPrint = strPrint & String(tab,vbTab) & expression & vbCrLf
    End If

    strPrint = strPrint & String(tab,vbTab) & ")" & vbCrLf

    If Not ret Then
        Response.Write strPrint
    Else
        print_r_helper = strPrint
    End If

End Function
%>
