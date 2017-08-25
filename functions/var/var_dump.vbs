<%
'=======================================================================
'変数に関する情報をダンプする
'=======================================================================
'【引数】
'  val   = mixed 破棄する変数。
'【戻り値】
'  値を返しません。
'【処理】
'  ・この関数は、指定した式に関してその型や値を含む構造化された情報を 返します。
'  ・配列の場合、その構造を表示するために各値について再帰的に 探索されます。
'=======================================================================
Sub var_dump(expression)
    var_dump_helper expression,0
End Sub

'***************************
Sub var_dump_helper(expression,tab)

    If VarType(tab) <> 2 Then tab = 0

    Dim strTab : strTab = String(tab,vbTab)

    If IsObject(expression) Then
        Response.Write "Dictionary Object(" & expression.count & ")" & vbCrLf
    ElseIf IsArray(expression) Then
        Response.Write "Array(" & (uBound(expression)+1) & ")" & vbCrLf
    End If

    Response.Write strTab & "(" & vbCrLf

    Dim a,i
    i = 0
    If IsObject(expression) Then
        For Each a In expression
            Response.Write strTab
            If IsArray(a) or IsObject(a) Then
                Response.Write vbTab & "[] => "
                call var_dump_helper(a,tab + 1)
            ElseIf isArray(expression(a)) or isObject( expression(a) ) Then
                Response.Write vbTab & "[""" & a & """] => "
                call var_dump_helper(expression(a),tab + 1)

            Else
               Response.Write vbTab & "[""" & a & """]" & " => " & _
                              gettype(expression(a)) & "(" & expression(a) & ")" & vbCrLf
            End If
        Next
    ElseIf IsArray(expression) Then
        For Each a In expression
            Response.Write strTab
            If IsArray(a) or IsObject(a) Then
                Response.Write vbTab & "[" & i & "] => "
                call var_dump_helper(a,tab + 1)
            Else
                Response.Write vbTab & "[" & i & "] => " & _
                               gettype(a) & "(" & a & ")" & vbCrLf
            End If

            i =  i+1
        Next
    Else
        Response.Write strTab & gettype(expression) & "(" & expression & ")" & vbCrLf
    End If

    Response.Write strTab & ")" & vbCrLf

End Sub
%>
