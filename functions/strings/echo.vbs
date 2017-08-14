Sub echo(str)

    If isObject(str) then
        Response.Write "Object"
    ElseIf IsArray(str) then
        Response.Write "Array"
    Else
        Response.Write str
    End if

End Sub
