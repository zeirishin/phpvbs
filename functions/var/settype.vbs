
Function settype(ByRef str,strType)

    settype = true

    Select Case strType
    Case "bool"
        str = CBool(str)
    Case "boolean"
        str = CBool(str)
    Case "byte"
        str = CByte(str)
    Case "currency"
        str = CCur(str)
    Case "date"
        str = CDate(str)
    Case "float"
        str = CDbl(str)
    Case "double"
        str = CDbl(str)
    Case "int"
        str = Cint(str)
    Case "integer"
        str = Cint(str)
    Case "long"
        str = CLng(str)
    Case "single"
        str = CSng(str)
    Case "string"
        str = Cstr(str)
    Case "array"
        If not isArray(str) Then
            str = array(str)
        End If
    Case "null"
        str = null
    Case Else
        settype = false
    End Select

End Function
