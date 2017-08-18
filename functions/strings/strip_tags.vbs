Function strip_tags( str )

    Dim objRegExp
    Dim plane

    plane = Trim( str & " )

    If Len( plane ) > 0 Then

        Set objRegExp = New RegExp
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        objRegExp.Pattern= "</?[^>]+>"
        plane = objRegExp.Replace(str, ")
        Set objRegExp = Nothing

    End If

    strip_tags = plane

End Function
