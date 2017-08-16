Function parse_str(str,ByRef arr)

    Dim glue1 : glue1 = "="
    Dim glue2 : glue2 = "&"
    Dim array2,array3 : set array3 = Server.CreateObject("Scripting.Dictionary")
    Dim x,tmp,counter,tmp_ar

    array2 = Split(str,glue2)
    If uBound( array2 ) > 0 Then
        For x = 0 to uBound( array2 )
            tmp = Split( array2(x), glue1 )
            If uBound( tmp ) > 0 Then

                tmp(0) = urldecode( tmp(0) )
                tmp(1) = Replace( urldecode(tmp(1)), "+", " ")

                If array3.Exists( tmp(0) ) Then
                    If isArray( array3.Item( tmp(0) ) ) Then
                        tmp_ar = array_values( array3.Item( tmp(0) ) )
                        [] tmp_ar, tmp(1)
                        array3.Item( tmp(0) ) = tmp_ar
                    Else
                        array3.Item( tmp(0) ) = array(array3.Item( tmp(0) ), tmp(1))
                    End If
               Else
                    array3.Add urldecode(tmp(0)), Replace( urldecode(tmp(1)), "+", " ")
                End If
            End If
        Next
    End If

    If vartype(arr) = 0 Then
        [=] arr, array3
    Else
        [=] parse_str, array3
    End If

End Function
