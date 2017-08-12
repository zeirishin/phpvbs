
Function http_build_query(formdata , numeric_prefix , arg_separator )

    If Not isArray(formdata) and Not isObject(formdata) Then Exit Function

    Dim i,key
    Dim url
    Dim separator

    separator = "&"
    If len(arg_separator) > 0 then
        separator = arg_separator
    end if

    If isArray(formdata) Then
        For key = 0 to uBound(formdata)
            If isArray(formdata(key)) or isObject(formdata(key)) Then
                url = url & separator & http_build_query(formdata(key) , _
                                            numeric_prefix , arg_separator )
            else
                url = url & separator & _
                    key & "=" & Server.URLEncode(formdata(key))
            end if
        Next
    ElseIf isObject(formdata) Then

        For Each i In formdata
            if isArray(i) or isObject(i) then
                url = url & separator & http_build_query(i , numeric_prefix , arg_separator )

            elseif isArray(formdata(i)) or isObject(formdata(i)) Then

                If isArray( formdata(i) ) Then
                    For Each key In formdata(i)
                        If isObject(key) or isArray(key) Then
                            url = url & separator & http_build_query(key , numeric_prefix , arg_separator )
                        Else
                            url = url & separator & _
                                i & "=" & Server.URLEncode(key)
                        End If
                    Next
                Else
                    url = url & separator & http_build_query(formdata(i) , numeric_prefix , arg_separator )
                End If
            else
                if isArray( formdata ) and len(numeric_prefix) > 0 then
                        url = url & separator & _
                            numeric_prefix & i & "=" & Server.URLEncode(formdata(i))
                else
                    url = url & separator & _
                        i & "=" & Server.URLEncode(formdata(i))
                end if
            end if
        Next
    End If

    If len( url ) > 0 Then url = Mid(url,2)

    http_build_query = url

End Function
