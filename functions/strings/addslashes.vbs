Function addslashes(ByVal str)

    If isNull(str) Then
        str = "
    End If

    str = Replace(str,"￥","￥￥")
    str = Replace(str,"","￥")
    str = Replace(str,"'","￥'")

    addslashes = str

End Function
