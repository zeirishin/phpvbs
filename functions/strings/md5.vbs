Function md5(str)

    Dim bobj
    Set bobj = Server.CreateObject("basp21")
    md5 = bobj.MD5(str)

End Function
