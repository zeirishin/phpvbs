Function sprintf(format , args)

    If is_empty(args) Then args = "
    Dim bobj : set bobj = Server.CreateObject("basp21")
    sprintf = bobj.sprintf(format,args)

End Function
