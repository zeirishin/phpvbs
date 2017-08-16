Function stripslashes(ByVal str)
    str = preg_replace("/￥￥(.)/","$1",str,",")
    stripslashes = str
End Function
