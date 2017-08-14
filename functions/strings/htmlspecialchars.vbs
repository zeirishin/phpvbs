Const ENT_NOQUOTES = 0
Const ENT_COMPAT   = 2
Const ENT_QUOTES   = 3
Function htmlSpecialChars(ByVal str)

    if len( str ) > 0 then
        str = Server.HTMLEncode(str)
        str = Replace(str,"'","&#039;")
    end if
    htmlspecialchars = str

End Function
