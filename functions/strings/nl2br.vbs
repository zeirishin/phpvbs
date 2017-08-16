Function nl2br(str)
    nl2br = preg_replace("/([^>])" & vbCrLf & "/","$1
", str,",")
End Function
