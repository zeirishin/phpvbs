<%
'=======================================================================
' 文字エンコーディングを変換する
'=======================================================================
'【引数】
'  str          = string    変換する文字列。
'  to_encoding  = string    str  の変換後の文字エンコーディング。
'  from_encoding= string    変換前の文字エンコーディング名を指定します。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・文字列 str の文字エンコーディングを、 オプションで指定した from_encoding  から to_encoding  に変換します。
'=======================================================================
Function mb_convert_encoding(str,to_encoding,from_encoding)

    Dim bobj : set bobj = Server.CreateObject("basp21")
    mb_convert_encoding = bobj.Kconv(str,_
                          mb_convert_encoding_helper(to_encoding),_
                          mb_convert_encoding_helper(from_encoding))
End Function
    
'*******************************************
Function mb_convert_encoding_helper(encoding)

    Dim tmp
    Select Case lcase(encoding)
    Case "shift_jis","sjis"
        tmp = 1
    Case "euc","euc-jp"
        tmp = 2
    Case "jis"
        tmp = 3
    Case "ucs2"
        tmp = 4
    Case "utf-8","utf8"
        tmp = 5
    Case "auto"
        tmp = 0
    End Select

End Function
%>
