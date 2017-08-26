<%
'=======================================================================
' バイナリセーフで大文字小文字を区別しない文字列比較を、最初の n 文字について行う
'=======================================================================
'【引数】
'  str1     =   string  最初の文字列。
'  str2     =   string  次の文字列。
'  intlen   =   string  比較する文字列の長さ。
'【戻り値】
' str1  が str2  より短い場合に < 0 を返し、str1  が str2  より大きい場合に > 0、等しい場合に 0 を返します。
'【処理】
'  ・この関数は、strcasecmp() に似ていますが、 各文字列から比較する文字数(の上限)(len ) を指定できるという違いがあります。
'  ・どちらかの文字列が len より短い場合、その文字列の長さが比較時に使用されます。
'=======================================================================
Function strncasecmp(ByVal str1,ByVal str2,intlen)

    If len(str1) > intlen Then str1 = Left(str1,intlen)
    If len(str2) > intlen Then str2 = Left(str2,intlen)

    strncasecmp = StrComp(str1,str2,vbTextCompare)
End Function
%>
