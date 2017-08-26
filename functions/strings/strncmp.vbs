<%
'=======================================================================
' 最初の n 文字についてバイナリセーフな文字列比較を行う
'=======================================================================
'【引数】
'  str1     =   string  最初の文字列。
'  str2     =   string  次の文字列。
'  intlen   =   string  比較する文字数。
'【戻り値】
' str1  が str2  より短い場合に < 0 を返し、str1  が str2  より大きい場合に > 0、等しい場合に 0 を返します。
'【処理】
'  ・ この関数は strcmp() に似ていますが、 各文字列から(最大)文字数(len ) を比較に使用するところが異なります。
'  ・ 比較は大文字小文字を区別することに注意してください。 
'=======================================================================
Function strncmp(ByVal str1,ByVal str2,intlen)

    If len(str1) > intlen Then str1 = Left(str1,intlen)
    If len(str2) > intlen Then str2 = Left(str2,intlen)

    strncmp = StrComp(str1,str2,vbBinaryCompare)
End Function
%>
