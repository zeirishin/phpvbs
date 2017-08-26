<%
'=======================================================================
' 特殊文字を HTML エンティティに変換する
'=======================================================================
'【引数】
'  str  = string    変換される文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・文字の中には HTML において特殊な意味を持つものがあり、 それらの本来の値を表示したければ HTML の表現形式に変換してやらなければなりません。
'  ・この関数は、これらの変換を行った結果の文字列を返します。 
'=======================================================================
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
%>
