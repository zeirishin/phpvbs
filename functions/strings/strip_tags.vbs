<%
'=======================================================================
' 文字列から HTMLタグを取り除く
'=======================================================================
'【引数】
'  str              = string 入力文字列。
'  allowable_tags   = string オプションの2番目の引数により、取り除かないタグを指定できます。
'【戻り値】
'  タグを除去した文字列を返します。
'【処理】
'  ・指定した文字列 ( str ) から全ての HTMLタグを取り除きます。
'=======================================================================
Function strip_tags( str )

    Dim objRegExp
    Dim plane

    plane = Trim( str & "" )

    If Len( plane ) > 0 Then

        Set objRegExp = New RegExp
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        objRegExp.Pattern= "</?[^>]+>"
        plane = objRegExp.Replace(str, "")
        Set objRegExp = Nothing

    End If

    strip_tags = plane

End Function
%>
