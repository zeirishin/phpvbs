<%
'=======================================================================
' 数字を千位毎にグループ化してフォーマットする
'=======================================================================
'【引数】
'  number           = float     フォーマットする数値。
'  decimals         = int       小数点以下の桁数。
'  dec_point        = string    小数点を表す区切り文字。
'  thousands_sep    = string    千位毎の区切り文字。thousands_sep は最初の文字だけが使用されます。 例えば、数字の 1000 に対する thousands_sep として bar を使用した場合、number_format() は 1b000 を返します。
'【戻り値】
'  変更後の文字列を返します。
'【処理】
'  ・パラメータが 1 つだけ渡された場合、 number  は千位毎にカンマ (",") が追加され、 小数なしでフォーマットされます。
'  ・パラメータが 2 つ渡された場合、number は decimals 桁の小数の前にドット (".") 、 千位毎にカンマ (",") が追加されてフォーマットされます。
'  ・パラメータが 4 つ全て渡された場合、number はドット (".") の代わりに dec_point が decimals 桁の小数の前に、千位毎にカンマ (",") の代わりに thousands_sep が追加されてフォーマットされます。 
'=======================================================================
Function number_format( number, decimals, dec_point, thousands_sep )

    Dim n,c,d,t,i,s
    n = number
    c = [?]( isNumeric(decimals),decimals,2 )
    c = abs( c )

    d = [?]( len(dec_point) = 0,",",dec_point)
    t = [?]( len(thousands_sep) = 0, ".", left(thousands_sep,1) )

    n = FormatNumber (n, c,true,false,true)
    n = Replace(n,",",d)
    n = Replace(n,".",t)

    number_format = n

End Function
%>
