<%
'=======================================================================
'変数が整数型かどうかを検査する
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str 整数型 の場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・与えられた変数の型が整数型かどうかを検査します。
'=======================================================================
Function is_int(str)

    is_int = false
    if Not isNumeric(str) Then Exit Function
    if str < 0 Then Exit Function
    is_int = (varType(str) = 2 or varType(str) = 3)

End Function
%>
