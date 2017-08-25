<%
'=======================================================================
' 値が数値でないかどうかを判定する
'=======================================================================
'【引数】
'  val = float   調べる値。
'【戻り値】
'  val  が '非数値 (not a number)' の場合に TRUE、そうでない場合に FALSE を返します。
'【処理】
'  ・val  が '非数値 (not a number)' であるかどうかを調べます。たとえば acos(1.01) の結果などがこれにあたります。
'=======================================================================
Function is_nan(val)
    is_nan = not isNumeric(val)
End Function
%>
