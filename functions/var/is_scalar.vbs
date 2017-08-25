<%
'=======================================================================
'変数がスカラかどうかを調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str がスカラの場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・ 指定した変数がスカラかどうかを調べます。
'  ・ スカラ変数には integer、float、string あるいは boolean が含まれます。
'  ・ array、object および resource はスカラではありません。 
'=======================================================================
Function is_scalar(str)

    is_scalar = false
    If isArray(str) or isObject(str) Then Exit Function
    if isNull(str) Then Exit Function
    is_scalar = true

End Function
%>
