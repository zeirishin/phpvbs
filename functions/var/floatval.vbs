'=======================================================================
'変数の float 値を取得する
'=======================================================================
'【引数】
'  str  = mixed  あらゆるスカラ型を指定できます。配列あるいはオブジェクトに floatval() を使用することはできません。
'【戻り値】
'  指定した変数の float 値を返します。
'【処理】
'  ・変数 str の float 値を返します。
'=======================================================================
Function floatval(str)

    floatval = false
    If isArray(str) or isObject(str) Then Exit Function
    If not isNumeric(str) Then Exit Function
    floatval = CDbl(str)

End Function
