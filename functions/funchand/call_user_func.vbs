'=======================================================================
'最初の引数で指定したユーザ関数をコールする
'=======================================================================
'【引数】
'  callback     = mixed  コールする関数。このパラメータに array(classname, methodname) を指定することにより、 クラスメソッドも静的にコールすることができます。
'  parameter    = mixed  この関数に渡す、ゼロ個以上のパラメータ。
'【戻り値】
'  関数の結果、あるいはエラー時に FALSE を返します。
'【処理】
'  ・パラメータ callback で指定した ユーザ定義のコールバック関数をコールします。
'=======================================================================
Function call_user_func(callback,parameter)

    Dim thisFunc,retval
    If isArray(callback) Then
        thisFunc  = callback(0) & "." & callback(1)
    Else
        thisFunc = callback
    End If

    execute("retval = " & thisFunc & "(parameter)")
    call_user_func = retval
End Function
