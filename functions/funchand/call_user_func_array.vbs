<%
'=======================================================================
'パラメータの配列を指定してユーザ関数をコールする
'=======================================================================
'【引数】
'  callback     = mixed  コールする関数。このパラメータに array($classname, $methodname) を指定することにより、 クラスメソッドも静的にコールすることができます。
'  param_arr    = array  関数に渡すパラメータを指定する配列。
'【戻り値】
'  関数の結果、あるいはエラー時に FALSE を返します。
'【処理】
'  ・param_arr  にパラメータを指定して、 function  で指定したユーザ定義関数をコールします。
'=======================================================================
Function call_user_func_array(callback,param_arr)

    Dim thisFunc,thisParam,retval
    If isArray(callback) Then
        thisFunc  = callback(0) & "." & callback(1)
    Else
        thisFunc = callback
    End If

    If isArray(param_arr) Then
        Dim key
        For Each key In parameter
            If len( thisParam ) > 0 Then
                thisParam = thisParam & "," & key
            Else
                thisParam = key
            End IF
        Next
    Else
        thisParam = param_arr
    End If
    execute("retval = " & thisFunc & "(" & thisParam & ")")
    call_user_func_array = retval

End Function
%>
