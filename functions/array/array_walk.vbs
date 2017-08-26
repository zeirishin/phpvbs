<%
'=======================================================================
'配列の全ての要素にユーザ関数を適用する
'=======================================================================
'【引数】
'  arr      = array     入力の配列。
'  callback = callback  引数を二つとります。 array  パラメータの値が最初の引数、 キー/添字は二番目の引数となります。funcname  により配列の値そのものを変更する必要がある場合、 funcname  の最初の引数は 参照  として渡す必要があります。この場合、配列の要素に加えた変更は、 配列自体に対して行われます。 
'  userdata = array     userdata  パラメータが指定された場合、 コールバック関数 funcname  への三番目の引数として渡されます。
'【戻り値】
'  成功した場合に TRUE を返します。
'【処理】
'  ・arr の各要素に callback  関数を適用します。
'=======================================================================
Function array_walk(ByRef arr, callback, userdata)

    Dim key

    If Len( callback ) = 0 Then Exit Function

    If isArray( arr ) Then

        For key = 0 to uBound( arr )
            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    ElseIf isObject( arr ) Then

        Dim return_val

        For Each key In arr
            execute("call " & callback & "(arr.Item(key),key,userdata)")
        Next

    End If

    array_walk = true

End Function
%>
