<%
'=======================================================================
'複数の多次元の配列をソートする
'=======================================================================
'【引数】
'  arr  = array  ソートしたい配列。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  array_multisort() は、多次元の配列をその次元の一つでソートする際に使用可能です。
'=======================================================================
Function array_multisort(ByRef arr)

    array_multisort = false
    If not isArray(arr) Then Exit Function

    Dim key
    For key = 0 to uBound(arr)
        array_multisort arr(key)
    Next

    sort arr,0

    array_multisort = true

End Function

%>
