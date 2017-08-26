<%
'=======================================================================
'配列の先頭から要素を一つ取り出す
'=======================================================================
'【引数】
'  ary     = Array 配列
'【戻り値】
'  ・ary  の最初の値を取り出して返します。
'  ・array  が空の場合 (または配列でない場合)、 NULL が返されます。
'【処理】
'  ・配列 ary  は、要素一つ分だけ短くなり、全ての要素は前にずれます。 
'  ・数値添字の配列のキーはゼロから順に新たに振りなおされますが、 リテラルのキーはそのままになります。
'=======================================================================
Function array_shift(ByRef ary)

    If Not isArray(ary) and Not isObject(ary) then
        array_shift = null
        Exit Function
    End If

    Dim i,key : i = 0

    If isArray(ary) Then
        array_shift = ary(0)

        For i = 0 to uBound(ary)-1
            ary(i) = ary(i+1)
        Next
        Redim Preserve ary(UBound(ary) - 1)

    ElseIf isObject(ary) Then
        For Each key In ary
            array_shift = ary(key)
            ary.Remove(key)
            Exit For
        Next
    End if

End Function
%>
