<%
'=======================================================================
'配列の一部を展開する
'=======================================================================
'【引数】
'  mAry     = Array 入力の配列。
'  offset   = int   offset  が負の値ではない場合、要素位置の計算は、 配列 array  の offset から始められます。 offset  が負の場合、要素位置の計算は array  の最後から行われます。
'  level    = int  level が指定され、正の場合、 連続する複数の要素が返されます。level が指定され、負の場合、配列の末尾から連続する複数の要素が返されます。 省略された場合、offset  から配列の最後までの全ての要素が返されます。
'【戻り値】
'  ・切り取った部分を返します。
'【処理】
'  ・mAry から引数 offset  および level で指定された連続する要素を返します。
'=======================================================================
Function array_slice(mAry,offset,level)

    array_slice = false

    If Not isArray(mAry) Then Exit Function
    If Not isNumeric(offset) Then Exit Function
    If Not isNumeric(level) Then level = uBound(mAry)

    Dim s,e,arynum
    arynum = uBound(mAry)

    If offset >= 0 Then _
        s = offset _
    Else _
        s = arynum + offset + 1

    If level >= 0 Then _
        e = s + level _
    Else _
        e = arynum + level

    If e > arynum Then e = arynum

    Dim i,counter
    counter = 0
    ReDim tmp_ar(e-s)
    For i = s to e
        tmp_ar(counter) = mAry(i)
        counter = counter +1
    Next

    array_slice = tmp_ar

End Function
%>
