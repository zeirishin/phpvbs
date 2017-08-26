<%
'=======================================================================
'配列の末尾から要素を取り除く
'=======================================================================
'【引数】
'  mAry = array  値を取り出す配列。
'【戻り値】
'  配列 mAry の最後の値を取り出して返します。
'  mAry が空 (または、配列でない) の場合、 NULL が返されます。
'【処理】
'  ・array  の最後の値を取り出して返します。
'  ・配列 array  は、要素一つ分短くなります。
'=======================================================================
Function array_pop(ByRef mAry)

    If Not isArray(mAry) Then
        array_pop = null
        Exit Function
    End If

    Dim intCounter
    intCounter = uBound( mAry )
    array_pop = mAry( intCounter )
    ReDim Preserve mAry(intCounter - 1)

End Function
%>
