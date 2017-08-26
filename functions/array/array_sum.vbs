<%
'=======================================================================
'配列の中の値の合計を計算する
'=======================================================================
'【引数】
'  mAry         = Array 入力の配列。
'【戻り値】
'  ・値の合計を整数または float として返します。
'【処理】
'  ・配列の中の値の合計を整数または float として返します。
'=======================================================================
Function array_sum(mAry)

    array_sum = 0
    If Not isArray(mAry) and Not isObject(mAry) Then Exit Function

    Dim key
    If isObject(mAry) Then
        For Each key in mAry
            array_sum = array_sum + mAry(key)
        Next
    Else
        For Each key in mAry
            array_sum = array_sum + key
        Next
    End If

End Function
%>
