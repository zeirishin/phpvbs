'=======================================================================
'指定長、指定した値で配列を埋める
'=======================================================================
'【引数】
'  mAry         = array  値を埋めるもととなる配列。
'  pad_size     = int    新しい配列のサイズ。
'  pad_value    = mixed  mAry が pad_size より小さいときに、 埋めるために使用する値。
'【戻り値】
'  pad_size  で指定した長さになるように値 pad_value  で埋めて mAry のコピーを返します。
'  pad_size  が正の場合、配列の右側が埋められます。 
'  負の場合、配列の左側が埋められます。 
'  pad_size  の絶対値が mAry の長さ以下の場合、埋める処理は行われません。
'【処理】
'  pad_size  で指定した長さになるように値 pad_value  で埋めて mAry のコピーを返します。
'=======================================================================
Function array_pad(ByVal mAry, pad_size, pad_value)

    If Not isArray( mAry ) Then Exit Function
    If Not isNumeric( pad_size ) Then Exit Function

    Dim pad,aryCounter,newLength,i,intCounter

    If pad_size < 0 Then
        newLength = pad_size * -1
    Else
        newLength = pad_size
    End If
    newLength = newLength -1

    aryCounter = uBound(mAry)
    If newLength > aryCounter Then

        ReDim pad(newLength)
        intCounter = 0
        For i = 0 to newLength
            If pad_size < 0 Then
                If newLength - aryCounter > i Then
                    pad(i) = pad_value
                Else
                    pad(i) = mAry(intCounter)
                    intCounter = intCounter + 1
                End If
            Else
                If i > aryCounter Then
                    pad(i) = pad_value
                Else
                    pad(i) = mAry(i)
                End If
            End If
        Next
    Else
        pad = mAry
    End If

    array_pad = pad

End Function
