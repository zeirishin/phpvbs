'=======================================================================
'配列の全ての値を返す
'=======================================================================
'【引数】
'  mAry     = array 配列。
'【戻り値】
'  数値添字の値の配列を返します。
'【処理】
'  ・配列から全ての値を取り出し、数値添字をつけた配列を返します。
'=======================================================================
Function array_values(mAry)

    Dim tmp_ar
    Dim key,counter : counter= 0

    If isObject(mAry) Then

        ReDim tmp_ar(mAry.Count -1)

        For Each key In mAry
            If isObject(mAry(key)) Then
                set tmp_ar(counter) = mAry(key)
            Else
                tmp_ar(counter) = mAry(key)
            End if
            counter = counter + 1
        Next

    ElseIf isArray(mAry) Then
        tmp_ar = mAry
    End If

    array_values = tmp_ar

End Function
