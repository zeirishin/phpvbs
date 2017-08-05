Sub [](ByRef mAry, ByVal mVal)

    If IsArray(mAry) Then
        Dim counter : counter = UBound(mAry) + 1
        ReDim Preserve mAry(counter)
        mAry(counter) = mVal
    Else
        mAry = Array(mVal)
    End If

End Sub
