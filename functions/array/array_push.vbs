<%
'=======================================================================
'一つ以上の要素を配列の最後に追加する
'=======================================================================
'【引数】
'  mAry     = array  配列
'  mVal     = mixed  追加する要素
'【戻り値】
'  処理後の配列の中の要素の数を返します。
'【処理】
'  ・渡された変数を mAry の最後に加えます。
'  ・配列 mAry の長さは渡された変数の数だけ増加します。
'=======================================================================
Function array_push(ByRef mAry, ByVal mVal)

    Dim intCounter
    Dim intElementCount

    If IsArray(mAry) Then
        If IsArray(mVal) Then

            intElementCount = UBound(mAry)
            ReDim Preserve mAry(intElementCount + UBound(mVal) + 1)

            For intCounter = 0 to UBound(mVal)
                mAry(intElementCount + intCounter + 1) = mVal(intCounter)
            Next

        Else
            ReDim Preserve mAry(UBound(mAry) + 1)
            mAry(UBound(mAry)) = mVal
        End If
    Else

        If IsArray(mVal) Then
            mAry = mVal
        Else
            mAry = Array(mVal)
        End If
    End If

    array_push = UBound(mAry)

End Function
%>
