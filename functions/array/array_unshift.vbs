<%
'=======================================================================
'一つ以上の要素を配列の最初に加える
'=======================================================================
'【引数】
'  mAry     = Array 配列
'  mVal     = mixed 追加する要素
'【戻り値】
'  ・処理後の mAry  の要素の数を返します。
'【処理】
'  ・リストの要素は全体として加えられるため、 加えられた要素の順番は変わらないことに注意してください。 
'  ・配列の数値添字はすべて新たにゼロから振りなおされます。 
'  ・リテのキーについては変更されません。
'=======================================================================
Function array_unshift(ByRef mAry, ByVal mVal)

    Dim intCounter
    Dim intElementCount

    If IsArray(mAry) Then
        If IsArray(mVal) Then

            ret = array_push(mVal,mAry)
            mAry = mVal

        Else

            ReDim Preserve mAry(UBound(mAry) + 1)

            For intCounter = UBound(mAry) to 1 Step -1
                mAry(intCounter) = mAry(intCounter -1)
            Next

            mAry(0) = mVal

        End If
    Else
        If IsArray(mVal) Then
            mAry = mVal
        Else
            mAry = Array(mVal)
        End If
    End If

    array_unshift = UBound(mAry)

End Function
%>
