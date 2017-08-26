<%
'=======================================================================
'一つの要素を配列の最後に追加する
'=======================================================================
'【引数】
'  mAry     = mixed  配列
'  mVal     = mixed  追加する要素
'【戻り値】
'  値を返しません。
'【処理】
'  ・渡された変数を mAry  の最後に加えます。
'=======================================================================
Sub [](ByRef mAry, ByVal mVal)

    If IsArray(mAry) Then
        Dim counter : counter = UBound(mAry) + 1
        ReDim Preserve mAry(counter)
        mAry(counter) = mVal
    Else
        mAry = Array(mVal)
    End If

End Sub
%>
