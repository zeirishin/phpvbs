<%
'=======================================================================
'配列の要素をフィルタリングする
'=======================================================================
'【引数】
'  mAry         = aarray    処理する配列。
'  callback     = callback  使用するコールバック関数。コールバック関数が与えられなかった場合、 input のエントリの中で FALSE と等しいもの (boolean への変換 を参照ください) がすべて削除されます

'【戻り値】
'  フィルタリングされた結果の配列を返します。
'【処理】
'  ・mAry のエントリの中で FALSE と等しいもの がすべて削除されます。
'=======================================================================
Function array_filter(ByRef mAry,callback)

    If isArray(mAry) Then

        Dim intCounter,i,strType,callback_ret
        intCounter = uBound(mAry)

        For i = 0 to intCounter
            callback_ret = true
            If Len( callback ) > 0 Then _
                execute("callback_ret = " & callback & "(mAry(i))")

            If callback_ret = true and ( mAry(i) = empty or isNull(mAry(i)) ) Then
                mAry = array_remove(mAry,i)
                call array_filter(mAry,callback)
                Exit For
            End If
        Next

    ElseIf isObject(mAry) Then
        Dim j
        For Each j IN mAry
            callback_ret = true
            If Len( callback ) > 0 Then _
                execute("callback_ret = " & callback & "(mAry(i))")

            If callback_ret = true and ( mAry(j) = empty or isNull(mAry(j)) ) Then _
                mAry.Remove j
        Next

    End If

    array_filter = true

End Function
%>
