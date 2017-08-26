<%
'=======================================================================
'コールバック関数を用いて配列を普通の値に変更することにより、配列を再帰的に減らす
'=======================================================================
'【引数】
'  mAry     = array     入力の配列。
'  callback = callback  コールバック関数。
'  initial  = int       オプションの intial が利用可能な場合、処理の最初で使用されたり、 配列が空の場合の最終結果として使用されます。
'【戻り値】
'   結果の値を返します。
'   配列が空で initial が渡されなかった場合は、 array_reduce() は NULL を返します。 
'【処理】
'  ・配列 mAry の各要素に callback 関数を繰り返し適用し、 配列を一つの値に減らします。
'=======================================================================
Function array_reduce(ByVal mAry, callback, ByVal initial)

    array_reduce = null
    If len( initial ) > 0 Then array_reduce = initial
    If not isArray( mAry ) and not isObject( mAry ) Then Exit Function

    Dim acc : acc = initial
    Dim key

    If isObject( mAry ) Then
        For Each key In mAry
            execute("acc = " & callback & "(acc, mAry(key))")
        Next

    ElseIf isArray( mAry ) Then

        Dim lon : lon = uBound( mAry )
        For key = 0 to lon
            execute("acc = " & callback & "(acc, mAry(key))")
        Next
    End If

    array_reduce = acc

End Function
%>
