<%
'=======================================================================
'二つ以上の配列を再帰的にマージする
'=======================================================================
'【引数】
'  mAry1    = array  マージするもとの配列。
'  mAry2    = array  再帰的にマージしていく配列。
'【戻り値】
'  すべての引数の配列をマージした結果の配列を返します。
'【処理】
'  ・ 一つ以上の配列の要素をマージし、 前の配列の最後にもう一つの配列の値を追加します。 
'  ・ マージした後の配列を返します。
'  ・ 入力配列が同じ文字列のキーを有している場合、 これらのキーの値は配列に一つのマージされます。
'  ・ これは再帰的に行われます。 
'  ・ このため、値の一つが配列自体を指している場合、 この関数は別の配列の対応するエントリもマージします。 
'  ・ しかし、配列が同じ数値キーを有している場合、 後の値は元の値を上書せず、追加されます。 
'=======================================================================
Function array_merge_recursive(mAry1,mAry2)

    Dim j
    Dim retAry : set retAry = Server.CreateObject("Scripting.Dictionary")

    If isObject( mAry1 ) Then
        For Each j In mAry1
            if Not retAry.Exists(j) then retAry.Add j, mAry1(j)
        Next

    ElseIf isArray( mAry1 ) Then
        For j = 0 to uBound( mAry1 )
            retAry.Add j, mAry1(j)
        Next

    End If

    If isObject( mAry2 ) Then
        For Each j In mAry2
            If isObject( mAry2(j) ) Then

                set retAry(j) = array_merge_recursive(retAry(j),mAry2(j))

            Elseif retAry.Exists(j) then
                retAry.Item(j) = array(retAry.Item(j) , mAry2(j))
            Else
                retAry.Add j, mAry2(j)
            End If
        Next

    ElseIf isArray( mAry2 ) Then
        For j = 0 to uBound( mAry2 )
            if retAry.Exists(j) then
                retAry.Item(j) = array(retAry.Item(j) , mAry2(j))
            Else
                retAry.Add j, mAry2(j)
            End If
        Next
    End If

    set array_merge_recursive = retAry

End Function
%>
