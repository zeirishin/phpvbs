<%
'=======================================================================
'ひとつまたは複数の配列をマージする
'=======================================================================
'【引数】
'  mAry1    = array  最初の配列。
'  mAry2    = array  再帰的にマージしていく配列。
'【戻り値】
'  結果の配列を返します。
'【処理】
'  ・前の配列の後ろに配列を追加することにより、 ひとつまたは複数の配列の要素をマージし、得られた配列を返します。
'  ・入力配列が同じキー文字列を有していた場合、そのキーに関する後に指定された値が、 前の値を上書きします。
'  ・しかし、配列が同じ添字番号を有していても 値は追記されるため、このようなことは起きません。
'  ・配列が一つだけ指定され、その配列が数字で添字指定されていた場合、 キーの添字が連続となるように振り直されます。 
'=======================================================================
Function array_merge(mAry1,mAry2)

    Dim j,k
    Dim ret,retAry

    If isArray(mAry1) AND isArray(mAry2) Then

        If is_empty(mAry1) Then
            array_merge = mAry2
            Exit Function
        End If

        If is_empty(mAry2) Then
            array_merge = mAry1
            Exit Function
        End If

        Dim cnt : cnt = 0
        Dim uBoundCnt : uBoundCnt = count(mAry1,0) + count(mAry2,0)
        ReDim retAry(uBoundCnt-1)

        For Each j In mAry1
            If isObject(j) Then
                set retAry(cnt) = j
            Else
                retAry(cnt) = j
            End If
            cnt = cnt + 1
        Next

        For Each j In mAry2
            If isObject(j) Then
                set retAry(cnt) = j
            Else
                retAry(cnt) = j
            End If
            cnt = cnt + 1
        Next

        array_merge = retAry

    Else
        If Not isObject(retAry) Then
            set retAry = Server.CreateObject("Scripting.Dictionary")
        End If
    
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

                    set retAry(j) = Server.CreateObject("Scripting.Dictionary")
                    set ret = array_merge(retAry(j),mAry2(j))

                    For Each k in ret
                        if retAry(j).Exists(k) then
                            retAry(j).Item(k) = ret(k)
                        Else
                            retAry(j).Add k, ret(k)
                        End If
                    Next
                Elseif retAry.Exists(j) then
                    retAry.Item(j) = mAry2(j)
                Else
                    retAry.Add j, mAry2(j)
                End If
            Next
    
        ElseIf isArray( mAry2 ) Then
            For j = 0 to uBound( mAry2 )
                if retAry.Exists(j) then
                    retAry.Item(j) = mAry2(j)
                Else
                    retAry.Add j, mAry2(j)
                End If
            Next
        End If
    
        set array_merge = retAry
    End If

End Function
%>
