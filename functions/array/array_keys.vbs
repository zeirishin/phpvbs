<%
'=======================================================================
'配列のキーをすべて返す
'=======================================================================
'【引数】
'  mAry         = array  返すキーを含む配列。
'  search_value = mixed  指定した場合は、これらの値を含むキーのみを返します。
'  strict       = mixed  検索時に型比較を行います。
'【戻り値】
'  mAry のすべてのキーを配列で返します。
'【処理】
'  ・配列 mAry から全てのキー (数値および文字列) を返します。
'  ・オプション search_value  が指定された場合、 指定した値に関するキーのみが返されます。
'  ・指定されない場合は、mAry から全てのキーが返されます。
'  ・strict  パラメータを使って、 比較に型も比較することができます。
'=======================================================================
Function array_keys(mAry,search_value,strict)

    Dim tmp_arr
    Dim key
    Dim include
    Dim addArr
    Dim cnt : cnt = 0

    addArr = true
    If [==](search_value,empty) Then
        addArr = false
        ReDim tmp_arr( count(mAry,0)-1 )
    End If

    If isObject( mAry ) Then

        For Each key In mAry
            include = true
            If [!=](search_value,empty) Then
                If strict = true Then
                    If [!=](mAry(key) , search_value) or (varType(mAry(key)) <> varType(search_value)) Then
                        include = false
                    End If
                ElseIf [!=](mAry(key) , search_value) Then
                    include = false
                End If
            End If

            If include = true Then
                If addArr Then
                    [] tmp_arr, key
                Else
                    tmp_arr(cnt) = key
                    cnt = cnt + 1
                End If
            End If
        Next

    ElseIf isArray(mAry) Then

        For cnt = 0 to uBound(mAry)

            include = true
            If [!=](search_value,empty) Then

                If strict = true Then
                    If [!=](mAry(cnt) , search_value) or (varType(mAry(cnt)) <> varType(search_value)) Then
                        include = false
                    End If
                ElseIf [!=](mAry(cnt) , search_value) Then
                    include = false
                End If
            End If

            If include = true Then
                If addArr Then
                    [] tmp_arr, cnt
                Else
                    tmp_arr(cnt) = cnt
                End If
            End if
        Next
    End If

    array_keys = tmp_arr

End Function
%>
