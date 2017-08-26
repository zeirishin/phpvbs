<%
'=======================================================================
'ユーザー定義の比較関数を使用して、配列を値でソートする
'=======================================================================
'【引数】
'  ary          = Array   入力の配列。
'  cmp_function = int     比較関数は、最初の引数が 2 番目の引数より小さいか、等しいか、大きい場合に、 それぞれゼロ未満、ゼロに等しい、ゼロより大きい整数を返す 必要があります。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、ユーザー定義の比較関数により配列をその値でソートします。 
'  ・ソートしたい配列を複雑な基準でソートする必要がある場合、 この関数を使用するべきです。
'=======================================================================
Function usort(ByRef arr, cmp_function)

'ユーザー定義関数の例
'きちんとtrue falseを返さないと動かない
'Function cmp(a,b)
'
'    If [==](a,b) Then
'        cmp = false
'    Else
'        If (a < b) Then
'            cmp = false
'        Else
'            cmp = true
'        End If
'    End If
'End Function



    usort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)
        j = i -1
        Do While usort_helper(temp,arr(j),cmp_function)
            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    usort = true


End Function

'*****************************************************
Function usort_helper(temp,arr,cmp_function)

    Dim output
    execute ("output = " & cmp_function & "(temp,arr)")
    usort_helper = output

End Function
'******************************************
%>
