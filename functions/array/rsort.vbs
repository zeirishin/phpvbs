<%
'=======================================================================
'配列を逆順にソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、配列を逆順に(高位から低位に)ソートします。
'=======================================================================
Const SORT_REGULAR = 0
Const SORT_NUMERIC = 1
Const SORT_STRING  = 2
Function rsort(ByRef arr, sort_flags)

    rsort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)

        j = i -1
        Do While rsort_helper(temp,arr(j),sort_flags)

            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    rsort = true

End Function

'******************************************
Function rsort_helper(temp , arr, sort_flags)

    rsort_helper = true
    If isArray(temp) or isObject(temp) Then Exit Function

    rsort_helper = false
    If isArray(arr) or isObject(arr) Then Exit Function

    If varType(sort_flags) <> 2 Then sort_flags = 0

    If sort_flags = SORT_REGULAR Then
        rsort_helper = (temp > arr)
    ElseIf sort_flags = SORT_NUMERIC Then
        rsort_helper = (intval(temp) > intval(arr))
    ElseIf sort_flags = SORT_STRING Then
        rsort_helper = (Cstr(temp) > Cstr(arr))
    End If

End Function
'******************************************
%>
