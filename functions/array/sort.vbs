<%
'=======================================================================
'配列をソートする
'=======================================================================
'【引数】
'  ary        = Array   ソートする配列
'  sort_flags = int     ソートの動作を修正するために使用することが可能です。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・配列をソートします。
'　・この関数が正常に終了すると、 各要素は低位から高位へ並べ替えられます。
'　・http://www.thinkit.co.jp/article/62/3/
'=======================================================================
Const SORT_REGULAR = 0
Const SORT_NUMERIC = 1
Const SORT_STRING  = 2
Function sort(ByRef arr, sort_flags)

    sort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)

        j = i -1
        Do While sort_helper(temp,arr(j),sort_flags)

            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    sort = true

End Function

'******************************************
Function sort_helper(temp , arr, sort_flags)

    sort_helper = false
    If isArray(temp) or isObject(temp) Then Exit Function

    sort_helper = true
    If isArray(arr) or isObject(arr) Then Exit Function

    If varType(sort_flags) <> 2 Then sort_flags = 0

    If sort_flags = SORT_REGULAR Then
        sort_helper = (temp < arr)
    ElseIf sort_flags = SORT_NUMERIC Then
        sort_helper = (intval(temp) < intval(arr))
    ElseIf sort_flags = SORT_STRING Then
        sort_helper = (Cstr(temp) < Cstr(arr))
    End If

End Function
'******************************************
%>
