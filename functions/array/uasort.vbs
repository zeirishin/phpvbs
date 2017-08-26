<%
'=======================================================================
'ユーザ定義の比較関数で配列をソートし、連想インデックスを保持する
'=======================================================================
'【引数】
'  ary          = Array   入力の配列。
'  cmp_function = int     ユーザ定義の比較関数の例については、 usort() および uksort()  を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・ この関数は、配列インデックスが関連する配列要素との関係を保持するような配列をソートします。
'  ・ 主に実際の配列の順序に意味がある連想配列をソートするためにこの関数は使用されます。 
'=======================================================================
Function uasort(ByRef arr, cmp_function)

    uasort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    usort keys,cmp_function

    For Each key In keys
        found = array_keys(arr,key,true)
        If isArray(found) Then
            For cnt = 0 to uBound(found)
                If Not new_arr.Exists(found(cnt)) Then
                    new_arr.Add found(cnt), arr(found(cnt))
                End If
            Next
        Else
            new_arr.Add found, arr(found)
        End If
    Next

    set arr = new_arr

    uasort = true

End Function
%>
