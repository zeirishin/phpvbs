<%
'=======================================================================
'配列をキーでソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・キーとデータの関係を維持しつつ、配列をキーでソートします。 
'  ・この関数は、主として連想配列において有用です。
'=======================================================================
Function ksort(ByRef arr, sort_flags)

    ksort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    sort keys,sort_flags

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    ksort = true

End Function
%>
