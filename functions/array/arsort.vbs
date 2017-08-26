<%
'=======================================================================
'連想キーと要素との関係を維持しつつ配列を逆順にソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、 連想配列において各配列のキーと要素との関係を維持しつつ配列をソートします。
'  ・この関数は、 主に実際の要素の並び方が重要である連想配列をソートするために使われます。
'=======================================================================
Function arsort(ByRef arr, sort_flags)

    arsort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    rsort keys,sort_flags

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

    arsort = true

End Function
%>
