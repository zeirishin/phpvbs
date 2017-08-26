<%
'=======================================================================
'ユーザ定義の比較関数を用いて、キーで配列をソートする
'=======================================================================
'【引数】
'  ary          = Array   入力の配列。
'  cmp_function = int     比較用のコールバック関数。関数 cmp_function は、 array のキーペアによって満たされる 2 つのパラメータを受け取ります。 この比較関数が返す値は、最初の引数が二番目より小さい場合は負の数、 等しい場合はゼロ、そして大きい場合は正の数でなければなりません。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・uksort() は、 ユーザ定義の比較関数を用いて配列のキーをソートします。
'  ・ソートしたい配列を複雑な基準でソートする必要がある場合には、 この関数を使う必要があります。
'=======================================================================
Function uksort(ByRef arr, cmp_function)

    uksort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    usort keys,cmp_function

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    uksort = true

End Function
%>
