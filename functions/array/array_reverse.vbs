<%
'=======================================================================
'要素を逆順にした配列を返す
'=======================================================================
'【引数】
'  ary              = Array 入力の配列。
'【戻り値】
'  ・逆転させた配列を返します。
'【処理】
'  ・配列を受け取って、要素の順番を逆にした新しい配列を返します。
'=======================================================================
Function array_reverse(mAry)

    Dim arr_len,i

    If isArray(mAry) Then

        Dim tmp_ar()
        Dim newkey

        arr_len = uBound(mAry)
        ReDim tmp_ar(arr_len)

        For i = 0 to arr_len
            newkey = arr_len -i
            tmp_ar(i) = mAry(newkey)
        Next

        array_reverse = tmp_ar

    ElseIf isObject(mAry) Then

        Dim tmpObj,j,cnt

        cnt = 0
        set tmpObj = Server.CreateObject("Scripting.Dictionary")
        arr_len = mAry.Count-1

        ReDim index_values(arr_len),index_keys(arr_len)

        For Each j In mAry
            index_values(cnt) = mAry(j)
            index_keys(cnt)   = j
            cnt = cnt + 1
        Next

        For i = cnt-1 To 0 Step -1

            If Not tmpObj.Exists(Cstr(index_keys(i))) Then
                tmpObj.add Cstr(index_keys(i)),index_values(i)
            End if
        Next

        set array_reverse = tmpObj
    End If
End Function
%>
