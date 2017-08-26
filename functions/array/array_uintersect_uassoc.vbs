<%
'=======================================================================
'データと添字の比較にコールバック関数を用い、 追加された添字の確認も含めて配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1                = Array     最初の配列。
'  mAry2                = Array     2 番目の配列。
'  data_compare_func    = callback  比較用のコールバック関数。比較は、ユーザが指定したコールバック関数を利用して行われます。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'  key_compare_func     = callback  キーの比較用のコールバック関数。
'【戻り値】
'  ・他の全ての引数に現れる mAry1 の値を含む配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の共通項を計算します。
'  ・キーが比較に使用されることに注意してください。 
'  ・データと添字は、それぞれ個別のコールバック関数を用いて比較されます。
'=======================================================================
Function array_uintersect_uassoc(mAry1,mAry2,data_compare_func,key_compare_func)

'Callbackの例
'function rmul(v, w)
'    rmul = 0
'    If isObject(v) or isArray(v) Then
'        rmul = 1
'    Elseif isObject(w) or isArray(w) Then
'        rmul = 1
'    End If
'    If rmul = 1 then Exit FUnction
'    If v = w Then
'        rmul = 0
'    Else
'        rmul = 1
'    End If
'End Function

    Dim key,key_c
    Dim found,key_found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then

        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    ElseIf isObject(mAry1) Then

        For Each key In mAry1
            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    End If

    set array_uintersect_uassoc = output

End Function
%>
