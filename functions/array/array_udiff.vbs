'=======================================================================
'データの比較にコールバック関数を用い、配列の差を計算する
'=======================================================================
'【引数】
'  mAry1     = Array 最初の配列。
'  mAry2     = Array 2 番目の配列。
'  mAry1     = callback 比較用のコールバック関数。ユーザが指定したコールバック関数を用いてデータの比較を行います。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。

'【戻り値】
'  ・他の引数のいずれにも存在しない mAry1 の値の全てを有する配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の差を計算します。 
'  ・この関数は array_diff() と異なり、 データの比較に内部関数を利用します。
'=======================================================================
Function array_udiff(mAry1,mAry2,data_compare_func)

    Dim arr_udif,key_c,key,found

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = retAry : Exit Function
    End If

    If isArray(mAry1) and isArray(mAry2) Then

        For Each key In mAry1

            found = 0
            For Each key_c In mAry2
                execute("found = " & data_compare_func & "(key, key_c)")
                If found <> 0 Then
                    Exit For
                End If
            Next

            If found > 0 Then
                [] arr_udif, mAry1(key)
            ElseIf found < 0 Then
                [] arr_udif, mAry2(key_c)
            End If
        Next

        array_udiff = arr_udif

    ElseIf isObject(mAry1) and isObject(mAry2) Then

        set arr_udif = Server.CreateObject("Scripting.Dictionary")

        For Each key In mAry1

            found = 0
            For Each key_c In mAry2
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                If found <> 0 Then
                    Exit For
                End If
            Next

            If found > 0 Then
                If arr_udif.Exists(key) Then
                    arr_udif.Item(key) = mAry1(key)
                Else
                    arr_udif.Add key, mAry1(key)
                End If
            ElseIf found < 0 Then
                If arr_udif.Exists(key_c) Then
                    arr_udif.Item(key_c) = mAry2(key_c)
                Else
                    arr_udif.Add key_c, mAry2(key_c)
                End If
            End If

        Next

        set array_udiff = arr_udif

    End If

End Function
