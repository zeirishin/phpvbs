'=======================================================================
'データの比較にコールバック関数を用い、 追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1     = Array 最初の配列。
'  mAry2     = Array 2 番目の配列。
'  mAry1     = callback 比較用のコールバック関数。ユーザが指定したコールバック関数を用いてデータの比較を行います。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。

'【戻り値】
'  ・他の引数のいずれにも存在しない mAry1 の値の全てを有する配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の差を計算します。 
'  ・この関数は array_diff_assoc() と異なり、 データの比較に内部関数を利用します。
'=======================================================================
Function array_udiff_assoc(ByVal mAry1,ByVal mAry2, data_compare_func)

    Dim arr_udif
    set arr_udif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) then
        set array_diff_uassoc = arr_udif
        Exit Function
    End If

    Dim key,found
    For Each key In mAry1
        arr_udif.Add key, mAry1(key)
    Next

    If Not isObject(mAry2) Then Exit Function

    For Each key In mAry2
        If arr_udif.Exists( key ) Then
            execute("found = " & data_compare_func & "(arr_udif(key), mAry2(key))")
            If found = 0 Then
                If arr_udif.Exists( key ) Then arr_udif.Remove key
            ElseIf found < 0 Then
                If isObject(mAry2(key)) Then
                    set arr_udif.Item( key ) = mAry2(key)
                Else
                    arr_udif.Item( key ) = mAry2(key)
                End If
            End if
        End If
    Next

    set array_udiff_assoc = arr_udif

End Function
