<%
'=======================================================================
'ユーザが指定したコールバック関数を利用し、 追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1            = array     比較元の配列。
'  mAry2            = array     比較する対象となる配列。
'  key_compare_func = callback  使用するコールバック関数。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・ユーザが指定したコールバック関数を用いて添字を比較します。
'=======================================================================
Function array_diff_uassoc(ByVal mAry1,ByVal mAry2,key_compare_func)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = arr_dif : Exit Function
    End If

    Dim j,k,callback_ret
    For Each j in mAry1

        arr_dif.Add j, mAry1(j)

        For Each k In mAry2
            If mAry1(j) = mAry2(k) Then
                execute("callback_ret = " & key_compare_func & "(j,k)")
                If callback_ret = 0 Then
                    If arr_dif.Exists(j) Then arr_dif.Remove j
                ElseIf callback_ret < 0 Then
                    arr_dif.Remove j
                    If arr_dif.Exists(k) Then
                        arr_dif.Item( k ) = mAry2(k)
                    Else
                        arr_dif.Add k ,mAry2(k)
                    End If
                End If
            End If
        Next
    Next

    set array_diff_uassoc = arr_dif

End Function
%>
