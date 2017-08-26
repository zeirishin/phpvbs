<%
'=======================================================================
'ユーザが指定したコールバック関数を利用し、 追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1                = array     比較元の配列。
'  mAry2                = array     比較する対象となる配列。
'  date_compare_func    = callback  使用するコールバック関数。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'  key_compare_func     = callback  キー（添字）の比較は、コールバック関数
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・ユーザが指定したコールバック関数を用いて添字を比較します。
'=======================================================================
Function array_udiff_uassoc(ByVal mAry1,ByVal mAry2, data_compare_func,key_compare_func)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = arr_dif : Exit Function
    End If

    Dim j,k,key_result,data_result,found
    For Each j in mAry1

        found = false
        For Each k In mAry2
            execute("key_result  = " & key_compare_func & "(j,k)")
            execute("data_result = " & data_compare_func & "(mAry1(j),mAry2(k))")

            If key_result = 0 and data_result = 0 Then
                found = true
                Exit For
            End If
        Next

        If Not found Then
             arr_dif.Add j , mAry1(j)
        End If
    Next

    set array_udiff_uassoc = arr_dif

End Function
%>
