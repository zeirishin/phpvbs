<%
'=======================================================================
'キーを基準にして配列の差を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元の配列。
'  mAry2    = array  比較する対象となる配列。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・この関数は array_diff() に似ていますが、 値ではなくキーを用いて比較するという点が異なります。
'=======================================================================
Function array_diff_key(ByVal mAry1,ByVal mAry2)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) then
        set array_diff_uassoc = arr_dif
        Exit Function
    End If

    Dim key
    For Each key In mAry1
        arr_dif.Add key, mAry1(key)
    Next

    If isObject(mAry2) Then
        For Each key In mAry2
            If arr_dif.Exists( key ) Then arr_dif.Remove key
        Next
    End If

    set array_diff_key = arr_dif

End Function
%>
