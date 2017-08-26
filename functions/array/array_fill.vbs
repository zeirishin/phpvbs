<%
'=======================================================================
'配列を指定した値で埋める
'=======================================================================
'【引数】
'  start_index  = int       返される配列の最初のインデックス。
'  num          = int       挿入する要素数。
'  val          = string    要素に使用する値。
'【戻り値】
'  値を埋めた配列を返します。
'【処理】
'  ・パラメータ value  を値とする num  個のエントリからなる配列を埋めます。 
'  ・この際、キーは、start_index  パラメータから開始します。
'=======================================================================
Function array_fill(start_index, num, val)

    If Not isNumeric(num) or num < 1 then Exit Function

    Dim intCounter,ary()
    Dim i

    intCounter = start_index + num -1
    ReDim ary(intCounter)

    For i = start_index to intCounter
        ary(i) = val
    Next

    array_fill = ary

End Function
%>
