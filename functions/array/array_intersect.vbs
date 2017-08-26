<%
'=======================================================================
'配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1    = array  値を調べるもととなる配列。
'  mAry2    = array  値を比較する対象となる配列。
'【戻り値】
'  mAry1 の値のうち、すべての引数に存在する値のものを含む配列を返します。
'【処理】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'=======================================================================
Function array_intersect(mAry1,mAry2)

    Dim key
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            If len(array_search(mAry1(key),mAry2,false)) > 0 Then
                output.Add key, mAry1(key)
            End If
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            If len(array_search(mAry1(key),mAry2,false)) > 0 Then
                output.Add key, mAry1(key)
            End If
        Next
    End If

    set array_intersect = output

End Function
%>
