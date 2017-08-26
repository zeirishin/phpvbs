<%
'=======================================================================
'追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元の配列。
'  mAry2    = array  比較する対象となる配列。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・array_diff() とは異なり、 配列のキーを用いて比較を行います。
'=======================================================================
Function array_diff_assoc(ByVal mAry1,ByVal mAry2)

    Dim retAry
    set retAry = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = retAry : Exit Function
    End If

    Dim j,k
    For Each j in mAry1

        retAry.Add j, mAry1(j)

        For Each k In mAry2
            if j = k and mAry1(j) = mAry2(k) Then retAry.Remove k
        Next
    Next

    set array_diff_assoc = retAry

End Function
%>
