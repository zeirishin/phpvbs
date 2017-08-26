<%
'=======================================================================
'配列の差を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元の配列。
'  mAry2    = array  比較する対象となる配列。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・二つの要素は、(string) elem1 = (string) elem2  の場合のみ等しいと見直されます。
'  ・言い換えると、文字列表現が同じ場合となります。 
'=======================================================================
Function array_diff(ByVal mAry1,ByVal mAry2)

    Dim arr_dif,key_c,key,found
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then
        set mAry1 = array2Dic(mAry1)
    End If

    If isArray(mAry2) Then
        set mAry2 = array2Dic(mAry2)
    End If


    For Each key In mAry1

        found = false
        For Each key_c In mAry2
            If mAry1(key) = mAry2(key_c) Then
                found = true
                Exit For
            End If
        Next

        If Not found Then
            arr_dif.add key, mAry1(key)
        End If
    Next


    set array_diff = arr_dif

End Function

'=======================================================================
'配列をディクショナリに変換する
'=======================================================================
'【引数】
'  arr  = array  配列
'【戻り値】
'  ディクショナリオブジェクト。
'【処理】
'  ・渡された配列を ディクショナリオブジェクトに変換します。
'=======================================================================
Function array2Dic(ByVal myAry)

    Dim i,tmpObj
    set tmpObj = Server.CreateObject("Scripting.Dictionary")
    For i = 0 to uBound(myAry)
        tmpObj.add i, myAry(i)
    Next
    set array2Dic = tmpObj

End Function
%>
