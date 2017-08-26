<%
'=======================================================================
'配列の値の数を数える
'=======================================================================
'【引数】
'  mAry     = array  値を数える配列。
'【戻り値】
'   mAry のキーとその登場回数を組み合わせた連想配列を作成します。
'【処理】
'  ・配列 mAry の値をキーとし、mAry におけるその値の出現回数を値とした配列を返します。
'【エラー / 例外】
'  ・string あるいは integer 以外の型の要素が登場すると致命的なエラーが発生します。
'=======================================================================
Function array_count_values(mAry)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")
    if Not isArray(mAry) and Not isObject(mAry) Then
        Set array_count_values = obj
        Exit Function
    End If

    Dim j,k
    Dim intCounter


    For Each j In mAry

        intCounter = 0

        For Each k In mAry
            If j = k Then intCounter = intCounter + 1
        Next

        If Not obj.Exists(j) Then obj.Add j, intCounter

    Next

    Set array_count_values = obj

End Function
%>
