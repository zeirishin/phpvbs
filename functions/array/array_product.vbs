<%
'=======================================================================
'配列の値の積を計算する
'=======================================================================
'【引数】
'  mAry     = array  配列
'【戻り値】
'  積を、integer あるいは float で返します。
'【処理】
'  ・配列の各要素の積を計算します。
'=======================================================================
Function array_product(mAry)

    If Not isArray( mAry ) Then Exit Function

    Dim j,product
    product = 1

    For Each j In mAry
        If isNumeric(j) Then product = product * j
    Next

    array_product = product

End Function
%>
