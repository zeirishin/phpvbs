<%
'=======================================================================
'指定したキーまたは添字が配列にあるかどうかを調べる
'=======================================================================
'【引数】
'  key      = mixed  配列
'  sAry     = array  キーが存在するかどうかを調べたい配列。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・指定した key  が配列に設定されている場合、 array_key_exists() は TRUE を返します。 
'  ・key  は配列添字として使用できる全ての値を使用可能です。
'=======================================================================
Function array_key_exists(key, sAry)

    array_key_exists = false
    If isObject(sAry) Then
        if sAry.Exists( key ) then array_key_exists = true
    ElseIf isArray(sAry) and isNumeric(key) Then
        If (uBound(sAry)+1) > key Then
            If Not isNull(sAry(key)) Then array_key_exists = true
        End If
    End If

End Function
%>
