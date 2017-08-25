<%
'=======================================================================
'変数の整数としての値を取得する
'=======================================================================
'【引数】
'  var = mixed 文字列
'【戻り値】
'  整数
'【処理】
'  ・var  の integer としての値を返します。
'=======================================================================
Function intval(str)

    intval = 1
    If IsObject(str) or IsArray(str) Then Exit Function
    If str = true Then Exit Function

    intval = 0
    If is_empty(str) or Not isNumeric(str) Then Exit Function

    str = int(str)
    If str > 32767 Then
        intval = 32767
    Else
        intval = Cint(str)
    End If

End Function
%>
