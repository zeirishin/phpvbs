<%
'=======================================================================
'キーを指定して、配列を値で埋める
'=======================================================================
'【引数】
'  keys     = array     キーとして使用する値の配列。
'  val      = string    文字列か、あるいは値の配列。
'【戻り値】
'  値を埋めた配列を返します。
'【処理】
'  ・パラメータ val  で指定した値で配列を埋めます。 
'  ・キーとして、配列 keys  で指定した値を使用します。
'=======================================================================
Function array_fill_keys(keys, val)

    Dim ary_fill,i
    set ary_fill = Server.CreateObject("Scripting.Dictionary")
    set array_fill_keys = ary_fill
    if Not isArray(keys) then Exit Function
    If isArray(val) Then
        If uBound(val) > uBound(keys) then Exit Function
    End If

    If IsArray(val) Then
        For i = 0 to uBound(keys)
            If Not ary_fill.Exists( keys(i) ) Then ary_fill.Add keys(i), val(i)
        Next
    Else
        For i = 0 to uBound(keys)
            If Not ary_fill.Exists( keys(i) ) Then ary_fill.Add keys(i), val
        Next
    End If

    set array_fill_keys = ary_fill

End Function
%>
