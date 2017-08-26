<%
'=======================================================================
'一方の配列をキーとして、もう一方の配列を値として、ひとつの配列を生成する
'=======================================================================
'【引数】
'  keys     = array  キーとして使用する配列。
'  values   = array  値として使用する配列。
'【戻り値】
'  作成した配列を返します。
'  互いの配列の要素の数が合致しない場合に FALSE を返します。
'【処理】
'  ・keys  配列の値をキーとして、
'  ・また values  配列の値を対応する値として生成した 配列 を作成します。
'=======================================================================
Function array_combine(keys,values)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")

    If uBound(keys) <> uBound(values) Then
        set array_combine = obj
        Exit Function
    End If

    Dim i
    For i = 0 to uBound(keys)
        If obj.Exists( keys(i) ) Then
            obj.Item( keys(i) ) = values(i)
        Else
            obj.Add keys(i) , values(i)
        End If
    Next

    set array_combine = obj

End Function
%>
