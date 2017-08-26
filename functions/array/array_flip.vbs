<%
'=======================================================================
'配列のキーと値を反転する
'=======================================================================
'【引数】
'  trans    = array  反転を行うキー/値の組。
'【戻り値】
'  成功した場合に反転した配列、失敗した場合に 空のオブジェクト を返します。
'【処理】
'  ・配列を反転して返します。
'  ・すなわち、trans  のキーが値となり、 trans  の値がキーとなります。
'=======================================================================
Function array_flip(trans)

    Dim aryObj : set aryObj = Server.CreateObject("Scripting.Dictionary")

    If Not isArray(trans) and Not isObject(trans) Then
        set array_flip = aryObj
        Exit Function
    End If


    If isArray(trans) Then

        Dim i
        For i = 0 to uBound(trans)
            aryObj( trans(i) ) = i
        Next

    Elseif isObject(trans) Then

        Dim j
        For Each j In trans
            aryObj( trans(j) ) = j
        Next

    End If

    set array_flip = aryObj
End Function
%>
