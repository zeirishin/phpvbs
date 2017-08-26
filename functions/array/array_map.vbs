<%
'=======================================================================
'指定した配列の要素にコールバック関数を適用する
'=======================================================================
'【引数】
'  callback = callback  配列の各要素に適用するコールバック関数。
'  arr      = array     コールバック関数を適用する配列。
'【戻り値】
'  arr の各要素に callback  関数を適用した後、 その全ての要素を含む配列を返します。
'【処理】
'  ・arr の各要素に callback  関数を適用します。
'=======================================================================
Function array_map(callback, arr)

    Dim key
    Dim tmp_ar

    If isArray( arr ) Then

        If Len( callback ) = 0 Then
            array_map = arr
            Exit Function
        End If

        ReDim tmp_ar( uBound(arr) )
        For key = 0 to uBound( arr )
            If isObject( arr(key) ) Then
                execute("set tmp_ar(key) = " & callback & "(arr(key))")
            Else
                execute("tmp_ar(key) = " & callback & "(arr(key))")
            End If
        Next

        array_map = tmp_ar

    ElseIf isObject( arr ) Then

        If Len( callback ) = 0 Then
            set array_map = arr
            Exit Function
        End If

        Dim return_val

        set tmp_ar = Server.CreateObject("Scripting.Dictionary")
        For Each key In arr
            return_val = ""
            If isObject( arr.Item(key) ) Then
                execute("set return_val = " & callback & "(arr.Item(key))")
            Else
                execute("return_val = " & callback & "(arr.Item(key))")
            End If
            tmp_ar.Add key, return_val
        Next

        set array_map = tmp_ar

    End If

End Function
%>
