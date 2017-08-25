<%
Const PREG_GREP_INVERT    = 1
'=======================================================================
'パターンにマッチする配列の要素を返す
'=======================================================================
'【引数】
'  pattern  = string    検索するパターンを表す文字列。
'  input    = array     入力の配列。
'  flags    = array     PREG_GREP_INVERT  を設定すると、この関数は 与えた pattern  にマッチ しない  要素を返します。
'【戻り値】
'  input  配列のキーを使用した配列を返します。
'【処理】
'  ・ input  配列の要素のうち、 指定した pattern  にマッチするものを要素とする配列を返します。 
'=======================================================================
Function preg_grep(pattern, input, flags)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")

    If not isArray(input) and not isObject(input) Then
        set preg_grep =  obj
        Exit Function
    End If

    Dim key
    If isArray(input) Then
        For key = 0 to uBound(input)
            If flags = PREG_GREP_INVERT Then
                If Not preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            Else
                If preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            End If
        Next
    ElseIf isObject(input) Then
        For Each key In input
            If flags = PREG_GREP_INVERT Then
                If Not preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            Else
                If preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            End If
        Next
    End If

    set preg_grep =  obj

End Function
%>
