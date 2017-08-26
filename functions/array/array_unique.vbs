<%
'=======================================================================
'配列から重複した値を削除する
'=======================================================================
'【引数】
'  mAry     = Array 入力の配列。
'【戻り値】
'  ・処理済の配列を返します。
'【処理】
'  ・値に重複のない新規配列を返します。
'=======================================================================
function array_unique(arr)

    Dim key,key_c
    Dim found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(arr) Then
        For key = 0 to uBound(arr)
            found = array_search(arr(key),output,false)
            If found = false and varType(found) = 11 Then
                output.Add key, arr(key)
            End If
        Next
    ElseIf isObject(arr) Then
        For Each key In arr
            found = array_search(arr(key),output,false)
            If found = false and varType(found) = 11 Then
                    output.Add key, arr(key)
            End If
        next
    End If

    set array_unique = output
end function
%>
