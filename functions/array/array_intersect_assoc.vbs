<%
'=======================================================================
'追加された添字の確認も含めて配列の共通項を確認する
'=======================================================================
'【引数】
'  mAry1    = array  値を調べるもととなる配列。
'  mAry2    = array  値を比較する対象となる配列。
'【戻り値】
'  mAry1 の値のうち、すべての引数に存在するものを含む連想配列を返します。
'【処理】
'  ・全ての引数に現れる mAry1 の全ての値を含む配列を返します。 
'  ・array_intersect() と異なり、 キーが比較に使用されることに注意してください。
'=======================================================================
Function array_intersect_assoc(mAry1,mAry2)

    Dim intersect : set intersect = Server.CreateObject("Scripting.Dictionary")
    Dim key,counter

    If isArray(mAry2) Then
        counter = uBound(mAry2)
    Else
        counter = null
    End If

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            intersect.Add key, mAry1(key)
            If counter >= key or isNull(counter) Then
                If mAry2(key) <> mAry1(key) Then
                    intersect.Remove key
                End If
            End If
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            intersect.Add key, mAry1(key)
            If isNull(counter) or (isNumeric(key) and counter >= key) Then
                If mAry2(key) <> mAry1(key) Then
                    intersect.Remove key
                End If
            Else
               intersect.Remove key
            End If
        Next
    End If

    set array_intersect_assoc = intersect

End Function
%>
