<%
'=======================================================================
'キーを基準にし、コールバック関数を用いて 配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元となる最初の配列。
'  mAry2    = array  キーを比較する対象となる最初の配列。
'  key_compare_func = callback  比較に使用する、ユーザ定義のコールバック関数。
'【戻り値】
'  mAry1 の値のうち、 すべての引数に存在するキーのものを含む連想配列を返します。
'【処理】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'=======================================================================
Function array_intersect_ukey(mAry1,mAry2,key_compare_func)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim key,k,compare

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            End If

        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            End If

        Next
    End If

    set array_intersect_ukey = result

End Function
%>
