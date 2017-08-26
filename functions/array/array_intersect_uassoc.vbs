<%
'=======================================================================
'追加された添字の確認も含め、コールバック関数を用いて 配列の共通項を確認する
'=======================================================================
'【引数】
'  mAry1            = array     比較元となる最初の配列。
'  mAry2            = array     キーを比較する対象となる最初の配列。
'  key_compare_func = callback  比較に使用する、ユーザ定義のコールバック関数。
'【戻り値】
'  mAry1 の値のうち、 すべての引数に存在するもののみを返します。
'【処理】
'  ・全ての引数に現れる mAry1 の全ての値を含む配列を返します。 array_intersect() と異なり、 キーが比較に使用されることに注意してください。
'  ・比較は、ユーザが指定したコールバック関数を利用して行われます。 
'  ・この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'=======================================================================
Function array_intersect_uassoc(mAry1,mAry2,key_compare_func)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim key,k,found,compare

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            found = false

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            End If

            If found = true Then
                result.Add k, mAry1(key)
            End if
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            found = false

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            End If

            If found = true Then
                result.Add k, mAry1(key)
            End if
        Next
    End If

    set array_intersect_uassoc = result

End Function

%>
