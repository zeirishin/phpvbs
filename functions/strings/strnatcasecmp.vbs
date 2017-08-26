<%
'=======================================================================
' "自然順"アルゴリズムにより大文字小文字を区別しない文字列比較を行う
'=======================================================================
'【引数】
'  str1 =   string  最初の文字列。
'  str2 =   string  次の文字列。
'【戻り値】
'  他の文字列比較関数と同様に、この関数は、 str1  が str2  より小さいに場合に < 0、str1  が str2  より大きい場合に > 0 、等しい場合に 0 を返します。
'【処理】
'  ・この関数は、人間が行うような手法でアルファベットまたは数字の 文字列の順序を比較するアルゴリズムを実装します。この手法は、"自然順" と言われます。
'  ・この関数の動作は、 strnatcmp() に似ていますが、 比較が大文字小文字を区別しない違いがあります。
'=======================================================================
Function strnatcasecmp( str1, str2 )

    Dim array1,array2
    array1 = strnatcmp_helper(str1)
    array2 = strnatcmp_helper(str2)

    Dim intlen,text,result,r
    intlen = uBound(array1)
    text   = true

    result = -1
    r      = 0

    if intlen > uBound(array2) Then
        intlen = uBound(array2)
        result = 1
    End If

    Dim i
    strnatcasecmp = false
    For i = 0 to intlen
        If not isNumeric( array1(i) ) Then
            If Not isNumeric( array2(i) ) Then
                text = true

                r = strcasecmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcasecmp = r
                End If

            ElseIf text Then
                strnatcasecmp = 1
            Else
                strnatcasecmp = 1
            End If

        ElseIf not isNumeric( array2(i) ) Then
            If text Then
                strnatcasecmp = -1
            Else
                strnatcasecmp = 1
            End If
        Else
            If text Then
                r = array1(i) - array2(i)
                If r <> 0 Then
                    strnatcasecmp = r
                End If
            Else
                r = strcasecmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcasecmp = r
                End If
            End If

            text = false
        End If

        if [!==](strnatcasecmp,false) Then Exit Function
    Next

    strnatcasecmp = result

End Function
%>
