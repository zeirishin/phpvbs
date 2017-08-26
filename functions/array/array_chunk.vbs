<%
'=======================================================================
'配列を分割する
'=======================================================================
'【引数】
'  mAry     = Array         処理を行う配列。
'  size     = int           各部分のサイズ。
'  preserve_keys = bool     TRUE の場合はキーをそのまま保持します。 デフォルトは FALSE で、各部分のキーをあらためて数字で振りなおします。
'【戻り値】
'  数値添字の多次元配列を返します。添え字はゼロから始まり、 各次元の要素数が size  となります。
'【処理】
'  ・配列を、size  個ずつの要素に分割します。 
'  ・最後の部分の要素数は size  より小さくなることもあります。
'=======================================================================
Function array_chunk(mAry,size)

    If not isNumeric(size) Then Exit Function
    If size < 1 then Exit Function

    Dim x,i,c : x = 0 : c = -1
    Dim l : l = uBound(mAry)
    Dim n : n = int(l / size)
    ReDim tmpAry(n)

    For i = 0 to l
        x = i Mod size

        If x >= 1 Then
            If isObject(mAry(i)) Then
                set tmpAry(c)(x) = mAry(i)
            Else
                tmpAry(c)(x) = mAry(i)
            End If
        Else
            c = c +1
            If n <> c Then
                toReDim tmpAry(c),size -1
            Else
                toReDim tmpAry(c),l -i
            End If

            If isObject(mAry(i)) Then
                set tmpAry(c)(0) = mAry(i)
            Else
                tmpAry(c)(0) = mAry(i)
            End If

        End If
    Next

    array_chunk = tmpAry

End Function

'=======================================================================
'配列を作成する
'=======================================================================
'【引数】
'  mAry     = mixed  配列
'  mVal     = mixed  追加する要素の数
'【戻り値】
'  値を返しません。
'【処理】
'  ・mAryを配列にします。
'=======================================================================
Sub toReDim(ByRef mAry, ByVal mVal)

    If isArray(mAry) Then
        ReDim Preserve mAry(mVal)
    Else
        ReDim mAry(mVal)
    End If

End Sub

%>
