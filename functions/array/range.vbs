<%
'=======================================================================
'ある範囲の整数を有する配列を作成する
'=======================================================================
'【引数】
'  low  = mixed 下限値。
'  high = mixed 上限値。
'  step = mixed step  が指定されている場合、それは 要素毎の増加数となります。step  は正の数でなければなりません。デフォルトは 1 です。
'【戻り値】
'  low  から high  までの整数の配列を返します。 low > high の場合、順番は high から low となります。
'【処理】
'  ・ある範囲の整数を有する配列を作成します。
'=======================================================================
Function range(low,high,step)

    Dim matrix
    Dim inival, endval, plus
    Dim walker : If len(step) > 0 Then walker = step Else walker = 1
    Dim chars : chars = false

    If isNumeric(low) and isNumeric(high) Then
        inival = low
        endval = high
    ElseIf Not isNumeric(low) and Not isNumeric(high) Then
        chars  = true
        inival = Asc(low)
        endval = Asc(high)
    Else
        inival = [?](isNumeric(low),low,0)
        endval = [?](isNumeric(high),high,0)
    End If

    plus = true
    If inival > endval Then plus = false

    If plus Then
        Do While inival <= endval
            [] matrix, [?](chars,Chr(inival),inival)
            inival = inival + walker
        Loop
    Else
        Do While inival >= endval
            [] matrix, [?](chars,Chr(inival),inival)
            inival = inival - walker
        Loop
    End If

    range = matrix

End Function
%>
