'=======================================================================
' 二つの文字列のレーベンシュタイン距離を計算する
'=======================================================================
'【引数】
'  str1 = string    レーベンシュタイン距離を計算する文字列のひとつ。
'  str2 = string    レーベンシュタイン距離を計算する文字列のひとつ。
'【戻り値】
'  この関数は、引数で指定した二つの文字列のレーベンシュタイン距離を返します。
'  引数文字列の一つが 255 文字の制限より長い場合に -1 を返します。
'【処理】
'  ・レーベンシュタイン距離は、str1  を str2  に変換するために置換、挿入、削除 しなければならない最小の文字数として定義されます。
'  ・アルゴリズムの複雑さは、 O(m*n) です。
'  ・ここで、n および m はそれぞれ str1  および str2  の長さです (O(max(n,m)**3) となる similar_text() よりは良いですが、 まだかなりの計算量です)。
'  ・上記の最も簡単な形式では、この関数はパラメータとして引数を二つだけとり、 str1  から str2  に変換する際に必要な 挿入、置換、削除演算の数のみを計算します。
'=======================================================================
Function levenshtein( str1, str2 )

    Dim s,l,t,i,j,m,n,u
    Dim a,tmp

    s = str_split(str1,1)
    u = str_split(str2,1)

    If isArray(s) Then l = count(s,"") Else l = 0
    If isArray(u) Then t = count(u,"") Else t = 0

    If is_empty(l) or is_empty(t) Then
        If [>](l , t) Then levenshtein = l
        If [>](t , l) Then levenshtein = t
        If isEmpty(levenshtein) Then levenshtein = 0
        Exit Function
    End If

    ReDim a(l)
    For i = 0 to l
        toReDim a(i),t
    Next

    For i = l to 0 Step -1
       a(i)(0) = i
    Next

    For i = t to 0 Step -1
       a(0)(i) = i
    Next

    i = 0
    m = l

    Do While(i < m)

        j = 0
        n = t

        Do While(j < n)
            tmp = a(i)(j + 1) + 1
            If tmp > a(i+1)(j) + 1 Then tmp = a(i+1)(j) + 1
            If tmp > a(i)(j) + intval([!=](s(i) ,u(j))) Then tmp = a(i)(j) + intval([!=](s(i) ,u(j)))
            a(i+1)(j+1) = tmp

            j = j + 1
        Loop

        i = i +1
    Loop

    levenshtein = a(l)(t)

End Function
