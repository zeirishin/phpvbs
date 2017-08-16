Function soundex(str)

    Dim i,j, l, r, p, m, s
    p = [?](Not isNumeric(p) , 4 , [?](p > 10 , 10 , [?](p < 4 , 4 , p) ) )

    set m = Server.CreateObject("Scripting.Dictionary")
    m.Add "BFPV", 1
    m.Add "CGJKQSXZ", 2
    m.add "DT", 3
    m.add "L", 4
    m.add "MN", 5
    m.add "R", 6

    s = Ucase( str )
    s = preg_replace("/[^A-Z]/",",s,",")
    s = str_split(s,1)
    r = array( array_shift(s) )

    For i = 0 to uBound(s)
        For Each j In m
            if inStr(j,s(i)) and r( uBound(r) ) <> m.Item(j) Then
                array_push r,m(j)
                Exit For
            End If
        Next
    Next

    If uBound(r) + 1 > p Then
        r = array_slice(r,0,p-1)
    End If

    Dim newArray()
    ReDim newArray(p - (uBound(r)+1))

    soundex = join(r,") & join( newArray, "0" )

End Function
