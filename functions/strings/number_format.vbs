Function number_format( number, decimals, dec_point, thousands_sep )

    Dim n,c,d,t,i,s
    n = number
    c = [?]( isNumeric(decimals),decimals,2 )
    c = abs( c )

    d = [?]( len(dec_point) = 0,",",dec_point)
    t = [?]( len(thousands_sep) = 0, ".", left(thousands_sep,1) )

    n = FormatNumber (n, c,true,false,true)
    n = Replace(n,",",d)
    n = Replace(n,".",t)

    number_format = n

End Function
