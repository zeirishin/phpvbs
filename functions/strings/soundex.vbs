<%
'=======================================================================
' 文字列の soundex キーを計算する
'=======================================================================
'【引数】
'  format   = string 入力文字列。
'【戻り値】
'  メタ文字をクォートした文字列を返します。
'【処理】
'  ・ str  の soundex キーを計算します。
'  ・ soundex キーには、似たような発音の単語に関して同じ soundex キーが生成されるという特性があります。 このため、発音は知っているが、スペルがわからない場合に、 データベースを検索することを容易にすることができます。
'  ・ soundex 関数は、ある文字から始まる 4 文字の文字列を返します。
'  ・ この soundex 関数についての説明は、Donald Knuth の "The Art Of Computer Programming, vol. 3: Sorting And Searching", Addison-Wesley (1973), pp. 391-392 にあります。 
'=======================================================================
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
    s = preg_replace("/[^A-Z]/","",s,"","")
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

    soundex = join(r,"") & join( newArray, "0" )

End Function
%>
