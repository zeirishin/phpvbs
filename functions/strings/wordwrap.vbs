<%
'=======================================================================
' 文字列分割文字を使用して指定した文字数数に文字列を分割する
'=======================================================================
'【引数】
'  str      = string    入力文字列。
'  width    = int       カラムの幅。デフォルトは 75。
'  break    = string    オプションのパラメータ break  を用いて行を分割します。 デフォルトは 'vbCrLf' です。
'  cut      = bool      cut  を TRUE に設定すると、 文字列は常に指定した幅でラップされます。このため、 指定した幅よりも長い単語がある場合には、分割されます (2 番目の例を参照ください)。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・ 指定した文字数で、指定した文字を用いて文字列を分割します。
'=======================================================================
Function wordwrap( str, int_width, str_break, cut )

    If len(int_width) = 0 Then int_width = 75
    If len(str_break) = 0 Then str_break = vbCrLf

    Dim m : m = int_width
    Dim b : b = str_break
    Dim c : c = cut

    Dim i,j, l, s, r
    Dim matches

    If m < 1 Then
        wordwrap = str
        Exit Function
    End If

    r = split(str,vbCrLf)
    l = uBound(r)
    i = -1

    Do While i < l
        i = i +1

        s = r(i)
        r(i) = ""

        Do While len(s) > m
            j = [==](c, 2)
            If is_empty(j) Then
                If preg_match("/¥S*(¥s)?$/",Left(s,m+1),matches,"","") Then
                    If len( trim(matches(0)) ) = 0 Then
                        j = m
                    Else
                        j = len( Left(s,m+1) ) - len(matches(0))
                    End If
                End If

                If is_empty(j) Then
                    j = [?]([==](c, true),m,false)
                End If

                If is_empty(j) Then
                    call preg_match("/^¥S*/",Mid(s,m),matches,"","")
                    j = len( Left(s,m) ) + len(matches(0))
                End If
            End If

            r(i) = r(i) & Left(s, j)
            s = Mid(s,j+1)
            r(i) = r(i) & [?](len(s), b , "")
        Loop

        r(i) = r(i) & s

    Loop

    wordwrap = join(r,vbCrLf)

End Function
%>
