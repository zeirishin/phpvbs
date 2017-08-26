<%
'=======================================================================
' 文字列の一部分を返す
'=======================================================================
'【引数】
'  str          = string    入力文字列。
'  start        = string     start  が正の場合、返される文字列は、 string  の 0 から数えて start 番目から始まる文字列となります。 例えば、文字列'abcdef'において位置 0にある文字は、'a'であり、 位置2には'c'があります。start が負の場合、返される文字列は、 string の後ろから数えて start 番目から始まる文字列となります。 
'  intLength    = string    入力文字列。
'【戻り値】
'  文字列の一部を返します。
'【処理】
'  ・文字列 str  の、start  で指定された位置から length  バイト分の文字列を返します。
'=======================================================================
Function substr(ByVal str,ByVal start,ByVal intLength)

	intStart = start
    If len(intStart) = 0 Then intStart = 0
    if intStart = 0 And Len(intLength) < 1 Then
    	substr = str
    	Exit Function
    End If
    
    If len(intLength) = 0 Then intLength = abs(start)
    
    If intStart < 0 Then
        intStart = len(str)+1 + intStart
    Else
    	intStart = intStart + 1
    End If

    Dim tmp
    If intStart <> 0 Then
	    tmp = Mid(str,intStart)
	Else 
		tmp = str
	End If

    Dim intLen
    intLen = len(tmp)


    If len(intLength) > 0 Then

        If intLen >= abs(intLength) Then
            If intLength > 0 Then

        		tmp = Left(tmp,intLength)

            Else
                tmp = Left(tmp,len(tmp) + intLength)
            End If
        Else
            tmp = False
        End If
    End If

    substr = tmp

End Function
%>
