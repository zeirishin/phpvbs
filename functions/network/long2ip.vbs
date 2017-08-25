<%
'=======================================================================
' (IPv4) インターネットアドレスをインターネット標準ドット表記に変換する
'=======================================================================
'【引数】
'  proper_address = string   正しい形式のアドレス。
'【戻り値】
'  インターネットの IP アドレスを表す文字列を返します。
'【処理】
'  ・関数long2ip() は、適切なアドレス表現からドット表記 (例:aaa.bbb.ccc.ddd)のインターネットアドレスを生成します。
'=======================================================================
Function long2ip( proper_address )

    long2ip = false

    If isNumeric(proper_address) Then
        If proper_address >= 0 and proper_address <= 4294967295 Then

            long2ip = ( proper_address / pow(256,3) ) & "." & _
                      ( (proper_address Mod pow(256,3)) / pow(256,2) ) & "." & _
                      ( ((proper_address Mod pow(256,3)) / pow(256,2)) / pow(256,1) ) & "." & _
                      ( (((proper_address Mod pow(256,3)) / pow(256,2)) / pow(256,1)) / pow(256,0) )

        End If
    End If
End Function
%>
