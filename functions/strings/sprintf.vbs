<%
'=======================================================================
' フォーマットされた文字列を返す
'=======================================================================
'【引数】
'  format   = string フォーマット文字列
'  args     = mixed  数値や文字列
'【戻り値】
'  フォーマット文字列 format  に基づき生成された文字列を返します。
'【処理】
'  ・フォーマット文字列 format  に基づき生成された文字列を返します。
'=======================================================================
Function sprintf(format , args)

    If is_empty(args) Then args = ""
    Dim bobj : set bobj = Server.CreateObject("basp21")
    sprintf = bobj.sprintf(format,args)

End Function
%>
