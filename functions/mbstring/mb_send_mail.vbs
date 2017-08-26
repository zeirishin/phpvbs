'=======================================================================
' メールを送信する
'=======================================================================
'【引数】
'  to       = string    文字列
'  subject  = string    件名
'  message  = string    本文
'  file     = string    添付ファイル
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・email を送信します。
'=======================================================================
function mb_send_mail( strto, subject, message, from, strFile )

    Dim basp
    Dim msg

    if inStr(subject,vbCrLf) then
        Response.Write "題名に改行を含めることはできません。"
        Response.End
    end if

	Set basp = Server.CreateObject("basp21")
	msg = basp.SendMail(SMTP_SERVER, strto, from, subject, message, strFile)

    mb_send_mail = msg

end function
