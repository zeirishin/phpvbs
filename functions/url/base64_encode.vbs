<%
'=======================================================================
'  MIME base64 方式でデータをエンコードする
'=======================================================================
'【引数】
'  data = mixed  エンコードするデータ。
'【戻り値】
'  エンコードされたデータを文字列で返します。
'【処理】
'  ・指定した data  を base64 でエンコードします。
'=======================================================================
Function base64_encode(data)

    Dim obj
    set obj=server.createobject("basp21")
    base64_encode = obj.Base64(data,0)
    set obj = nothing

    'basp21を使用しない場合
'    Dim ST, DM, EL, bin
'  
'    Set ST = CreateObject("ADODB.Stream")
'    ST.Type = adTypeText
'    ST.Charset = "Shift-JIS"
'    ST.Open
'    ST.WriteText PlainText
'    ST.Position = 0
'    ST.Type = adTypeBinary
'    bin = ST.Read
'    ST.Close
' 
'    Set DM = CreateObject("Microsoft.XMLDOM")
'    Set EL = DM.CreateElement("tmp")
'    EL.DataType = "bin.base64"
'    EL.NodeTypedValue = bin
'    base64_encode = EL.Text

End Function
%>
