Function base64_decode(data)

    Dim obj
    set obj=server.createobject("basp21")
    base64_decode = obj.Base64(data,1)
    set obj = nothing

    'BASP21を使用しない場合
'    Dim ST, DM, EL
'    Dim bin
' 
'    Set DM = CreateObject("Microsoft.XMLDOM")
'    Set EL = DM.createElement("tmp")
'    EL.DataType = "bin.base64"
'    EL.Text = Base64Text
'    bin = EL.NodeTypedValue
' 
'    Set ST = CreateObject("ADODB.Stream")
'    ST.Open
'    ST.Charset = "Shift-JIS"
'    ST.Type = adTypeBinary
'    ST.Write bin
'    ST.Position = 0
'    ST.Type = adTypeText
'    base64_decode = ST.ReadText
'    ST.Close

End Function
