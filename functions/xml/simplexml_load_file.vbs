<%
'/**
' *
' * php.asp :  Rapid Development Framework
' * Copyright 2005-2017, zeirishin
' *
' * @filesource
' * @copyright		Copyright 2005-2017, zeirishin
' * @link			http://blog.verygoodtown.com/
' * @package		php.vbs
' */

'=======================================================================
'XMLファイルをパースし、配列に代入する
'=======================================================================
'【引数】
'  filename     = String XML ファイルへのパス。
'  encode       = String XML ファイルの文字コード。
'【戻り値】
'  ドキュメント内のデータを返します。
'  エラー時には エラーメッセージを返します。
'【処理】
'  ・指定したファイルの中の整形式 XML 配列に変換します。
'=======================================================================
Function simplexml_load_file(filename,encode)

    Dim objDoc,result
    Dim ret,strXml,rtResult,xPE

    Set ret = Server.CreateObject("Scripting.Dictionary")

    If encode <> "Shift_JIS" and encode <> "sjis" Then _
        Set objDoc = Server.CreateObject("MSXML2.DOMDocument") _
    Else _
        Set objDoc = Server.CreateObject("MSXML.DOMDocument")

    objDoc.async = true

    If inStr(filename,"http://") = 1 Then

        Dim file
        set file = new File_System
        strXml = file.file_get_contents(filename)
        rtResult = objDoc.LoadXML(strXml)
    Else
        rtResult = objDoc.Load(filename)
    End If

    Set xPE = objDoc.parseerror
    If xPE.errorcode <> 0 then
        ret("error") = xPE
        set simplexml_load_file = ret
        Exit Function
    End If

    If rtResult = True Then
        call simplexml_parse(objDoc.childNodes, ret)
    Else
        ret("error") = "XMLを取得できません。"
    End If

    Set simplexml_load_file = ret
    Set objDoc = Nothing

End Function

Function simplexml_parse(objNode,ByRef ret)

    Dim obj,tmp_ob,tmp_ar()
    Dim intCounter,objData,att
    Dim counter,j

    If Not isObject(ret) Then Set ret = Server.CreateObject("Scripting.Dictionary")
    Set counter = Server.CreateObject("Scripting.Dictionary")

    ReDim tmp_ar(objNode.length-1)
    For j = 0 to (objNode.length-1)
        tmp_ar(j) = objNode(j).nodeName
    Next

    Set tmp_ob = array_count_values(tmp_ar)

    For Each obj In tmp_ob
        counter.Add obj, 0
    Next

    For Each obj In objNode

        objData = obj.nodeName

        If obj.nodeTypeString = "element" Then
            If obj.attributes.length > 0 Then
                Set ret(objData & "_attr") = Server.CreateObject("Scripting.Dictionary")

                For Each att IN obj.attributes
                    ret(objData & "_attr").Add _
                        preg_replace("/(.+)="".*""/","$1",att.xml), att.value
                Next
            End If
        End If

        If obj.hasChildNodes Then

            If obj.childNodes.length = 1 and (obj.childNodes(0).nodeName = "#text" or _
               obj.childNodes(0).nodeName = "#cdata-section") Then
                If Not ret.Exists( objData ) Then ret.Add objData , obj.text
            Else

                If Not isObject(ret(objData)) Then _
                    Set ret(objData) = Server.CreateObject("Scripting.Dictionary")
                If tmp_ob(objData) = 1 Then
                    call simplexml_parse(obj.childNodes, ret(objData))
                Else

                    If Not isObject(ret(objData)(counter(objData))) Then _
                        Set ret(objData)(counter(obj.nodeName)) = _
                            Server.CreateObject("Scripting.Dictionary")

                    call simplexml_parse(obj.childNodes, ret(objData)(counter(objData)))
                    counter(objData) = counter(objData) +1
                End If
           End If

        Else
            If Not ret.Exists( objData ) Then ret.Add objData , obj.text
        End If

    Next
End Function

%>
