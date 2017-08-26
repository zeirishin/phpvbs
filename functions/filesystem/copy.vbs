<%
'=======================================================================
' ファイルをコピーする
'=======================================================================
'【引数】
'  source  = string   コピー元ファイルへのパス。
'  dest    = string   コピー先のパス。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・ ファイル source  を dest  にコピーします。
'  ・ ファイルを移動したいならは、rename() 関数を使用してください。 
'=======================================================================
Function copy(source, dest)

    copy = false
    Dim fileObj : set fileObj = new File_System
    copy = fileObj.copy(source,dest)
    set fileObj = nothing

End Function
%>
