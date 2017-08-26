'=======================================================================
' ファイルの内容を全て文字列に読み込む
'=======================================================================
'【引数】
'  filename  = string データを読み込みたいファイルの名前。
'【戻り値】
'  読み込んだデータを返します。失敗した場合は FALSE を返します。
'【処理】
'  ファイルの内容を文字列に読み込む
'=======================================================================
Function file_get_contents(filename)

    file_get_contents = false
    Dim fileObj : set fileObj = new File_System
    file_get_contents = fileObj.file_get_contents(filename)
    set fileObj = nothing

End Function
