'=======================================================================
' ファイル全体を読み込んで配列に格納する
'=======================================================================
'【引数】
'  filename  = string ファイルへのパス。
'  flags     = int    オプションのパラメータ flags  は、以下の定数のうちのひとつ、あるいは複数の組み合わせとなります。
'【戻り値】
'  ファイルを配列に入れて返します。 配列の各要素はファイルの各行に対応します。
'  改行記号はついたままとなります。 失敗すると file() は FALSE を返します。
'【処理】
'  ファイル全体を配列に読み込みます。
'=======================================================================
Const FILE_IGNORE_NEW_LINES = 2
Const FILE_SKIP_EMPTY_LINES = 4
Function file(filename,flags)

    Dim req,tmp,key
    req = file_get_contents(filename)

    If flags = FILE_SKIP_EMPTY_LINES Then
        var_dump req
        tmp = preg_replace("/^" & vbCrLf & "/is","",req,"","")
        tmp = Split(tmp,vbCrLf)
    Else
        tmp = Split(req,vbCrLf)

        If flags = FILE_IGNORE_NEW_LINES Then
            For key = 0 to uBound(tmp)
                tmp(key) = tmp(key) & vbCrLf
            Next
        End If
    End If

    file = tmp

End Function
