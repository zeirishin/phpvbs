<%
'=======================================================================
' ファイルパスに関する情報を返す
'=======================================================================
'【引数】
'  path     = string    調べたいパス。
'  options  = string    どの要素を返すのかをオプションのパラメータ options  で指定します。これは PATHINFO_DIRNAME、 PATHINFO_BASENAME、 PATHINFO_EXTENSION および PATHINFO_FILENAME の組み合わせとなります。 デフォルトではすべての要素を返します。
'【戻り値】
'   以下の要素を含む連想配列を返します。 dirname (ディレクトリ名)、basename (ファイル名) そして、もし存在すれば extension (拡張子)。
'   options を使用すると、 すべての要素を選択しない限りこの関数の返り値は文字列となります。 
'【処理】
'  pathinfo() は、path  に関する情報を有する連想配列を返します。
'=======================================================================
Const PATHINFO_DIRNAME = 1
Const PATHINFO_BASENAME = 2
Const PATHINFO_EXTENSION = 4
Const PATHINFO_FILENAME = 3
Function pathinfo(ByVal path,ByVal options)

    Dim fileObj : set fileObj = new File_System
    set pathinfo = fileObj.pathinfo(path,options)
    set fileObj = nothing

End Function
%>
