Class File_System

    Private fso
    Private ts

    '
    'Initialize Class
    ' 
    '@access private
    '
    Private Sub Class_Initialize()
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
    End Sub

    '
    'Terminate Class
    ' 
    '@access private
    '
    Private Sub Class_Terminate()
        Set fso = Nothing
    End Sub

    '=======================================================================
    ' パス中のファイル名の部分を返す
    '=======================================================================
    '【引数】
    '  path      = string   パス。
    '  suffix    = string   ファイル名が、 suffix  で終了する場合、 この部分もカットされます。
    '【戻り値】
    '  指定した path  のベース名を返します。
    '【処理】
    '  ・この関数は、ファイルへのパスを有する文字列を引数とし、 ファイルのベース名を返します。
    '=======================================================================
    Function basename(path, suffix)

        Dim b
        b = preg_replace("/^.*[¥/¥¥]/g","",path,null,null)

        If len(suffix) > 0 Then
            If Right(b,len(suffix)) = suffix Then
                b = Left(b,len(b) - len(suffix))
            End If
        End If

        basename = b

    End Function

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
    Public Function copy(source,dest)
        fso.CopyFile source,dest
    End Function

    '=======================================================================
    ' オープンされたファイルポインタをクローズする
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  
    '【処理】
    '  ファイルをクローズします。
    '=======================================================================
    Public function fclose
        ts.close
        Set ts = Nothing
    end function

    '=======================================================================
    ' ファイルポインタがファイル終端に達しているかどうか調べる
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  ファイルポインタが EOF に達しているかまたはエラー (ソケットタイムアウトを含みます) の場合に TRUE 、 その他の場合に FALSE を返します。
    '【処理】
    '  ファイルポインタがファイル終端に達しているかどうかを調べます。
    '=======================================================================
    Public function feof
        feof = ts.AtEndofStream
    end function

    '=======================================================================
    ' ファイルポインタから1文字取り出す
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  ファイルポインタから 1 文字読み出し、 その文字からなる文字列を返します。EOF の場合に FALSE を返します。
    '【処理】
    '  指定したファイルポインタから 1 文字読み出します。
    '=======================================================================
    Public function fgetc
        If ts.AtEndofStream Then
            fgetc = false
        Else
            fgetc = ts.Read(1)
        End If
    end function

    '=======================================================================
    ' ファイルポインタから行を取得し、CSVフィールドを処理する
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  読み込んだフィールドの内容を含む数値添字配列を返します。
    '【処理】
    '  fgets() に動作は似ていますが、 fgetcsv() は行を CSV  フォーマットのフィールドとして読込み処理を行い、 読み込んだフィールドを含む配列を返すという違いがあります。
    '=======================================================================
    Public Function fgetcsv(delim)

        Dim tmp,d
        If len(delim) > 0 Then d = delim Else d = ","
        tmp = ts.ReadLine
        fgetcsv = fgetcsv_helper(tmp,d)

    End Function

    '************************************
    Public Function fgetcsv_helper(str,d)

        Dim matchAll,key
        Dim data,field,record : field = 0 : record = 0
        ReDim data(0)

        If preg_match_all(_
        "/" & d & "|" & vbCr & "?" & vbLf & "|[^" & d & """" & vbCrLf & "][^" & d & "" & vbCrLf & "]*|""(?:[^""]|"""")*""/",_
        str, matchAll,PREG_PATTERN_ORDER,"") Then
            For Each key In matchAll(0)
                Select Case key
                Case d
                    field = field + 1
                Case vbCrLf
                    [] data , ""
                    record = record +1
                Case Else
                    If left(key,1) = """" Then
                        key = Replace(substr(key,2,-1),"""""","""")
                    End if
                    [] data(record), key
                End Select
            Next
        End If

        fgetcsv_helper = data

    End Function

    '=======================================================================
    ' ファイルポインタから 1 行取得する
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  ファイルポインタから 1 行取得する
    '【処理】
    '  ファイルポインタから 1 行取得します。
    '=======================================================================
    Public function fgets
        fgets = ts.ReadLine
    end function

    '=======================================================================
    ' ファイルポインタから 1 行取り出し、HTML タグを取り除く
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  HTML や PHP コードを取り除いた文字列を返します。
    '【処理】
    '  fgets() と同じですが、 fgetss() は読み込んだテキストから HTML および PHP のタグを取り除こうとすることが異なります。
    '=======================================================================
    Public function fgetss
        fgets = strip_tags(ts.ReadLine)
    end function

    '=======================================================================
    ' ファイルまたはディレクトリが存在するかどうか調べる
    '=======================================================================
    '【引数】
    '  path      = string   ファイルあるいはディレクトリへのパス。
    '【戻り値】
    '  ファイルあるいはディレクトリが存在するかどうかを調べます。
    '【処理】
    '  ファイルあるいはディレクトリが存在するかどうかを調べます。
    '=======================================================================
    Public Function file_exists(ByVal filename)

        file_exists = false
        filename = fileMapPath(filename)
        If fso.FileExists(filename) Then file_exists = true
        If fso.FolderExists(filename) Then file_exists = true

    End Function

    '=======================================================================
    'ファイルの内容を全て文字列に読み込む
    '=======================================================================
    '【引数】
    '  filename  = string データを読み込みたいファイルの名前。
    '【戻り値】
    '  読み込んだデータを返します。失敗した場合は FALSE を返します。
    '【処理】
    '  ファイルの内容を文字列に読み込む
    '=======================================================================
    Public function file_get_contents(filename)

        Dim ts
        Dim contents

        if left(filename,7) <> "http://" and file_exists( filename ) then
            Set TS = fso.OpenTextFile( fileMapPath(filename),1)

            '空のファイルの場合、エラーになってしまう
            If TS.AtEndOfStream <> True Then
               contents = TS.ReadAll
            End If

            file_get_contents = contents
            Exit Function
        end if

        if left(filename,7) <> "http://" then
            file_get_contents = false
            Exit Function
        end if

        Dim objWinHttp
        'Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
        'Set objWinHttp = Server.CreateObject("MSXML2.XMLHTTP")
        Set objWinHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
'        on error resume next

        objWinHttp.Open "GET", filename, false
        objWinHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        objWinHttp.Send
        'Response.Write objWinHttp.Status & " " & objWinHttp.StatusText

        file_get_contents = objWinHttp.ResponseText

        Set objWinHttp = Nothing

        '有効にしたエラー処理を無効にする
'        on error goto 0

    end function

    '=======================================================================
    ' ファイルの最終アクセス時刻を取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルの最終アクセス時刻を返し、エラーの場合は FALSE を返します。
    '【処理】
    '  指定したファイルの最終アクセス時刻を取得します。
    '=======================================================================
    Public Function fileatime(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        fileatime = f.DateLastAccessed

    End Function

    '=======================================================================
    ' ファイルの更新時刻を取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルの最終更新時刻を返し、エラーの場合は FALSE  を返します。
    '【処理】
    '  この関数は、ファイルのブロックデータが書き込まれた時間を返します。 これは、ファイルの内容が変更された際の時間です。
    '=======================================================================
    Public Function filemtime(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filemtime = f.DateLastModified

    End Function

    '*************************************
    Private Function fileMapPath(filename)

        Dim tmp
        tmp = Left(filename,3)
        tmp = Lcase(tmp)
        If tmp <> "d:¥" and tmp <> "c:¥" and left(filename,7) <> "http://" then
                fileMapPath = Server.MapPath(filename)
        Else
            fileMapPath = filename
        End If

    End Function

    '=======================================================================
    ' ファイルのサイズを取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルのサイズを返し、エラーの場合は FALSE を返します (また E_WARNING レベルのエラーを発生させます) 。
    '【処理】
    '  指定したファイルのサイズを取得します。
    '=======================================================================
    Public Function filesize(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filesize = f.Size

    End Function

    '=======================================================================
    ' ファイルタイプを取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルのタイプを返します。
    '【処理】
    '  指定したファイルのタイプを返します。
    '=======================================================================
    Public Function filetype(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filetype = f.Type

    End Function

    '=======================================================================
    ' ファイルまたは URL をオープンする
    '=======================================================================
    '【引数】
    '  filename  =  string データを読み込みたいファイルの名前。
    '  mode      =  string ストリームに要するアクセス形式を指定します
    '【戻り値】
    '  成功した場合にファイルポインタリソースを返します。
    '【処理】
    '  fopen() は、filename  で指定されたリソースをストリームに結び付けます。
    '=======================================================================
    Public function fopen(filename, mode)

        Dim filePath
        filePath = fileMapPath(filename)

        If left(filePath,len("http://")) = "http://" Then
            fopen = file_get_contents(filePath)
            Exit Function
        End If

        Select Case mode
        Case "r"
            '読み込みのみでオープンします。
            Set ts = fso.OpenTextFile(filePath,1,false)
        Case "w"
            '書き込みでオープンします。
            Set ts = fso.OpenTextFile(filePath,2,true)
        Case "a"
            '追記でオープンします。
            Set ts = fso.OpenTextFile(filePath,8,true)
        Case "x"
            '書き込みでオープンします。ファイルが存在した場合はfalseを返します。
            If is_file(filePath) Then
                fopen = false
            Else
                Set ts = fso.OpenTextFile(filePath,2,true)
            End If
        Case Else
            'empty
            ts = false
        End Select

    end function

    '=======================================================================
    ' 行を CSV 形式にフォーマットし、ファイルポインタに書き込む
    '=======================================================================
    '【引数】
    '  fields  =  string    値の配列。
    '  delimiter  =  string オプションの delimiter  はフィールド区切り文字 (一文字だけ) を指定します。デフォルトはカンマ (,) です。
    '  enclosure  =  string オプションの enclosure  はフィールドを囲む文字 (一文字だけ) を指定します。デフォルトは二重引用符 (") です。
    '【戻り値】
    '  書き込んだ文字列の長さを返します。失敗した場合は FALSE を返します。
    '【処理】
    '  fputcsv() は、行（fields  配列として渡されたもの）を CSV としてフォーマットし、それを ファイルに書き込みます (いちばん最後に改行を追加します)。
    '=======================================================================
    Public function fputcsv(fields,delimiter,enclosure)

        fputcsv = false
        If len(delimiter) = 0 Then delimiter = ","
        If len(enclosure) = 0 Then enclosure = """"

        Dim key,replaced
        For key = 0 to uBound(fields)
            replaced = false
            If inStr(fields(key),delimiter) or inStr(fields(key),enclosure) or inStr(fields(key),vbCrLf) Then
                fields(key) = Replace(fields(key),enclosure,enclosure & enclosure)
                fields(key) = enclosure & fields(key) & enclosure
            End If
        Next

        Dim str : str = join(fields,delimiter)
        ts.WriteLine str
        fputcsv = len(str)
    end function

    '=======================================================================
    ' fwrite() のエイリアス
    '=======================================================================
    '【引数】
    '  str  =  string 書き込む文字列。
    '【説明】
    '  この関数は次の関数のエイリアスです。 fwrite().
    '=======================================================================
    Public function fputs(str)
        fputs = fwrite(str)
    end function

    '=======================================================================
    ' バイナリセーフなファイル書き込み処理
    '=======================================================================
    '【引数】
    '  str     =  string   書き込む文字列。
    '【戻り値】
    '  
    '【処理】
    '  string の内容を ファイル・ストリームに書き込みます。
    '=======================================================================
    Public function fwrite(str)
        ts.WriteLine str
    end function

    '=======================================================================
    ' ファイルがディレクトリかどうかを調べる
    '=======================================================================
    '【引数】
    '  filename = string    ファイルへのパス。filename  が相対パスの場合は、現在の作業ディレクトリからの相対パスとして処理します。
    '【戻り値】
    '  ファイルがが存在して、かつそれがディレクトリであれば TRUE、それ以外の場合は FALSE を返します。
    '【処理】
    '  指定したファイルがディレクトリかどうかを調べます。
    '=======================================================================
    Public Function is_dir(filename)

        Dim fn
        is_dir = false
        fn = fileMapPath(filename)

        If fso.FolderExists(fn) Then is_dir = true

    End Function

    '=======================================================================
    ' 通常ファイルかどうかを調べる
    '=======================================================================
    '【引数】
    '  filename = string    ファイルへのパス。
    '【戻り値】
    '  ファイルが存在し、かつそれが通常のファイルである場合に TRUE、 それ以外の場合に FALSE を返します。
    '【処理】
    '  指定したファイルが通常のファイルかどうかを調べます。
    '=======================================================================
    Public Function is_file(ByVal filename)

        is_file = false
        filename = fileMapPath(filename)
        If fso.FileExists(filename) Then is_file = true

    End Function

    '=======================================================================
    ' ディレクトリを作る
    '=======================================================================
    '【引数】
    '  pathname = string    ディレクトリのパス。
    '【戻り値】
    '  成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  指定したディレクトリを作成します。
    '=======================================================================
    Public Function mkdir(ByVal pathname)

        mkdir = false
        pathname = fileMapPath(pathname)
        If not fso.FolderExists(pathname) Then
            mkdir = fso.CreateFolder(pathname)
        End If

    End Function

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
    Public Function pathinfo(path,options)

        Dim obj : set obj = Server.CreateObject("Scripting.Dictionary")
        Dim tmp


        obj("dirname") = dirname(path)
        obj("basename") = basename(path,"")
        obj("extension") = obj("basename")

        If inStr(obj("basename"),".") Then
            tmp = Split(obj("basename"),".")
            obj("extension") = tmp( uBound(tmp) )
        End if

        obj("filename") = Replace(obj("basename"),"." & obj("extension"),"")

        If len(options) > 0 Then

            If options = PATHINFO_DIRNAME Then
                pathinfo = obj("dirname")
            ElseIf options = PATHINFO_BASENAME Then
                pathinfo = obj("basename")
            ElseIf options = PATHINFO_EXTENSION Then
                pathinfo = obj("extension")
            ElseIf options = PATHINFO_FILENAME Then
                pathinfo = obj("filename")
            End if
            Exit Function
        End If

        set pathinfo = obj
    End Function

    '=======================================================================
    ' ディレクトリを削除する
    '=======================================================================
    '【引数】
    '  dirname     = string    ディレクトリへのパス。
    '【戻り値】
    '   成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  dirname で指定されたディレクトリを 削除しようと試みます。
    '  ディレクトリは空でなくてはならず、また 適切なパーミッションが設定されていなければなりません。
    '=======================================================================
    Public Function rmdir(ByVal dirname)

        dirname = fileMapPath(dirname)
        fso.DeleteFolder dirname
        rmdir = true

    End Function
    
    '=======================================================================
    ' ファイルを削除する
    '=======================================================================
    '【引数】
    '  filename     = string    ファイルへのパス。
    '【戻り値】
    '   成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  filename  を削除します。 Unix C 言語の関数 unlink() と動作は同じです。
    '=======================================================================
    Public Function unlink(ByVal filename)

        filename = fileMapPath(filename)
        fso.DeleteFile filename
        unlink = true

    End Function

End Class
