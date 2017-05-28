Option Explicit
Dim globalResult As Long  'グローバル変数を定義
'/**
'*フォルダ内のエクセルファイル数を返す関数（サブフォルダ検索はオプション）
'*WorkSheetFunction "A2" = sagasu("A1",hierarchy=true )
'*@pram {Variant} path ファイルを列挙したいフォルダへのパスを文字列もしくはセルで指定
'*@pram {Optional Boolean} hierarchy サブフォルダまで探すかどうかのオプション引数通常はFalse
'*@return {Number} sagasu　入っているファイル数
'*/

Function sagasu(ByVal path As Variant, Optional ByVal hierarchy As Boolean = False) As Long
    Dim result As Long '結果を返す為の変数を数値型で用意
    result = 0 '結果用変数を初期化
        If hierarchy Then
            result = deepFindFiles(path)
        Else
            result = findFiles(path)
        End If
    sagasu = result
End Function

'/**
'*フォルダ内のエクセルファイル数を返す関数（サブフォルダは検索しません）
'*WorkSheetFunction "A2" = findFiles("A1","*.xls")
'*@pram {Variant} path ファイルを列挙したいフォルダへのパスを文字列もしくはセルで指定
'*@pram {Optional Variant} fileType 探したいファイルのタイプを文字列もしくはセルで指定
' fileTypeへは探したいファイル名でも可、オプション引数の為指定がなければエクセルファイル
'*@return {Number} findFiles　入っているファイル数
'*/
Function findFiles(ByVal path As Variant, Optional ByVal fileType As Variant = "*.xls") As Long
  
    Dim result As Long '結果を返す為の変数を数値型で用意
    result = 0 '結果用変数を初期化
    Dim buf As String 'ファイル名を格納する一時変数
 
 '引数がセルの場合はセル内容を取り出し再代入
    If TypeName(path) = "Range" Then
        path = path.Value
    End If
    
 '引数がセルの場合はセル内容を取り出し再代入
    If TypeName(fileType) = "Range" Then
        fileType = fileType.Value
    End If
    If path = "" Then
        result = 0
        GoTo last
    End If
 'dir関数を用いてファイルタイプのファイルを列挙して条件合致でresultをインクリメント
    buf = Dir(path & fileType)
        Do While buf <> ""
            result = result + 1
            buf = Dir()
        Loop
last:
    findFiles = result
End Function

'/**
'*フォルダ内のエクセルファイル数を返す関数（サブフォルダを検索します）
'*WorkSheetFunction "A2" = deepFindFiles("A1","Excel")
'*@pram {Variant} path ファイルを列挙したいフォルダへのパスを文字列もしくはセルで指定
'*@pram {Optional Variant} fileType 探したいファイルのタイプを文字列もしくはセルで指定
'                                オプション引数の為指定がなければエクセルファイル
'*@return {Number} findFiles　入っているファイル数
'*/
Function deepFindFiles(ByVal path As Variant, Optional ByVal fileType As Variant = "Excel") As Long
    Dim result As Long '結果用変数を定義
    '引数がセルの場合はセル内容を取り出し再代入
    If TypeName(path) = "Range" Then
        path = path.Value
    End If
    '引数がセルの場合はセル内容を取り出し再代入
    If TypeName(fileType) = "Range" Then
        fileType = fileType.Value
    End If
    If path = "" Then
        result = 0
        GoTo last
    End If
    '各変数を初期化
    globalResult = 0
    result = 0
    '深い階層をチェックする為の関数を呼び出し
    result = deepFindFilesMain(path, fileType)
last:
    deepFindFiles = result
End Function
'/**
'*深い階層を探索してファイル数を返す関数（サブフォルダを検索します）
'* result = deepFindFilesMain(path,fileType)
'*@pram {String} path ファイルを列挙したいフォルダへのパスを文字列もしくはセルで指定
'*@pram {String} fileType 探したいファイルのタイプを文字列もしくはセルで指定
'*@return {Number} findFiles　入っているファイル数
'*/
Function deepFindFilesMain(path As Variant, fileType As Variant) As Long
    'ファイルシステムオブジェクトを使いファイル探索
    Dim fso As Object  'ファイルシステムオブジェクト用変数
    Dim fol As Variant  'フォルダオブジェクトを格納する為の変数
    Dim file As Variant  'ファイルオブジェクトを格納する為の変数
    '
    Set fso = CreateObject("Scripting.FileSystemObject")
        For Each fol In fso.GetFolder(path).subFolders
            Call deepFindFilesMain(fol.path, fileType)
        Next fol
        For Each file In fso.GetFolder(path).Files
            If InStr(file.Type, fileType) > 0 Then
                globalResult = globalResult + 1
            End If
        Next file
   Set fso = Nothing
   deepFindFilesMain = globalResult
End Function
