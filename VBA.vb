Function arrpush(targetRange As Long, shName As String, Optional zero As Boolean = False, _
Optional padi As Long = 0, Optional row As Long = 1, Optional fmt As Boolean = False, Optional targetbook As Workbook) As String()
'/**
'*配列格納関数（オプション付き）
'*returnArray = arrpusharrpush(1,"sheetName",False,0,1,False,"workBookName",1)
'*@param {Number} targetRange ターゲット列
'*@param {String} shName ターゲットシート名
'*@param {Boolean} zero ゼロ埋めするかしないか
'*@param {Number} padi ゼロ埋めで切り出す文字数
'*@param {Number} row 開始行
'*@param {Boolean} fmt yyMMddフォーマットにするかしないか
'*@param {WorkbookObject} targetWorkbook WorkbookObject
'*@return {String()} arrpush 文字列の配列
    Dim tarsh As Worksheet
    Dim lastRow As Long
    Dim fileName As String
    Dim namearr() As String
    Dim i As Integer
    Dim j As Integer
    j = 0
    If targetbook Is Nothing Then
        Set targetbook = ThisWorkbook
    End If
    Set tarsh = targetbook.Worksheets(shName)
    lastRow = tarsh.Cells(Rows.Count, targetRange).End(xlUp).row
        For i = row To lastRow Step 1
            fileName = tarsh.Cells(i, targetRange).value
                If zero Then
                    fileName = "0000000" & fileName
                    fileName = Right(fileName, padi)
                End If
                If fmt Then
                    fileName = Format(fileName, "yyMMdd")
                End If
            ReDim Preserve namearr(j)
            namearr(j) = fileName
            j = j + 1
        Next i
    arrpush = namearr
End Function

