Module enCode
    '    Sub UTF8() 'https://msdn.microsoft.com/zh-tw/library/x14b16ab.aspx
    '        On Error GoTo eH
    '        Dim rst As DAO.Recordset, rst1 As DAO.Recordset, x As String
    '        Dim u8 As System.Text.Encoding = System.Text.Encoding.UTF8
    '        If db Is Nothing Then
    '            db = od.OpenDatabase("codes.mdb") ', False, True)
    '        End If
    '        rst = db.OpenRecordset("漢字總表_的複本TEST")
    '        rst1 = db.OpenRecordset("網路碼 的複本")
    '        With rst
    '            Do Until .EOF
    '                x = rst.Fields("漢字").Value
    '                'x = rst.Fields("字").Value
    '                With rst1
    '                    .AddNew()
    '                    rst1.Fields("字").Value = x
    '                    Dim bytes As Byte() = u8.GetBytes(x)
    '                    ' Display all the encoded bytes.
    '                    x = PrintHexBytes(bytes)
    '                    '.Edit()
    '                    '.Fields("utf8").Value = x
    '                    .Fields("網路碼").Value = x
    '                    .Update()
    '                End With
    '                .MoveNext()
    '            Loop
    '        End With
    '        'PrintCountsAndBytes(myChars, u8)
    '        'Encoding.UTF8() 'https://msdn.microsoft.com/zh-tw/library/system.text.encoding.utf8(v=vs.110).aspx
    '        Exit Sub
    'eH:
    '        Select Case Err.Number
    '            Case Else
    '                MsgBox(Err.Number & Err.Description)

    '        End Select
    '    End Sub
    Public Function PrintHexBytes(bytes() As Byte) As String
        PrintHexBytes = ""
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            'Console.WriteLine("<none>")
            MsgBox("<none>")
        Else
            Dim i As Integer
            For i = 0 To bytes.Length - 1
                PrintHexBytes &= "%" & Hex(bytes(i))
                'PrintHexBytes &= Hex(bytes(i))
                'Console.Write("{0:X2} ", bytes(i))
            Next i
            'Console.WriteLine()
        End If
    End Function 'PrintHexBytes 
    '    Function 查詢字串轉換_百度碼(ByVal w As String) '2008/12/18'今因百度較能索引spread故
    '        Dim x As String = "", i As Integer, rst As DAO.Recordset, q As String, bai As String, wb As String = ""
    '        'Static db As DAO.Database
    '        On Error GoTo ErrMsg
    '        If db Is Nothing Then
    '            db = od.OpenDatabase("codes.mdb") ', False, True)
    '        End If
    '        Dim qd As New DAO.QueryDef, p As DAO.Parameter
    '        qd = db.QueryDefs("check百度碼")
    '        p = qd.Parameters("q")
    '        For i = 1 To Len(w)
    '            p.Value = Mid(w, i, 1) : rst = qd.OpenRecordset
    '            If rst.RecordCount = 0 Then
    '                'If IsNull(DLookup("百度碼", "百度碼", "字 = """ & Mid(w, i, 1) & """")) Then
    '                rst = db.OpenRecordset("百度碼")
    '                With rst
    '                    .AddNew()
    '                    .Fields("字").Value = Mid(w, i, 1)
    '                    'setOX()
    '                    'OX.ClipPut(Mid(w, i, 1))
    '                    Clipboard.SetText(Mid(w, i, 1))
    '                    'AppActivate(pID) '(Mid(InStrRev(GetDefaultBrowserEXE, "\"))
    'Rep:                Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com"))
    '                    q = InputBox("請輸入百度碼", , Mid(w, i, 1))
    '                    AppActivate(pID)
    '                    If q = "" Or q = Mid(w, i, 1) Then
    '                        GoTo Rep
    '                    ElseIf InStr(q, "word=") Then
    '                        wb = "word="
    '                    ElseIf InStr(q, "wd=") Then
    '                        wb = "wd="
    '                    End If
    '                    'bai = Mid(q, InStr(q, wb) + Len(wb), 6)
    '                    bai = Mid(q, InStr(q, wb) + Len(wb))
    '                    If InStr(bai, "&") = 0 Then
    '                        .Fields("百度碼").Value = bai
    '                    Else
    '                        .Fields("百度碼").Value = Left(bai, InStr(bai, "&") - 1)
    '                    End If
    '                    .Update()
    '                    .Close()
    '                    dbCompactDatabase = True
    '                    p.Value = Mid(w, i, 1) : rst = qd.OpenRecordset
    '                End With
    '                GoTo t
    '            ElseIf IsDBNull(rst.Fields("百度碼").Value) Then
    '                With rst
    '                    .Edit()
    '                    Clipboard.SetText(Mid(w, i, 1))
    'Rep1:               Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com"))
    '                    q = InputBox("請輸入百度碼", , Mid(w, i, 1))
    '                    'AppActivate(Replace(GetDefaultBrowserEXE, " -- ""%1""", ""))
    '                    AppActivate(pID)
    '                    If q = "" Or q = Mid(w, i, 1) Then
    '                        GoTo Rep1
    '                    ElseIf InStr(q, "word=") Then
    '                        wb = "word="
    '                    ElseIf InStr(q, "wd=") Then
    '                        wb = "wd="
    '                    End If
    '                    'bai = Mid(q, InStr(q, wb) + Len(wb), 9)
    '                    bai = Mid(q, InStr(q, wb) + Len(wb))
    '                    If InStr(bai, "&") = 0 Then
    '                        .Fields("百度碼").Value = bai
    '                    Else
    '                        .Fields("百度碼").Value = Left(bai, InStr(bai, "&") - 1)
    '                    End If
    '                    .Update()
    '                    dbCompactDatabase = True
    '                    GoTo t
    '                End With
    '            Else
    't:              x = x & rst.Fields("百度碼").Value 'DLookup("百度碼", "百度碼", "字 = """ & Mid(w, i, 1) & """")
    '            End If
    '        Next i
    '        查詢字串轉換_百度碼 = x
    '        Exit Function
    'ErrMsg:
    '        Select Case Err.Number
    '            Case 5 '找不到處理序
    '                Resume Next
    '            Case Else
    '                MsgBox(Err.Number & " : " & Err.Description, , "查詢字串轉換_百度碼Error")
    '        End Select


    '    End Function

    '    Function 查詢字串轉換_網路碼(ByVal w As String) '2008/12/18'今因百度較能索引spread故
    '        Dim x As String = "", i As Integer, rst As DAO.Recordset, q As String, bai As String, wb As String = "", od As New DAO.DBEngine
    '        Static db As DAO.Database
    '        On Error GoTo ErrMsg
    '        If db Is Nothing Then
    '            db = od.OpenDatabase("codes.mdb") ', False, True)
    '        End If
    '        Dim qd As New DAO.QueryDef, p As DAO.Parameter
    '        qd = db.QueryDefs("check網路碼")
    '        p = qd.Parameters("q")
    '        For i = 1 To Len(w)
    '            p.Value = Mid(w, i, 1) : rst = qd.OpenRecordset
    '            If rst.RecordCount = 0 Then
    '                'If IsNull(DLookup("百度碼", "百度碼", "字 = """ & Mid(w, i, 1) & """")) Then
    '                'rst = db.OpenRecordset("網路碼")
    '                With rst
    '                    .AddNew()
    '                    .Fields("字").Value = Mid(w, i, 1)
    '                    'setOX()
    '                    'OX.ClipPut(Mid(w, i, 1))
    '                    Clipboard.SetText(Mid(w, i, 1))
    '                    'AppActivate(pID) '(Mid(InStrRev(GetDefaultBrowserEXE, "\"))
    'Rep:                Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com"))
    '                    q = InputBox("請輸入網路碼", , Mid(w, i, 1))
    '                    AppActivate(pID)
    '                    If q = "" Or q = Mid(w, i, 1) Then
    '                        GoTo Rep
    '                    ElseIf InStr(q, "word=") Then
    '                        wb = "word="
    '                    ElseIf InStr(q, "wd=") Then
    '                        wb = "wd="
    '                    End If
    '                    'bai = Mid(q, InStr(q, wb) + Len(wb), 6)
    '                    bai = Mid(q, InStr(q, wb) + Len(wb))
    '                    If InStr(bai, "&") = 0 Then
    '                        .Fields("網路碼").Value = bai
    '                    Else
    '                        .Fields("網路碼").Value = Left(bai, InStr(bai, "&") - 1)
    '                    End If
    '                    .Update()
    '                    .Close()
    '                    p.Value = Mid(w, i, 1) : rst = qd.OpenRecordset
    '                End With
    '                GoTo t
    '            ElseIf IsDBNull(rst.Fields("網路碼").Value) Then
    '                With rst
    '                    .Edit()
    '                    Clipboard.SetText(Mid(w, i, 1))
    'Rep1:               Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com"))
    '                    q = InputBox("請輸入網路碼", , Mid(w, i, 1))
    '                    'AppActivate(Replace(GetDefaultBrowserEXE, " -- ""%1""", ""))
    '                    AppActivate(pID)
    '                    If q = "" Or q = Mid(w, i, 1) Then
    '                        GoTo Rep1
    '                    ElseIf InStr(q, "word=") Then
    '                        wb = "word="
    '                    ElseIf InStr(q, "wd=") Then
    '                        wb = "wd="
    '                    End If
    '                    'bai = Mid(q, InStr(q, wb) + Len(wb), 9)
    '                    bai = Mid(q, InStr(q, wb) + Len(wb))
    '                    If InStr(bai, "&") = 0 Then
    '                        .Fields("網路碼").Value = bai
    '                    Else
    '                        .Fields("網路碼").Value = Left(bai, InStr(bai, "&") - 1)
    '                    End If
    '                    .Update()
    '                    GoTo t
    '                End With
    '            Else
    't:              x = x & rst.Fields("百度碼").Value 'DLookup("百度碼", "百度碼", "字 = """ & Mid(w, i, 1) & """")
    '            End If
    '        Next i
    '        查詢字串轉換_網路碼 = x
    '        Exit Function
    'ErrMsg:
    '        Select Case Err.Number
    '            Case 5 '找不到處理序
    '                Resume Next
    '            Case Else
    '                MsgBox(Err.Number & " : " & Err.Description, , "查詢字串轉換_百度碼Error")
    '        End Select


    '    End Function

    Function 查詢字串轉換_網路碼(w As String)
        Dim u8 As System.Text.Encoding = System.Text.Encoding.UTF8
        Dim bytes As Byte() = u8.GetBytes(w)
        查詢字串轉換_網路碼 = PrintHexBytes(bytes)
    End Function
    Function 查詢字串轉換_Big5碼(w As String) '國語會碼
        Dim u8 As System.Text.Encoding =
            System.Text.Encoding.GetEncoding("big5")
        Dim bytes As Byte() = u8.GetBytes(w)
        查詢字串轉換_Big5碼 = PrintHexBytes(bytes)
    End Function
    '    Function 查詢字串轉換_網路碼_舊(w As String)
    '        Dim x As String = "", i As Integer, rst As DAO.Recordset, q As String, qd As DAO.QueryDef, p As DAO.Parameter, hz As String, hzascw As Integer
    '        On Error GoTo ErrMsg
    '        If db Is Nothing Then
    '            db = od.OpenDatabase("codes.mdb") ', False, True)
    '        End If
    '        For i = 1 To Len(w)
    '            qd = db.QueryDefs("check網路碼") : p = qd.Parameters("q")
    '            'If IsNull(DLookup("網路碼", "網路碼", "字 = """ & hz & """")) Then
    '            'If -10176 <= AscW(Mid(w, i, 1)) <= -10130 Or AscW(Mid(w, i, 1)) = -10114 Then '判斷漢字是否長2字元
    '            hzascw = AscW(Mid(w, i, 1))
    '            If 55360 <= hzascw And hzascw <= 55406 Or AscW(Mid(w, i, 1)) = 55422 Then '判斷漢字是否長2字元'VB與VBA值不同
    '                hz = Mid(w, i, 2)
    '                i = i + 1
    '            Else
    '                hz = Mid(w, i, 1)
    '            End If
    '            p.Value = hz : rst = qd.OpenRecordset
    '            If rst.RecordCount = 0 Then
    '                'rst = CurrentDb.OpenRecordset("網路碼")
    '                With rst
    '                    .AddNew()
    '                    .Fields("字").Value = hz
    '                    Clipboard.SetText(hz)

    'Rep:                'Shell(Replace(GetDefaultBrowserEXE, """%1", "https://www.google.com.tw/"))
    '                    q = InputBox("請輸入網路碼", , hz)
    '                    '今發現FierFox在網址列輸入漢字亦可轉碼,不必借助google轉,故改用此式.但太慢,且僅按F6亦不成,故二者相輔相成!
    '                    If q = "" Or q = hz Then GoTo Rep
    '                    If InStr(q, "google") Then
    '                        If InStr(q, "&") = 0 Then
    '                            If InStr(Mid(q, InStr(q, "q=") + 1), "&") = 0 Then
    '                                q = Mid(q, InStr(q, "q=") + Len("q="))
    '                            Else
    '                                q = Mid(q, InStr(q, "q=") + Len("q="), InStr(InStr(q, "q="), q, "&") - InStr(q, "q=") - Len("q=")) '直接貼上google轉換的網址並取其值2007/10/8
    '                            End If
    '                        ElseIf InStrRev(q, "&") < InStrRev(q, "q=") Then
    '                            q = Mid(q, InStr(q, "q=") + Len("q="))
    '                        ElseIf InStrRev(q, "&") > InStrRev(q, "q=") Then
    '                            q = Mid(q, InStr(q, "q=") + Len("q="), InStr(InStr(q, "q="), q, "&") - InStr(q, "q=") - Len("q=")) '直接貼上google轉換的網址並取其值2007/10/8

    '                        End If

    '                        'Else
    '                        '.Fields("網路碼") = Mid(q, InStr(q, ":") + 1)

    '                    End If
    '                    .Fields("網路碼").Value = q
    '                    .Update()
    '                    dbCompactDatabase = True
    '                    '.Close()
    '                    'CurrentDb.Close()
    '                    .Requery()
    '                End With
    '                GoTo t
    '            Else
    't:              x = x & rst.Fields("網路碼").Value 'x = x & DLookup("網路碼", "網路碼", "字 = """ & hz & """")
    '            End If
    '        Next i
    '        查詢字串轉換_網路碼_舊 = x
    '        Exit Function
    'ErrMsg:
    '        Select Case Err.Number
    '            Case 5 '找不到處理序
    '                Resume Next
    '            Case Else
    '                MsgBox(Err.Number & " : " & Err.Description, , "查詢字串轉換__網路碼Error")
    '        End Select
    '    End Function

    '    Sub 壓縮檔案(filename As String)
    '        On Error GoTo eH
    '        db.Close()
    '        od.CompactDatabase(filename, "temp.mdb") ' Replace(SourceFile, filename, "temp.mdb"))  '壓縮為暫存檔'當有引用,則不壓縮時會造成錯誤!
    '        'Kill(SourceFile) '刪除選取檔案
    '        My.Computer.FileSystem.DeleteFile(filename) 'https://msdn.microsoft.com/zh-tw/library/ms127977(v=vs.100).aspx
    '        'Name Replace(SourceFile, filename, "temp.mdb") As SourceFile
    '        My.Computer.FileSystem.RenameFile("temp.mdb", filename) 'https://msdn.microsoft.com/zh-tw/library/microsoft.visualbasic.fileio.filesystem.renamefile(v=vs.100).aspx
    '        Exit Sub
    'eH:
    '        Select Case Err.Number
    '            Case Else
    '                MsgBox(Err.Number & Err.Description)
    '        End Select
    '    End Sub
End Module
