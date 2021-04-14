Module Module1
    'Public dbCompactDatabase As Boolean
    Public pID As Integer ', db As DAO.Database, od As New DAO.DBEngine

    Sub 查詢百度百科() '2008/12/23 Ctrl+F1 '因為百科內字號別名資料不少 ;今加國學'12/28
        Dim baidu As String
        baidu = 查詢字串轉換_網路碼(Form1.TextBox1.Text) '查詢字串轉換_百度碼(Form1.TextBox1.Text) 'baidu = 查詢字串轉換_百度碼(Screen.ActiveControl.seltext)
        ' ''FollowHyperlink "http://baike.baidu.com/notexists", , , , "word=" & baidu, msoMethodGet
        ' ''FollowHyperlink "http://baike.baidu.com/w", , , , "ct=17&lm=0&tn=baiduWikiSearch&pn=0&rn=10&word=" & baidu & "&submit=search", msoMethodGet
        ' ''FollowHyperlink "http://guoxue.baidu.com/s", , , , "tn=baiduguoxue&ie=gb2312&bs=&cl=3&wd=" & baidu & "&si=guoxue.baidu.com&ct=2097152", msoMethodGet

        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/notexists?word=" & baidu))
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/w?&word=" & baidu))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/search/none?word=" & baidu & "&pn=0&rn=10&enc=utf8"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/search/none?word=" & baidu & "&pn=0&rn=10&enc=utf8"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/w?ct=17&lm=0&tn=baiduWikiSearch&pn=0&rn=10&word=" & baidu & "&submit=search"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://guoxue.baidu.com/s?tn=baiduguoxue&ie=gb2312&bs=&cl=3&wd=" & baidu & "&si=guoxue.baidu.com&ct=2097152"))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.hk/custom?domains=guoxue.com&q=" & 查詢字串轉換_網路碼(Screen.ActiveControl.seltext) & "&sa=%E5%9B%BD%E5%AD%A6%E6%90%9C%E7%B4%A2&client=pub-7066894251315332&forid=1&ie=UTF-8&oe=UTF-8&safe=active&sitesearch=guoxue.com&cof=GALT%3A%23008000%3BGL%3A1%3BDIV%3A%23336699%3BVLC%3A663399%3BAH%3Acenter%3BBGC%3AFFFFFF%3BLBGC%3AFFFFFF%3BALC%3A0000FF%3BLC%3A0000FF%3BT%3A000000%3BGFNT%3A0000FF%3BGIMP%3A0000FF%3BLH%3A50%3BLW%3A149%3BL%3Ahttp%3A%2F%2Fwww.guoxue.com%2Fpic%2Fgxbd.gif%3BS%3Ahttp%3A%2F%2Fwww.guoxue.com%3BFORID%3A1&hl=zh-CN"))
        ''or: http://baike.baidu.com/list-php/dispose/searchword.php?word=%D7%DE%20%BD%F5%D5%C2&pic=1
    End Sub
    Sub 查詢百度()
        Dim baidu As String
        'baidu = Form1.TextBox1.Text '新百度碼實則已採用網路碼了
        baidu = 查詢字串轉換_網路碼(Form1.TextBox1.Text) '查詢字串轉換_百度碼(Form1.TextBox1.Text)
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com/s?&wd=" & baidu))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&rsv_idx=1&tn=baidu&wd=" & baidu & "&rsv_pq=808cafa70002b1f5&rsv_t=86e2wQLHYztZWmhzpgxdUc4SV52Z47aJocHjNCt1NLpUxjoggzxXt2Cr1t0&rsv_enter=1&rsv_sug3=2&rsv_sug1=1&rsv_sug2=0&inputT=450&rsv_sug4=1271"))

    End Sub
    Sub 查詢Bing()
        On Error GoTo ErrMsg
        Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.bing.com/search?q=" & 查詢字串轉換_網路碼(Form1.TextBox1.Text)))
        'Form1.TextBox1.Copy()
        Exit Sub
ErrMsg:
        Select Case Err.Number
            Case 462 '遠端伺服器不存在或無法使用
                End 'word經重新開啟
            Case Else
                MsgBox(Err.Number & " : " & Err.Description)
        End Select
    End Sub
    Sub 查詢Yahoo()
        On Error GoTo ErrMsg
        Shell(Replace(GetDefaultBrowserEXE, """%1", "https://tw.search.yahoo.com/search?p=" & 查詢字串轉換_網路碼(Form1.TextBox1.Text)))
        'Form1.TextBox1.Copy()
        Exit Sub
ErrMsg:
        Select Case Err.Number
            Case 462 '遠端伺服器不存在或無法使用
                End 'word經重新開啟
            Case Else
                MsgBox(Err.Number & " : " & Err.Description)
        End Select
    End Sub
    Sub 查詢Google() '快速鍵'Ctrl+shift+g
        On Error GoTo ErrMsg
        ''FollowHyperlink "http://tw.search.yahoo.com/search", , , , "fr=slv1-ptec&p=" & Screen.ActiveControl.seltext
        ''FollowHyperlink "http://www.google.com.tw/search", , , , "q=" & Screen.ActiveControl.seltext, msoMethodGet
        ''Shell Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.tw/search?q=" & Screen.ActiveControl.seltext)
        ''If Tasks.Exists("skqs professional version") Then
        pID = Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.tw/search?q=" & 查詢字串轉換_網路碼(Form1.TextBox1.Text)))
        'Else
        ''    'Shell "C:\Program Files\Opera\opera.exe" & " http://www.google.com.tw/search?q=" & Screen.ActiveControl.seltext
        ''    Shell "C:\Program Files\Opera\opera.exe" & " http://www.google.com.tw/search?q=" & Screen.ActiveControl.seltext, vbNormalFocus
        ''End If
        'DoCmd.RunCommand(acCmdCopy)
        'Form1.TextBox1.Copy()
        '按下掃描鍵_不最小化()
        Exit Sub
ErrMsg:
        Select Case Err.Number
            Case 462 '遠端伺服器不存在或無法使用
                End 'word經重新開啟
            Case Else
                MsgBox(Err.Number & " : " & Err.Description)
        End Select
    End Sub

    Function GetDefaultBrowserEXE() '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
        Dim objShell
        objShell = CreateObject("WScript.Shell")
        'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
        '取得註冊表中的值
        GetDefaultBrowserEXE = objShell.RegRead _
                ("HKCR\http\shell\open\command\")


    End Function

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
        Dim u8 As System.Text.Encoding = System.Text.Encoding.GetEncoding("big5")
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
