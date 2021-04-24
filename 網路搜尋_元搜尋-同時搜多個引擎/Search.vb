Module Search
    'Public dbCompactDatabase As Boolean
    'Public pID As Integer ', db As DAO.Database, od As New DAO.DBEngine
    Dim stxt As String = Strings.Left(Clipboard.GetText, 200)
    Sub Searches()
        查詢百度()
        查詢百度百科()
        查詢Yahoo()
        查詢Bing()
        查詢Google()
        End
    End Sub
    Sub 查詢百度百科() '2008/12/23 Ctrl+F1 '因為百科內字號別名資料不少 ;今加國學'12/28
        Dim baidu As String
        baidu = 查詢字串轉換_網路碼(stxt) 'Form1.TextBox1.Text) '查詢字串轉換_百度碼(Form1.TextBox1.Text) 'baidu = 查詢字串轉換_百度碼(Screen.ActiveControl.seltext)
        BrowserOps.openUrl(BrowserApp,
            "https://baike.baidu.com/item/" & baidu)
        ' ''FollowHyperlink "http://baike.baidu.com/notexists", , , , "word=" & baidu, msoMethodGet
        ' ''FollowHyperlink "http://baike.baidu.com/w", , , , "ct=17&lm=0&tn=baiduWikiSearch&pn=0&rn=10&word=" & baidu & "&submit=search", msoMethodGet
        ' ''FollowHyperlink "http://guoxue.baidu.com/s", , , , "tn=baiduguoxue&ie=gb2312&bs=&cl=3&wd=" & baidu & "&si=guoxue.baidu.com&ct=2097152", msoMethodGet

        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/notexists?word=" & baidu))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://baike.baidu.com/w?&word=" & baidu))
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
        baidu = 查詢字串轉換_網路碼(stxt) 'Form1.TextBox1.Text) '查詢字串轉換_百度碼(Form1.TextBox1.Text)
        BrowserOps.openUrl(BrowserApp,
                           "https://www.baidu.com/s?&wd=" & baidu)
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com/s?&wd=" & baidu))
        'Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&rsv_idx=1&tn=baidu&wd=" & baidu & "&rsv_pq=808cafa70002b1f5&rsv_t=86e2wQLHYztZWmhzpgxdUc4SV52Z47aJocHjNCt1NLpUxjoggzxXt2Cr1t0&rsv_enter=1&rsv_sug3=2&rsv_sug1=1&rsv_sug2=0&inputT=450&rsv_sug4=1271"))

    End Sub
    Sub 查詢Bing()
        On Error GoTo ErrMsg
        BrowserOps.openUrl(BrowserApp,
             "http://www.bing.com/search?q=" &
             查詢字串轉換_網路碼(stxt)) 'Form1.TextBox1.Text))
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        '"http://www.bing.com/search?q=" &
        '查詢字串轉換_網路碼(Form1.TextBox1.Text)))
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
        BrowserOps.openUrl(BrowserApp,
            "https://tw.search.yahoo.com/search?p=" &
            查詢字串轉換_網路碼(stxt)) 'Form1.TextBox1.Text))
        'Shell(Replace(GetDefaultBrowserEXE, """%1",
        '"https://tw.search.yahoo.com/search?p=" &
        '查詢字串轉換_網路碼(Form1.TextBox1.Text)))
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
        BrowserOps.openUrl(BrowserApp,
            "http://www.google.com.tw/search?q=" &
            查詢字串轉換_網路碼(stxt)) 'Form1.TextBox1.Text))

        ''FollowHyperlink "http://tw.search.yahoo.com/search", , , , "fr=slv1-ptec&p=" & Screen.ActiveControl.seltext
        ''FollowHyperlink "http://www.google.com.tw/search", , , , "q=" & Screen.ActiveControl.seltext, msoMethodGet
        ''Shell Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.tw/search?q=" & Screen.ActiveControl.seltext)
        ''If Tasks.Exists("skqs professional version") Then
        'pID = Shell(Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.tw/search?q=" & 查詢字串轉換_網路碼(Form1.TextBox1.Text)))
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

End Module
