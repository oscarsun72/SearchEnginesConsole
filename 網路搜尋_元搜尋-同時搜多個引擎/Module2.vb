Module Module2
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
End Module
