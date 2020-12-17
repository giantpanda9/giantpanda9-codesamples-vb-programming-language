Public Class Form1

    Dim ApPat, FSO, Folder, TextStream, temp, File, MyTime, StringN, tmp0, Numb(1000), Strg(1000), checkN, curTemp
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ' On Error GoTo vse
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Dim i
        ApPat = Application.StartupPath
        FSO = CreateObject("Scripting.FileSystemObject")
        TextStream = FSO.CreateTextFile(ApPat + "\recent.ini", True)
        TextStream.Close()
        File = FSO.GetFile(ApPat + "\recent.ini")
        TextStream = File.OpenAsTextStream(2)
        For i = 0 To ListBox3.Items.Count - 1
            TextStream.Write(ListBox3.Items.Item(i) & vbCrLf)
        Next i
        TextStream.Close()
        End
vse:
        End
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        On Error GoTo vse
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Dim i, mng
        If TextBox2.Text <> vbNullString Then
            ApPat = Application.StartupPath
            FSO = CreateObject("Scripting.FileSystemObject")
            If CheckBox1.Checked = True Then
                Folder = FSO.GetFolder(ApPat + "\" + TextBox2.Text)
                Folder.Delete()
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + ": Project successfully overwriten")
            End If
            'On Error GoTo ErrorFolderAlreadyExists
            If FSO.FolderExists(ApPat + "\" + TextBox2.Text) Then
                Folder = FSO.GetFolder(ApPat + "\" + TextBox2.Text)
                Folder.Delete()
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + ": Project successfully overwriten")
            Else
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + ": Project saved successfully")
            End If
            FSO.CreateFolder(ApPat + "\" + TextBox2.Text)
            Folder = FSO.GetFolder(ApPat + "\" + TextBox2.Text)
            TextStream = Folder.CreateTextFile(TextBox2.Text & ".@code", True)
            TextStream.Close()
            File = FSO.GetFile(ApPat + "\" + TextBox2.Text + "\" + TextBox2.Text + ".@code")
            TextStream = File.OpenAsTextStream(2)
            TextStream.Write(TextBox1.Text + vbCrLf)
            TextStream.Close()
            For i = 0 To ListBox3.Items.Count - 1
                mng = TextBox2.Text
                ListBox3.Items.Remove(mng)

            Next i
            ListBox3.Items.Add(TextBox2.Text)
            Button4.Enabled = True
           

        Else
            MyTime = DateTime.Now
            MsgBox("Give your program a name:", MsgBoxStyle.Exclamation, "@-Language")
            ListBox1.Items.Add(MyTime + ": Program is unnamed")
        End If
        GoTo vse
        'ErrorFolderAlreadyExists:
        'Folder = FSO.GetFolder(ApPat + "\" + TextBox2.Text)
        'Folder.Delete()
        'MyTime = DateTime.Now
        'MsgBox("There is an expected error occured! Try again!", MsgBoxStyle.Exclamation, "@-Language")
        'ListBox1.Items.Add(MyTime + ": Project folder was already exists and has been removed! Try again!")
vse:
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        On Error GoTo vse
        Dim i, mng
        ListBox2.Items.Clear()
        ListBox1.Items.Clear()
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        If TextBox2.Text <> vbNullString Then
            ApPat = Application.StartupPath
            FSO = CreateObject("Scripting.FileSystemObject")
            On Error GoTo ErrorNoProject
            Folder = FSO.GetFolder(ApPat + "\" + TextBox2.Text)
            File = FSO.GetFile(ApPat + "\" + TextBox2.Text + "\" + TextBox2.Text + ".@code")
            TextStream = File.OpenAsTextStream(1)
            temp = TextStream.ReadAll()
            TextStream.Close()
            TextBox1.Text = temp
            For i = 0 To ListBox3.Items.Count - 1
                mng = TextBox2.Text
                ListBox3.Items.Remove(mng)

            Next i
            ListBox3.Items.Add(TextBox2.Text)
            Button4.Enabled = True
        Else
            MyTime = DateTime.Now
            MsgBox("What's a program do you want to open?", MsgBoxStyle.Exclamation, "@-Language")
            ListBox1.Items.Add(MyTime + ": Can't open program. Program name is not available.")
        End If
        GoTo vse
ErrorNoProject:
        MyTime = DateTime.Now
        MsgBox("No project file!", MsgBoxStyle.Critical, "@-Language")
        ListBox1.Items.Add(MyTime + ": No project.")
vse:
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Clear()
        TextBox2.Clear()
        ListBox2.Items.Clear()
        ListBox1.Items.Clear()
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        On Error GoTo vse
        If TextBox2.Text <> vbNullString Then
            StringN = 0
            ApPat = Application.StartupPath
            FSO = CreateObject("Scripting.FileSystemObject")
            'On Error GoTo ErrorNoProg
            Folder = FSO.GetFolder(ApPat + "\" + TextBox2.Text)
            File = FSO.GetFile(ApPat + "\" + TextBox2.Text + "\" + TextBox2.Text + ".@code")
            TextStream = File.OpenAsTextStream(1)
            ListBox2.Items.Clear()
            temp = TextStream.ReadLine()
            ListBox2.Items.Add(temp)
            If UCase(temp) = "START" Then
                ListBox2.Items.Clear()
                ListBox2.Items.Add(temp)
                Do

                    temp = TextStream.ReadLine()
                    ' Memory Allocation
                    If UCase(temp) = "NUMBERS" Or UCase(temp) = "STRINGS" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 2 Or checkN < 2 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If


                        'TriByte Operators
                    ElseIf UCase(temp) = "ADD" Or UCase(temp) = "SUB" Or UCase(temp) = "MUL" Or UCase(temp) = "DIV" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 4 Or checkN < 4 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                        'CONCAT
                    ElseIf UCase(temp) = "CONDOG" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 4 Or checkN < 4 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                        'IFNXTN\S
                    ElseIf UCase(temp) = "IFNXTN" Or UCase(temp) = "IFNXTS" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 4 Or checkN < 4 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                        '4Byte Operators
                    ElseIf UCase(temp) = "PUTS" Or UCase(temp) = "PUTN" Or UCase(temp) = "GETS" Or UCase(temp) = "GETN" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 2 Or checkN < 2 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                    ElseIf UCase(temp) = "APPS" Or UCase(temp) = "APPN" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 3 Or checkN < 3 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                    ElseIf UCase(temp) = "COMM" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"


                        If checkN > 4 Or checkN < 4 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                    ElseIf UCase(temp) = "OUTS" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"
                        If checkN > 2 Or checkN < 2 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                    ElseIf UCase(temp) = "OUTL" Then
                        curTemp = UCase(temp)
                        ListBox2.Items.Add(temp)
                        checkN = 0
                        Do
                            temp = TextStream.ReadLine()
                            ListBox2.Items.Add(temp)
                            checkN = checkN + 1
                        Loop Until temp = "@"


                        If checkN > 2 Or checkN < 2 Then
                            MyTime = DateTime.Now
                            ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + curTemp + " parameters quantity are incorrect. Scan has been stopped.")
                            GoTo vse
                        Else : Continue Do
                        End If

                    ElseIf UCase(temp) = "STOP" Then
                        GoTo endLoop
                    Else
                        MyTime = DateTime.Now
                        ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Incorrect operator. Scan has been stopped.")
                        GoTo vse
                    End If

endLoop:
                Loop Until UCase(temp) = "STOP"
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Scan was succesful. Debug your program before ''Make @exe + Run''. Use ''Debug'' button to do it.")
                Button5.Enabled = True
                ListBox2.Items.Add(temp)

            Else
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": No START instruction. Scan stopped.")
            End If

            TextStream.Close()


        Else
            MyTime = DateTime.Now
            MsgBox("What's a program d'u want to scan?:", MsgBoxStyle.Exclamation, "@-Language")
            ListBox1.Items.Add(MyTime + ": Can't scan program. Name is N/A.")
        End If
        GoTo vse
ErrorNoProg:
        MyTime = DateTime.Now
        MsgBox("Program file doesn't exist, write a program first.", MsgBoxStyle.Critical, "@-Language")
        ListBox1.Items.Add(MyTime + ": No program. Nothing to scan.")

vse:

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        On Error GoTo Vse
        Dim i
        i = 1
        For i = 1 To ListBox2.Items.Count - 1
            'Memory Allocation
            If UCase(ListBox2.Items.Item(i)) = "NUMBERS" Or UCase(ListBox2.Items.Item(i)) = "STRINGS" Then
                On Error GoTo ErrorMemAlloc
                If -1000 <= ListBox2.Items.Item(i + 1) <= +1000 And ListBox2.Items.Item(i + 2) = "@" Then
                    Continue For
                Else
ErrorMemAlloc:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If

            'TriByte Operators
            If UCase(ListBox2.Items.Item(i)) = "ADD" Or UCase(ListBox2.Items.Item(i)) = "SUB" Or UCase(ListBox2.Items.Item(i)) = "MUL" Or UCase(ListBox2.Items.Item(i)) = "DIV" Then

                On Error GoTo ErrorTriByteOp
                If -1000 <= ListBox2.Items.Item(i + 1) <= +1000 And -1000 <= ListBox2.Items.Item(i + 2) <= +1000 And -1000 <= ListBox2.Items.Item(i + 3) <= +1000 And ListBox2.Items.Item(i + 4) = "@" Then
                    Continue For
                Else
ErrorTriByteOp:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If

            'CONDOG
            If UCase(ListBox2.Items.Item(i)) = "CONDOG" Then

                On Error GoTo ErrorConDog
                If -1000 <= ListBox2.Items.Item(i + 1) <= +1000 And -1000 <= ListBox2.Items.Item(i + 2) <= +1000 And -1000 <= ListBox2.Items.Item(i + 3) <= +1000 And ListBox2.Items.Item(i + 4) = "@" Then
                    Continue For
                Else
ErrorConDog:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If

            'IFNXTS\N
            If UCase(ListBox2.Items.Item(i)) = "IFNXTS" Or UCase(ListBox2.Items.Item(i)) = "IFNXTN" Then

                On Error GoTo ErrorIfNxtSN
                If -1000 <= ListBox2.Items.Item(i + 1) <= +1000 And -1000 <= ListBox2.Items.Item(i + 2) <= +1000 And ListBox2.Items.Item(i + 4) = "@" Then
                    Continue For
                Else
ErrorIfNxtSN:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If

            '4Byte operators
            If UCase(ListBox2.Items.Item(i)) = "PUTS" Or UCase(ListBox2.Items.Item(i)) = "PUTN" Or UCase(ListBox2.Items.Item(i)) = "GETS" Or UCase(ListBox2.Items.Item(i)) = "GETN" Then

                On Error GoTo Error4ByteOp
                If -1000 <= ListBox2.Items.Item(i + 1) <= +1000 And ListBox2.Items.Item(i + 2) = "@" Then
                    Continue For
                Else
Error4ByteOp:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If


            'App Operators
            If UCase(ListBox2.Items.Item(i)) = "APPN" Or UCase(ListBox2.Items.Item(i)) = "APPN" Then

                On Error GoTo ErrorAppOp
                If -1000 <= ListBox2.Items.Item(i + 1) <= +1000 And -1000 <= ListBox2.Items.Item(i + 2) <= +1000 And ListBox2.Items.Item(i + 3) = "@" Then
                    Continue For
                Else
ErrorAppOp:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If


            'COMM
            If UCase(ListBox2.Items.Item(i)) = "COMM" Then

                On Error GoTo ErrorCommOp
                If ListBox2.Items.Item(i + 4) = "@" Then
                    Continue For
                Else
ErrorCommOp:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If


            'OUTS
            If UCase(ListBox2.Items.Item(i)) = "OUTS" Then

                On Error GoTo ErrorCommOuts
                If ListBox2.Items.Item(i + 2) = "@" Then
                    Continue For
                Else
ErrorCommOuts:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If
            'OUTL
            If UCase(ListBox2.Items.Item(i)) = "OUTS" Then

                On Error GoTo ErrorCommOutl
                If ListBox2.Items.Item(i + 2) = "@" Then
                    Continue For
                Else
ErrorCommOutl:
                    MyTime = DateTime.Now
                    ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Operator " + UCase(ListBox2.Items.Item(i)) + " parameters are incorrect or there is no ''@'' - symbol here. Debug stopped.")
                    GoTo Vse
                End If
            End If
        Next i
        Button6.Enabled = True
        MyTime = DateTime.Now
        ListBox1.Items.Add(MyTime + "_" + TextBox2.Text + ": Debug succeed ")
Vse:
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i

        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        ListBox3.Items.Add("Recent projects are:")
        TextBox1.Clear()
        TextBox2.Clear()
        ListBox2.Items.Clear()
        ListBox1.Items.Clear()
        ApPat = Application.StartupPath
        ListBox3.Items.Clear()
        On Error GoTo nextg
        FSO = CreateObject("Scripting.FileSystemObject")
        File = FSO.GetFile(ApPat + "\recent.ini")
        TextStream = File.OpenAsTextStream(1)

        Do
            temp = TextStream.ReadLine
            ListBox3.Items.Add(temp)
        Loop Until TextStream.AtEndOfStream
nextg:
        If ListBox3.Items.Count = 0 Then ListBox3.Items.Add("Recent projects are:")
        TextStream.Close()
        PictureBox1.LoadAsync(ApPat + "\pic\logo.gif")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        On Error GoTo vse
        Dim yesno
        If TextBox2.Text <> vbNullString Then
            ApPat = Application.StartupPath
            FSO = CreateObject("Scripting.FileSystemObject")
            'On Error GoTo ErrorResolve
            If FSO.FolderExists(ApPat + "\" + "\Release_" + TextBox2.Text) Then
                Folder = FSO.GetFolder(ApPat + "\" + "\Release_" + TextBox2.Text)
                Folder.Delete()
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + ": @exe file has been made successfully.")
            Else
                MyTime = DateTime.Now
                ListBox1.Items.Add(MyTime + ": @exe file has been remade successfully.")
            End If
            FSO.CreateFolder(ApPat + "\" + "\Release_" + TextBox2.Text)
            File = FSO.GetFile(ApPat + "\" + TextBox2.Text + "\" + TextBox2.Text + ".@code")
            File.Copy(ApPat + "\Release_" + TextBox2.Text + "\prog.dat")
            File = FSO.GetFile(ApPat + "\" + "makeexe\doglangapp.exe")
            File.Copy(ApPat + "\Release_" + TextBox2.Text + "\" + TextBox2.Text + ".exe")
            Folder = FSO.GetFolder(ApPat + "\" + "\Release_" + TextBox2.Text)
            TextStream = Folder.CreateTextFile("prog.ini", True)
            TextStream.Close()
            File = FSO.GetFile(ApPat + "\" + "\Release_" + TextBox2.Text + "\prog.ini")
            TextStream = File.OpenAsTextStream(2)
            TextStream.Write(TextBox2.Text)
            TextStream.Close()
            yesno = MsgBox("Exe-File is succefully created. Do you want to run it?", MsgBoxStyle.YesNo, "@-Language")
            If yesno = vbYes Then
                Shell(ApPat + "\Release_" + TextBox2.Text + "\" + TextBox2.Text + ".exe", AppWinStyle.NormalFocus)
            Else
                GoTo vse
            End If
        Else
            MyTime = DateTime.Now
            MsgBox("What program do you want to compile and run?", MsgBoxStyle.Exclamation, "@-Language")
            ListBox1.Items.Add(MyTime + ": Can't compile program. progran name is not available.")
        End If
        GoTo vse
        'ErrorResolve:
        'Folder = FSO.GetFolder(ApPat + "\" + "\Release_" + TextBox2.Text)
        '    Folder.Delete()
        ' MyTime = DateTime.Now
        'MsgBox("There is an expected error occured! Try again!", MsgBoxStyle.Exclamation, "@-Language")
        'ListBox1.Items.Add(MyTime + ": Release folder was already exists and has been removed! Try again!")
vse:
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox3.SelectedIndexChanged
        Dim itm
        itm = ListBox3.SelectedItem
        If itm <> vbNullString Then
            TextBox2.Text = itm
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ListBox3.ClearSelected()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        ListBox3.Items.Clear()
        ListBox3.Items.Add("Recent projects are:")
    End Sub
End Class
