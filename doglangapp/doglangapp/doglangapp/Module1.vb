Imports System.Windows.Forms
Imports System.Web.Hosting
Module Module1

    Dim ApPat, FSO, File, TextStream, temp, s1, s2, s3
    Dim a, b, c As Double
    Dim a1, b1, c1 As String

    Sub Main()
        'On Error GoTo vse
        Dim nxt As String
        Dim arrN(1000) As Double
        Dim arrS(1000) As String
        nxt = vbNullString
        ApPat = Application.StartupPath
        FSO = CreateObject("Scripting.FileSystemObject")
        File = FSO.GetFile(ApPat + "\prog.ini")
        TextStream = File.OpenAsTextStream(1)
        temp = TextStream.ReadAll()
        Console.Title = temp
        TextStream.Close()

        File = FSO.GetFile(ApPat + "\prog.dat")
        TextStream = File.OpenAsTextStream(1)
        temp = TextStream.ReadLine()
        If UCase(temp) = "START" Then
            Do
             
                temp = TextStream.ReadLine()
                'If temp = nxt Then
                'nxt = vbNullString
                ' Continue Do
                'End If

                If UCase(temp) = "CONDOG" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        b = TextStream.ReadLine()
                        c = TextStream.ReadLine()
                        arrS(c) = arrS(b) + arrS(a)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "IFNXTN" Then
                        If UCase(temp) <> nxt Then
                            a = TextStream.ReadLine()
                            b = TextStream.ReadLine()
                            c1 = TextStream.ReadLine()
                            If arrN(a) = arrN(b) Then
                                'nxt = c1
                                s1 = Len(c1)
                                s2 = Left(c1, s1 - 1)
                                nxt = s2
                            End If
                        Else
                            nxt = vbNullString
                        End If
                    End If

                    If UCase(temp) = "IFNXTS" Then
                        If UCase(temp) <> nxt Then
                            a = TextStream.ReadLine()
                            b = TextStream.ReadLine()
                            c1 = TextStream.ReadLine()
                            If arrS(a) = arrS(b) Then
                                s1 = Len(c1)
                                s2 = Left(c1, s1 - 1)
                                nxt = s2
                            End If
                        Else
                            nxt = vbNullString
                        End If
                    End If

                    If UCase(temp) = "ADD" Then
                        If UCase(temp) <> nxt Then
                            a = TextStream.ReadLine()
                            b = TextStream.ReadLine()
                            c = TextStream.ReadLine()
                            arrN(c) = arrN(a) + arrN(b)
                        Else
                            nxt = vbNullString
                        End If
                    End If

                    If UCase(temp) = "SUB" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        b = TextStream.ReadLine()
                        c = TextStream.ReadLine()
                        arrN(c) = arrN(a) - arrN(b)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "MUL" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        b = TextStream.ReadLine()
                        c = TextStream.ReadLine()
                        arrN(c) = arrN(a) * arrN(b)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "DIV" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        b = TextStream.ReadLine()
                        c = TextStream.ReadLine()
                        arrN(c) = arrN(a) / arrN(b)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "PUTN" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        Console.WriteLine(arrN(a).ToString)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "PUTS" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        Console.WriteLine(arrS(a).ToString)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "GETN" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        arrN(a) = Console.ReadLine
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "GETS" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        arrS(a) = Console.ReadLine
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "APPN" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        b = TextStream.ReadLine()
                        arrN(a) = b
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "APPS" Then
                    If UCase(temp) <> nxt Then
                        a = TextStream.ReadLine()
                        b1 = TextStream.ReadLine()
                        arrS(a) = b1
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "OUTS" Then
                    If UCase(temp) <> nxt Then
                        a1 = TextStream.ReadLine()
                        Console.Write(a1)
                    Else
                        nxt = vbNullString
                    End If
                End If

                    If UCase(temp) = "OUTL" Then
                    If UCase(temp) <> nxt Then
                        a1 = TextStream.ReadLine()
                        Console.WriteLine(a1)
                    Else
                        nxt = vbNullString
                    End If
                End If

            Loop Until UCase(temp) = "STOP"
        Else : GoTo vse
        End If
vse:    Console.WriteLine()
        Console.Write("Press any key to quit.")
        Console.ReadKey()


    End Sub

End Module
