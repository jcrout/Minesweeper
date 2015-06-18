Public Class SafeFile

    Public ErrorTitle As String
    Public ErrorStringStart As String

    Public Sub New(Optional ByRef sTitle As String = "stop messing up!", Optional ByRef sErrorStringStartPhrase As String = "Error: ")
        ErrorTitle = sTitle
        ErrorStringStart = sErrorStringStartPhrase
    End Sub

    Public Function CheckFileExists(ByVal sFile As String, Optional ByVal blnShowError As Boolean = False) As Boolean
        Try
            If System.IO.File.Exists(sFile) Then Return True
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not check file exists for file " & sFile & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not find file " & Chr(34) & sFile & Chr(34), ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return False
    End Function

    Public Function DeleteFile(ByVal sFile As String, Optional ByVal blnShowError As Boolean = True) As Boolean
        Try
            If System.IO.File.Exists(sFile) = True Then
                System.IO.File.Delete(sFile)
                Return True
            Else
                'If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not find file " & sFile, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not delte file " & Chr(34) & sFile & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Function RenameFile(ByVal sOldFileName As String, ByVal sNewFileName As String, Optional ByRef blnHandleSameFileName As Boolean = False, Optional ByVal blnShowError As Boolean = True) As Boolean
        Try
            If sNewFileName.Contains("\") Then sNewFileName = sNewFileName.Substring(sNewFileName.LastIndexOf("\") + 1, sNewFileName.Length - sNewFileName.LastIndexOf("\"))
            If System.IO.File.Exists(sNewFileName) Then
                If blnHandleSameFileName = True Then
                    Dim BaseName As String = sNewFileName.Substring(0, sNewFileName.LastIndexOf("."))
                    Dim ExtName As String = sNewFileName.Substring(sNewFileName.LastIndexOf("."), sNewFileName.Length - sNewFileName.LastIndexOf("."))
                    Dim i As Short = 2
                    Do
                        If i = 10000 Then Exit Do
                        If System.IO.File.Exists(BaseName & " (" & i & ")" & ExtName) = False Then
                            sNewFileName = BaseName & " (" & i & ")" & ExtName
                            Exit Do
                        Else
                            i += 1
                        End If
                    Loop
                Else
                    If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not rename file " & Chr(34) & sOldFileName & Chr(34) & " to " & Chr(34) & sNewFileName & Chr(34) & " because a file with that name already exists.", ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
            End If
            My.Computer.FileSystem.RenameFile(sOldFileName, sNewFileName)
            Return True
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not rename file " & Chr(34) & sOldFileName & Chr(34) & " to " & Chr(34) & sNewFileName & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Function SW_CreateText(ByVal sFile As String, ByRef sw As System.IO.StreamWriter, Optional ByVal blnDeletePrevFile As Boolean = True, Optional ByVal blnShowError As Boolean = True) As System.IO.StreamWriter
        Try
            If blnDeletePrevFile = True Then If DeleteFile(sFile, blnShowError) = False Then Return Nothing
            sw = System.IO.File.CreateText(sFile)
            Return sw
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not create text file from streamwriter " & Chr(34) & sFile & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function SR_OpenText(ByVal sFile As String, ByRef sr As System.IO.StreamReader, Optional ByVal blnShowError As Boolean = True) As System.IO.StreamReader
        Try
            sr = System.IO.File.OpenText(sFile)
            Return sr
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not open text file from streamreader " & Chr(34) & sFile & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function GetArrayLinesFromStream(ByVal sFile As String, Optional ByRef blnRemoveEmptyLines As Boolean = False, Optional ByVal blnTrim As Boolean = False, Optional ByVal blnSort As Boolean = False, Optional ByVal blnShowError As Boolean = True) As List(Of String)
        Try
            Dim arrTemp As New List(Of String)
            Using sr As System.IO.StreamReader = System.IO.File.OpenText(sFile)
                Do While sr.Peek <> -1
                    Dim sLine As String = sr.ReadLine
                    If blnRemoveEmptyLines = True Then If sLine = "" Then Continue Do
                    If blnTrim = True Then arrTemp.Add(sLine.Trim) Else arrTemp.Add(sLine)
                Loop
            End Using
            If blnSort = True Then arrTemp.Sort()
            Return arrTemp
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not read file " & Chr(34) & sFile & Chr(34) & " in streamreader." & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

End Class