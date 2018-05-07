Public Class ListAllFile

    Dim sSourceFolderpath As String
    Dim sFileArray(0) As String
    Dim sInt As Integer = 0
    Dim sFileInformation(0) As String
    Dim sFileInfoInt As Integer = 1

    Private Sub GetFileFolder(ByVal sSourceFolderpath)

        Dim sFile As String() = System.IO.Directory.GetFiles(sSourceFolderpath)
        For Each Filename In sFile
            ReDim Preserve sFileArray(sInt)
            sFileArray(sInt) = Filename
            sInt = sInt + 1
        Next

        Dim sFolder As String() = System.IO.Directory.GetDirectories(sSourceFolderpath)
        For Each Folder In sFolder
            GetFileFolder(Folder)
        Next Folder
    End Sub

    Private Sub GetAllFileInfo(ByVal sFileArray As String())

        sFileInformation(0) = "FileName, Extension, Size, Byte, Folder Path"

        For Each tempFile In sFileArray

            If Not IsNothing(tempFile) Then

                Try

                    Dim sFileinfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(tempFile)
                    ReDim Preserve sFileInformation(sFileInfoInt)

                    sFileInformation(sFileInfoInt) = sFileinfo.Name.Replace(",", "_") & "," & _
                                                        sFileinfo.Extension & "," & _
                                                        GetSize(sFileinfo.Length) & "," & _
                                                        sFileinfo.DirectoryName.Replace(",", "_")
                    sFileInfoInt = sFileInfoInt + 1

                Catch ex As System.IO.PathTooLongException

                    ReDim Preserve sFileInformation(sFileInfoInt)

                    sFileInformation(sFileInfoInt) = tempFile.Split("\")(tempFile.Split("\").Length - 1).Replace(",", "_") & "," & "." & _
                                                        tempFile.Split("\")(tempFile.Split("\").Length - 1).Replace(",", "_").Split(".")(tempFile.Split("\")(tempFile.Split("\").Length - 1).Replace(",", "_").Split(".").Length - 1) & "," & _
                                                        "PathTooLongException, PathTooLongException" & "," & tempFile.Replace(",", "_")
                    sFileInfoInt = sFileInfoInt + 1

                Catch ex As Exception

                    MessageBox.Show(ex.Message)

                End Try

            End If

        Next

    End Sub

    Public Function GetSize(ByVal size As Long) As String
        Dim sSize As String = ","
        Select Case size
            Case 0 To 1024
                sSize = size & "," & "Byte"
            Case 1024 To 1048576
                sSize = Format(size / 1024, "0.00").ToString() & "," & "KB"
            Case 1048576 To 1073741824
                sSize = Format(size / 1048576, "0.00").ToString() & "," & "MB"
            Case 1073741824 To 1099511627776
                sSize = Format(size / 1073741824, "0.00").ToString() & "," & "GB"
            Case Else
                sSize = "Size capacity more than,GB"
        End Select
        Return sSize
    End Function

    Private Sub CreateCSV()
        Dim SaveFileDialog1 As New SaveFileDialog
        SaveFileDialog1.Filter = "CSV File|*.csv"
        SaveFileDialog1.FileName = "ListAllFile" & "_" & DateTime.Now.Year & "_" & DateTime.Now.Month & "_" & DateTime.Now.Day & "_" & DateTime.Now.Hour & "_" & DateTime.Now.Minute & "_" & DateTime.Now.Second
        SaveFileDialog1.Title = "Save a CSV File"
        SaveFileDialog1.RestoreDirectory = True
        Me.Show()
        SaveFileDialog1.ShowDialog()

        If SaveFileDialog1.FileName.Contains(".csv") Then
            Dim stramwriter As System.IO.StreamWriter = System.IO.File.CreateText(SaveFileDialog1.FileName)

            For Each sRecord In sFileInformation
                stramwriter.WriteLine(sRecord)
            Next

            stramwriter.Flush()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        GetFileFolder(sSourceFolderpath)
        GetAllFileInfo(sFileArray)
        CreateCSV()
        Me.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim FolderBrowserDialog As New FolderBrowserDialog()
        FolderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop
        FolderBrowserDialog.SelectedPath = Environment.SpecialFolder.Desktop
        FolderBrowserDialog.Description = "Select Folder"
        If FolderBrowserDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            sSourceFolderpath = FolderBrowserDialog.SelectedPath
        End If

    End Sub
End Class
