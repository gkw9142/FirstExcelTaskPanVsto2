Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports OpenFileDialog = System.Windows.Forms.OpenFileDialog
Imports zip = System.IO.Compression
Public Class Calendar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myStream As System.IO.Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog
        Dim ZipArchive As zip.ZipArchive = Nothing
        Dim ZipArchiveDataMashup As zip.ZipArchive = Nothing
        Dim StreamDataMashup As System.IO.Stream = Nothing

        Dim ZipEntryPowerQuery As zip.ZipArchiveEntry = Nothing

        openFileDialog1.InitialDirectory = "c:\"
        'openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        'openFileDialog1.FilterIndex = 2
        'openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    ' Insert code to read the stream here.
                    ZipArchive = New zip.ZipArchive(myStream, zip.ZipArchiveMode.Read)
                    Dim intLen As Int16 = ZipArchive.GetEntry("DataMashup").Length
                    StreamDataMashup = ZipArchive.GetEntry("DataMashup").Open
                    ''DataMashup 需要考虑如何偏移8位出来 StreamDataMashup.Seek(8, 0)
                    ''ZipArchiveDataMashup = New zip.ZipArchive(StreamDataMashup, zip.ZipArchiveMode.Read)

                    Dim ByteDataMashup(intLen) As Byte
                    StreamDataMashup.Read(ByteDataMashup, 0, intLen)

                End If
            Catch Ex As Exception
                System.Windows.Forms.MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Sub
    'Private Sub MonthCalendar1_DateChanged(sender As Object, e As Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
    '    Dim iSheet As Excel.Worksheet
    '    iSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add()

    '    iSheet.Range("A1").Value = e.Start.ToLongDateString

    'End Sub



End Class
