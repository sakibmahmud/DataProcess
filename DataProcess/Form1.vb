
Imports System.Data.OleDb
Imports System.Data.OleDb.OleDbCommand
Imports System.Data.SqlClient
Imports System.Net.Mime.MediaTypeNames
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Linq
Imports Microsoft
Imports Microsoft.Office.Interop
Imports System.ComponentModel



Public Class Form1

    Private Property current_crop As String
    Private Property areas As New List(Of Integer)
    'Private Property dt As New DataTable
    Private Property fpath As String

    Private Property crop_name As String

    Private Property dt As New DataTable
    Private Property ra As Integer

    Private Property y As Integer





    Private Sub Form1_Load() Handles Me.Load

        Label1.Visible = True
        Me.BackColor = Color.Azure
        Button1.BackColor = Color.BurlyWood
        Button2.BackColor = Color.Coral

    End Sub


    'This function is to import file and save the filepath
    Private Sub Btn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim dialog As New OpenFileDialog
        dialog.Filter = "Excel files |*.xls;*.xlsx"
        dialog.InitialDirectory = "C:\Documents"
        dialog.RestoreDirectory = True
        Try
            If dialog.ShowDialog() = DialogResult.OK Then
                ' Dim dt1 As New DataTable
                fpath = dialog.FileName.ToString
                'MsgBox(fpath)

                MsgBox("File Successfully loaded")

            End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical)
        End Try


    End Sub

    'Taking input here from relevant textboxes
    Private Sub DataInput()
        ' Dim dt1 As New DataTable


        If TextBox1.Text = "" Then
            MsgBox("Please Enter Risk Area")

        ElseIf TextBox2.Text = "" Then
            MsgBox("Please Enter Year")

        ElseIf TextBox3.Text = "" Then
            MsgBox("Please Enter Crop")

        Else

            If Integer.TryParse(TextBox1.Text, vbNull) And Integer.TryParse(TextBox2.Text, vbNull) And System.Text.RegularExpressions.Regex.IsMatch(TextBox3.Text, "^[A-Za-z]+$") Then

                ra = TextBox1.Text
                y = TextBox2.Text
                current_crop = TextBox3.Text
                crop_name = current_crop
                current_crop = TextBox3.Text.Substring(0, 1).ToUpper() + TextBox3.Text.Substring(1).ToLower()
                dt = ImportExcelToDataTable(System.IO.Path.GetFullPath(fpath))
                If dt.Rows.Count > 0 Then
                    DataView()
                End If

                'makeTable(dt)

            Else
                MsgBox("Please Enter Correct Value in Year, Risk Area and Crop")


            End If



        End If


    End Sub



    Public Shared Function ImportExcelToDataTable(ByVal filepath As String) As DataTable
        ' Dim dt As New DataTable
        Dim rowColumn As New DataSet()
        'MsgBox(current_crop)
        Try
            Dim n As Single
            Dim i As Integer
            Dim r As Integer
            ' Dim areas As New List(Of Integer)
            'Dim yrs As New List(Of Integer)
            Dim cmnd As New OleDbCommand
            Dim cmnd2 As New OleDbCommand
            'Dim cmnd3 As New OleDbCommand
            Dim ds As New DataSet()
            Dim constring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filepath & ";Extended Properties=""Excel 12.0 xml;HDR=YES;IMEX=1;"""
            Dim con As New OleDbConnection(constring & "")
            con.Open()

            Dim myTableName = con.GetSchema("Tables").Rows(0)("TABLE_NAME")
            cmnd2.CommandText = String.Format("SELECT DISTINCT RA FROM [{0}] WHERE curr_crop = '" & Form1.current_crop & "' AND prev_crop ='Wht-HRS' AND year BETWEEN " & Form1.y & " AND " & 2018 & " AND RA BETWEEN " & Form1.ra & " AND " & 22, myTableName)
            cmnd2.Connection = con
            Dim dr = cmnd2.ExecuteReader()

            While dr.Read()
                Form1.areas.Add(dr("RA"))
            End While


            ' MsgBox(Form1.areas.Count())
            If Form1.areas.Count() > 0 Then

                For r = 0 To Form1.areas.Count() - 1 Step 1
                    Form1.ra = Form1.areas(r)
                    ' MsgBox(ra)
                    For i = Form1.y To 2018 Step 1
                        cmnd.CommandText = String.Format("SELECT ave_yld FROM [{0}] WHERE prev_crop='Wht-HRS' AND curr_crop='" & Form1.current_crop & "' AND year=" & i & " AND RA=" & Form1.ra, myTableName)
                        cmnd.Connection = con
                        n = cmnd.ExecuteScalar()
                        If n = 0 Then
                            ' GoTo out_for
                            Continue For

                        End If

                        Dim sqlquery1 As String = String.Format("SELECT RA, year,prev_crop,curr_crop, ave_yld, [{0}].ave_yld/" & n & " AS relative_yld FROM [{0}] WHERE RA=" & Form1.ra & " AND year = " & i & " AND [{0}].ave_yld IN (SELECT [{0}].ave_yld FROM [{0}] WHERE curr_crop='" & Form1.current_crop & "' AND year=" & i & " AND RA=" & Form1.ra & " AND prev_code <17) ORDER BY curr_crop, year", myTableName)
                        Dim da1 As New OleDbDataAdapter(sqlquery1, con)

                        'da.Fill(ds)
                        da1.Fill(rowColumn)


                    Next
                    'out_for:

                    'MsgBox("ok")
                    If rowColumn.Tables.Count > 0 Then
                        If rowColumn.Tables(0).Rows.Count > 0 Then
                            Form1.dt = rowColumn.Tables(0)
                        End If

                    Else
                        Continue For
                    End If



                Next

                con.Close()
                Return Form1.dt


            Else
                MsgBox("No relevant data found")
                con.Close()
                Return Form1.dt

            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical)
            Return Form1.dt
        End Try
    End Function


    ' displaying data in datatable using a data gridview
    Private Sub DataView()
        'DataGridView1.DataSource = dt2
        ' ExportTable(dt, current_crop)
        BackgroundWorker1.RunWorkerAsync()

    End Sub

   
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged


    End Sub


    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub


    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        MyBase.OnPaint(e)
        Dim borderWidth As Integer = 1
        Dim theColor As Color = Color.BurlyWood
        ControlPaint.DrawBorder(e.Graphics, e.ClipRectangle, theColor, borderWidth, ButtonBorderStyle.Solid, theColor, borderWidth, ButtonBorderStyle.Solid, theColor, borderWidth, ButtonBorderStyle.Solid, theColor, borderWidth, ButtonBorderStyle.Solid)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            DataInput()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical)
        End Try

    End Sub


    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BackgroundWorker1.ReportProgress(5)
        Dim _excel As New Excel.Application

        If _excel Is Nothing Then
            MsgBox("Excel is not installed properly")
        End If

        Dim wbook As Excel.Workbook

        Dim wsheet As Excel.Worksheet

        Dim i As Integer

        wbook = _excel.Workbooks.Add()


        wsheet = CType(_excel.Worksheets.Add(, , areas.Count() - 1), Excel.Worksheet)
        wsheet = wbook.ActiveSheet()




        Dim dt1 As System.Data.DataTable = dt

        Dim dc As System.Data.DataColumn

        Dim dr As System.Data.DataRow

        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        _excel.ScreenUpdating = False
        _excel.DisplayAlerts = False


        For i = 0 To areas.Count() - 1 Step 1
            wsheet = wbook.Worksheets(i + 1)
            colIndex = 0
            For Each dc In dt1.Columns
                colIndex = colIndex + 1
                wsheet.Cells(1, colIndex) = dc.ColumnName

            Next

        Next
        BackgroundWorker1.ReportProgress(45)
        'System.Threading.Thread.Sleep(1000)
        Dim n As Integer = 0

        'wsheet.Name = "RA" & Form1.areas(0).ToString
        For Each dr In dt1.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            wsheet = wbook.Worksheets(1)
            For Each dc In dt1.Columns
                colIndex = colIndex + 1
                If areas(n) = dr("RA") Then
                    wsheet = wbook.Worksheets(n + 1)
                    wsheet.Name = "RA" & dr("RA")
                    wsheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Else

                    n = n + 1
                    rowIndex = 1

                End If

            Next
            '   MsgBox(dr("RA"))
            ' wsheet.Name = dr("RA").ToString
        Next
        '  BackgroundWorker1.ReportProgress(100, "Complete!"

        wsheet.Columns.AutoFit()
        BackgroundWorker1.ReportProgress(80)

        Dim strFileName As String = crop_name & ".xlsx"
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If
        'BackgroundWorker1.ReportProgress(100)
        wbook.SaveAs(strFileName)
        BackgroundWorker1.ReportProgress(100)
        wbook.Close()
        _excel.Quit()
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wsheet)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbook)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel)
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        'System.Threading.Thread.Sleep(100)
        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        'set datasource of DatagridView
        'ProgressBar1.Style = ProgressBarStyle.Continuous
        MsgBox("File Exported")

    End Sub


End Class
