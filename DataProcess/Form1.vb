
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

    Private Property columnList As New List(Of String)

    'Private Property _excel As New Excel.Application

    'Private Property wbook As Excel.Workbook


   
    'UI of the form
    Private Sub Form1_Load() Handles Me.Load

        Label1.Visible = True
        Me.BackColor = Color.Azure
        Button1.BackColor = Color.BurlyWood
        Button2.BackColor = Color.Coral

    End Sub


    'This function is to import file after clicking a button
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
                'dt = ImportExcelToDataTable(System.IO.Path.GetFullPath(fpath))
                BackgroundWorker2.RunWorkerAsync()

                ' If dt.Rows.Count > 0 Then
                'DataView()
                ' End If


                'makeTable(dt)

            Else
                MsgBox("Please Enter Correct Value in Year, Risk Area and Crop")

        End If




        End If


    End Sub



    ' displaying data in datatable using a data gridview
    Private Sub DataView()
        'DataGridView1.DataSource = dt2
        ' ExportTable(dt, current_crop)
        ' MsgBox("ok")
        'wbook As Excel.Workbook
        'Test()
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

    ' Outputting the result in a matrix format and exporting output excel files
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        MsgBox("Hello from Test")
        columnList.Add("Canola")
        columnList.Add("Barley")
        columnList.Add("Oats")
        columnList.Add("Rye-F")
        columnList.Add("Wht-HRW")
        columnList.Add("Wht-DURUM")
        columnList.Add("Wheat-CPS")
        columnList.Add("Flax")
        columnList.Add("Lentils")
        columnList.Add("Faba Bean")
        columnList.Add("Pease,Field")
        columnList.Add("Wheat-ES")
        columnList.Add("Wheat-SWS")
        columnList.Add("Mustard-Ye")
        columnList.Add("Wht-HRS")
        BackgroundWorker1.ReportProgress(5)
        Dim _excel As New Excel.Application

        If _excel Is Nothing Then
            MsgBox("Excel is not installed properly")
        End If

        Dim wbook As Excel.Workbook
        'Dim wsheet As Excel.Worksheet

        Dim wsheet As Excel.Worksheet
        Dim i As Integer

        'BackgroundWorker1.ReportProgress(100)
        wbook = _excel.Workbooks.Add()

        wsheet = CType(_excel.Worksheets.Add(, , areas.Count() - 1), Excel.Worksheet)
        wsheet = wbook.ActiveSheet()


        'wsheet = CType(_excel.Worksheets.Add(, , areas.Count() - 1), Excel.Worksheet)

        Dim dt1 As System.Data.DataTable = dt

        Dim dc As System.Data.DataColumn


        Dim dr As System.Data.DataRow
        Dim drows As New List(Of Single)


        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        _excel.ScreenUpdating = False
        _excel.DisplayAlerts = False


        Dim p As Integer

        For i = 0 To areas.Count() - 1 Step 1
            wsheet = wbook.Worksheets(i + 1)
            colIndex = 1
            For p = 0 To columnList.Count() - 1 Step 1
                colIndex = colIndex + 1
                wsheet.Cells(1, colIndex) = columnList(p)

            Next
            wsheet.Columns.AutoFit()

        Next

        For i = 0 To areas.Count() - 1 Step 1
            wsheet = wbook.Worksheets(i + 1)
            colIndex = 1
            rowIndex = 1
            For q = 2001 To 2018 Step 1
                rowIndex = rowIndex + 1
                wsheet.Cells(rowIndex, colIndex) = q

            Next
        Next
        BackgroundWorker1.ReportProgress(45)
        'System.Threading.Thread.Sleep(1000)
        Dim n As Integer = 0
        Dim t As Integer
        Dim yr As Integer
        Dim row_val As Integer = 1
        Dim col_val As Integer
        Dim queryResults
        'wsheet.Name = "RA" & Form1.areas(0).ToString
        MsgBox(areas.Count())
        For r = 0 To areas.Count() - 1
            wsheet = wbook.Worksheets(r + 1)
            wsheet.Name = "RA" & areas(r)
            For yr = 2001 To 2018 Step 1
                row_val = row_val + 1
                col_val = 1
                Dim m As Integer = areas(r)
                'wsheet = wbook.Worksheets(1)
                For t = 0 To columnList.Count() - 1 Step 1

                    col_val = col_val + 1
                    queryResults = From ry In dt1.AsEnumerable
                              Where (ry("year") = yr And ry("prev_crop") = columnList(t) And ry("curr_crop") = current_crop And ry("RA") = m)
                              Select ry

                    For Each result In queryResults
                        wsheet.Cells(row_val, col_val) = result("relative_yld") '2nd issue

                    Next


                Next
                '   MsgBox(dr("RA"))
                ' wsheet.Name = dr("RA").ToString
            Next
            row_val = 1

        Next

        wsheet.Columns.AutoFit()
        BackgroundWorker1.ReportProgress(80)

        Dim strFileName As String = crop_name & ".xlsx"
      
        'BackgroundWorker1.ReportProgress(100)
        wbook.SaveAs(strFileName)
        wbook.Close()
        _excel.Quit()
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wsheet)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbook)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        areas.Clear()
        columnList.Clear()
        BackgroundWorker1.ReportProgress(100)


    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        'System.Threading.Thread.Sleep(100)
        ProgressBar1.Value = e.ProgressPercentage

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        'set datasource of DatagridView
        'ProgressBar1.Style = ProgressBarStyle.Continuous
        MsgBox("File Exported")
        ProgressBar1.Value = 0
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        current_crop = Nothing
        dt.Reset()


    End Sub

    'Running a thread in the background to run query on the files and saving into data table
    Public Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
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
            Dim constring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fpath & ";Extended Properties=""Excel 12.0 xml;HDR=YES;IMEX=1;"""
            Dim con As New OleDbConnection(constring & "")
            con.Open()

            Dim myTableName = con.GetSchema("Tables").Rows(0)("TABLE_NAME")
            cmnd2.CommandText = String.Format("SELECT DISTINCT RA FROM [{0}] WHERE curr_crop = '" & current_crop & "' AND prev_crop ='Wht-HRS' AND year BETWEEN " & y & " AND " & 2018 & " AND RA BETWEEN " & ra & " AND " & 22, myTableName)
            cmnd2.Connection = con
            Dim dr = cmnd2.ExecuteReader()

            While dr.Read()
                areas.Add(dr("RA"))
            End While


            'MsgBox(areas.Count())
            If areas.Count() > 0 Then

                For r = 0 To areas.Count() - 1 Step 1
                    ra = areas(r)
                    ' MsgBox(ra)
                    For i = y To 2018 Step 1
                        cmnd.CommandText = String.Format("SELECT ave_yld FROM [{0}] WHERE prev_crop='Wht-HRS' AND curr_crop='" & current_crop & "' AND year=" & i & " AND RA=" & ra, myTableName)
                        cmnd.Connection = con
                        n = cmnd.ExecuteScalar()
                        If n = 0 Then
                            ' GoTo out_for
                            Continue For

                        End If

                        Dim sqlquery1 As String = String.Format("SELECT RA, year,prev_crop,curr_crop, ave_yld, [{0}].ave_yld/" & n & " AS relative_yld FROM [{0}] WHERE RA=" & ra & " AND year = " & i & " AND [{0}].ave_yld IN (SELECT [{0}].ave_yld FROM [{0}] WHERE curr_crop='" & current_crop & "' AND year=" & i & " AND RA=" & ra & " AND prev_code <17) ORDER BY curr_crop, year", myTableName)
                        Dim da1 As New OleDbDataAdapter(sqlquery1, con)

                        'da.Fill(ds)
                        da1.Fill(rowColumn)


                    Next
                    'out_for:

                    'MsgBox("ok")
                    If rowColumn.Tables.Count > 0 Then
                        If rowColumn.Tables(0).Rows.Count > 0 Then
                            dt = rowColumn.Tables(0)
                            'MsgBox(dt.Rows.Count)

                        End If

                    Else
                        Continue For
                    End If



                Next

                con.Close()

                'Return dt


            Else
                MsgBox("No relevant data found")
                con.Close()
                ' Return dt

            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical)
            'Return dt
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted

        'set datasource of DatagridView
        'ProgressBar1.Style = ProgressBarStyle.Continuous

        'MsgBox(dt.Rows.Count)

        If dt.Rows.Count > 0 Then
            DataView()
        End If


    End Sub

    'A test function to test the output module
    Private Sub Test()
        MsgBox("Hello from Test")
        columnList.Add("Canola")
        columnList.Add("Barley")
        columnList.Add("Oats")
        columnList.Add("Rye-F")
        columnList.Add("Wht-HRW")
        columnList.Add("Wht-DURUM")
        columnList.Add("Wheat-CPS")
        columnList.Add("Flax")
        columnList.Add("Lentils")
        columnList.Add("Faba Bean")
        columnList.Add("Pease,Field")
        columnList.Add("Wheat-ES")
        columnList.Add("Wheat-SWS")
        columnList.Add("Mustard-Ye")
        columnList.Add("Wht-HRS")

        Dim _excel As New Excel.Application
        If _excel Is Nothing Then
            MsgBox("Excel is not installed properly")
        End If

        Dim wbook As Excel.Workbook
        'Dim wsheet As Excel.Worksheet

        Dim wsheet As Excel.Worksheet
        Dim i As Integer

        'BackgroundWorker1.ReportProgress(100)
        wbook = _excel.Workbooks.Add()

        wsheet = CType(_excel.Worksheets.Add(, , areas.Count() - 1), Excel.Worksheet)
        wsheet = wbook.ActiveSheet()


        'wsheet = CType(_excel.Worksheets.Add(, , areas.Count() - 1), Excel.Worksheet)

        Dim dt1 As System.Data.DataTable = dt

        Dim dc As System.Data.DataColumn


        Dim dr As System.Data.DataRow
        Dim drows As New List(Of Single)
        
        ' Dim queryResults = From ry In dt1.AsEnumerable
        'Where(ry("year") = 2017 And ry("prev_crop") = columnList(0) And ry("curr_crop") = current_crop)
        '                  Select ry

        ' For Each result In queryResults
        'MsgBox(result("relative_yld"))

        'Next


        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        _excel.ScreenUpdating = False
        _excel.DisplayAlerts = False

        Dim p As Integer

        For i = 0 To areas.Count() - 1 Step 1
            wsheet = wbook.Worksheets(i + 1)
            colIndex = 1
            For p = 0 To columnList.Count() - 1 Step 1
                colIndex = colIndex + 1
                wsheet.Cells(1, colIndex) = columnList(p)

            Next
            wsheet.Columns.AutoFit()

        Next

        For i = 0 To areas.Count() - 1 Step 1
            wsheet = wbook.Worksheets(i + 1)
            colIndex = 1
            rowIndex = 1
            For q = 2001 To 2018 Step 1
                rowIndex = rowIndex + 1
                wsheet.Cells(rowIndex, colIndex) = q

            Next
        Next

        Dim n As Integer = 0
        Dim t As Integer
        Dim yr As Integer
        Dim row_val As Integer = 1
        Dim col_val As Integer
        Dim queryResults
        'wsheet.Name = "RA" & Form1.areas(0).ToString
        MsgBox(areas.Count())
        For r = 0 To areas.Count() - 1
            wsheet = wbook.Worksheets(r + 1)
            wsheet.Name = "RA" & areas(r)
            For yr = 2001 To 2018 Step 1
                row_val = row_val + 1
                col_val = 1
                Dim m As Integer = areas(r)
                'wsheet = wbook.Worksheets(1)
                For t = 0 To columnList.Count() - 1 Step 1

                    col_val = col_val + 1
                    queryResults = From ry In dt1.AsEnumerable
                              Where (ry("year") = yr And ry("prev_crop") = columnList(t) And ry("curr_crop") = current_crop And ry("RA") = m)
                              Select ry

                    For Each result In queryResults
                            wsheet.Cells(row_val, col_val) = result("relative_yld") '2nd issue
                        
                    Next


                Next
                '   MsgBox(dr("RA"))
                ' wsheet.Name = dr("RA").ToString
            Next
            row_val = 1

        Next
        '  BackgroundWorker1.ReportProgress(100, "Complete!"

        wsheet.Columns.AutoFit()
        Dim strFileName As String = crop_name & "sakib" & ".xlsx"
        If System.IO.File.Exists(strFileName) Then
            System.IO.File.Delete(strFileName)
        End If
        'BackgroundWorker1.ReportProgress(100)
        wbook.SaveAs(strFileName)
        wbook.Close()
        _excel.Quit()

        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wsheet)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wbook)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excel)
        GC.Collect()
        GC.WaitForPendingFinalizers()







    End Sub
End Class
