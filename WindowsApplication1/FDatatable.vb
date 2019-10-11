Imports Oracle.DataAccess.Client

Public Class FDatatable

    Dim dt As New DataTable
    Dim sql, str As String
    Dim edit As Boolean = False
    Dim connString As String = "DATA SOURCE=DESKTOP-KEM64H6.mshome.net:1521/XE;PERSIST SECURITY INFO=True;USER ID=HENDRO; Password=orcl"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadGrid()
    End Sub

    Sub loadGrid()
        Dim conn As New OracleConnection(connString) ' VB.NET

        Try
            conn.Open()

            Dim d = New DataTable
            Dim adp As New OracleDataAdapter("select * from jabatan", conn)
            adp.Fill(d)

            DataGridView1.DataSource = d
        Catch ex As Exception ' catches any error
            MessageBox.Show(ex.Message.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim conn As New OracleConnection(connString) ' VB.NET

        Try
            conn.Open()

            If edit = False Then
                str = "insert into jabatan values ('" & tb_id.Text.Trim & "', '" & tb_jabatan.Text.Trim & "', " & _
                    "" & Val(tb_gaji.Text.Trim) & ", " & Val(tb_tanak.Text.Trim) & ", " & Val(tb_tdinas.Text.Trim) & ",  " & _
                    "" & Val(tb_tkesehatan.Text.Trim) & ")"
            Else
                str = "update jabatan set jabatan =  '" & tb_jabatan.Text.Trim & "', gapok =  " & Val(tb_gaji.Text.Trim) & " , " & _
                    "tanak = " & Val(tb_tanak.Text.Trim) & ", tdinas = " & Val(tb_tdinas.Text.Trim) & ", " & _
                    "tkesehatan = " & Val(tb_tkesehatan.Text.Trim) & " where idjabatan = '" & tb_id.Text.Trim & "'"
            End If
            
            Dim cmd As New OracleCommand(str, conn)
            cmd.ExecuteNonQuery()

            loadGrid()
            clear()
        Catch ex As Exception ' catches any error
            MessageBox.Show(ex.Message.ToString())
        Finally
            conn.Close()
        End Try
    End Sub

    Sub clear()
        tb_id.Text = Nothing : tb_id.Enabled = True
        tb_jabatan.Text = Nothing
        tb_gaji.Text = Nothing
        tb_tanak.Text = Nothing
        tb_tdinas.Text = Nothing
        tb_tkesehatan.Text = Nothing
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        With DataGridView1
            Dim rowIndex As Integer = .CurrentCell.RowIndex
            tb_id.Text = .Rows(rowIndex).Cells(0).Value
            tb_id.Enabled = False
            tb_jabatan.Text = .Rows(rowIndex).Cells(1).Value
            tb_gaji.Text = .Rows(rowIndex).Cells(2).Value
            tb_tanak.Text = .Rows(rowIndex).Cells(3).Value
            tb_tdinas.Text = .Rows(rowIndex).Cells(4).Value
            tb_tkesehatan.Text = .Rows(rowIndex).Cells(5).Value
        End With

        edit = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If tb_id.Text.Trim <> "" Then Exit Sub

        If MsgBox("Hapus " & tb_jabatan.Text.Trim & "?", vbYesNo + vbCritical, "Perhatian") = vbYes Then
            Dim conn As New OracleConnection(connString) ' VB.NET

            Try
                conn.Open()
                str = "delete from jabatan where idjabatan  = '" & tb_id.Text.Trim & "'"
                Dim cmd As New OracleCommand(str, conn)
                cmd.ExecuteNonQuery()

                loadGrid()
                clear()
            Catch ex As Exception ' catches any error
                MessageBox.Show(ex.Message.ToString())
            Finally
                conn.Close()
            End Try
        End If
    End Sub
End Class
