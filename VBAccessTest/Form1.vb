Imports System.Data.OleDb
Public Class Form1
    Dim conn As New OleDb.OleDbConnection
    Dim dbProvider As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
    Dim dbSource As String = "Data Source = D:\AccessDB\Test.mdb"
    Dim adapter As OleDb.OleDbDataAdapter
    Dim ds As New DataSet
    Dim currentid As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ds = New DataSet
        adapter = New OleDbDataAdapter("insert into [tbl_accounts] ([username], [password]) VALUES ('" &
                                       TextBox1.Text & "','" & TextBox2.Text & "')", conn)
        adapter.Fill(ds, "tbl_accounts")

        GetRecords()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = dbProvider & dbSource
        GetRecords()
    End Sub

    Public Sub GetRecords()
        ds = New DataSet

        conn.Open()
        adapter = New OleDbDataAdapter("SELECT * FROM tbl_accounts", conn)
        adapter.Fill(ds, "tbl_accounts")

        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "tbl_accounts"

        conn.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        currentid = DataGridView1.Item(0, i).Value.ToString()

        ds = New DataSet
        adapter = New OleDbDataAdapter("DELETE FROM tbl_accounts WHERE id = " & currentid, conn)
        adapter.Fill(ds, "tbl_accounts")

        GetRecords()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index

        currentid = DataGridView1.Item(0, i).Value.ToString()
        TextBox1.Text = DataGridView1.Item(1, i).Value.ToString()
        TextBox2.Text = DataGridView1.Item(2, i).Value.ToString()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ds = New DataSet
        adapter = New OleDbDataAdapter("UPDATE [tbl_accounts] SET [username] = '" & TextBox1.Text &
                                       "', [password] = '" & TextBox2.Text &
                                       "' where id = " & currentid, conn)
        adapter.Fill(ds, "tbl_accounts")
        TextBox1.Clear()
        TextBox2.Clear()

        GetRecords()
    End Sub
End Class
