Attribute VB_Name = "ConnectionDB"

Public conn As ADODB.Connection

Public Sub Connect()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=;Initial Catalog=;User ID=;Password=;"
    conn.Open
End Sub

Public Sub Disconnect()
    If Not conn Is Nothing Then
        conn.Close
        Set coon = Nothing
    End If
End Sub
