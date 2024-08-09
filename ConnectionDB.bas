Attribute VB_Name = "ConnectionDB"

Public conn As ADODB.Connection

Public Sub Connect()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=LAPTOP-ABU5HG8F;Initial Catalog=HubOfReading;User ID=jonhy-d;Password=Merry01.;"
    conn.open
End Sub

Public Sub Disconnect()
    If Not conn Is Nothing Then
        conn.Close
        Set coon = Nothing
    End If
End Sub
