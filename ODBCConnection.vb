Imports Microsoft.Win32

Public Class ODBCConnection
    Public Function CreateDSN(ByVal DB_Name As String, _
                            ByVal DSN As String, _
                            ByVal Description As String, _
                            ByVal Driver_Name As String, _
                            ByVal Driver_Path As String, _
                            ByVal Last_User As String, _
                            ByVal Server_Name As String, _
                            ByRef Status As String _
                            ) As Boolean

        Dim regHandle As RegistryKey

        Dim reg As RegistryKey = Registry.LocalMachine
        Dim regKey1 As String = "SOFTWARE\ODBC\ODBC.INI\" & DSN
        Dim regKey2 As String = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"

        Try
            regHandle = reg.CreateSubKey(regKey1)
            regHandle.SetValue("Database", DB_Name)
            regHandle.SetValue("Description", Description)
            regHandle.SetValue("Driver", Driver_Path)
            regHandle.SetValue("LastUser", Last_User)
            regHandle.SetValue("Server", Server_Name)
            regHandle.SetValue("Trusted_Connection", "Yes")
            regHandle.Close()
            reg.Close()

            regHandle = reg.CreateSubKey(regKey2)
            regHandle.SetValue(DSN, Driver_Name)
            regHandle.Close()
            reg.Close()
        Catch err As Exception
            MsgBox(err.Message)
        End Try
    End Function
End Class
