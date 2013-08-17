Public Class cLogger
    Dim sLogData As String

    Public Enum LogType
        ERRMSG
        SERIAL
        NOTICE
        DEBUG
    End Enum

    Public Sub Log(ByVal sLogText As String)
        Log(sLogText, LogType.NOTICE)
    End Sub

    Public Sub Log(ByVal sLogText As String, ByVal eType As LogType)
        sLogData = sLogData & sLogText & vbNewLine
    End Sub

    Public Function getLogdata() As String
        Return sLogData
    End Function
End Class
