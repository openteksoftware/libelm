Public Class cLogger
    Public Enum LogType
        ERRMSG
        SERIAL
        NOTICE
        DEBUG
    End Enum

    Public Sub Log(ByVal sLogText As String)

    End Sub

    Public Sub Log(ByVal sLogText As String, ByVal eType As LogType)

    End Sub
End Class
