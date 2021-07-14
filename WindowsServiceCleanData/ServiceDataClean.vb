Imports System.IO
Imports DataLibraryClean

Public Class ServiceDataClean
    Private WithEvents myTimer As System.Timers.Timer
    Private configFile As String = "c:/DataCleaner/configArenas.txt"
    Private Const TimeParam As String = "TimeService"
    Private Const FileError As String = "c:/DataCleaner/logError.txt"
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        Dim dataClean As New DataCleaner

        Me.myTimer = New System.Timers.Timer()

        Me.myTimer.Enabled = True
        'ejecute cada X minutos ahora leído del archivo
        Dim timeMinutes As Integer = CInt(dataClean.getParam(configFile, TimeParam))
        Me.myTimer.Interval = 1000 * 60 * timeMinutes

        Me.myTimer.Start()
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub

    Protected Sub myTimer_Elapsed(ByVal sender As Object, e As EventArgs) Handles myTimer.Elapsed
        Try
            Dim dataClean As New DataCleaner
            dataClean.CleanData(configFile)
        Catch ex As Exception
            WriteErrorLog(ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub WriteErrorLog(ByVal mensaje As String, ByVal trace As String)
        Dim ruta As String = FileError
        Dim escritor As StreamWriter

        escritor = New StreamWriter(FileError)
        Dim fecha As DateTime = DateTime.Now
        escritor.WriteLine(fecha)
        escritor.WriteLine(mensaje)
        escritor.Write(trace)

        escritor.Flush()
        escritor.Close()
    End Sub

End Class
