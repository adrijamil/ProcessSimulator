Public Class SimCase
    Public FlowSheets As New Collection
    Private myIO As New InputOutput
    Public Tags As New Collection



    ReadOnly Property IO As InputOutput
        Get
            IO = myIO
        End Get

    End Property



    Public Sub New()
        myIO.SimulationCase = Me
    End Sub

End Class
