Public Class JREInstallObject
    Dim app_name As String
    Dim version_number As String
    Dim uninstall_string As String

    Sub New(ByVal name As String, ByVal version As String, ByVal uninstall As String)

        app_name = name
        version_number = version
        uninstall_string = uninstall

    End Sub

    Public ReadOnly Property Version
        Get
            Return version_number
        End Get
    End Property
    Public ReadOnly Property Name
        Get
            Return app_name
        End Get
    End Property
    Public ReadOnly Property UninstallString
        Get
            Return uninstall_string
        End Get
    End Property
End Class
