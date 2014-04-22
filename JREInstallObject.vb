Public Class JREInstallObject

    'Instance variables
    Dim app_name As String
    Dim version_number As String
    Dim uninstall_string As String
    Dim _installed As Boolean = True

    'Constructor to create an instance of the JREInstallObject class
    Sub New(ByVal name As String, ByVal version As String, ByVal uninstall As String)

        app_name = name
        version_number = version
        uninstall_string = uninstall

    End Sub

    'Return the version number
    Public ReadOnly Property Version As String
        Get
            Return version_number
        End Get
    End Property

    'Return the Application Name
    Public ReadOnly Property Name As String
        Get
            Return app_name
        End Get
    End Property

    'Get the MSI uninstall string
    Public ReadOnly Property UninstallString As String
        Get
            Return uninstall_string
        End Get
    End Property

    'Allow an object to be marked as uninstalled without removing it from the collection
    Public Property Installed As Boolean
        Get
            Return _installed
        End Get
        Set(value As Boolean)
            _installed = value
        End Set
    End Property

End Class
