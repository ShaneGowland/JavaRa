Option Strict Off
Imports Microsoft.Win32
Module routines_registry

    'Delete a registry key
    Public Sub delete_key(ByVal key As String)
        'Check if the key exists first, reducing processing time and potential for errors.
        If RegKeyExists(key) = True Then
            Try

                'Establish the subkey name
                Dim subkey_name As String
                subkey_name = key.Remove(0, key.LastIndexOf("\"))
                subkey_name = subkey_name.Trim("\")

                'Establish which hive we are working with
                Dim reg_hive As String = "NULL"
                Dim parent_key As String = "NULL"

                'Establish registry hive
                If key.StartsWith("HKLM\") Then
                    reg_hive = "HKLM"
                    parent_key = key.Replace("HKLM\", "")
                    parent_key = parent_key.Replace(subkey_name, "")
                    parent_key = parent_key.Trim("\")
                End If

                'Classes root conditional
                If key.StartsWith("HKCR\") Then
                    reg_hive = "HKCR"
                    parent_key = key.Replace("HKCR\", "")
                    parent_key = parent_key.Replace(subkey_name, "")
                    parent_key = parent_key.Trim("\")
                End If

                'Current Users hive conditional
                If key.StartsWith("HKCU\") Then
                    reg_hive = "HKCU"
                    parent_key = key.Replace("HKCU\", "")
                    parent_key = parent_key.Replace(subkey_name, "")
                    parent_key = parent_key.Trim("\")
                End If

                'HKEY USERS hive conditional
                If key.StartsWith("HKUS\") Then
                    reg_hive = "HKUS"
                    parent_key = key.Replace("HKUS\", "")
                    parent_key = parent_key.Replace(subkey_name, "")
                    parent_key = parent_key.Trim("\")
                End If

                'Assign the registry key variable. Set a default value to prevent compiler warning.
                Dim regKey As RegistryKey : regKey = Registry.LocalMachine.OpenSubKey("NULL", True)

                'Set for HKEY_LOCAL_MACHINE
                If reg_hive = "HKLM" Then
                    regKey = Registry.LocalMachine.OpenSubKey(parent_key, True)
                End If

                'Set for HKEY_CLASSES_ROOT
                If reg_hive = "HKCR" Then
                    regKey = Registry.ClassesRoot.OpenSubKey(parent_key, True)
                End If

                'Set for HKEY_CURRENT_USERS
                If reg_hive = "HKCU" Then
                    regKey = Registry.CurrentUser.OpenSubKey(parent_key, True)
                End If

                'Set for HKEY_USERS
                If reg_hive = "HKUS" Then
                    regKey = Registry.Users.OpenSubKey(parent_key, True)
                End If

                'Run the deletion routine
                Try
                    regKey.DeleteSubKey(subkey_name, True)
                    write_log(get_string("Removed registry subkey:") & " " & subkey_name)
                    UI.removal_count = UI.removal_count + 1
                Catch ex As InvalidOperationException 'If it contains children, must throw more powerful function.
                    regKey.DeleteSubKeyTree(subkey_name)
                    write_log(get_string("Removed registry subkey tree:") & " " & subkey_name)
                    UI.removal_count = UI.removal_count + 1
                End Try

                'Close the loop
                regKey.Close()

            Catch ex As Exception
                write_error(ex)
            End Try
        Else
            'Registry key doesn't exist
        End If
    End Sub

    'Create a text file containing a list of all installed JRE versions
    Public Sub output_jre_version()
        'Define the path for the log
        Dim version_output_path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\version-output.log"

        'Delete any previous instances of this file
        DeleteIfPermitted(version_output_path)

        'Declare the textwriter   
        Try
            Dim SW As IO.TextWriter
            SW = IO.File.AppendText(version_output_path)

            SW.WriteLine("Installed JRE Versions:" & Environment.NewLine & "========================")


            'Loop through all installed JREs
            For Each InstalledJRE As JREInstallObject In UI.JREObjectList
                Try
                    If InstalledJRE.Installed Then
                        SW.WriteLine(InstalledJRE.Name & " version: " & InstalledJRE.Version)
                    End If
                Catch ex As Exception
                    write_error(ex)
                End Try
            Next

            SW.Close()

            'show in notepad
            Process.Start(version_output_path)

        Catch ex As Exception
            write_error(ex)
        End Try
    End Sub

    'Remove all startup entries corresponding to the JRE
    Public Sub delete_jre_startup_entries()
        Dim startup_reg As New List(Of String) : startup_reg.Add("jusched-Java Quick Start") : startup_reg.Add("SunJavaUpdateSched") : startup_reg.Add("Java(tm) Plug-In 2 SSV Helper")
        For Each item As String In startup_reg
            'Delete the registry keys
            On Error Resume Next
            My.Computer.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(item)
            On Error Resume Next
            My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(item)
            On Error Resume Next
            My.Computer.Registry.LocalMachine.OpenSubKey("Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(item)
            On Error Resume Next
            My.Computer.Registry.CurrentUser.OpenSubKey("Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue(item)
        Next
    End Sub

    'Enumerate all installed instances of the JRE
    Public Sub get_jre_uninstallers()

        'Reset the list of installed JRE, allowing this code to be called multiple times
        UI.JREObjectList.Clear()

        'Create a variable to store value information 
        Dim sk As RegistryKey

        'Create a list of possible installed-programs sources
        Dim regpath As New List(Of RegistryKey)
        regpath.Add(Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
        regpath.Add(Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
        regpath.Add(Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"))
        regpath.Add(Registry.CurrentUser.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"))

        'Keep track of the image index
        Dim image_index As Integer = 0

        'Loop through possible locations for lists of apps
        Try
            For Each reg_location As RegistryKey In regpath

                'Declare the path to this individual list of installed programs
                Dim rk As RegistryKey = reg_location

                'The real deal
                Dim skname() = rk.GetSubKeyNames

                'Iterate through the keys located here
                For counter As Integer = 0 To skname.Length - 1

                    'Filter out empty keys
                    sk = rk.OpenSubKey(skname(counter))

                    If sk.GetValue("DisplayName") Is Nothing = False Then

                        'Write the display name
                        Dim name As String = CStr((sk.GetValue("DisplayName")))
                        Dim version As String
                        Dim uninstall As String


                        'Write the version
                        If sk.GetValue("DisplayVersion") Is Nothing = False Then
                            version = (CStr((sk.GetValue("DisplayVersion"))))
                        Else
                            version = (get_string("Data Unavailable"))
                        End If

                        'Save the tag for the uninstall path
                        If sk.GetValue("UninstallString") Is Nothing Then
                            uninstall = ""
                        Else
                            uninstall = (CStr((sk.GetValue("UninstallString"))))
                        End If


                        'Check if entry is for Java 6
                        If name.StartsWith("Java(TM) 6") Or name.ToString.StartsWith("Java 6 Update") = True Then
                            UI.JREObjectList.Add(New JREInstallObject(name, version, uninstall))
                        End If

                        'Check if entry is for Java 7
                        If name.StartsWith("Java(TM) 7") Or name.ToString.StartsWith("Java 7 Update") = True Then
                            UI.JREObjectList.Add(New JREInstallObject(name, version, uninstall))
                        End If

                        'Check if entry is for Java 8
                        If name.StartsWith("Java(TM) 8") Or name.ToString.StartsWith("Java 8 Update") = True Then
                            UI.JREObjectList.Add(New JREInstallObject(name, version, uninstall))
                        End If

                    End If
                Next
            Next

        Catch ex As Exception
            write_error(ex)
        End Try
    End Sub

End Module
