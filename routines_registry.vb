Option Strict Off
Imports Microsoft.Win32
Module routines_registry
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


            Dim uninstallkey As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            Dim uninstallkey64 As String = "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            Dim rk As RegistryKey = Registry.LocalMachine.OpenSubKey(uninstallkey)
            Dim rk64 As RegistryKey = Registry.LocalMachine.OpenSubKey(uninstallkey64)
            Dim sk As RegistryKey
            Dim skname() = rk.GetSubKeyNames
            Dim skname64() = rk64.GetSubKeyNames

            'Iterate for x32 installs
            For counter As Integer = 0 To skname.Length - 1

                sk = rk.OpenSubKey(skname(counter))
                If sk.GetValue("DisplayName") Is Nothing = False Then

                    'Write the display name
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java(TM) 6") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java(TM) 7") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java 7") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java 6") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If

                End If
            Next

            'Iterate for x64 installs
            For counter As Integer = 0 To skname64.Length - 1

                sk = rk64.OpenSubKey(skname64(counter))
                If sk.GetValue("DisplayName") Is Nothing = False Then

                    'Write the display name
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java(TM) 6") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java(TM) 7") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java 6") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If
                    If ((sk.GetValue("DisplayName"))).ToString.StartsWith("Java 7") Then
                        SW.WriteLine(sk.GetValue("DisplayName").ToString & " version: " & sk.GetValue("DisplayVersion").ToString)
                    End If

                End If
            Next

            SW.Close()

            'show in notepad
            Process.Start(version_output_path)

        Catch ex As Exception
            write_error(ex)
        End Try
    End Sub
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
End Module
