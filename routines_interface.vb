Imports Microsoft.Win32

Module routines_interface
    'Shared routine to bring a panel to the front
    Public Sub show_panel(ByVal pnl As Panel, Optional ByVal scrollbar As Boolean = False)
        'Load the correct panel
        pnl.BringToFront()
        pnl.Dock = DockStyle.Fill

        'Ensure the page footer is correct
        UI.Panel7.BringToFront()
    End Sub
    'Restore the home page UI
    Public Sub return_home()
        'fill with new panel
        UI.pnlTopDock.BringToFront()
        UI.pnlTopDock.Dock = DockStyle.Fill
        UI.lvTools.Dock = DockStyle.Fill
        UI.Panel7.BringToFront()
    End Sub
    'Clean JRE related temp files
    Public Sub clean_jre_temp_files()
        write_log(get_string("== Cleaning JRE temporary files =="))

        'Java Cache
        Try
            Dim path1 As String = (System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Sun\Java\Deployment\").Replace("\Local\", "\LocalLow\")
            For Each foundFile As String In IO.Directory.GetFiles(path1, "*.*", IO.SearchOption.AllDirectories)

                'Only catch the cache directories
                If foundFile.Contains("\cache\") Then
                    IO.File.Delete(foundFile)
                    write_log(get_string("Deleted file:" & " " & foundFile))
                End If
                If foundFile.Contains("\SystemCache\") Then
                    write_log(get_string("Deleted file:" & " " & foundFile))
                End If
            Next
        Catch ex As Exception
            write_log(ex.Message)
        End Try

        'Insert a blank space into the log
        write_log(" ")

    End Sub
    'Delete old JRE Firefox extensions
    Public Sub delete_jre_firefox_extensions()
        Try

            'Store the base path to the firefox directory
            Dim folder_base_path As String = Nothing
            If My.Computer.FileSystem.DirectoryExists("C:\Program Files (x86)\Mozilla Firefox\extensions\") Then
                folder_base_path = ("C:\Program Files (x86)\Mozilla Firefox\extensions\")
            ElseIf My.Computer.FileSystem.DirectoryExists("C:\Program Files\Mozilla Firefox\extensions\") Then
                folder_base_path = ("C:\Program Files\Mozilla Firefox\extensions\")
            Else
                Exit Sub
            End If

            'Create a variable to contain the highest version folder
            Dim version_to_keep As String = 0

            'scan extensions folder for jre
            For Each foundFolder As String In My.Computer.FileSystem.GetDirectories(folder_base_path)

                'Filter out non-JRE folders
                If foundFolder.Contains("{CAFEEFAC-0016-0000-") = True And foundFolder.Contains("-ABCDEFFEDCBA}") Then

                    'Get the filename only
                    foundFolder = foundFolder.Remove(0, foundFolder.LastIndexOf("\") + 1)

                    'trim the ends
                    foundFolder = foundFolder.Replace("-ABCDEFFEDCBA}", "")
                    foundFolder = foundFolder.Replace("{CAFEEFAC-0016-0000-", "")

                    'Store only the highest version
                    If foundFolder > version_to_keep Then
                        version_to_keep = foundFolder
                    End If

                End If

            Next

            'Create a variable to decide which folder to keep
            Dim folder_to_keep As String = folder_base_path & "{CAFEEFAC-0016-0000-" & version_to_keep & "-ABCDEFFEDCBA}"

            'Delete the JRE folders that should not be kept
            For Each delFolder As String In My.Computer.FileSystem.GetDirectories(folder_base_path)

                'Ensure that the directory belongs to JRE
                If delFolder.Contains("{CAFEEFAC-0016-0000-") = True And delFolder.Contains("-ABCDEFFEDCBA}") Then

                    'Exclude the newest version
                    If delFolder <> folder_to_keep Then
                        My.Computer.FileSystem.DeleteDirectory(delFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    End If
                End If

            Next

        Catch ex As Exception
            write_log(ex.Message)
        End Try
    End Sub
    'Uninstall all JRE's with their uninstallers
    Public Sub uninstall_all(Optional ByVal silent As Boolean = False)
        Try
            Dim Software As String = Nothing
            Dim SoftwareKey As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
            Using rk As RegistryKey = Registry.LocalMachine.OpenSubKey(SoftwareKey)
                For Each skName In rk.GetSubKeyNames
                    Dim name = Registry.LocalMachine.OpenSubKey(SoftwareKey).OpenSubKey(skName).OpenSubKey("InstallProperties").GetValue("DisplayName")
                    Dim uninstallString = Registry.LocalMachine.OpenSubKey(SoftwareKey).OpenSubKey(skName).OpenSubKey("InstallProperties").GetValue("UninstallString")
                    'Check for /SILENT and add appropriate options to uninstallstring
                    If silent = True Then
                        uninstallString = uninstallString & " /qn /Norestart"
                    End If
                    'Check if entry is for Java 6
                    If name.ToString.StartsWith("Java(TM) 6") Or name.ToString.StartsWith("Java 6 Update") = True Then
                        Try
                            Shell(uninstallString, AppWinStyle.NormalFocus, True)
                        Catch ex As Exception
                            write_log(ex.Message)
                        End Try
                    End If

                    'Check if entry is for Java 7
                    If name.ToString.StartsWith("Java(TM) 7") Or name.ToString.StartsWith("Java 7") Then

                        Try
                            Shell(uninstallString, AppWinStyle.NormalFocus, True)
                        Catch ex As Exception
                            write_log(ex.Message)
                        End Try

                    End If

                Next
            End Using
        Catch ex As Exception
            write_log(ex.Message)
        End Try
    End Sub
    'Cleanup old JRE registry keys
    Public Sub cleanup_old_jre()
        Try

            'Check if JavaRa defs are present
            If UI.stay_silent = False Then
                If My.Computer.FileSystem.FileExists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def") = False Then
                    MessageBox.Show(get_string("The JavaRa definitions have been removed.") & Environment.NewLine & get_string("You need to download them before continuing."), get_string("Definitions Not Available"))
                    show_panel(UI.pnlUpdate)
                    Exit Sub
                End If
            End If

            'Write that the user started the removal process to the logger
            write_log(get_string("User initialised redundant data purge.") & Environment.NewLine & "......................" & Environment.NewLine)

            'Load definition file and unmanaged code into memory
            Dim r As IO.StreamReader
            Dim rule As String
            r = New IO.StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def")

            'Quickly iterate through definitions for progress bar.
            Dim total_lines As Integer = 1
            Do While (r.Peek() > -1)
                rule = r.ReadLine.ToString
                If rule.StartsWith("linecount=") Then
                    total_lines = CInt(rule.Replace("linecount=", ""))
                    Exit Do
                End If
            Loop

            'Set variable for progress of actual routine
            Dim current_line As Integer = 0

            'Iterate through definitions and perform operations
            Dim r2 As IO.StreamReader

            r2 = New IO.StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def")
            Do While (r2.Peek() > -1)
                rule = (r2.ReadLine.ToString)

                'Delete specified registry key
                If rule.StartsWith("[key]") = True Then
                    'Call the function that deletes registry keys.
                    delete_key(rule.Replace("[key]", ""))
                End If

                'Update the user interface and prevent application hangs.
                current_line = current_line + 1

                If UI.stay_silent = False Then
                    'Ensure does not overflow
                    If current_line < total_lines Then
                        UI.ProgressBar1.Value = (current_line / total_lines) * 100
                    Else
                        'ProgressBar1.Value = 100
                    End If
                    Application.DoEvents()
                End If

            Loop

            'Write the results (quantity) to the log file.
            write_log(get_string("Cleanup routine completed successfully.") & " " & UI.removal_count & " " & get_string("items have been deleted."))
            If UI.stay_silent = False Then
                MessageBox.Show(get_string("Cleanup routine completed successfully.") & " " & UI.removal_count & " " & get_string("items have been deleted."), get_string("Removal Routine Complete"))
            End If

        Catch ex As Exception
            write_log(ex.Message)
        End Try

        'Close the program if in silent mode
        If UI.stay_silent = True Then
            UI.Close()
        End If
    End Sub
    'Purge entire JRE install
    Public Sub purge_jre()
        Try

            'Check if JavaRa defs are present
            If UI.stay_silent = False Then
                If My.Computer.FileSystem.FileExists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def") = False Then
                    MessageBox.Show(get_string("The JavaRa definitions have been removed.") & Environment.NewLine & get_string("You need to download them before continuing."), get_string("Definitions Not Available"))
                    show_panel(UI.pnlUpdate)
                    Exit Sub
                End If
            End If

            'Write that the user started the removal process to the logger
            write_log(get_string("User initialised redundant data purge.") & Environment.NewLine & "......................" & Environment.NewLine)

            'Load definition file and unmanaged code into memory
            Dim r As IO.StreamReader
            Dim rule As String
            r = New IO.StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def")

            'Quickly iterate through definitions for progress bar.
            Dim total_lines As Integer = 1
            Do While (r.Peek() > -1)
                rule = r.ReadLine.ToString
                If rule.StartsWith("linecount=") Then
                    total_lines = CInt(rule.Replace("linecount=", ""))
                    Exit Do
                End If
            Loop

            'Set variable for progress of actual routine
            Dim current_line As Integer = 0

            'Iterate through definitions and perform operations
            Dim r2 As IO.StreamReader
            r2 = New IO.StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def")
            Do While (r2.Peek() > -1)
                rule = (r2.ReadLine.ToString)

                'Delete specified registry key
                If rule.StartsWith("[key]") = True Then
                    'Call the function that deletes registry keys.
                    delete_key(rule.Replace("[key]", ""))
                End If

                'Remove any files specified
                If rule.StartsWith("[dir]") Then
                    delete_dir(rule.Replace("[dir]", ""))
                End If

                If rule.StartsWith("[file]") Then
                    delete_file(rule.Replace("[file]", ""))
                End If

                'Update the user interface and prevent application hangs.
                current_line = current_line + 1

                If UI.stay_silent = False Then
                    'Ensure does not overflow
                    If current_line < total_lines Then
                        UI.ProgressBar1.Value = CInt((current_line / total_lines) * 100)
                    Else
                        'ProgressBar1.Value = 100
                    End If
                    Application.DoEvents()
                End If
            Loop

            'Delete the files
            Try
                If IO.Directory.Exists("C:\Program Files\Java\") Then
                    IO.Directory.Delete("C:\Program Files\Java\", True)
                End If
                If IO.Directory.Exists("C:\Program Files (x86)\Java\") Then
                    IO.Directory.Delete("C:\Program Files (x86)\Java\", True)
                End If
                If IO.Directory.Exists("C:\Users\" & Environment.UserName & "\AppData\LocalLow\Sun\Java") Then
                    IO.Directory.Delete("C:\Users\" & Environment.UserName & "\AppData\LocalLow\Sun\Java", True)
                End If
            Catch ex As Exception

            End Try

            'Write the results (quantity) to the log file.
            write_log(get_string("Removal routine completed successfully.") & " " & UI.removal_count & " " & get_string("items have been deleted."))
            If UI.stay_silent = False Then
                MessageBox.Show(get_string("Removal routine completed successfully.") & " " & UI.removal_count & " " & get_string("items have been deleted."), get_string("Removal Routine Complete"))
            End If

        Catch ex As Exception
            write_log(ex.Message)
        End Try

        'Close the program if in silent mode
        If UI.stay_silent = True Then
            UI.Close()
        End If
    End Sub

    'Method for deleting files with wildcard data
    Private Sub delete_file(ByVal file As String)

        'Change environmental variable paths
        file = file.Replace("%ProgramFiles%", "C:\Program Files")
        file = file.Replace("%CommonAppData%", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
        file = file.Replace("%Windows%", "C:\Windows")
        file = file.Replace("%AppData%", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
        file = file.Replace("%LocalAppData%", Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))

        'Delete file
        If IO.File.Exists(file) Then
            IO.File.Delete(file)
        ElseIf file.Contains("C:\Program Files") Then

            Try
                IO.File.Delete(file.Replace("C:\Program Files", "C:\Program Files (x86)"))
            Catch ex As Exception
            End Try

        End If
    End Sub

    'Routine to convert wildcards in a directory path 
    Private Sub delete_dir(ByVal file As String)

        'Change path for program files
        If file.Contains("%ProgramFiles%") Then
            file = file.Replace("%ProgramFiles%", "C:\Program Files")
            LoopDelete(file)
            If IO.Directory.Exists("C:\Program Files (x86)") Then
                LoopDelete(file.Replace("C:\Program Files\", "C:\Program Files (x86)\"))
            End If
            Exit Sub
        End If

        'Change path for WinDir
        If file.Contains("%Windows%") Then
            file = file.Replace("%Windows%", "C:\Windows")
            LoopDelete(file)
            Exit Sub
        End If

        'Change path for AppData
        If file.Contains("%AppData%") Then
            file = file.Replace("%AppData%", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
            LoopDelete(file)
            Exit Sub
        End If

        'Common AppData
        If file.Contains("%CommonAppData%") Then
            file = file.Replace("%CommonAppData%", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
            LoopDelete(file)
            Exit Sub
        End If

        'Local application data
        If file.Contains("%LocalAppData%") Then
            file = file.Replace("%LocalAppData%", Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
            LoopDelete(file)
            Exit Sub
        End If

    End Sub

    'Sub to loop and delete all files in specified directory
    Private Sub LoopDelete(ByVal file As String)

        'Loop and delete
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(file)
            Try
                If IO.File.Exists(foundFile) Then
                    IO.File.Delete(foundFile)
                End If
            Catch ex As Exception
                write_log(ex.Message)
            End Try
        Next
    End Sub
End Module
