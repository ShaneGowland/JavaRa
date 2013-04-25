Imports System.Net
Imports Microsoft.Win32
Public Class UI

#Region "Global Variables/Declarations"

    'Store the version number
    Dim version_number As String = "020100"

    'Downloader variables
    Dim whereToSave As String 'Where the program save the file
    Delegate Sub ChangeTextsSafe(ByVal length As Long, ByVal position As Integer, ByVal percent As Integer, ByVal speed As Double)
    Delegate Sub DownloadCompleteSafe(ByVal cancelled As Boolean)
    Dim theResponse As HttpWebResponse
    Dim theRequest As HttpWebRequest

    'Local variables used for command-line args
    Friend stay_silent As Boolean = False 'Self explanatory.

    'Determine if program should restart when closed
    Dim reboot As Boolean = False

    'Variable to count the number of removed items
    Friend removal_count As Integer = 0

#End Region

#Region "Opening/Closing JavaRa"
    'Form closing/saving event
    Private Sub UI_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Save the settins to config.ini
        Try
            'Remove the previous instance
            If My.Computer.FileSystem.FileExists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\config.ini") Then
                My.Computer.FileSystem.DeleteFile(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\config.ini")
            End If

            'Declare the textwriter
            Dim SW As IO.TextWriter
            SW = IO.File.AppendText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\config.ini")

            'Do not save log if language is English
            If language = "English" = False Then : SW.WriteLine("Language:" & language) : End If

            'Save the "save log" setting
            If boxSaveLog.Checked = True Then : SW.WriteLine("SaveLog:True") : End If

            'Save the update check settings
            If boxUpdateCheck.Checked = False Then : SW.WriteLine("UpdateCheck:False") : End If

            'Close the textwriter
            SW.Close()
        Catch ex As Exception
        End Try
    End Sub
    'Form_Load event
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'read config.ini to acquire settings
        If My.Computer.FileSystem.FileExists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\config.ini") = True Then
            Dim r As IO.StreamReader
            Dim rule As String
            Try
                r = New IO.StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\config.ini")
                Do While (r.Peek() > -1)
                    rule = (r.ReadLine.ToString)

                    'Read language settings
                    If (rule.StartsWith("Language:")) = True Then
                        language = rule.Replace("Language:", "")
                    End If

                    'Read the "save log" settings
                    If rule = ("SaveLog:True") Then
                        boxSaveLog.Checked = True
                    End If

                    'Update check
                    If rule = ("UpdateCheck:False") Then
                        boxUpdateCheck.Checked = False
                    End If

                Loop
                r.Close()
            Catch ex As Exception
                write_log(ex.Message)
            End Try
        End If

        'Show available language files
        Try
            For Each script As String In My.Computer.FileSystem.GetFiles(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\localizations")
                'Check whether it is actually a script
                If System.IO.Path.GetExtension(script) = ".locale" = True Then

                    'Get the name of the file
                    Dim file_name As String = System.IO.Path.GetFileNameWithoutExtension(script)

                    'Ensure the file is a language file
                    If file_name.StartsWith("lang.") Then
                        'Add it to the list of available languages
                        boxLanguage.Items.Add(file_name.Replace("lang.", ""))

                    End If
                End If
            Next
        Catch ex As Exception
            write_log(ex.Message)
        End Try

        'Define the language JavaRa should load in
        If language = Nothing Then
            language = "English"
            boxLanguage.Text = "English"
        Else
            boxLanguage.Text = language
            Call translate_strings()
        End If

        'Decide if UI should be displayed
        If My.Application.CommandLineArgs.Count > 0 Then
            For Each value As String In My.Application.CommandLineArgs

                'Make value uppercase
                value = value.ToUpper

                'Check for the presense of the /Silent option
                If value = "/SILENT" Then
                    Me.Visible = False
                    Me.Opacity = 0
                    Me.ShowInTaskbar = False
                    stay_silent = True
                End If
            Next
            If stay_silent = False Then
                'Render the ui if no /Silent tag is detected
                Call render_ui()
            End If
            'Parse command line arguements
            Select Case My.Application.CommandLineArgs(0).ToString.ToUpper
                Case "/PURGE"
                    If stay_silent = True Then
                        Call purge_jre()
                        Me.Close()
                    Else
                        Call purge_jre()
                    End If
                Case "/CLEAN"
                    If stay_silent = True Then
                        Call cleanup_old_jre()
                        Me.Close()
                    Else
                        Call cleanup_old_jre()
                    End If
                Case "/UNINSTALLALL"
                    If stay_silent = True Then
                        Call uninstall_all(stay_silent)
                        Me.Close()
                    Else
                        Call uninstall_all()
                    End If
                Case "/UPDATEDEFS"
                    MsgBox("Upper case supported")
                    If stay_silent = True Then
                        download_defs()
                        Me.Close()
                    Else
                        btnUpdateDefs.PerformClick()
                    End If
                Case "/SILENT"
                    MessageBox.Show("Syntax Error. /Silent is a secondary option to be used in combination with other command line options." _
                                    & "It should not be the first option used, nor should it be used alone. use /? for more information.", "Syntax Error.")
                    Me.Close()
                Case "/?"
                    MessageBox.Show("/Purge -" & vbTab & vbTab & "Removes all JRE related registry keys, files and directories." & vbCrLf _
                                    & "/Clean -" & vbTab & vbTab & "Removes only JRE registry keys from previous version." & vbCrLf _
                                    & "/Uninstallall -" & vbTab & "Run the built-in uninstaller for all versions of JRE detected." & vbCrLf _
                                    & "/Updatedefs -" & vbTab & "Downloads a new copy of the JavaRa definitions." & vbCrLf _
                                    & "/Silent -" & vbTab & vbTab & "Hides the graphical user interface and suppresses all dialog" & vbCrLf _
                                    & vbTab & vbTab & "messages and error reports." & vbCrLf _
                                    & "/? -" & vbTab & vbTab & "Displays this help dialog" & vbCrLf & vbCrLf _
                                    & " Example: Javara /Updatedefs /Silent" & vbCrLf _
                                    & " Example: Javara /Uninstallall /Silent" & vbCrLf, "Command Line Parameters")
                    Me.Close()
                Case Else
                    MessageBox.Show("Unsupported argument. Please use /Purge, /Clean, /Uninstallall, or /Updatedefs" & vbCrLf _
                                    & "with, or without /Silent.", "Option Not Supported.")
                    Me.Close()
            End Select
        Else
            'Render the UI if no command line arguments were used
            Call render_ui()
        End If

            'Check silently for updates
            If boxUpdateCheck.Checked Then
                Dim trd As Threading.Thread = New Threading.Thread(AddressOf check_for_update)
                trd.IsBackground = True
                trd.Start()
            End If

    End Sub
    'Render the grid of icons on the main GUI
    Private Sub render_ui()
        'Render the user interface
        Try

            'First, clear the existing UI
            lvTools.Items.Clear()
            lbTasks.Items.Clear()

            Dim i As Bitmap
            Dim exePath As String

            'Add icon for Java updater
            i = My.Resources.update
            ExecutableImages.Images.Add(get_string("Update Java Runtime"), i)
            exePath = get_string("Update Java Runtime")
            lvTools.Items.Add(get_string("Update Java Runtime"), exePath)

            'Add icon for redundant cleaner
            i = My.Resources.clear
            ExecutableImages.Images.Add(get_string("Remove Java Runtime"), i)
            exePath = get_string("Remove Java Runtime")
            lvTools.Items.Add(get_string("Remove Java Runtime"), exePath)

            'Add icon for definition updates
            i = My.Resources.download
            ExecutableImages.Images.Add(get_string("Update JavaRa Definitions"), i)
            exePath = get_string("Update JavaRa Definitions")
            lvTools.Items.Add(get_string("Update JavaRa Definitions"), exePath)

            'Add icon for additional tools
            i = My.Resources.tools
            ExecutableImages.Images.Add(get_string("Additional Tasks"), i)
            exePath = get_string("Additional Tasks")
            lvTools.Items.Add(get_string("Additional Tasks"), exePath)

            'Populate the list of tasks
            lbTasks.Items.Add(get_string("Remove startup entry"))
            lbTasks.Items.Add(get_string("Check Java version"))
            lbTasks.Items.Add(get_string("Remove Outdated JRE Firefox Extensions"))
            lbTasks.Items.Add(get_string("Clean JRE Temp Files"))

        Catch ex As Exception
            write_log(ex.Message)
        End Try

        'Set the user interface
        Me.Width = 510
        Me.Height = 260
        Call return_home() 'Sets the GUI to the start position


    End Sub
    'Perform the reboot
    Public Sub reboot_app()
        reboot = True
        Me.Close()
    End Sub
    'After the form has closed, reboot it
    Private Sub UI_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Allow for a reboot
        If reboot = True Then
            Process.Start(System.Reflection.Assembly.GetExecutingAssembly. _
          GetModules()(0).FullyQualifiedName)
        End If

        'Log language file errors
        'Output the locale errors  
        If IO.File.Exists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\localizations\output_strings.true") And (language = "English" = False) Then
            For Each untranslated_string As String In routines_locales.untranslated_strings
                Try
                    Dim SW As IO.TextWriter
                    SW = IO.File.AppendText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\debug-errors." & language.ToUpper & ".log")
                    SW.WriteLine(untranslated_string)
                    SW.Close()
                Catch ex As Exception
                End Try
            Next
        End If

    End Sub
#End Region

#Region "Additional Tools and cleaning functions"
    'Iterates through checkedlistbox and decides which functions need to be run
    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        'Check if anything is selected
        If lbTasks.CheckedItems.Count = 0 Then : ToolTip1.Show(get_string("You didn't select any tasks to be performed."), lblTitle) : Exit Sub : End If

        'Check for startup entry
        If lbTasks.CheckedItems.Contains(get_string("Remove startup entry")) Then
            delete_jre_startup_entries()
        End If

        'Check Java version
        If lbTasks.CheckedItems.Contains(get_string("Check Java version")) Then
            output_jre_version()
        End If

        'Remove firefox extensions
        If lbTasks.CheckedItems.Contains(get_string("Remove Outdated JRE Firefox Extensions")) Then
            delete_jre_firefox_extensions()
        End If

        'Clean JRE temp files
        If lbTasks.CheckedItems.Contains(get_string("Clean JRE Temp Files")) Then
            Call clean_jre_temp_files()
        End If

        'Show some user feedback
        ToolTip1.Show(get_string(get_string("Selected tasks completed successfully.")), lblTitle)
    End Sub
    'Read the list of defs and remove the files and registry keys specified
    Private Sub btnCleanup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCleanup.Click
        Call cleanup_old_jre()
    End Sub
#End Region

#Region "Update Java Downloader/Backgroundworker"
    'Downloads the specified file in a thread.
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'Creating the request and getting the response
        Dim theResponse As HttpWebResponse
        Dim theRequest As HttpWebRequest
        Try 'Checks if the file exist
            theRequest = WebRequest.Create(Me.txtFileName.Text)
            theResponse = theRequest.GetResponse
        Catch ex As Exception
            MessageBox.Show(get_string("An error occurred while downloading file. Possible causes:") & ControlChars.CrLf & _
                            get_string("1) File doesn't exist") & ControlChars.CrLf & _
                           get_string("2) Remote server error"), get_string("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Dim cancelDelegate As New DownloadCompleteSafe(AddressOf DownloadComplete)
            Me.Invoke(cancelDelegate, True)

            Exit Sub
        End Try
        Dim length As Long = theResponse.ContentLength 'Size of the response (in bytes)
        Dim safedelegate As New ChangeTextsSafe(AddressOf ChangeTexts)
        Me.Invoke(safedelegate, length, 0, 0, 0) 'Invoke the TreadsafeDelegate
        Dim writeStream As New IO.FileStream(Me.whereToSave, IO.FileMode.Create)
        'Replacement for Stream.Position (webResponse stream doesn't support seek)
        Dim nRead As Integer
        'To calculate the download speed
        Dim speedtimer As New Stopwatch
        Dim currentspeed As Double = -1
        Dim readings As Integer = 0
        Do
            If BackgroundWorker1.CancellationPending Then 'If user abort download
                Exit Do
            End If
            speedtimer.Start()
            Dim readBytes(4095) As Byte
            Dim bytesread As Integer = theResponse.GetResponseStream.Read(readBytes, 0, 4096)
            nRead += bytesread
            Dim percent As Short = (nRead * 100) / length
            Me.Invoke(safedelegate, length, nRead, percent, currentspeed)
            If bytesread = 0 Then Exit Do
            writeStream.Write(readBytes, 0, bytesread)
            speedtimer.Stop()
            readings += 1
            If readings >= 5 Then 'For increase precision, the speed it's calculated only every five cycles
                currentspeed = 20480 / (speedtimer.ElapsedMilliseconds / 1000)
                speedtimer.Reset()
                readings = 0
            End If
        Loop
        'Close the streams
        theResponse.GetResponseStream.Close()
        writeStream.Close()
        If Me.BackgroundWorker1.CancellationPending Then
            IO.File.Delete(Me.whereToSave)
            Dim cancelDelegate As New DownloadCompleteSafe(AddressOf DownloadComplete)
            Me.Invoke(cancelDelegate, True)
            Exit Sub
        End If
        Try
            Dim completeDelegate As New DownloadCompleteSafe(AddressOf DownloadComplete)
            Me.Invoke(completeDelegate, False) : Catch ex As Exception : End Try
    End Sub
    'Code that runs when background worker has completed.
    Public Sub DownloadComplete(ByVal cancelled As Boolean)
        Me.txtFileName.Enabled = True
        btnDownload.Enabled = True
        If cancelled Then
            MessageBox.Show(get_string("Download has been cancelled."), get_string("Cancelled."), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else

        End If
        Me.ProgressBar1.Value = 0 : Me.ProgressBar3.Value = 0
        'Just in case of a minor exception
        Me.lblStep1.Text = ""
        Me.txtFileName.Enabled = True

        'Call the exe on completion
        If Me.whereToSave = My.Computer.FileSystem.SpecialDirectories.Temp & "\java-installer.exe" Then
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\java-installer.exe") Then
                Shell(My.Computer.FileSystem.SpecialDirectories.Temp & "\java-installer.exe", AppWinStyle.NormalFocus, True)
            End If
        Else
            ToolTip1.Show(get_string("JavaRa definitions updated successfully"), ToolStrip1)
        End If

        Me.Cursor = Cursors.Default
    End Sub
    'Update the progress bar while a file is being downloaded. Shared between multiple downloaders.
    Public Sub ChangeTexts(ByVal length As Long, ByVal position As Integer, ByVal percent As Integer, ByVal speed As Double)
        Me.ProgressBar2.Value = percent
        ProgressBar3.Value = percent
    End Sub
    'Start the background worker by supplying essential "what to download?" information.
    Private Sub btnDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownload.Click
        Me.Cursor = Cursors.WaitCursor
        'Confirm the connection to the server  
        Try
            theRequest = WebRequest.Create("http://content.thewebatom.net/files/confirm.txt")
            theResponse = theRequest.GetResponse
            'Set the path to the rules
            txtFileName.Text = "http://javadl.sun.com/webapps/download/AutoDL?BundleId=49024"
        Catch ex As Exception
            MessageBox.Show(get_string("Could not make a connection to download server. Please see our online help for assistance.") & Environment.NewLine & get_string("This error can be caused by incorrect proxy settings or a security product conflict."), get_string("An error was encountered."), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Cursor = Cursors.Default
            btnDownload.Enabled = False
            Exit Sub
        End Try

        'Set some paths and variables
        Me.whereToSave = My.Computer.FileSystem.SpecialDirectories.Temp & "\java-installer.exe"
        Me.txtFileName.Enabled = False
        Me.btnDownload.Enabled = False
        Me.BackgroundWorker1.RunWorkerAsync() 'Start download
    End Sub
    'Update the JavaRa definitions
    Private Sub btnUpdateDefs_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateDefs.Click
        download_defs()
    End Sub
#End Region

#Region "Uninstall/Reinstall Java Runtime environment"
    'Read the list of defs and remove the files and registry keys specified
    Private Sub btnRemoveKeys_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveKeys.Click
        Call purge_jre()
    End Sub
    'Run the uninstaller depending on which combobox item is selected.
    Private Sub btnRunUninstaller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRunUninstaller.Click
        'Check for blank selection
        If cboVersion.Text = "" Then
            MessageBox.Show(get_string("Please select a version of JRE to remove."))
        End If

        'Iterate through all installed programs and find Java.
        Try
            Dim Software As String = Nothing
            Dim SoftwareKey As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
            Using rk As RegistryKey = Registry.LocalMachine.OpenSubKey(SoftwareKey)
                For Each skName In rk.GetSubKeyNames
                    Dim name = Registry.LocalMachine.OpenSubKey(SoftwareKey).OpenSubKey(skName).OpenSubKey("InstallProperties").GetValue("DisplayName")
                    Dim uninstallString = Registry.LocalMachine.OpenSubKey(SoftwareKey).OpenSubKey(skName).OpenSubKey("InstallProperties").GetValue("UninstallString")

                    'Check if entry is for Java 6
                    If cboVersion.Text = get_string("Java Runtime Environment 6") Then
                        If name.ToString.StartsWith("Java(TM) 6") Or name.ToString.StartsWith("Java 6 Update") = True Then
                            Try
                                Shell(uninstallString, AppWinStyle.NormalFocus, True)
                            Catch ex As Exception
                                MessageBox.Show(get_string("Could not locatate uninstaller for") & " " & cboVersion.Text, get_string("Uninstaller not found"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                        End If
                    End If

                    'Check if entry is for Java 7
                    If cboVersion.Text = get_string("Java Runtime Environment 7") Then
                        If name.ToString.StartsWith("Java(TM) 7") Or name.ToString.StartsWith("Java 7") Then

                            Try
                                Shell(uninstallString, AppWinStyle.NormalFocus, True)
                            Catch ex As Exception
                                MessageBox.Show(get_string("Could not locatate uninstaller for") & " " & cboVersion.Text, get_string("Uninstaller not found"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try

                        End If
                    End If
                Next
            End Using
        Catch ex As Exception
            write_log(ex.Message)
        End Try

    End Sub
    'Launch the appropriate method for checking the JRE version installed.
    Private Sub btnUpdateNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateNext.Click
        'Decide which method the user wishes to check for updates with
        If boxOnlineCheck.Checked = True Then
            If MessageBox.Show(get_string("Would you like to open the Oracle online JRE version checker?"), get_string("Launch online check"), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then

                'This needs to be localized
                If language = "French" Then
                    Process.Start("http://java.com/fr/download/installed.jsp")
                ElseIf language = "Brazilian" Then
                    Process.Start("http://java.com/pt_BR/download/installed.jsp")
                ElseIf language = "German" Then
                    Process.Start("http://java.com/de/download/installed.jsp")
                ElseIf language = "Polish" Then
                    Process.Start("http://www.java.com/pl/download/installed.jsp")
                ElseIf language = "Russian" Then
                    Process.Start("http://www.java.com/ru/download/installed.jsp")
                ElseIf language = "Spanish" Then
                    Process.Start("http://www.java.com/es/download/installed.jsp")
                Else
                    Process.Start("http://java.com/en/download/installed.jsp")
                End If
                Exit Sub
            End If


            'fill with new panel
            show_panel(pnlCompleted)

            'Change the label
            lblCompleted.Text = get_string("step 2 - completed.")
        End If

        If boxDownloadJRE.Checked = True Then

            'fill with new panel
            show_panel(pnlDownload)

            'Update the label
            lblDownloadNewVersion.Text = get_string("step 2 - download new version")

        End If

        If boxJucheck.Checked Then
            Try
                'Create a variable to contain the path
                Dim path As String = Nothing

                'Get the path to jucheck.exe (x64 systems)
                If My.Computer.FileSystem.FileExists("C:\Program Files (x86)\Common Files\Java\Java Update\jucheck.exe") Then
                    path = "C:\Program Files (x86)\Common Files\Java\Java Update\jucheck.exe"
                End If

                'x86 systems.
                If My.Computer.FileSystem.FileExists("C:\Program Files\Common Files\Java\Java Update\jucheck.exe") Then
                    path = "C:\Program Files\Common Files\Java\Java Update\jucheck.exe"
                End If

                'Inform user if JRE does not exist
                If path = Nothing Then
                    MessageBox.Show(get_string("Java Update Checking Utility (jucheck.exe) could not be found on your system."), get_string("Could not locate jucheck.exe"))
                Else
                    Shell(path, AppWinStyle.NormalFocus, True)
                End If

                'Show the next panel
                show_panel(pnlCleanup)

                'Change the label
                lblCompleted.Text = get_string("step 3 - completed.")

                'Catch any errors 
            Catch ex As Exception
                write_log(ex.Message)
            End Try
        End If
    End Sub
    'Decide which step of the process the downloader should be.
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If lblDownloadNewVersion.Text = get_string("step 3 - download new version") Then
            show_panel(pnlRemoval)
        End If
        'Java update mode
        If lblDownloadNewVersion.Text = get_string("step 2 - download new version") Then
            show_panel(pnlUpdateJRE)
        End If
    End Sub
    'This launches the manual download page. Auto-jumped to the Windows section.
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://www.java.com/en/download/manual.jsp#win")
    End Sub
#End Region

#Region "User interface navigation"
    'Open the appropriate subtool
    Private Sub lvTools_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvTools.Click
        'Handles click events on the main GUI and loads the correct panel.
        For Each file As ListViewItem In lvTools.Items
            If file.Selected = True Then
                Dim path As String = file.SubItems(0).Text
                If path = get_string("Additional Tasks") Then
                    show_panel(PanelExtra)
                End If
                If path = get_string("Remove Java Runtime") Then

                    'Determine which uninstallers to show
                    Try
                        Dim Software As String = Nothing
                        Dim SoftwareKey As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
                        Using rk As RegistryKey = Registry.LocalMachine.OpenSubKey(SoftwareKey)
                            For Each skName In rk.GetSubKeyNames
                                Dim name = Registry.LocalMachine.OpenSubKey(SoftwareKey).OpenSubKey(skName).OpenSubKey("InstallProperties").GetValue("DisplayName")
                                Dim uninstallString = Registry.LocalMachine.OpenSubKey(SoftwareKey).OpenSubKey(skName).OpenSubKey("InstallProperties").GetValue("UninstallString")

                                'Check if entry is for Java 6
                                If name.ToString.StartsWith("Java(TM) 6") Or name.ToString.StartsWith("Java 6 Update") = True Then
                                    cboVersion.Items.Add(get_string("Java Runtime Environment 6"))
                                End If

                                'Check if entry is for Java 7
                                If name.ToString.StartsWith("Java(TM) 7") Or name.ToString.StartsWith("Java 7") = True Then
                                    cboVersion.Items.Add(get_string("Java Runtime Environment 7"))
                                End If
                            Next
                        End Using

                    Catch ex As Exception

                    End Try

                    show_panel(Step1)

                    'Make sure there are some uninstallers found. If not, display a message.
                    'The opposite output needs to be explicitly coded, as users may run this function twice.
                    If cboVersion.Items.Count = 0 Then
                        cboVersion.Enabled = False
                        lblStep1.Text = get_string("No uninstaller found. Please press 'next' to continue")
                        btnRunUninstaller.Enabled = False
                    Else
                        btnRunUninstaller.Enabled = True
                        cboVersion.Enabled = True
                        lblStep1.Text = get_string("We recommend that you try running the Java Runtime Environment's built-in") & Environment.NewLine & get_string("uninstaller before you continue.")
                    End If

                End If
                If path = get_string("Update JavaRa Definitions") Then
                    show_panel(pnlUpdate)
                End If
                If path = get_string("Update Java Runtime") = True Then
                    show_panel(pnlUpdateJRE)
                End If
            End If
        Next
    End Sub
    'Load support page in web browser control.
    Private Sub txtSupport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSupport.Click
        show_panel(pnlBrowser)
        WebBrowser1.DocumentText = My.Resources.credits
    End Sub
    'Navigate to step #2 of removal process.
    Private Sub btnStep1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep1.Click
        show_panel(pnlRemoval)
    End Sub
    'Make sure the correct "step-label" is displayed when moving the step 3
    Private Sub btnStep2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep2.Click
        show_panel(pnlDownload)
    End Sub
    'Decide if the user wishes to close JavaRa or continue using it
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If btnCloseWiz.Checked = True Then
            Call return_home()
            'Reset the label text
            lblCompleted.Text = get_string("step 4 - completed.")
        Else
            Me.Close()
        End If
    End Sub
    'Navigate to step #4
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'Clear progress bars
        ProgressBar2.Value = 0
        ProgressBar3.Value = 0

        show_panel(pnlCompleted)

        'Display correct label on completion page
        If lblDownloadNewVersion.Text = get_string("step 3 - download new version") Then
            lblCompleted.Text = get_string("step 4 - completed.")
        ElseIf lblDownloadNewVersion.Text = get_string("step 2 - download new version") Then
            lblCompleted.Text = get_string("step 3 - completed.")
        End If
    End Sub
    'Return to panel #1
    Private Sub btnPrev1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrev1.Click
        show_panel(Step1)
    End Sub
    'Show the settings panel in the UI
    Private Sub btnSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSettings.Click
        show_panel(PanelSettings)
    End Sub
    'Return to the home page whenever a back button is pressed
    Private Sub Return_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click, Button11.Click, Button12.Click, btnUpdateJavaPrevious.Click, Button8.Click, Button2.Click
        Call return_home()
    End Sub
    'Load the about page
    Private Sub btnAbout_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbout.Click
        show_panel(pnlBrowser)
        WebBrowser1.DocumentText = My.Resources.about
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        show_panel(pnlUpdateJRE)
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        show_panel(pnlCompleted)
    End Sub
#End Region

#Region "Configuration & Updating"
    'Change the current language selection
    Private Sub btnSaveLang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveLang.Click
        'Temporarily store the previous language
        Dim initial_lang As String = language

        'Update the language global
        language = boxLanguage.Text

        'Ensure that user has actually changed the settings
        If initial_lang = language Then

        Else

            If initial_lang = "English" Then

                'Initial language is English; we can directly translate the control.
                Call translate_strings()

                'Re-render the UI
                Call render_ui()
            Else
                'If the initial language wasn't English then reboot
                MessageBox.Show(get_string("Your changes will take effect when JavaRa is restarted."))
                Call reboot_app()
                Me.Close()
            End If
        End If
    End Sub
    'Prevent the language settings from doing dumb stuff.
    Private Sub boxLanguage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles boxLanguage.SelectedIndexChanged
        If boxLanguage.Text <> language Then
            btnSaveLang.Enabled = True
        Else
            btnSaveLang.Enabled = False
        End If
    End Sub
    'This delegate allows the update notification to show across threads.
    Public Delegate Sub ShowNotification(ByVal data As Boolean)
    Public Sub UI_Thread_Show_Notification(ByVal data As Boolean)
        If MessageBox.Show(get_string("A new version of JavaRa is available! Visit download page?"), get_string("Update Available"), MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Process.Start("http://singularlabs.com/software/javara/javara-download/")
        End If
    End Sub
    'Invokation method so that dialog can be shown on UI thread.
    Public Sub show_threaded_dialog()
        lblTitle.BeginInvoke(New ShowNotification(AddressOf UI_Thread_Show_Notification), False)
    End Sub
    'Method that performs the update check
    Public Sub check_for_update()

        'Download the current version definition file
        Try

            'Store the path to the version tag online.
            Dim version_tag_url As String = "http://content.thewebatom.net/updates/javara/version.tag"

            'Variable to hold the version tag after it's downloaded
            Dim version_tag_path As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\javaraversion.tag"

            'Download and store the file containing the version number
            Dim req As Net.WebRequest
            Dim resp As IO.Stream
            Dim out As IO.BinaryWriter
            req = Net.HttpWebRequest.Create(version_tag_url)
            resp = req.GetResponse().GetResponseStream()
            out = New IO.BinaryWriter(New IO.FileStream(version_tag_path, IO.FileMode.OpenOrCreate))
            Dim buf(4096) As Byte
            Dim k As Int32 = resp.Read(buf, 0, 4096)
            Do While k > 0
                out.Write(buf, 0, k)
                k = resp.Read(buf, 0, 4096)
            Loop

            'Close the downloader methods
            resp.Close()
            out.Close()

            'Run a version comparison
            If IO.File.Exists(version_tag_path) Then
                Dim r As IO.StreamReader
                Dim newVersionNumber As String
                r = New IO.StreamReader(version_tag_path)
                Do While (r.Peek() > -1)
                    newVersionNumber = (r.ReadLine.ToString)
                    Dim oldVersionNumber As String = version_number 'Read from global variable
                    If newVersionNumber.Length = oldVersionNumber.Length = False Then
                        'Version strings are a different length. Exit gracefully.
                        End
                    End If
                    If CInt(newVersionNumber) > CInt(oldVersionNumber) Then

                        'A new version of JavaRa is available.
                        Call show_threaded_dialog()

                    End If
                Loop
            End If

        Catch ex As Exception
            'Discard silently
        End Try

    End Sub

    Private Sub download_defs()
        Me.Cursor = Cursors.WaitCursor
        'Confirm the connection to the server 
        Dim webclient As New WebClient, sFilename As String, sURI As String
        Try
            theRequest = WebRequest.Create("http://content.thewebatom.net/files/confirm.txt")
            theResponse = theRequest.GetResponse
            'Set the path to the rules
            sURI = "http://content.thewebatom.net/updates/javara/"
            sFilename = "JavaRa.def"
            txtFileName.Text = sURI & sFilename
        Catch ex As Exception
            MessageBox.Show(get_string("Could not make a connection to download server. Please see our online help for assistance.") & Environment.NewLine & get_string("This error can be caused by incorrect proxy settings or a security product conflict."), get_string("An error was encountered."), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Cursor = Cursors.Default
            btnDownload.Enabled = False
            Exit Sub
        End Try

        'Set some paths and variables
        Me.whereToSave = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa.def"
        Me.txtFileName.Enabled = False
        Me.btnUpdateDefs.Enabled = False
        Me.btnDownload.Enabled = False
        If stay_silent = True Then
            webclient.DownloadFile(txtFileName.Text, Me.whereToSave)
        Else
            Me.BackgroundWorker1.RunWorkerAsync() 'Start download               
        End If
    End Sub

#End Region

    Private Sub lvTools_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles lvTools.SelectedIndexChanged

    End Sub
End Class

