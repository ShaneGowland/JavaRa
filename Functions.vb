Imports Microsoft.Win32
Module Functions

    'Determine if a file is in use
    Public Function FileInUse(ByVal sFile As String) As Boolean
        If System.IO.File.Exists(sFile) Then
            Try
                Dim F As Short = CShort(FreeFile())
                FileOpen(F, sFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                FileClose(F)
            Catch
                Return True
            End Try
        End If
        Return False
    End Function

    'Delete a file if it isn't in use
    Public Sub DeleteIfPermitted(ByVal path As String)
        Try
            If My.Computer.FileSystem.FileExists(path) = True Then
                If FileInUse(path) = False Then
                    My.Computer.FileSystem.DeleteFile(path)
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    'Determine if a registry key exists
    Public Function RegKeyExists(ByVal key As String) As Boolean
        Try

            'Filter any trailing backslashes, as they make everything go to hell.
            If key.EndsWith("\") Then
                key = key.TrimEnd(CChar("\"))
            End If

            'preserve the original path
            Dim original_key As String = key

            'Check for the existence of the subkey, assigning the key value to a predefined registry hive
            Dim regKey As Microsoft.Win32.RegistryKey
            If key.StartsWith("HKLM\") Then
                key = key.Replace("HKLM\", "")
                regKey = Registry.LocalMachine.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKCU\") Then
                key = key.Replace("HKCU\", "")
                regKey = Registry.CurrentUser.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKCR\") Then
                key = key.Replace("HKCR\", "")
                regKey = Registry.ClassesRoot.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKCC\") Then
                key = key.Replace("HKCC\", "")
                regKey = Registry.CurrentConfig.OpenSubKey(key, True)
            ElseIf key.StartsWith("HKUS\") Then
                key = key.Replace("HKUS\", "")
                regKey = Registry.Users.OpenSubKey(key, True)
            Else
                'prevent a null reference exeption in rare circumstances
                regKey = Nothing
            End If

            'Conditional to check if key exists
            If regKey Is Nothing Then 'doesn't exist. Check for values.

                'Values require the hive name to be specified. Do the replacement.
                If original_key.StartsWith("HKLM\") Then
                    original_key = original_key.Replace("HKLM\", "HKEY_LOCAL_MACHINE\")
                ElseIf key.StartsWith("HKCU\") Then
                    original_key = original_key.Replace("HKCU\", "HKEY_CURRENT_USER\")
                ElseIf original_key.StartsWith("HKCR\") Then
                    original_key = original_key.Replace("HKCR\", "HKEY_CLASSES_ROOT\")
                ElseIf original_key.StartsWith("HKUS\") Then
                    original_key = original_key.Replace("HKUS\", "HKEY_USERS\")
                ElseIf original_key.StartsWith("HKCC\") Then
                    original_key = original_key.Replace("HKCC\", "HKEY_CURRENT_CONFIG\")
                End If

                'Work out the valuename and subkey
                Dim last_index As Integer = original_key.LastIndexOf("\")
                Dim valueName As String = original_key.Remove(0, (last_index + 1))
                original_key = original_key.Remove(last_index)

                'Do the conditional on the specified value in the subkey
                Dim regValue As Object = Registry.GetValue(original_key, valueName, Nothing)
                If regValue Is Nothing Then
                    Return False
                Else
                    Return True
                End If

            Else
                'It was the subkey after all, return true.
                Return True
            End If

        Catch ex As Exception
            'An exception has occurred. Return false to be on the safe side.
            Return False
        End Try
    End Function

    'Method for logging important events
    Public Sub write_log(ByVal message As String)
            Try
                Dim SW As IO.TextWriter
                SW = IO.File.AppendText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) & "\JavaRa-" & sanitze_str(CStr(Date.Today)) & ".log")
                SW.WriteLine(message)
                SW.Close()
            Catch ex As Exception
            End Try
    End Sub

    'Method that logs exceptions 
    Public Sub write_error(ByVal error_object As Exception)

        Try
            'Get the exception data
            Dim source As String = Reflection.Assembly.GetCallingAssembly.GetName.Name
            Dim message As String = error_object.Message
            Dim stack As String = error_object.StackTrace

            'Write it
            Dim exception_text As String = ("Exception encountered in module [" & source & "]" & _
                         Environment.NewLine & "Message: " & message & _
                         Environment.NewLine & stack & Environment.NewLine)
            write_log(exception_text)
        Catch ex As Exception
        End Try
    End Sub

    'Remove unsafe characters from a string. Useful for formatting dates as NTFS-compatible filenames.
    Public Function sanitze_str(ByVal filename As String, Optional ByVal allowSpaces As Boolean = False) As String
        filename = filename.Replace("[", "")
        filename = filename.Replace("*]", "")

        'Allow spaces
        If allowSpaces = False Then
            filename = filename.Replace(" ", "")
        End If

        filename = filename.Replace(")", "")
        filename = filename.Replace("(", "")
        filename = filename.Replace("/", "-")
        filename = filename.Replace("\", "-")
        filename = filename.Replace(":", "")
        filename = filename.Replace(".", "")
        filename = filename.Replace("'", "")

        Return filename
    End Function

End Module
