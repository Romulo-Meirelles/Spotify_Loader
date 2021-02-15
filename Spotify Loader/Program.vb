Imports System.Security
Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Text
Imports System.Net

''' <summary>
''' Feito por Romulo Meirelles
''' Pague-me um café! 
''' </summary>
''' <remarks>
''' </remarks>
''' 
Public Module Program
    Private Close_Application As Boolean = True
    Private DIRETORIO_SPOTIFY_LOCAL_UPDATE = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData).ToString & "\Spotify\Update"
    Private DIRETORIO_HOST = Environment.GetFolderPath(Environment.SpecialFolder.System).ToString & "\drivers\etc\hosts"
    Private DIRETORIO_SPOTIFY = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).ToString & "\Spotify"
    Private SPOTIFY_VERSION = "1.0.96.181"

    Sub Main()
        Try

                        If (Not System.IO.Directory.Exists(DIRETORIO_SPOTIFY)) Then
                            MsgBox("Spotify não esta instalado!", MsgBoxStyle.Exclamation, "Error!")
                            Exit Sub
                        End If

                        If (Not System.IO.File.Exists(DIRETORIO_SPOTIFY & "\Spotify.exe")) Then
                            MsgBox("Spotify.exe não encontrado!", MsgBoxStyle.Exclamation, "Error!")
                            Exit Sub
                        End If

                        Dim FileVersion As FileVersionInfo = FileVersionInfo.GetVersionInfo(DIRETORIO_SPOTIFY & "\Spotify.exe")

                        If FileVersion.FileVersion <> SPOTIFY_VERSION Then
                            MsgBox("VERSÃO INCOPATÍVEL COM O LOADER! " & vbCrLf & " VERSÃO NECESSÁRIA: " & SPOTIFY_VERSION & "", MsgBoxStyle.Exclamation, "Error!")
                            Exit Sub
                        End If



                        Call DenyFolderUpdate()
                        Call ADDHOSTBLOCK()
                        Threading.Thread.Sleep(1000)
                        Process.Start(DIRETORIO_SPOTIFY & "\Spotify.exe")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error!")
            Exit Sub
        End Try



        While Close_Application
            Try
                Threading.Thread.Sleep(1000)
                Dim MyProcess() As Process = Process.GetProcessesByName("Spotify") '<--- Process Name sem .exe
                If MyProcess.Length <= 0 Then

                    Call AllowFolderUpdate()
                    Call REMOVEHOSTBLOCK()
                    Close_Application = False
                End If

            Finally
            End Try
        End While

    End Sub


    Private Sub ADDHOSTBLOCK()
        Try
            ServicePointManager.SecurityProtocol = DirectCast(3072, SecurityProtocolType)
            Dim WEB As New Net.WebClient
            WEB.Encoding = Encoding.UTF8
            Dim Host_Sites As String = WEB.DownloadString("https://raw.githubusercontent.com/Romulo-Meirelles/Spotify_Loader/gh-pages/Block_Hosts")

            SetAttr(DIRETORIO_HOST, FileAttribute.Normal)

            Dim HOSTS As String = File.ReadAllText(DIRETORIO_HOST)

            If Not HOSTS.Contains("<***SPOTIFY***>") Then
                HOSTS += vbCrLf & vbCrLf & "<***SPOTIFY***><***SPOTIFY***>"
            End If

            Dim BLOCKS = HOSTS.Insert(HOSTS.IndexOf(">") + 1, vbCrLf & Host_Sites)
            File.WriteAllText(DIRETORIO_HOST, BLOCKS)
            SetAttr(DIRETORIO_HOST, FileAttribute.ReadOnly)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error!")
            End
        End Try


    End Sub
    Private Sub REMOVEHOSTBLOCK()

        Try
            SetAttr(DIRETORIO_HOST, FileAttribute.Normal)
            Dim HOSTS As String = File.ReadAllText(DIRETORIO_HOST)
            If Not HOSTS.Contains("<***SPOTIFY***>") Then
                SetAttr(DIRETORIO_HOST, FileAttribute.ReadOnly)
                Exit Sub
            End If

            Dim BLOCKS = HOSTS.Remove(HOSTS.IndexOf(">") + 1, HOSTS.LastIndexOf("<") - HOSTS.IndexOf(">") - 1)
            File.WriteAllText(DIRETORIO_HOST, BLOCKS)
            SetAttr(DIRETORIO_HOST, FileAttribute.ReadOnly)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error!")
            End
        End Try


    End Sub

    Private Sub DenyFolderUpdate()

        Try

            If (Not System.IO.Directory.Exists(DIRETORIO_SPOTIFY_LOCAL_UPDATE)) Then
                System.IO.Directory.CreateDirectory(DIRETORIO_SPOTIFY_LOCAL_UPDATE)
            End If

            Dim sec As FileSecurity = File.GetAccessControl(DIRETORIO_SPOTIFY_LOCAL_UPDATE)
            sec.AddAccessRule(New FileSystemAccessRule(GETUSERNAME, FileSystemRights.FullControl, AccessControlType.Deny))
            File.SetAccessControl(DIRETORIO_SPOTIFY_LOCAL_UPDATE, sec)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error!")
            End
        End Try


    End Sub

    Private Sub AllowFolderUpdate()

        Try

            If (Not System.IO.Directory.Exists(DIRETORIO_SPOTIFY_LOCAL_UPDATE)) Then
                System.IO.Directory.CreateDirectory(DIRETORIO_SPOTIFY_LOCAL_UPDATE)
            End If

            Dim DIRETORIO = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData).ToString & "\Spotify\Update"
            Dim sec As FileSecurity = File.GetAccessControl(DIRETORIO_SPOTIFY_LOCAL_UPDATE)
            sec.RemoveAccessRule(New FileSystemAccessRule(GETUSERNAME, FileSystemRights.FullControl, AccessControlType.Deny))
            File.SetAccessControl(DIRETORIO_SPOTIFY_LOCAL_UPDATE, sec)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error!")
            End
        End Try


    End Sub
    Public Function GETUSERNAME() As String

        Dim attr As FileAttributes = File.GetAttributes(DIRETORIO_SPOTIFY_LOCAL_UPDATE)
        If attr.HasFlag(FileAttributes.Archive) Then
            If Not File.Exists(DIRETORIO_SPOTIFY_LOCAL_UPDATE) Then
                MsgBox("O arquivo informado não existe.")
                Return Nothing
            End If
        Else
            If Not Directory.Exists(DIRETORIO_SPOTIFY_LOCAL_UPDATE) Then
                MsgBox("O diretório informado não existe.")
                Return Nothing
            End If
        End If
        Try
            Dim fileSecurity As FileSecurity = File.GetAccessControl(DIRETORIO_SPOTIFY_LOCAL_UPDATE)
            Dim identityReference As IdentityReference = fileSecurity.GetOwner(GetType(SecurityIdentifier))
            Dim ntAccount As NTAccount = TryCast(identityReference.Translate(GetType(NTAccount)), NTAccount)

            Return ntAccount.Value.ToString()
        Catch ex As Exception
            MsgBox("Erro : " + ex.Message)
            Return Nothing
        End Try
    End Function
End Module
