Imports System
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Collections
Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Reflection
Imports System.IO
Imports System.Xml

Namespace OffLine.Installer
    <RunInstaller(True)> _
    Public Class Installer
        Inherits System.Configuration.Install.Installer
        Public Sub New()
            MyBase.New()
            'This call is required by the Component Designer.
            'InitializeComponent()
            'Add initialization code after the call to InitializeComponent
            AddHandler Me.Committed, AddressOf MyInstaller_Committed
            AddHandler Me.Committing, AddressOf MyInstaller_Committing
        End Sub

        Private Sub MyInstaller_Committing(ByVal sender As Object, ByVal e As InstallEventArgs)
        End Sub

        Private Sub MyInstaller_Committed(ByVal sender As Object, ByVal e As InstallEventArgs)
            Try
                Directory.SetCurrentDirectory(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))
                Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\QLNS.exe")
            Catch
            End Try
        End Sub

        Public Overloads Overrides Sub Install(ByVal savedState As IDictionary)
            MyBase.Install(savedState)
            Try
                Dim param1 As String = Context.Parameters("Server")
                Dim param2 As String = Context.Parameters("Database")
                Dim param3 As String = Context.Parameters("User")
                Dim param4 As String = Context.Parameters("Password")
                Dim assemblypath As String = Context.Parameters("assemblypath")
                Dim appConfigPath As String = assemblypath & ".config"
                ' Write the path to the app.config file
                Dim doc As XmlDocument = New XmlDocument()
                doc.Load(appConfigPath)

                Dim configuration As XmlNode = Nothing
                For Each node As XmlNode In doc.ChildNodes
                    If node.Name = "configuration" Then configuration = node
                Next
                If Not (configuration Is Nothing) Then
                    Dim settingNode As XmlNode = Nothing
                    For Each node As XmlNode In configuration.ChildNodes
                        If node.Name = "appSettings" Then settingNode = node
                    Next
                    If Not (settingNode Is Nothing) Then
                        Dim NumNode As XmlNode = Nothing
                        For Each node As XmlNode In settingNode.ChildNodes
                            If Not (node.Attributes("key") Is Nothing) Then
                                If node.Attributes("key").Value = "Server" Then
                                    NumNode = node
                                    If Not (NumNode Is Nothing) Then
                                        Dim att As XmlAttribute = NumNode.Attributes("value")
                                        att.Value = param1
                                    End If
                                End If
                                If node.Attributes("key").Value = "Database" Then
                                    NumNode = node
                                    If Not (NumNode Is Nothing) Then
                                        Dim att As XmlAttribute = NumNode.Attributes("value")
                                        att.Value = param2
                                    End If
                                End If
                                If node.Attributes("key").Value = "User" Then
                                    NumNode = node
                                    If Not (NumNode Is Nothing) Then
                                        Dim att As XmlAttribute = NumNode.Attributes("value")
                                        att.Value = param3
                                    End If
                                End If
                                If node.Attributes("key").Value = "Password" Then
                                    NumNode = node
                                    If Not (NumNode Is Nothing) Then
                                        Dim att As XmlAttribute = NumNode.Attributes("value")
                                        att.Value = param4
                                    End If
                                End If
                            End If
                        Next
                        doc.Save(appConfigPath)
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Overloads Overrides Sub Commit(ByVal savedState As IDictionary)
            MyBase.Commit(savedState)
        End Sub

        Public Overloads Overrides Sub Rollback(ByVal savedState As IDictionary)
            MyBase.Rollback(savedState)
        End Sub
    End Class
End Namespace

