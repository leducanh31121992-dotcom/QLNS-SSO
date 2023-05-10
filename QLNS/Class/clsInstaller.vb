Imports System
Imports System.Diagnostics
'Imports System.Windows.Forms
'Imports System.Collections
'Imports System.ComponentModel
'Imports System.Configuration.Install
'Imports System.Reflection
'Imports System.IO
'Imports System.Xml
''Imports System.Text

'Namespace OffLine.Installer
'    ' Taken from:http://msdn2.microsoft.com/en-us/library/
'    ' system.configuration.configurationmanager.aspx
'    ' Set 'RunInstaller' attribute to true.

'    <RunInstaller(True)> _
'    Public Class InstallerClass
'        Inherits System.Configuration.Install.Installer
'        Public Sub New()
'            MyBase.New()
'            AddHandler Me.Committed, AddressOf MyInstaller_Committed
'            ' Attach the 'Committed' event.
'            ' Attach the 'Committing' event.
'            AddHandler Me.Committing, AddressOf MyInstaller_Committing
'        End Sub

'        ' Event handler for 'Committing' event.
'        Private Sub MyInstaller_Committing(ByVal sender As Object, ByVal e As InstallEventArgs)
'            'Console.WriteLine("");
'            'Console.WriteLine("Committing Event occurred.");
'            'Console.WriteLine("");
'        End Sub

'        ' Event handler for 'Committed' event.
'        Private Sub MyInstaller_Committed(ByVal sender As Object, ByVal e As InstallEventArgs)
'            Try
'                Directory.SetCurrentDirectory(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))
'                Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\QLNS.exe")
'                ' Do nothing... 
'            Catch
'            End Try
'        End Sub
'        '-----------------------------------------------------------------------
'        '' Override the 'Install' method.
'        'Public Overloads Overrides Sub Install(ByVal savedState As IDictionary)
'        '    MyBase.Install(savedState)
'        'End Sub
'        '-----------------------------------------------------------------------
'        ' Override the 'Install' method.
'        Public Overloads Overrides Sub Install(ByVal savedState As IDictionary)
'            MyBase.Install(savedState)
'            Try
'                ' In order to get the value from the textBox named 'EDITA1' I needed to add the line:
'                ' '/PathValue = [EDITA1]' to the CustomActionData property of the CustomAction We added. 
'                Dim filesPath As String = Context.Parameters("PathValue")

'                ' Get the path to the executable file that is being installed on the target computer
'                Dim assemblypath As String = Context.Parameters("assemblypath")
'                Dim appConfigPath As String = assemblypath & ".config"
'                ' Write the path to the app.config file
'                Dim doc As XmlDocument = New XmlDocument()
'                doc.Load(appConfigPath)

'                Dim configuration As XmlNode = Nothing
'                For Each node As XmlNode In doc.ChildNodes
'                    If node.Name = "configuration" Then configuration = node
'                Next
'                If Not (configuration Is Nothing) Then
'                    ' Get the 'appSettings' node
'                    Dim settingNode As XmlNode = Nothing
'                    For Each node As XmlNode In configuration.ChildNodes
'                        If node.Name = "appSettings" Then settingNode = node
'                    Next
'                    If Not (settingNode Is Nothing) Then
'                        ' Get the node with the attribute key="FilePath"
'                        Dim NumNode As XmlNode = Nothing
'                        For Each node As XmlNode In settingNode.ChildNodes
'                            If Not (node.Attributes("key") Is Nothing) Then
'                                If node.Attributes("key").Value = "FilePath" Then NumNode = node
'                            End If
'                        Next
'                        If Not (NumNode Is Nothing) Then
'                            Dim att As XmlAttribute = NumNode.Attributes("value")
'                            att.Value = filesPath
'                            ' Update the configuration file
'                            ' Save the configuration file
'                            doc.Save(appConfigPath)
'                        End If
'                    End If
'                End If
'            Catch ex As FormatException
'                Console.WriteLine(ex.Message)
'            End Try
'        End Sub
'        '-----------------------------------------------------------------------
'        'Public Overloads Overrides Sub Install(ByVal savedState As IDictionary)
'        '    MyBase.Install(savedState)

'        '    ' Retrieve configuration settings
'        '    Dim targetSite As String = Context.Parameters("TargetSite")
'        '    Dim targetVDir As String = Context.Parameters("TargetVdir")
'        '    Dim targetDir As String = Context.Parameters("TargetDir")

'        '    If (targetSite.StartsWith("/LM/")) Then targetSite = targetSite.Substring(4)
'        '    writeEventLogs()
'        '    'adjust the web.config file.
'        '    createConnectionString(targetSite, targetVDir, targetDir)
'        '    'create ASP.Net 2 site...
'        '    RegisterScriptMaps(targetSite, targetVDir)
'        'End Sub


'        'Private Sub RegisterScriptMaps(ByVal targetSite As String, ByVal targetVDir As String)
'        '    ' Calculate Windows path
'        '    Dim sysRoot As String = System.Environment.GetEnvironmentVariable("SystemRoot")
'        '    ' get the latest .net framework directory
'        '    Dim di As DirectoryInfo = New DirectoryInfo(sysRoot + "/Microsoft.NET/Framework")
'        '    Dim frameworkDir As DirectoryInfo() = di.GetDirectories("v2.0.*", SearchOption.TopDirectoryOnly)
'        '    Dim big As Integer = 0
'        '    For i As Int32 = 0 To i < frameworkDir.Length
'        '        Dim current As Integer = Convert.ToInt32(frameworkDir(i).Name.Substring(5))
'        '        If (current > big) Then big = current
'        '    Next
'        '    Dim latestFramework As String = di.FullName + "\\v2.0." + big.ToString()

'        '    ' Launch aspnet_regiis.exe utility to configure mappings
'        '    Dim info As ProcessStartInfo = New ProcessStartInfo()
'        '    info.FileName = latestFramework + "\\aspnet_regiis.exe"
'        '    info.Arguments = String.Format("-s {0}/ROOT/{1}", targetSite, targetVDir)
'        '    info.CreateNoWindow = True
'        '    info.UseShellExecute = False
'        '    Process.Start(info)
'        'End Sub

'        'Private Sub writeEventLogs()
'        '    'for debugging purposes write these values to eventlog.
'        '    'EventLog.WriteEntry("o!Web Server", Context.Parameters["Server"]);
'        '    'EventLog.WriteEntry("o!Web User", Context.Parameters["Username"]);
'        '    'EventLog.WriteEntry("o!Web Password", Context.Parameters["Password"]);
'        '    'EventLog.WriteEntry("o!Web Target Directory", Context.Parameters["TargetDir"]);
'        '    'EventLog.WriteEntry("o!Web Target Site", Context.Parameters["TargetSite"]);
'        '    'EventLog.WriteEntry("o!Web Virtual Directory", Context.Parameters["TargetVdir"]);
'        'End Sub

'        'Private Sub createConnectionString(ByVal targetSite As String, ByVal targetVDir As String, ByVal targetDir As String)
'        '    ' Retrieve "Friendly Site Name" from IIS for TargetSite
'        '    Dim entry As DirectoryEntry = New DirectoryEntry("IIS://LocalHost/" + targetSite)
'        '    Dim friendlySiteName As String = entry.Properties("ServerComment").Value.ToString()

'        '    'set the connection timeout property
'        '    entry.Properties("ConnectionTimeout")(0) = 1800
'        '    entry.CommitChanges()

'        '    entry = New DirectoryEntry("IIS://LocalHost/" + targetSite + "/ROOT/" + targetVDir)
'        '    entry.Properties("AspScriptTimeout")(0) = 1800
'        '    entry.Properties("AspSessionTimeout")(0) = 60
'        '    entry.CommitChanges()

'        '    Dim fSecurity As FileSecurity = File.GetAccessControl(targetDir + "App_Data")
'        '    'string DomainUserName = Environment.UserDomainName + "\\IUSR_" + Environment.MachineName;
'        '    Dim LocalUserName As String = Environment.MachineName + "\\IUSR_" + Environment.MachineName
'        '    Dim LocalGroupName As String = Environment.MachineName + "\\IIS_WPG"
'        '    ' Add the FileSystemAccessRule to the security settings. 
'        '    Try
'        '        fSecurity.AddAccessRule(New FileSystemAccessRule(LocalGroupName, FileSystemRights.FullControl, AccessControlType.Allow))
'        '        fSecurity.AddAccessRule(New FileSystemAccessRule(LocalUserName, FileSystemRights.FullControl, AccessControlType.Allow))
'        '        'fSecurity.AddAccessRule(new FileSystemAccessRule(DomainUserName, FileSystemRights.Write, AccessControlType.Allow))
'        '        File.SetAccessControl(targetDir + "App_Data", fSecurity)
'        '    Catch ex As Exception

'        '    End Try
'        '   ' Open Application's Web.Config
'        '    Dim config As Configuration = WebConfigurationManager.OpenWebConfiguration("/" + targetVDir, friendlySiteName)
'        '    config.AppSettings.Settings("installpath").Value = targetDir.Substring(0, targetDir.Length - 1)
'        '    config.AppSettings.Settings("DBuser").Value = login()
'        '    config.AppSettings.Settings("adminEmail").Value = Context.Parameters("adminEmail")
'        '    config.ConnectionStrings.ConnectionStrings("DBConnectionString").ConnectionString = connectionString("oDirectv3")
'        '    config.ConnectionStrings.ConnectionStrings("oDirectConnectionString").ConnectionString = connectionString("oDirectv3")
'        '    (CompilationSection)config.GetSection("system.web/compilation").Debug = false
'        '    'save web.config
'        '    config.Save()
'        'End Sub

'        'Private Function connectionString(ByVal db As String) As String
'        '    Dim server As String = String.Empty
'        '    If (Context.Parameters("Server") <> String.Empty) Then server &= "Data Source=" & Context.Parameters("Server") + ";Initial Catalog=" + db + ";Persist Security Info=True;" + login()
'        '    Return server
'        'End Function

'        'Private Function login() As String
'        '    'User ID=o;Password=abacus
'        '    If ((Context.Parameters("Username") <> String.Empty) And (Context.Parameters("Password") <> String.Empty)) Then login += "User ID=" + Context.Parameters("Username") + ";Password=" + Context.Parameters("Password")
'        '    Return login
'        'End Function

'        'Public Overloads Overrides Sub Uninstall(ByVal savedState As IDictionary)
'        '    MyBase.Uninstall(savedState)
'        'End Sub

'        '-----------------------------------------------------------------------
'        ' Override the 'Commit' method.
'        Public Overloads Overrides Sub Commit(ByVal savedState As IDictionary)
'            MyBase.Commit(savedState)
'        End Sub

'        ' Override the 'Rollback' method.
'        Public Overloads Overrides Sub Rollback(ByVal savedState As IDictionary)
'            MyBase.Rollback(savedState)
'        End Sub
'    End Class
'End Namespace