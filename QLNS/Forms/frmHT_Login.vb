Imports System.Collections.Generic
Imports System.Configuration
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Web
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json.Linq

Public Class frmHT_Login

#Region "--->Khai báo thuộc tính và khởi tạo đối tượng<---"
    Private _UserName As String
    Public Property UserName() As String
        Get
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property

    Private _Permits As String
    Public Property Permits() As String
        Get
            Return _Permits
        End Get
        Set(ByVal value As String)
            _Permits = value
        End Set
    End Property

    Private _Admin As Boolean
    Public Property Admin() As Boolean
        Get
            Return _Admin
        End Get
        Set(ByVal value As Boolean)
            _Admin = value
        End Set
    End Property

    Private _Systems As clsHeThong = New clsHeThong
    Private _Globals As Globals = New Globals
    Dim _DBAccess As DBAccess = New DBAccess()
    Public Delegate Sub ProgressChangedEventHandler()
    Public Progress_Changed As ProgressChangedEventHandler

    Private WithEvents webView As Microsoft.Web.WebView2.WinForms.WebView2

    ' Biến cờ chống gọi lặp nhiều lần sự kiện chuyển hướng WebView2
    Private isProcessingLogin As Boolean = False
#End Region

#Region "--->Events: Các sự kiện chính<---"
    Private Sub frmHT_Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Rào/ẩn các control đăng nhập cũ
        'If edt_username IsNot Nothing Then edt_username.Visible = False
        'If edt_password IsNot Nothing Then edt_password.Visible = False
        'If edt_maPOS IsNot Nothing Then edt_maPOS.Visible = False
        'If btn_login IsNot Nothing Then btn_login.Visible = False
        'If btn_reset IsNot Nothing Then btn_reset.Visible = False
        'If chkLuuThongTin IsNot Nothing Then chkLuuThongTin.Visible = False


        InitializeWebView()
    End Sub

    Private Sub frmHT_Login_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If Not (Progress_Changed Is Nothing) Then
            Progress_Changed()
        End If
    End Sub
#End Region

#Region "--->Xử lý SSO (WebView2 & Chuỗi API)<---"
    Private Async Sub InitializeWebView()
        Try
            ' Đọc cấu hình từ App.config
            Dim baseUrl As String = ConfigurationManager.AppSettings("SSO_BaseUrl")
            Dim clientId As String = ConfigurationManager.AppSettings("SSO_ClientId")
            Dim redirectUri As String = ConfigurationManager.AppSettings("SSO_RedirectUri")
            Dim scope As String = ConfigurationManager.AppSettings("SSO_Scope")
            Dim state As String = ConfigurationManager.AppSettings("SSO_State")
            Dim nonce As String = ConfigurationManager.AppSettings("SSO_Nonce")
            Dim codeChallenge As String = ConfigurationManager.AppSettings("SSO_CodeChallenge")
            Dim codeChallengeMethod As String = ConfigurationManager.AppSettings("SSO_CodeChallengeMethod")

            ' 1. Khởi động HttpListener để lắng nghe RedirectUri chạy cục bộ (Ví dụ: http://localhost:8080/callback/)
            Dim listener As New HttpListener()
            listener.Prefixes.Add(redirectUri)
            listener.Start()

            ' 2. Mở trình duyệt mặc định của hệ thống với URL đăng nhập SSO
            Dim authUrl As String = $"{baseUrl}/api/v1/oauth2/authorize?response_type=code&client_id={clientId}&redirect_uri={HttpUtility.UrlEncode(redirectUri)}&scope={HttpUtility.UrlEncode(scope)}&state={state}&nonce={nonce}&code_challenge={codeChallenge}&code_challenge_method={codeChallengeMethod}"

            Process.Start(New ProcessStartInfo(authUrl) With {.UseShellExecute = True})

            ' 3. Lắng nghe phản hồi từ trình duyệt gửi về qua Local Server một cách bất đồng bộ
            Dim unused = Task.Run(Async Function()
                                      Try
                                          While listener.IsListening
                                              Dim context As HttpListenerContext = Await listener.GetContextAsync()
                                              Dim requestUrl = context.Request.Url

                                              ' Trả về một trang thông báo nhỏ cho người dùng đóng tab trình duyệt
                                              Dim htmlContent As String =
                 "<!DOCTYPE html>" &
                 "<html>" &
                 "<head>" &
                 "    <meta charset='utf-8'>" &
                 "    <title>Xác thực thành công</title>" &
                 "    <style>" &
                 "        body {" &
                 "            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;" &
                 "            background-color: #f4f6f9;" &
                 "            display: flex;" &
                 "            justify-content: center;" &
                 "            align-items: center;" &
                 "            height: 100vh;" &
                 "            margin: 0;" &
                 "        }" &
                 "        .card {" &
                 "            background: #ffffff;" &
                 "            padding: 40px;" &
                 "            border-radius: 12px;" &
                 "            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);" &
                 "            text-align: center;" &
                 "            max-width: 400px;" &
                 "            width: 100%;" &
                 "        }" &
                 "        .icon {" &
                 "            font-size: 50px;" &
                 "            color: #28a745;" &
                 "            margin-bottom: 20px;" &
                 "        }" &
                 "        h2 {" &
                 "            color: #0366cc;" &
                 "            margin-bottom: 10px;" &
                 "            font-size: 22px;" &
                 "        }" &
                 "        p {" &
                 "            color: #555555;" &
                 "            font-size: 15px;" &
                 "            line-height: 1.5;" &
                 "            margin-top: 0;" &
                 "        }" &
                 "    </style>" &
                 "</head>" &
                 "<body>" &
                 "    <div class='card'>" &
                 "        <div class='icon'>&#10004;</div>" &
                 "        <h2>Đăng nhập thành công!</h2>" &
                 "        <p>Hệ thống đã xác thực tài khoản của bạn.<br>Bạn có thể đóng cửa sổ trình duyệt này và quay trở lại ứng dụng.</p>" &
                 "    </div>" &
                 "</body>" &
                 "</html>"

                                              Dim responseBytes = System.Text.Encoding.UTF8.GetBytes(htmlContent)
                                              context.Response.ContentLength64 = responseBytes.Length
                                              Await context.Response.OutputStream.WriteAsync(responseBytes, 0, responseBytes.Length)
                                              context.Response.OutputStream.Close()

                                              ' Lấy authorization code từ Query string
                                              Dim query As System.Collections.Specialized.NameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query)
                                              Dim authorizationCode As String = query("code")

                                              If Not String.IsNullOrEmpty(authorizationCode) Then
                                                  System.Diagnostics.Debug.WriteLine("code: " & authorizationCode)

                                                  ' Gọi API đổi token ngầm
                                                  Await ExecuteSSOFlowAsync(authorizationCode)
                                              End If

                                              ' Dừng listener sau khi đã nhận được code
                                              listener.Stop()
                                              Exit While
                                          End While
                                      Catch ex As Exception
                                          System.Diagnostics.Debug.WriteLine(ex.Message)
                                      End Try
                                  End Function)

        Catch ex As Exception
            MessageBox.Show("Không thể mở trình duyệt đăng nhập: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Function ExecuteSSOFlowAsync(authorizationCode As String) As Task
        Using client As New HttpClient()

            ' Bước 2: Đổi Code lấy Token
            Dim accessToken As String = Await Step2_ExchangeToken(client, authorizationCode)
            If String.IsNullOrEmpty(accessToken) Then
                Return
            End If

            ' Bước 3: Lấy thông tin UserInfo
            Dim username As String = Await Step3_GetUserInfo(client, accessToken)
            If Not String.IsNullOrEmpty(username) Then
                _UserName = username
                gUsername = username
            End If

            ' Bước 4: Lấy Permissions và Menu, đồng thời map vào Globals
            Await Step4_GetPermissionsAndMenu(client, accessToken)

            ' Kiểm tra phân quyền hợp lệ trước vào chương trình
            If String.IsNullOrEmpty(Globals.Roles) Then
                If Not Me.IsDisposed Then
                    Me.Invoke(Sub()
                                  MessageBox.Show("Thành viên đăng nhập hiện chưa được phân quyền thao tác chương trình!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                  Me.Close()
                              End Sub)
                End If
                Return
            End If

            ' Đóng form và trả về DialogResult.OK an toàn tuyệt đối
            If Not Me.IsDisposed Then
                Me.Invoke(Sub()
                              If Not Me.IsDisposed Then
                                  Me.DialogResult = Windows.Forms.DialogResult.OK
                                  Me.Close()
                              End If
                          End Sub)
            End If
        End Using
    End Function
#End Region
#Region "--->Các bước xử lý API SSO<---"
    ' --- BƯỚC 2: Đổi Code lấy Access Token ---
    Private Async Function Step2_ExchangeToken(client As HttpClient, authorizationCode As String) As Task(Of String)
        Dim baseUrl As String = ConfigurationManager.AppSettings("SSO_BaseUrl")
        Dim tokenEndpoint As String = $"{baseUrl}/api/v1/oauth2/token"

        Dim postData As New Dictionary(Of String, String) From {
            {"grant_type", "authorization_code"},
            {"code", authorizationCode},
            {"redirect_uri", ConfigurationManager.AppSettings("SSO_RedirectUri")},
            {"client_id", ConfigurationManager.AppSettings("SSO_ClientId")},
            {"client_secret", ConfigurationManager.AppSettings("SSO_ClientSecret")},
            {"code_verifier", ConfigurationManager.AppSettings("SSO_CodeVerifier")}
        }

        Try
            Dim content As New FormUrlEncodedContent(postData)
            Dim response As HttpResponseMessage = Await client.PostAsync(tokenEndpoint, content)

            Dim jsonResponse = Await response.Content.ReadAsStringAsync()

            If Not response.IsSuccessStatusCode Then
                Return Nothing
            End If

            Dim tokenObj As JObject = JObject.Parse(jsonResponse)
            Return tokenObj("access_token")?.ToString()

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' --- BƯỚC 3: Lấy thông tin UserInfo ---
    Private Async Function Step3_GetUserInfo(client As HttpClient, accessToken As String) As Task(Of String)
        Dim baseUrl As String = ConfigurationManager.AppSettings("SSO_BaseUrl")
        Dim userInfoEndpoint As String = $"{baseUrl}/api/v1/oauth2/userinfo"

        Try
            client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken)
            Dim response As HttpResponseMessage = Await client.GetAsync(userInfoEndpoint)

            Dim jsonResponse = Await response.Content.ReadAsStringAsync()
            'System.Diagnostics.Debug.WriteLine("user info: " & jsonResponse)

            If Not response.IsSuccessStatusCode Then
                Return Nothing
            End If

            Dim userObj As JObject = JObject.Parse(jsonResponse)
            Return userObj("preferred_username")?.ToString()

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' --- BƯỚC 4: Lấy danh sách Quyền, Menu và Map vào biến Global cũ ---
    Private Async Function Step4_GetPermissionsAndMenu(client As HttpClient, accessToken As String) As Task
        Dim baseUrl As String = ConfigurationManager.AppSettings("SSO_BaseUrl")
        Dim clientId As String = ConfigurationManager.AppSettings("SSO_ClientId")
        Dim permissionsEndpoint As String = $"{baseUrl}/api/v1/admin/user/permissions?client-id={clientId}"

        Try
            client.DefaultRequestHeaders.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken)
            Dim response As HttpResponseMessage = Await client.GetAsync(permissionsEndpoint)

            Dim jsonResponse = Await response.Content.ReadAsStringAsync()
            'System.Diagnostics.Debug.WriteLine("menu info: " & jsonResponse)

            If Not response.IsSuccessStatusCode Then
                ' Fallback về hàm cũ nếu lỗi gọi API
                Globals.Roles = _Systems.GetRoles(_UserName)
                Globals.Group = _Systems.GetGroup(_UserName)
                Globals.QuyenQuanLyCN = _Systems.Get_QuyenQuanLyCN(_UserName)
                Return
            End If

            ' Parse JSON phản hồi từ API theo đúng chuẩn cú pháp VB.NET
            Dim permObj As JObject = JObject.Parse(jsonResponse)
            Dim rolesArr As JArray = TryCast(permObj("roles"), JArray)
            Dim roleList As New List(Of String)()

            If rolesArr IsNot Nothing Then
                For Each r As JObject In rolesArr
                    Dim resourcesObj As JObject = TryCast(r("resources"), JObject)
                    If resourcesObj IsNot Nothing Then
                        For Each resProp As JProperty In resourcesObj.Properties()
                            roleList.Add(resProp.Name) ' Thêm resource ID
                            Dim actionsArr As JArray = TryCast(resProp.Value("actions"), JArray)
                            If actionsArr IsNot Nothing Then
                                For Each act As JToken In actionsArr
                                    roleList.Add(act.ToString()) ' Thêm action ID
                                Next
                            End If
                        Next
                    End If
                Next
            End If

            If roleList.Count > 0 Then
                Globals.Roles = ";" & String.Join(";", roleList) & ";"
            Else
                Globals.Roles = _Systems.GetRoles(_UserName)
            End If

            ' Lấy thêm Group và Quyền quản lý chi nhánh từ cơ sở dữ liệu hệ thống
            Globals.Group = _Systems.GetGroup(_UserName)
            Globals.QuyenQuanLyCN = _Systems.Get_QuyenQuanLyCN(_UserName)

        Catch ex As Exception
            ' Xử lý fallback an toàn
            Globals.Roles = _Systems.GetRoles(_UserName)
            Globals.Group = _Systems.GetGroup(_UserName)
            Globals.QuyenQuanLyCN = _Systems.Get_QuyenQuanLyCN(_UserName)
        End Try
    End Function

#End Region

End Class