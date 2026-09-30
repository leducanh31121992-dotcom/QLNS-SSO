Module modConstants
    'Môi trường CSDL chính thức
    Public Server As String = "10.63.48.63" 'Configuration.ConfigurationSettings.AppSettings("Server")
    Private DataBase As String = "QLNS_NHCSXH" 'Configuration.ConfigurationSettings.AppSettings("Database")
    Private User As String = "sa"  'Configuration.ConfigurationSettings.AppSettings("User")
    Private Pass As String = "Sql2017"  'Configuration.ConfigurationSettings.AppSettings("Password")
    Public QLNS_CONNSTR As String = "Data Source=" & Server & ";Initial Catalog=" & DataBase & ";Persist Security Info=True;User ID=" & User & ";Password=" & Pass

    ''Môi trường CSDL DEV
    'Public Server As String = "10.63.48.63" 'Configuration.ConfigurationSettings.AppSettings("Server")
    'Private DataBase As String = "QLNS_NHCSXH" 'Configuration.ConfigurationSettings.AppSettings("Database")
    'Private User As String = "sa"  'Configuration.ConfigurationSettings.AppSettings("User")
    'Private Pass As String = "Sql2017"  'Configuration.ConfigurationSettings.AppSettings("Password")
    'Public QLNS_CONNSTR As String = "Data Source=" & Server & ";Initial Catalog=" & DataBase & ";Persist Security Info=True;User ID=" & User & ";Password=" & Pass

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Public Server1 As String = ".\SQLEXPRESS"
    'Private DataBase1 As String = "QLNS_NHCSXH" 'Configuration.ConfigurationSettings.AppSettings("Database")
    'Private User1 As String = "chudv"  'Configuration.ConfigurationSettings.AppSettings("User")
    'Private Pass1 As String = "chudv2510"  'Configuration.ConfigurationSettings.AppSettings("Password")
    'Public QLNS_CONNSTR As String = "Data Source=" & Server1 & ";Initial Catalog=" & DataBase1 & ";Persist Security Info=True;User ID=" & User1 & ";Password=" & Pass1


    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Viết hướng dẫn
    'Public Server1 As String = "10.63.16.65"
    'Private DataBase1 As String = "QLNS_NHCSXH_DEV" 'Configuration.ConfigurationSettings.AppSettings("Database")
    'Private User1 As String = "qlns_temp_readonly"  'Configuration.ConfigurationSettings.AppSettings("User")
    'Private Pass1 As String = "VbsptemP"  'Configuration.ConfigurationSettings.AppSettings("Password")
    'Public QLNS_CONNSTR As String = "Data Source=" & Server1 & ";Initial Catalog=" & DataBase1 & ";Persist Security Info=True;User ID=" & User1 & ";Password=" & Pass1
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public HSC As String = "1"      ' Hội sở chính
    Public TINH As String = "2"     ' Tỉnh thành
    Public VPDD As String = "3"     ' Văn phòng đại diện
    Public TTCNTT As String = "5"   ' Trung tâm công nghệ thông tin
    Public TTDT As String = "7"     ' Trung tâm đào tạo
    Public SGD As String = "9"      ' Sở giao dịch
    Public HUYEN As String = "4"    ' Huyện
    'Theo CV1205/NHCS-TCCB: 30/7/2011 hoàn thành nhập dữ liệu, gửi dữ liệu định kỳ vào 07 hàng tháng
    Public vDateCheckAll As Date = CDate("October 31, 2012")

    Public gConfigFile As String = Application.StartupPath & "\" & "QLNS.cfg"
    Public Const gMaDonViTW As String = "000100"
    Public Const gMaNhomCanBoTCCB = "S02"
End Module
