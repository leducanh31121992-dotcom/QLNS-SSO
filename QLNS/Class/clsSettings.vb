Imports System.Configuration
Imports System.Xml

Public Class clsSettings
    Dim xDoc As New XmlDocument
    Dim XmlRdr As XmlReader
    Dim XmlWrt As XmlWriter

    Sub New()
        'check if file myxml.xml is existing
        If Not IO.File.Exists(gConfigFile) Then
            _CreateConfigFile()
        End If
        xDoc.Load(gConfigFile)
        XmlRdr = New XmlTextReader(gConfigFile)
    End Sub

    Sub _CreateConfigFile()
        If IO.File.Exists(gConfigFile) Then Exit Sub
        Dim settings As New XmlWriterSettings()
        'lets tell to our xmlwritersettings that it must use indention for our xml
        settings.Indent = True
        XmlWrt = XmlWriter.Create(gConfigFile, settings)

        With XmlWrt
            ' Write the Xml declaration.
            .WriteStartDocument()

            ' Write a comment.
            .WriteComment("Cấu hình các tham số cho phần mềm Quản lý nhân sự")

            ' Write the root element.
            .WriteStartElement("Main")

            .WriteStartElement("MAX_KY")
            .WriteValue(1)
            .WriteEndElement()
            .WriteStartElement("TINH_THUE_TNCN")
            .WriteValue(1)
            .WriteEndElement()
            .WriteStartElement("GIAMDOC")
            .WriteValue("...")
            .WriteEndElement()
            .WriteStartElement("PHOGIAMDOC")
            .WriteValue("...")
            .WriteEndElement()
            .WriteStartElement("TRUONGHC-TC")
            .WriteValue("...")
            .WriteEndElement()
            .WriteStartElement("KETOANTRUONG")
            .WriteValue("...")
            .WriteEndElement()
            .WriteStartElement("NGUOILAPBIEU")
            .WriteValue("...")
            .WriteEndElement()
            .WriteStartElement("TM-ExportFiles")
            .WriteValue("C:\")
            .WriteEndElement()

            .WriteEndElement() 'Main

            ' Close the XmlTextWriter.
            .WriteEndDocument()
            .Close()
        End With

    End Sub

    ''' <summary>
    ''' Đọc giá trị của key từ config file, nếu không tìm được key thì trả về chuỗi rỗng
    ''' </summary>
    Function ReadValue(key As String) As String
        Dim sReturn$ = ""
        Try
            Dim xNode As XmlNode = xDoc.GetElementsByTagName(key)(0)
            sReturn = xNode.InnerText
        Catch ex As Exception

        End Try
        Return sReturn
    End Function

    Sub WriteValue(key As String, Value As String)
        Dim xNode As XmlNode
        If Not CheckIfKeyExists(key) Then
            xNode = xDoc.GetElementsByTagName("Main")(0)
            Dim xNodeChild As XmlNode = xDoc.CreateNode(Xml.XmlNodeType.Element, key, "")
            xNode.AppendChild(xNodeChild)
        End If
        xNode = xDoc.GetElementsByTagName(key)(0)
        xNode.InnerText = Value
    End Sub

    Sub Save()
        xDoc.Save(gConfigFile)
    End Sub

    Function CheckIfKeyExists(key As String) As Boolean
        Dim bReturn As Boolean = (xDoc.GetElementsByTagName(key).Count > 0)
        Return bReturn
    End Function
End Class
