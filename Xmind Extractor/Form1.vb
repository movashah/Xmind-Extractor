Imports Microsoft.Win32
Imports System.IO
Imports System.IO.Compression
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Globalization
Imports System.Text
Imports OfficeOpenXml

Public Class Form1
    Dim LastPath As String = ""
    Dim xFormat As String = ""
    Dim xContent As String = ""
    Dim LineTedad As Integer = 30
    Dim Level As Integer = 0
    Dim LevelNum(100) As Integer
    Dim CurLine As Integer = 0
    Dim StopSample As Boolean = False
    Dim FullText As String = ""
    Dim XmindFile As String
    Dim OnvanAsli As String = ""
    Dim XMLsample As Boolean = False
    Dim TYelow As Color

    Private Sub CheckFramework()
        ' بررسی وجود دات‌نت فریم‌ورک 4.7.2 یا بالاتر
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\")
            If key Is Nothing Then
                ' فریم‌ورک 4.5 یا بالاتر نصب نیست
                ShowDownloadMessage()
                Return
            End If

            Dim releaseKey As Integer = Convert.ToInt32(key.GetValue("Release"))

            ' اعداد مربوط به نسخه‌ها: 
            ' 461808 برای 4.7.2
            ' 528040 برای 4.8
            If releaseKey < 461808 Then
                ' فریم‌ورک 4.7.2 یا بالاتر نصب نیست
                ShowDownloadMessage()
            Else
                ' فریم‌ورک مورد نیاز نصب است - ادامه اجرای برنامه
                ' MessageBox.Show("فریم‌ورک مورد نیاز نصب است.")
            End If
        End Using
    End Sub

    Private Sub ShowDownloadMessage()
        Dim result As DialogResult = MessageBox.Show(
        "این برنامه نیاز به Microsoft .NET Framework 4.7.2 یا بالاتر دارد." & vbCrLf &
        "آیا می‌خواهید هم‌اکنون به صفحه دانلود رسمی مایکروسافت هدایت شوید؟",
        "نیاز به پیش‌نیاز",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Exclamation)

        If result = DialogResult.Yes Then
            ' باز کردن لینک دانلود در مرورگر پیش‌فرض کاربر
            Process.Start("https://go.microsoft.com/fwlink/?linkid=863265")
        End If

        ' خروج از برنامه (چون پیش‌نیاز نصب نیست)
        Application.Exit()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Level = 0
        If Len(TextBox2.Text) = 0 Then TextBox2.Text = "."
        If Len(TextBox3.Text) = 0 Then TextBox3.Text = "/"
        If Len(TextBox5.Text) = 0 Or Not IsNumeric(TextBox5.Text) Then TextBox5.Text = "1"
        With OpenFileDialog1
            .Filter = "فایل‌های XMind (*.xmind)|*.xmind|همه فایل‌ها (*.*)|*.*"
            .FilterIndex = 1
            .FileName = ""
            'If Not Len(LastPath) > 0 Then .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) Else .InitialDirectory = LastPath
            .Title = "انتخاب فایل XMind"
            .RestoreDirectory = True
            If .ShowDialog() = DialogResult.OK Then
                Dim filePath As String = .FileName
                'MessageBox.Show("فایل انتخاب شد: " & filePath)
                LastPath = Path.GetDirectoryName(filePath)
                XmindFile = filePath
                GetXmind(filePath)
            End If
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FullText = ""
        Level = 0
        If RadioButton1.Checked Then 'docx
            Write2docx()
        ElseIf RadioButton2.Checked Then 'txt
            Write2txt()
            Try
                File.WriteAllText(XmindFile & ".txt", FullText, System.Text.Encoding.UTF8)
                Process.Start(XmindFile & ".txt")
            Catch ex As Exception
                MessageBox.Show("خطا در ذخيره‌سازي: " & ex.Message)
            End Try
        ElseIf RadioButton3.Checked Then 'xml
            Write2xml()
        ElseIf RadioButton4.Checked Then 'json
            Write2json()
        ElseIf RadioButton5.Checked Then
            Write2Excel()
        End If
    End Sub

    Public Sub ExtractXMLdocx(xmlContent As String)
        Try
            Dim doc As XDocument = XDocument.Parse(xContent)
            Dim result As New StringBuilder()
            For Each el In doc.Descendants()
                el.Name = el.Name.LocalName
            Next
            For Each sheet In doc.Descendants("sheet")
                LevelNum(Level) = 1
                Dim sheetTitle = sheet.Element("title")
                If sheetTitle IsNot Nothing Then
                    OnvanAsli += sheetTitle.Value.Trim() & " "
                End If
                Dim topic = sheet.Element("topic")
                If topic IsNot Nothing Then
                    Level += 1
                    LevelNum(Level) = 1
                    ProcessTopicXMLdocx(topic, result)
                End If
            Next
        Catch ex As Exception
            MsgBox($"خطا: {ex.Message}")
        End Try
    End Sub

    Private Sub ProcessTopicXMLdocx(topic As XElement, ByRef result As StringBuilder)
        Dim inlineTemp As String = "<w:p w:rsidR=""001843B4"" w:rsidRDefault=""001843B4"" w:rsidP=""00A66AB8""><w:pPr><w:pStyle w:val=""ListParagraph""/><w:numPr><w:ilvl w:val=""{سطح}""/><w:numId w:val=""28""/></w:numPr><w:rPr><w:rFonts w:ascii=""Lotus YA"" w:hAnsi=""Lotus YA""/></w:rPr></w:pPr><w:r w:rsidRPr=""00A66AB8""><w:rPr><w:rFonts w:ascii=""Lotus YA"" w:hAnsi=""Lotus YA""/><w:rtl/></w:rPr><w:t>{متن}</w:t></w:r></w:p>"
        If topic Is Nothing Then Exit Sub
        Dim title = topic.Element("title")
        If title IsNot Nothing Then
            Dim titleText = title.Value.Trim()
            If Not String.IsNullOrWhiteSpace(titleText) Then
                FullText += Replace(Replace(inlineTemp, "{سطح}", Level - 1), "{متن}", titleText)
            End If
        End If
        ' فرزندان
        Dim children = topic.Element("children")
        If children IsNot Nothing Then
            Dim topics = children.Element("topics")
            If topics IsNot Nothing Then
                Level += 1
                For Each child In topics.Elements("topic")
                    ProcessTopicXMLdocx(child, result)
                Next
                'LevelNum(level) = 0
                Level -= 1
            End If
        End If
    End Sub

    Public Function ExtractXMLjson(xmlContent As String) As String
        Try
            Dim doc As XDocument = XDocument.Parse(xContent)
            Dim result As New StringBuilder()
            For Each el In doc.Descendants()
                el.Name = el.Name.LocalName
            Next
            For Each sheet In doc.Descendants("sheet")
                Dim sheetTitle = sheet.Element("title")
                If sheetTitle IsNot Nothing Then
                    Dim CurTab As New String(vbTab, Level)
                    result.AppendLine(CurTab & Replace(Replace(sheetTitle.Value.Trim(), vbCr, ""), vbLf, ""))
                End If
                Dim topic = sheet.Element("topic")
                If topic IsNot Nothing Then
                    ProcessTopicXMLjson(topic, result)
                End If
            Next
            Return result.ToString()
        Catch ex As Exception
            MsgBox($"خطا: {ex.Message}")
        End Try
    End Function

    Private Sub ProcessTopicXMLjson(topic As XElement, ByRef result As StringBuilder)
        If topic Is Nothing Then Exit Sub
        Dim title = topic.Element("title")
        If title IsNot Nothing Then
            Dim titleText = title.Value.Trim()
            If Not String.IsNullOrWhiteSpace(titleText) Then
                Dim CurTab As New String(vbTab, Level)
                result.AppendLine(CurTab & Replace(Replace(titleText, vbCr, ""), vbLf, ""))
            End If
        End If
        Dim children = topic.Element("children")
        If children IsNot Nothing Then
            Dim topics = children.Element("topics")
            If topics IsNot Nothing Then
                Level += 1
                For Each child In topics.Elements("topic")
                    ProcessTopicXMLjson(child, result)
                Next
                Level -= 1
            End If
        End If
    End Sub

    Public Function ExtractXML(xmlContent As String) As String
        Try
            Dim doc As XDocument = XDocument.Parse(xContent)
            Dim result As New StringBuilder()
            For Each el In doc.Descendants()
                el.Name = el.Name.LocalName
            Next
            For Each sheet In doc.Descendants("sheet")
                LevelNum(Level) = 1
                Dim sheetTitle = sheet.Element("title")
                If sheetTitle IsNot Nothing Then
                    CurLine += 1
                    If CurLine >= LineTedad And XMLsample Then Exit For
                    result.AppendLine(sheetTitle.Value.Trim())
                End If
                Dim topic = sheet.Element("topic")
                If topic IsNot Nothing Then
                    Level += 1
                    LevelNum(Level) = 1
                    ProcessTopicXML(topic, result)
                End If
            Next
            Return result.ToString()
        Catch ex As Exception
            MsgBox($"خطا: {ex.Message}")
            Return ""
        End Try
    End Function

    Sub Write2Excel()
        If Len(xContent) > 0 Then
            If xFormat = "json" Then
                Try
                    Dim jsonArray As JArray = JArray.Parse(xContent)
                    For Each item As JObject In jsonArray
                        If item("rootTopic") IsNot Nothing Then
                            'ExtractTitlejson(item("rootTopic"))
                            ExtractCSV(item("rootTopic"))
                        End If
                    Next
                    'File.WriteAllText(XmindFile & ".xml", IndentedTextToXml(FullText), System.Text.Encoding.UTF8)
                    ConvertCSVToExcel(FullText, XmindFile & ".xlsx")
                    Process.Start(XmindFile & ".xlsx")
                Catch ex As Exception
                    MessageBox.Show("خطا در پردازش JSON: " & ex.Message)
                End Try
            ElseIf xFormat = "xml" Then
                Try
                    'File.WriteAllText(XmindFile & ".xml", IndentedTextToXml(ExtractXMLjson(xContent)), System.Text.Encoding.UTF8)
                    ConvertCSVToExcel(ExtractCSVxml(xContent), XmindFile & ".xlsx")
                    Process.Start(XmindFile & ".xlsx")
                Catch ex As Exception
                    MessageBox.Show("خطا در ذخيره: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub ProcessTopicXML(topic As XElement, ByRef result As StringBuilder)
        If topic Is Nothing Then Exit Sub
        If StopSample And XMLsample Then Exit Sub
        ' عنوان
        'LevelNum(Level) = 1
        Dim title = topic.Element("title")
        If title IsNot Nothing Then
            Dim titleText = title.Value.Trim()
            If Not String.IsNullOrWhiteSpace(titleText) Then
                Dim JayTab As String = vbTab
                If Len(TextBox4.Text) > 0 Then JayTab = TextBox4.Text
                Dim Tabs As New String(JayTab, Level * TextBox5.Text)
                Dim Shomareh As String = ""
                If CheckBox2.Checked Then
                    For i As Integer = Level To 1 Step -1
                        Shomareh += LevelNum(i) & TextBox3.Text
                    Next
                Else
                    For i As Integer = 1 To Level
                        Shomareh += LevelNum(i) & TextBox3.Text
                    Next
                End If
                If Strings.Right(Shomareh, 1) = TextBox3.Text Then Shomareh = Strings.Left(Shomareh, Len(Shomareh) - 1)
                Dim CurLvl As String = ""
                If Not CheckBox3.Checked Then Tabs = ""
                If Not CheckBox1.Checked Then Shomareh = ""
                If Len(Shomareh) > 0 Then CurLvl = Shomareh & TextBox2.Text & " "
                If Len(Tabs) > 0 Then CurLvl = Tabs & CurLvl
                'result.Append(New String(ControlChars.Tab, level))
                result.AppendLine(CurLvl & titleText)
                CurLine += 1
                If CurLine >= LineTedad Then StopSample = True
            End If
        End If
        ' فرزندان
        Dim children = topic.Element("children")
        If children IsNot Nothing Then
            Dim topics = children.Element("topics")
            If topics IsNot Nothing Then
                Level += 1
                LevelNum(Level) = 0
                For Each child In topics.Elements("topic")
                    LevelNum(Level) += 1
                    ProcessTopicXML(child, result)
                Next
                'LevelNum(level) = 0
                Level -= 1
            End If
        End If
    End Sub

    Sub Write2xml()
        If Len(xContent) > 0 Then
            If xFormat = "json" Then
                Try
                    Dim jsonArray As JArray = JArray.Parse(xContent)
                    For Each item As JObject In jsonArray
                        If item("rootTopic") IsNot Nothing Then
                            ExtractTitlejson(item("rootTopic"))
                        End If
                    Next
                    File.WriteAllText(XmindFile & ".xml", IndentedTextToXml(FullText), System.Text.Encoding.UTF8)
                    'File.WriteAllText(XmindFile & ".xml", FullText, System.Text.Encoding.UTF8)
                    Process.Start("notepad.exe", XmindFile & ".xml")
                Catch ex As Exception
                    MessageBox.Show("خطا در پردازش JSON: " & ex.Message)
                End Try
            ElseIf xFormat = "xml" Then
                Try
                    File.WriteAllText(XmindFile & ".xml", IndentedTextToXml(ExtractXMLjson(xContent)), System.Text.Encoding.UTF8)
                    Process.Start("notepad.exe", XmindFile & ".xml")
                Catch ex As Exception
                    MessageBox.Show("خطا در ذخيره: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Function IndentedTextToXml(text As String) As String
        Dim lines = text.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        If lines.Length = 0 Then Return "<root />"

        Dim root As New XElement("root")
        Dim doc As New XDocument(New XDeclaration("1.0", "utf-8", "yes"), root)

        ' نگهداری آخرین گره در هر سطح
        Dim lastNodeAtLevel As New List(Of XElement)()

        For Each line In lines
            Dim level = line.TakeWhile(Function(c) c = ControlChars.Tab).Count()
            Dim content = line.TrimStart(ControlChars.Tab)
            Dim newNode = <node title=<%= content %>/>

            If level = 0 Then
                ' گره ریشه
                root.Add(newNode)
                ' تنظیم lastNodeAtLevel برای سطح 0
                While lastNodeAtLevel.Count <= level
                    lastNodeAtLevel.Add(Nothing)
                End While
                lastNodeAtLevel(level) = newNode

            Else
                ' پیدا کردن والد (گره سطح بالاتر)
                If level - 1 < lastNodeAtLevel.Count AndAlso lastNodeAtLevel(level - 1) IsNot Nothing Then
                    lastNodeAtLevel(level - 1).Add(newNode)
                End If

                ' به‌روزرسانی lastNodeAtLevel برای این سطح
                While lastNodeAtLevel.Count <= level
                    lastNodeAtLevel.Add(Nothing)
                End While
                lastNodeAtLevel(level) = newNode
            End If
        Next

        Return doc.ToString()
    End Function

    Sub Write2json()
        If Len(xContent) > 0 Then
            If xFormat = "json" Then
                Try
                    Dim jsonArray As JArray = JArray.Parse(xContent)
                    For Each item As JObject In jsonArray
                        If item("rootTopic") IsNot Nothing Then
                            ExtractTitlejson(item("rootTopic"))
                        End If
                    Next
                    File.WriteAllText(XmindFile & ".json", TextToNestedJson(FullText), System.Text.Encoding.UTF8)
                    Process.Start("notepad.exe", XmindFile & ".json")
                Catch ex As Exception
                    MessageBox.Show("خطا در پردازش JSON: " & ex.Message)
                End Try
            ElseIf xFormat = "xml" Then
                Try
                    File.WriteAllText(XmindFile & ".json", TextToNestedJson(ExtractXMLjson(xContent)), System.Text.Encoding.UTF8)
                    Process.Start("notepad.exe", XmindFile & ".json")
                Catch ex As Exception
                    MessageBox.Show("خطا در ذخيره: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub ExtractTitlejson(token As JToken)
        If token Is Nothing Then Exit Sub
        If token.Type = JTokenType.Object Then
            Dim obj As JObject = CType(token, JObject)
            If obj("title") IsNot Nothing Then
                Dim titleValue As String = obj("title").ToString()
                If Not String.IsNullOrWhiteSpace(titleValue) Then
                    Dim CurTab As New String(vbTab, Level)
                    FullText += CurTab & Replace(Replace(titleValue, vbCr, ""), vbLf, "") & vbCrLf
                End If
            End If
            If obj("children") IsNot Nothing Then
                Dim children As JObject = obj("children")
                If children("attached") IsNot Nothing Then
                    Level += 1
                    For Each child As JToken In children("attached")
                        ExtractTitlejson(child)
                    Next
                    Level -= 1
                End If
            End If
        End If
    End Sub

    Public Function TextToNestedJson(text As String) As String
        Dim lines = text.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        If lines.Length = 0 Then Return "{}"

        Dim root As New Dictionary(Of String, Object)()
        Dim stack As New Stack(Of Dictionary(Of String, Object))()
        Dim levelStack As New Stack(Of Integer)()

        For Each line In lines
            ' شمارش تب‌ها
            Dim jlevel = line.TakeWhile(Function(c) c = ControlChars.Tab).Count()
            Dim content = line.TrimStart(ControlChars.Tab)

            If jlevel = 0 Then
                ' گره ریشه
                root(content) = New List(Of Object)()
                stack.Clear()
                stack.Push(root)
                levelStack.Clear()
                levelStack.Push(0)
            Else
                ' تنظیم پشته بر اساس سطح
                While levelStack.Count > 0 AndAlso jlevel <= levelStack.Peek()
                    stack.Pop()
                    levelStack.Pop()
                End While

                If stack.Count = 0 Then Continue For

                ' والد فعلی
                Dim parent As Dictionary(Of String, Object) = stack.Peek()
                Dim parentKey = parent.Keys.Last()

                If TypeOf parent(parentKey) Is List(Of Object) Then
                    Dim children = CType(parent(parentKey), List(Of Object))
                    Dim newNode As New Dictionary(Of String, Object)()
                    newNode(content) = New List(Of Object)()

                    children.Add(newNode)
                    stack.Push(newNode)
                    levelStack.Push(jlevel)
                End If
            End If
        Next

        Return JsonConvert.SerializeObject(root, Formatting.Indented)
    End Function

    Sub Write2docx()
        Dim inlineTemp As String = "<w:p w:rsidR=""001843B4"" w:rsidRDefault=""001843B4"" w:rsidP=""00A66AB8""><w:pPr><w:pStyle w:val=""ListParagraph""/><w:numPr><w:ilvl w:val=""{سطح}""/><w:numId w:val=""28""/></w:numPr><w:rPr><w:rFonts w:ascii=""Lotus YA"" w:hAnsi=""Lotus YA""/></w:rPr></w:pPr><w:r w:rsidRPr=""00A66AB8""><w:rPr><w:rFonts w:ascii=""Lotus YA"" w:hAnsi=""Lotus YA""/><w:rtl/></w:rPr><w:t>{متن}</w:t></w:r></w:p>"
        '{سطح} - {متن}
        'سطح از صفر آغاز مي‌شود
        'تاريخدراينجا - عنواندراينجا - متندراينجا
        'word/document.xml
        Try
            Dim fileBytes As Byte() = My.Resources.template
            File.WriteAllBytes(XmindFile & ".docx", fileBytes)
            Dim docxml As String = LoadTextFileFromZip(XmindFile & ".docx", "word/document.xml")
            Dim inMatn As String = ""
            inMatn = Get4docx()
            docxml = Replace(docxml, "متندراينجا", inMatn)
            docxml = Replace(docxml, "عنواندراينجا", OnvanAsli)
            docxml = Replace(docxml, "تاريخدراينجا", GetShamsiDate())
            Using archive As New ZipArchive(File.Open(XmindFile & ".docx", FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update)
                Dim oldEntry As ZipArchiveEntry = archive.GetEntry("word/document.xml")
                If oldEntry IsNot Nothing Then oldEntry.Delete()
                Dim newEntry As ZipArchiveEntry = archive.CreateEntry("word/document.xml")
                Using writer As New StreamWriter(newEntry.Open(), System.Text.Encoding.UTF8)
                    writer.Write(docxml)
                End Using
            End Using
            Process.Start(XmindFile & ".docx")
            fileBytes = My.Resources.LotusYA
            File.WriteAllBytes(LastPath & "\LotusYA.ttf", fileBytes)
            fileBytes = My.Resources.LotusYAb
            File.WriteAllBytes(LastPath & "\LotusYAb.ttf", fileBytes)
            fileBytes = My.Resources.VahidYA
            File.WriteAllBytes(LastPath & "\VahidYA.ttf", fileBytes)
            fileBytes = My.Resources.VazirYA
            File.WriteAllBytes(LastPath & "\VazirYA.ttf", fileBytes)
        Catch ex As Exception
            MessageBox.Show("خطا در ذخيره‌سازي: " & ex.Message)
        End Try
    End Sub

    Public Function GetShamsiDate() As String
        Dim pc As New PersianCalendar()
        Dim d As DateTime = DateTime.Now
        Dim months As String() = {"فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"}
        Return $"{pc.GetDayOfMonth(d)} {months(pc.GetMonth(d) - 1)} {pc.GetYear(d)}"
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            System.Diagnostics.Process.Start("http://www.movashah.ir")
        Catch
            ' روش جایگزین برای ویندوزهای جدیدتر
            Dim psi As New System.Diagnostics.ProcessStartInfo
            psi.UseShellExecute = True
            psi.FileName = "https://www.movashah.ir"
            System.Diagnostics.Process.Start(psi)
        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Dim Vaziat As Boolean = False
        If RadioButton2.Checked Then Vaziat = True
        CheckBox1.Enabled = Vaziat
        CheckBox2.Enabled = Vaziat
        CheckBox3.Enabled = Vaziat
        CheckBox4.Enabled = Vaziat
        TextBox2.Enabled = Vaziat
        TextBox3.Enabled = Vaziat
        TextBox4.Enabled = Vaziat
        TextBox5.Enabled = Vaziat
        TextBox6.Enabled = Vaziat
        Label1.Enabled = Vaziat
        Label2.Enabled = Vaziat
        Label3.Enabled = Vaziat
        Label4.Enabled = Vaziat
        Label5.Enabled = Vaziat
    End Sub

    Sub GetXmind(xFile As String)
        xContent = LoadTextFileFromZip(xFile, "content.json")
        xFormat = "json"
        If Not Len(xContent) > 0 Then
            xContent = LoadTextFileFromZip(xFile, "content.xml")
            xFormat = "xml"
        End If
        If Len(xContent) > 0 Then
            GetSample(xContent, xFormat)
        Else
            TextBox1.Text = "محتوايي متناسب با فايل Xmind يافت نشد!"
        End If
    End Sub

    Function Get4docx() As String
        OnvanAsli = ""
        If Len(xContent) > 0 Then
            If xFormat = "json" Then
                Try
                    Dim jsonArray As JArray = JArray.Parse(xContent)
                    For Each item As JObject In jsonArray
                        If item("rootTopic") IsNot Nothing Then
                            ExtractTitledocx(item("rootTopic"))
                        End If
                    Next
                    Return FullText
                Catch ex As Exception
                    MessageBox.Show("خطا در پردازش JSON: " & ex.Message)
                End Try
            ElseIf xFormat = "xml" Then
                ExtractXMLdocx(xContent)
                Return FullText
            End If
        End If
    End Function

    Private Sub ExtractTitledocx(token As JToken)
        Dim inlineTemp As String = "<w:p w:rsidR=""001843B4"" w:rsidRDefault=""001843B4"" w:rsidP=""00A66AB8""><w:pPr><w:pStyle w:val=""ListParagraph""/><w:numPr><w:ilvl w:val=""{سطح}""/><w:numId w:val=""28""/></w:numPr><w:rPr><w:rFonts w:ascii=""Lotus YA"" w:hAnsi=""Lotus YA""/></w:rPr></w:pPr><w:r w:rsidRPr=""00A66AB8""><w:rPr><w:rFonts w:ascii=""Lotus YA"" w:hAnsi=""Lotus YA""/><w:rtl/></w:rPr><w:t>{متن}</w:t></w:r></w:p>"
        If token Is Nothing Then Exit Sub
        If token.Type = JTokenType.Object Then
            Dim obj As JObject = CType(token, JObject)
            If obj("title") IsNot Nothing Then
                Dim titleValue As String = obj("title").ToString()
                If Not String.IsNullOrWhiteSpace(titleValue) Then
                    If Level = 0 Then
                        OnvanAsli += titleValue & " "
                    Else
                        FullText += Replace(Replace(inlineTemp, "{سطح}", Level - 1), "{متن}", titleValue)
                    End If
                End If
            End If
            If obj("children") IsNot Nothing Then
                Dim children As JObject = obj("children")
                If children("attached") IsNot Nothing Then
                    Level += 1
                    For Each child As JToken In children("attached")
                        ExtractTitledocx(child)
                    Next
                    Level -= 1
                End If
            End If
        End If
    End Sub

    Sub Write2txt()
        If Len(xContent) > 0 Then
            If Len(TextBox2.Text) = 0 Then TextBox2.Text = "."
            If Len(TextBox3.Text) = 0 Then TextBox3.Text = "/"
            If Len(TextBox5.Text) = 0 Or Not IsNumeric(TextBox5.Text) Then TextBox5.Text = "1"
            If xFormat = "json" Then
                Try
                    Dim jsonArray As JArray = JArray.Parse(xContent)
                    For Each item As JObject In jsonArray
                        If item("rootTopic") IsNot Nothing Then
                            LevelNum(Level) = 1
                            ExtractTitle(item("rootTopic"))
                        End If
                    Next
                Catch ex As Exception
                    MessageBox.Show("خطا در پردازش JSON: " & ex.Message)
                End Try
            ElseIf xFormat = "xml" Then
                FullText = ExtractXML(xContent)
            End If
        End If
    End Sub

    Private Sub ExtractTitle(token As JToken)
        If token Is Nothing Then Exit Sub
        If token.Type = JTokenType.Object Then
            Dim obj As JObject = CType(token, JObject)
            If obj("title") IsNot Nothing Then
                Dim titleValue As String = obj("title").ToString()
                If Not String.IsNullOrWhiteSpace(titleValue) Then
                    Dim JayTab As String = vbTab
                    If Len(TextBox4.Text) > 0 Then JayTab = TextBox4.Text
                    Dim Tabs As New String(JayTab, Level * TextBox5.Text)
                    Dim Shomareh As String = ""
                    If CheckBox2.Checked Then
                        For i As Integer = Level To 1 Step -1
                            Shomareh += LevelNum(i) & TextBox3.Text
                        Next
                    Else
                        For i As Integer = 1 To Level
                            Shomareh += LevelNum(i) & TextBox3.Text
                        Next
                    End If
                    If Strings.Right(Shomareh, 1) = TextBox3.Text Then Shomareh = Strings.Left(Shomareh, Len(Shomareh) - 1)
                    Dim CurLvl As String = ""
                    If Not CheckBox3.Checked Then Tabs = ""
                    If Not CheckBox1.Checked Then Shomareh = ""
                    If Len(Shomareh) > 0 Then CurLvl = Shomareh & TextBox2.Text & " "
                    If Len(Tabs) > 0 Then CurLvl = Tabs & CurLvl
                    FullText += CurLvl & titleValue & vbCrLf
                End If
            End If
            If obj("children") IsNot Nothing Then
                Dim children As JObject = obj("children")
                If children("attached") IsNot Nothing Then
                    Level += 1
                    LevelNum(Level) = 0
                    For Each child As JToken In children("attached")
                        LevelNum(Level) += 1
                        ExtractTitle(child)
                    Next
                    LevelNum(Level) = 0
                    Level -= 1
                End If
            End If
        End If
    End Sub

    Public Function LoadTextFileFromZip(zipPath As String, fileNameInZip As String) As String
        Try
            Using archive As ZipArchive = ZipFile.OpenRead(zipPath)
                Dim entry As ZipArchiveEntry = archive.GetEntry(fileNameInZip)

                If entry IsNot Nothing Then
                    Using reader As New StreamReader(entry.Open())
                        Return reader.ReadToEnd()  ' خواندن کل متن
                    End Using
                End If
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
        Return Nothing
    End Function

    Sub GetSample(xContent As String, xFormat As String)
        CurLine = 0
        If Len(xContent) > 0 Then
            TextBox1.Text = ""
            TextBox1.WordWrap = False
            If xFormat = "json" Then
                Dim titles As New List(Of String)
                Try
                    Dim jsonArray As JArray = JArray.Parse(xContent)
                    'ProgressBar1.Maximum = jsonArray.LongCount
                    For Each item As JObject In jsonArray
                        ' استخراج title از سطح اول
                        'ExtractTitleFromToken(item, titles)

                        ' استخراج title از rootTopic
                        If item("rootTopic") IsNot Nothing Then
                            LevelNum(Level) = 1
                            StopSample = False
                            ExtractTitleFromToken(item("rootTopic"), titles)
                        End If
                        'ProgressBar1.PerformStep()
                    Next
                Catch ex As Exception
                    MessageBox.Show("خطا در پردازش JSON: " & ex.Message)
                End Try

                'TextBox1.Text = titles
            ElseIf xFormat = "xml" Then
                XMLsample = True
                StopSample = False
                TextBox1.Text = ExtractXML(xContent)
                XMLsample = False
            End If
        End If
        'TextBox1.Text = xContent
        If LineTedad <= CurLine Then TextBox1.Text += "[ادامه دارد...]"
        If CurLine = 0 Then
            TextBox1.Text += "محتوايي يافت نشد!"
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub ExtractTitleFromToken(token As JToken, ByRef titles As List(Of String))
        If token Is Nothing Then Exit Sub
        If StopSample Then Exit Sub
        ' بررسی title در شیء جاری
        If token.Type = JTokenType.Object Then
            Dim obj As JObject = CType(token, JObject)

            ' اگر title وجود داشت، اضافه کن
            If obj("title") IsNot Nothing Then
                Dim titleValue As String = obj("title").ToString()
                If Not String.IsNullOrWhiteSpace(titleValue) Then
                    titles.Add(titleValue)
                    Dim JayTab As String = vbTab
                    If Len(TextBox4.Text) > 0 Then JayTab = TextBox4.Text
                    Dim Tabs As New String(JayTab, Level * TextBox5.Text)
                    Dim Shomareh As String = ""
                    If CheckBox2.Checked Then
                        For i As Integer = Level To 1 Step -1
                            Shomareh += LevelNum(i) & TextBox3.Text
                        Next
                    Else
                        For i As Integer = 1 To Level
                            Shomareh += LevelNum(i) & TextBox3.Text
                        Next
                    End If
                    If Strings.Right(Shomareh, 1) = TextBox3.Text Then Shomareh = Strings.Left(Shomareh, Len(Shomareh) - 1)
                    'Dim CurLvl As String = Tabs & Level & ". "
                    Dim CurLvl As String = ""
                    If Not CheckBox3.Checked Then Tabs = ""
                    If Not CheckBox1.Checked Then Shomareh = ""
                    If Len(Shomareh) > 0 Then CurLvl = Shomareh & TextBox2.Text & " "
                    If Len(Tabs) > 0 Then CurLvl = Tabs & CurLvl
                    TextBox1.Text += CurLvl & titleValue & vbCrLf
                    TextBox1.Refresh()
                    CurLine += 1
                    If CurLine >= LineTedad Then StopSample = True
                End If
            End If
            If obj("children") IsNot Nothing Then
                Dim children As JObject = obj("children")
                If children("attached") IsNot Nothing Then
                    Level += 1
                    LevelNum(Level) = 0
                    For Each child As JToken In children("attached")
                        LevelNum(Level) += 1
                        ExtractTitleFromToken(child, titles)
                    Next
                    LevelNum(Level) = 0
                    Level -= 1
                End If
            End If
        End If

        ' اگر آرایه است
        If token.Type = JTokenType.Array Then
            For Each child As JToken In token.Children()
                LevelNum(Level) += 1
                ExtractTitleFromToken(child, titles)
            Next
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        CheckFramework()
        TYelow = TextBox4.BackColor
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        With TextBox4
            If Len(.Text) > 0 Then
                .BackColor = Color.White
            Else
                .BackColor = TYelow
            End If
        End With
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        With TextBox6
            If Len(.Text) > 0 Then
                .BackColor = Color.White
            Else
                .BackColor = TYelow
            End If
        End With
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        With TextBox2
            If Len(.Text) > 0 Then
                .BackColor = Color.White
            Else
                .BackColor = TYelow
            End If
        End With
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        With TextBox3
            If Len(.Text) > 0 Then
                .BackColor = Color.White
            Else
                .BackColor = TYelow
            End If
        End With
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        With TextBox5
            If Len(.Text) > 0 Then
                .BackColor = Color.White
            Else
                .BackColor = TYelow
            End If
        End With
    End Sub

    Public Shared Sub ConvertCSVToExcel(ByVal csvFilePath As String, ByVal excelFilePath As String)
        ExcelPackage.License.SetNonCommercialPersonal("movashah.ir")

        Using package As New ExcelPackage()
            ' ایجاد یک کاربرگ
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Data")

            ' خواندن فایل CSV
            Dim lines As String() = File.ReadAllLines(csvFilePath, Encoding.UTF8)

            ' پر کردن داده‌ها
            For i As Integer = 0 To lines.Length - 1
                If Not String.IsNullOrWhiteSpace(lines(i)) Then
                    Dim cells As String() = lines(i).Split(","c)

                    For j As Integer = 0 To cells.Length - 1
                        worksheet.Cells(i + 1, j + 1).Value = cells(j).Trim()
                    Next
                End If
            Next

            ' تنظیم عرض ستون‌ها به صورت خودکار
            worksheet.Cells(worksheet.Dimension.Address).AutoFitColumns()

            ' ذخیره فایل
            Dim fileInfo As New FileInfo(excelFilePath)
            package.SaveAs(fileInfo)

            Console.WriteLine($"فایل با موفقیت در مسیر {excelFilePath} ذخیره شد.")
        End Using
    End Sub

    Public Shared Sub ConvertCSVToExcelSimple(ByVal csvPath As String)
        Dim excelPath As String = Path.ChangeExtension(csvPath, ".xlsx")
        ConvertCSVToExcel(csvPath, excelPath)
    End Sub

End Class
