Imports System.Data.OleDb
Public Class Form1
    'Dim AddMode As Boolean
    'Dim Ctrl As Integer
    Dim myConn As New OleDbConnection
    Dim myCmd As New OleDbCommand
    Dim myDA As New OleDbDataAdapter
    Dim myDR As OleDbDataReader
    Dim strSQL As String
    Private fname As String
    Private path As String = ""
    Private fnameWithoutExtension As String = ""
    Private Dist As String = ""
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dist = "c:\"
        Me.convButton.Enabled = False
        Me.schema_CheckBox.Checked = False

    End Sub
    Private Function repHtmlSpecChars(ByVal str As String) As String
        Dim spec As New List(Of Char)
        Dim RepSpec As New List(Of String)

        spec.AddRange(New Char() {"&", """", "'", ">", "<"})
        RepSpec.AddRange(New String() {"amp", "quot", "apos", "gt", "lt"})
        For i As Integer = 0 To spec.Count - 1
            str = str.Replace(spec(i), "&" + RepSpec(i) + ";")
        Next
        
        Return str
    End Function
    Private Sub convButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles convButton.Click
        If schema_CheckBox.Checked Then
            GenerateXmlSchema()
        End If
        GenerateXml()
        Me.convButton.Enabled = False
    End Sub
    Sub GenerateXmlSchema()
        Try
            Dim gc As Integer = 0
            Dim pc As Integer = 1
            Dim wr As IO.StreamWriter = Nothing

            IsConnected(Me.path)
            wr = New IO.StreamWriter(Dist + "\" + fname + "Schema.xml")
            wr.WriteLine("<?xml version=""1.0""?>")
            wr.WriteLine("<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">")
            wr.WriteLine(Chr(9) + "<xs:element name=""" + fnameWithoutExtension + """>")
            wr.WriteLine(dTab() + "<xs:complexType>")
            wr.WriteLine(ThreeTab() + "<xs:sequence>")
            For Each tbl As String In getallTables()
                ' for TABLE
                wr.WriteLine(fourTab() + "<xs:element name=""" + tbl + """>")
                wr.WriteLine(fiveTab() + "<xs:complexType>")
                wr.WriteLine(sixTab() + "<xs:sequence>")

                ' for RECORD ELEMENT
                wr.WriteLine(sevTab() + "<xs:element name=""RECORD"">")
                wr.WriteLine(Tab8() + "<xs:complexType>")
                wr.WriteLine(Tab9() + "<xs:sequence>")

                strSQL = "SELECT * FROM [" + tbl + "]"
                myCmd.CommandText = strSQL
                myCmd.Connection = myConn
                myDA.SelectCommand = myCmd
                myDR = myCmd.ExecuteReader()
                For i As Integer = 0 To myDR.FieldCount - 1
                    wr.WriteLine(Tab10() + "<xs:element name=""" + myDR.GetName(i) + """ type=""xs:" + HandleDataType(myDR.GetFieldType(i).Name.ToLower) + """/>")
                Next

                myDR.Close()

                ' for RECORD ELEMENT
                wr.WriteLine(Tab9() + "</xs:sequence>")
                wr.WriteLine(Tab8() + "</xs:complexType>")
                wr.WriteLine(sevTab() + "</xs:element>")

                ' for TABLE
                wr.WriteLine(sixTab() + "</xs:sequence>")
                wr.WriteLine(fiveTab() + "</xs:complexType>")
                wr.WriteLine(fourTab() + "</xs:element>")
            Next
            wr.WriteLine(ThreeTab() + "</xs:sequence>")
            wr.WriteLine(dTab() + "</xs:complexType>")
            wr.WriteLine(Chr(9) + "</xs:element>")

            wr.WriteLine("</xs:schema>")
            wr.Close()
        Catch ex As Exception
            MsgBox(" Error Ocurred")
        End Try

    End Sub
    Public Function getallTables() As String()
        Dim SchemaTable As DataTable
        Try
            'Get table and view names
            Dim tbles As New List(Of String)
            SchemaTable = myConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, Nothing})
            Dim int As Integer
            For int = 0 To SchemaTable.Rows.Count - 1
                If SchemaTable.Rows(int)!TABLE_TYPE.ToString = "TABLE" Then
                    tbles.Add(SchemaTable.Rows(int)!TABLE_NAME.ToString())
                End If

            Next
            Return tbles.ToArray
        Catch ex As Exception
            Return Nothing
        End Try


    End Function
    Sub GenerateXml()
        Try
            Dim gc As Integer = 0
            Dim pc As Integer = 1
            Dim gmax As Integer = 0
            Dim wr As IO.StreamWriter = Nothing

            IsConnected(Me.path)
            Dim xml As String = ""

            For Each tbl As String In getallTables()

                strSQL = "SELECT * FROM [" + tbl + "]"
                myCmd.CommandText = strSQL
                myCmd.Connection = myConn
                myDA.SelectCommand = myCmd
                myDR = myCmd.ExecuteReader()

                While myDR.Read()
                    gc = gc + 1

                    If gmax = 0 Then ' no split
                        If gc = 1 Then
                            wr = New IO.StreamWriter(Dist + "\" + fname + ".xml")
                            wr.WriteLine("<?xml version=""1.0""?>" + Chr(10) + "<" + fnameWithoutExtension + ">")
                        End If
                    Else
                        If gc = 1 Then
                            wr = New IO.StreamWriter(Dist + "\" + fname + "" + pc.ToString + ".xml")
                            wr.WriteLine("<?xml version=""1.0""?>" + Chr(10) + "<" + fnameWithoutExtension + ">")
                        ElseIf gc = gmax Then
                            If (Not wr Is Nothing) Then
                                wr.WriteLine("</" + fnameWithoutExtension + ">")
                                wr.Close()
                                wr = New IO.StreamWriter(Dist + "\" + fname + "" + pc.ToString + ".xml")
                                wr.WriteLine("<?xml version=""1.0""?>" + Chr(10) + "<" + fnameWithoutExtension + ">")
                                gc = 1
                                pc = pc + 1
                            End If

                        End If
                    End If

                    wr.WriteLine(Chr(9) + "<" + tbl + ">" + Chr(10) + dTab() + "<RECORD>")
                    For i As Integer = 0 To myDR.FieldCount - 1
                        wr.WriteLine(ThreeTab() + "<" + myDR.GetName(i) + ">" + repHtmlSpecChars(myDR.GetValue(i).ToString) + "</" + myDR.GetName(i) + ">")
                    Next
                    wr.WriteLine(dTab() + "</RECORD>" + Chr(10) + Chr(9) + "</" + tbl + ">")
                End While
                myDR.Close()

            Next

            If gmax = 0 Then
                wr.WriteLine("</" + fnameWithoutExtension + ">")
            End If
            wr.Close()
        Catch ex As Exception
            MsgBox(" Error Ocurred")
        End Try

    End Sub

    Function getFileExtension(ByVal file As String) As String
        Return file.Substring(file.LastIndexOf(".") + 1)
    End Function
    Function IsConnected(Optional ByVal path As String = Nothing) As Boolean
        Try
            'Checks first if already connected to database,if connected, it will be disconnected.
            If myConn.State = ConnectionState.Open Then myConn.Close()
            Dim provider As String
            If getFileExtension(path).ToLower = "accdb" Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0" ' access 2007
            Else
                provider = "Provider=Microsoft.Jet.OLEDB.4.0"
            End If
            myConn.ConnectionString = provider + ";Data Source=" + path + ";"
            myConn.Open()
            IsConnected = True
        Catch ex As Exception
            MsgBox(" Error Ocurred")
        End Try
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles selectButton.Click
        Me.OpenFileDialog1.ShowDialog()
        path = Me.OpenFileDialog1.FileName
        Dist = getFileFolder(path)
        fname = Me.OpenFileDialog1.SafeFileName
        fnameWithoutExtension = "book" 'remExt(fname)
        If (Not String.IsNullOrEmpty(path)) Then
            Me.convButton.Enabled = True
        End If
    End Sub

    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.FolderBrowserDialog1.ShowDialog()
        Dist = Me.FolderBrowserDialog1.SelectedPath
    End Sub

    Private Function remExt(ByVal file As String) As String
        Try
            Dim dotPosition As Integer = file.LastIndexOf(".")
            Dim filenameOnly As String = file.Substring(0, dotPosition)
            Return filenameOnly
        Catch ex As Exception
            Return file
        End Try
    End Function

    Private Function getFileFolder(ByVal file As String) As String
        Try
            Dim dotPosition As Integer = file.LastIndexOf("\")
            Dim filenameOnly As String = file.Substring(0, dotPosition)
            Return filenameOnly
        Catch ex As Exception
            Return file
        End Try
    End Function

    Private Function HandleDataType(ByVal typeName As String) As String
        If typeName = "int32" Then
            Return "integer"
        Else
            Return typeName
        End If
    End Function
    Private Function dTab() As String
        Return Chr(9) + Chr(9)
    End Function
    Private Function ThreeTab() As String
        Return Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function fourTab() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function fiveTab() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function sixTab() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function sevTab() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function Tab8() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function Tab9() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function Tab10() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function Tab11() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function
    Private Function Tab12() As String
        Return Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9)
    End Function

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("http://www.linkedin.com/in/yaburrub")
        Catch ex As Exception

        End Try

    End Sub
End Class
