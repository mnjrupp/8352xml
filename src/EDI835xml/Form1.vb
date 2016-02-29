Imports System.IO

Public Class Form1





    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.ToolStripStatusLabel1.Text = ""
        Call LoadSettingsFromFile()

    End Sub
    Private Sub LoadModelForm()
        Form2.Show()
    End Sub

    Private Sub LoadSettingsFromFile()

        Dim executingself As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim inifilepath = New System.Text.StringBuilder
        With inifilepath
            .Append(IO.Path.GetDirectoryName(executingself))
            .Append(IO.Path.GetFileName(executingself))
            .Append(".ini")
        End With
        If FileIO.FileSystem.FileExists(inifilepath.ToString()) Then
            Try
                Using srini As StreamReader = New StreamReader(inifilepath.ToString())
                    Dim strFile As String
                    ' Read in the complete ini config file
                    strFile = srini.ReadToEnd()
                    ' Need to split out by line the configs
                    Dim varFileArray() As String
                    Dim varSettingsArray() As String
                    Dim count As Integer
                    varFileArray = strFile.Split(vbCrLf)
                    For count = 0 To varFileArray.Length - 1
                        Select Case count
                            Case 0
                                Form2.TextBox1.Text = varFileArray(0)
                                varSettingsArray = varFileArray(0).Split(";")
                                Dim i As Integer
                                For i = 0 To varSettingsArray.Length - 1
                                    Dim comboItem As String = Mid(varSettingsArray(i).ToString(), _
                                                                  InStrRev(varSettingsArray(i).ToString(), "\", vbTextCompare) + 1)

                                    Form2.Combo1.Items.Add(comboItem)
                                Next
                            Case 1
                                varSettingsArray = varFileArray(1).Split(";")
                                With Me
                                    .Location = New Point(CInt(varSettingsArray(0)), CInt(varSettingsArray(1)))
                                End With
                        End Select

                    Next
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    Private Sub SaveSettingsToFile()

        Dim executingself As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim inifilepath = New System.Text.StringBuilder
        With inifilepath
            .Append(IO.Path.GetDirectoryName(executingself))
            .Append(IO.Path.GetFileName(executingself))
            .Append(".ini")
        End With

        Try
            ' Over write existing file in case settings have changed
            Using swrini As StreamWriter = New StreamWriter(inifilepath.ToString(), False)
                With swrini
                    .WriteLine(Form2.TextBox1.Text.ToString())
                    .Write(CStr(Me.Location.X))
                    .Write(";")
                    .Write(CStr(Me.Location.Y))
                    .Write(vbCrLf)
                End With
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Export2Access()
        Dim i As Integer, strName1 As String, strPath1 As String, strSize1 As String, strType1 As String
    End Sub

    Public Sub ConvertERA(m_import As Boolean, strPath As String, Optional xmlAccess As Integer = 2)
        Dim rtnStr As String
        Static i As Long

        If Len(strPath) > 1 Then
            Try
                Using srini As StreamReader = New StreamReader(strPath)
                    rtnStr = srini.ReadToEnd()

                End Using
            Catch ex As Exception

            End Try
        End If
    End Sub
End Class
