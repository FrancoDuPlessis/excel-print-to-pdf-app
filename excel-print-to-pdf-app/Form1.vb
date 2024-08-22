Imports System.IO

Public Class Form1

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim fileFolder = fso.GetAbsolutePathName(".") + "\filesToPrint\"
        If Not Directory.Exists(fileFolder) Then
            Directory.CreateDirectory(fileFolder)
        End If

    End Sub

    Private Sub btn_Print_Click(sender As Object, e As EventArgs) Handles btn_Print.Click

        If rbtn_printAll.Checked() Then
            PrintServices.XLS_print_to_PDF()
            MessageBox.Show("Finished")
        ElseIf rbtn_printFirstSheet.Checked() Then
            PrintServices.XLS_print_first_sheet_to_PDF()
            MessageBox.Show("Finished")
        End If

    End Sub
End Class
