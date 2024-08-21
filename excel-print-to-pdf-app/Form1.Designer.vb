<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Label1 = New Label()
        rbtn_printAll = New RadioButton()
        rbtn_printFirstSheet = New RadioButton()
        btn_Print = New Button()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(12, 12)
        Label1.Margin = New Padding(3)
        Label1.Name = "Label1"
        Label1.Size = New Size(119, 17)
        Label1.TabIndex = 0
        Label1.Text = "Print Excel to PDF"
        ' 
        ' rbtn_printAll
        ' 
        rbtn_printAll.AutoSize = True
        rbtn_printAll.Location = New Point(12, 35)
        rbtn_printAll.Name = "rbtn_printAll"
        rbtn_printAll.Size = New Size(101, 19)
        rbtn_printAll.TabIndex = 1
        rbtn_printAll.TabStop = True
        rbtn_printAll.Text = "Print all sheets"
        rbtn_printAll.UseVisualStyleBackColor = True
        ' 
        ' rbtn_printFirstSheet
        ' 
        rbtn_printFirstSheet.AutoSize = True
        rbtn_printFirstSheet.Location = New Point(128, 35)
        rbtn_printFirstSheet.Name = "rbtn_printFirstSheet"
        rbtn_printFirstSheet.Size = New Size(104, 19)
        rbtn_printFirstSheet.TabIndex = 2
        rbtn_printFirstSheet.TabStop = True
        rbtn_printFirstSheet.Text = "Print first sheet"
        rbtn_printFirstSheet.UseVisualStyleBackColor = True
        ' 
        ' btn_Print
        ' 
        btn_Print.Location = New Point(12, 64)
        btn_Print.Name = "btn_Print"
        btn_Print.Size = New Size(220, 31)
        btn_Print.TabIndex = 3
        btn_Print.Text = "Print"
        btn_Print.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(244, 107)
        Controls.Add(btn_Print)
        Controls.Add(rbtn_printFirstSheet)
        Controls.Add(rbtn_printAll)
        Controls.Add(Label1)
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        Name = "Form1"
        Text = "Excel to PDF"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents rbtn_printAll As RadioButton
    Friend WithEvents rbtn_printFirstSheet As RadioButton
    Friend WithEvents btn_Print As Button

End Class
