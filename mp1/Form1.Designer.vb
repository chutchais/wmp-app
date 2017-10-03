
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.UcTest1 = New mp.ucTest()
        Me.tbScript = New System.Windows.Forms.TextBox()
        Me.txtReturn = New System.Windows.Forms.TextBox()
        Me.UcMain1 = New mp.ucMain()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.txtSnippetSlug = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(697, 234)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(144, 33)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 347)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(144, 33)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 386)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(144, 33)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(12, 425)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(144, 33)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'UcTest1
        '
        Me.UcTest1.CurrentForm = Nothing
        Me.UcTest1.Location = New System.Drawing.Point(12, 255)
        Me.UcTest1.Name = "UcTest1"
        Me.UcTest1.ParentObjectName = Nothing
        Me.UcTest1.Size = New System.Drawing.Size(280, 91)
        Me.UcTest1.TabIndex = 4
        '
        'tbScript
        '
        Me.tbScript.Location = New System.Drawing.Point(397, 3)
        Me.tbScript.Multiline = True
        Me.tbScript.Name = "tbScript"
        Me.tbScript.Size = New System.Drawing.Size(444, 225)
        Me.tbScript.TabIndex = 5
        Me.tbScript.Text = resources.GetString("tbScript.Text")
        '
        'txtReturn
        '
        Me.txtReturn.Location = New System.Drawing.Point(397, 234)
        Me.txtReturn.Name = "txtReturn"
        Me.txtReturn.Size = New System.Drawing.Size(179, 20)
        Me.txtReturn.TabIndex = 6
        Me.txtReturn.Text = "ssdssx"
        '
        'UcMain1
        '
        Me.UcMain1.CurrentForm = Nothing
        Me.UcMain1.Location = New System.Drawing.Point(12, 3)
        Me.UcMain1.Name = "UcMain1"
        Me.UcMain1.ParentObjectName = Nothing
        Me.UcMain1.Size = New System.Drawing.Size(379, 203)
        Me.UcMain1.TabIndex = 7
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(803, 425)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(74, 30)
        Me.Button5.TabIndex = 8
        Me.Button5.Text = "Get Code"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'txtSnippetSlug
        '
        Me.txtSnippetSlug.Location = New System.Drawing.Point(450, 431)
        Me.txtSnippetSlug.Name = "txtSnippetSlug"
        Me.txtSnippetSlug.Size = New System.Drawing.Size(347, 20)
        Me.txtSnippetSlug.TabIndex = 9
        Me.txtSnippetSlug.Text = "test-snippet"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(337, 434)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Snippet Slug Name"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(889, 467)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSnippetSlug)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.UcMain1)
        Me.Controls.Add(Me.txtReturn)
        Me.Controls.Add(Me.tbScript)
        Me.Controls.Add(Me.UcTest1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "I am Back"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents UcTest1 As ucTest
    Friend WithEvents tbScript As TextBox
    Friend WithEvents txtReturn As TextBox
    Friend WithEvents UcMain1 As ucMain
    Friend WithEvents Button5 As Button
    Friend WithEvents txtSnippetSlug As TextBox
    Friend WithEvents Label1 As Label
End Class
