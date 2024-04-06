<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExitForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.ConfirmButton = New System.Windows.Forms.Button()
        Me.QuestionLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'CancelButton
        '
        Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton.Location = New System.Drawing.Point(66, 102)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(84, 23)
        Me.CancelButton.TabIndex = 0
        Me.CancelButton.Text = "&No"
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'ConfirmButton
        '
        Me.ConfirmButton.Location = New System.Drawing.Point(156, 102)
        Me.ConfirmButton.Name = "ConfirmButton"
        Me.ConfirmButton.Size = New System.Drawing.Size(84, 23)
        Me.ConfirmButton.TabIndex = 1
        Me.ConfirmButton.Text = "&Yes"
        Me.ConfirmButton.UseVisualStyleBackColor = True
        '
        'QuestionLabel
        '
        Me.QuestionLabel.AutoSize = True
        Me.QuestionLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QuestionLabel.Location = New System.Drawing.Point(9, 26)
        Me.QuestionLabel.Name = "QuestionLabel"
        Me.QuestionLabel.Size = New System.Drawing.Size(231, 13)
        Me.QuestionLabel.TabIndex = 2
        Me.QuestionLabel.Text = "Are you sure that you want to quit this program?"
        '
        'ExitForm
        '
        Me.AcceptButton = Me.ConfirmButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.CancelButton
        Me.ClientSize = New System.Drawing.Size(250, 134)
        Me.Controls.Add(Me.QuestionLabel)
        Me.Controls.Add(Me.ConfirmButton)
        Me.Controls.Add(Me.CancelButton)
        Me.Name = "ExitForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Exit?"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CancelButton As Button
    Friend WithEvents ConfirmButton As Button
    Friend WithEvents QuestionLabel As Label
End Class
