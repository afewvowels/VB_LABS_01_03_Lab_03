<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AnalysisForm
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
        Me.ClearResultsButton = New System.Windows.Forms.Button()
        Me.AnalysisResultsLabel = New System.Windows.Forms.Label()
        Me.AnalysisLabel = New System.Windows.Forms.Label()
        Me.AnalyzeResultsButton = New System.Windows.Forms.Button()
        Me.SubmitResultButton = New System.Windows.Forms.Button()
        Me.ResultTextBox = New System.Windows.Forms.TextBox()
        Me.EnterResultLabel = New System.Windows.Forms.Label()
        Me.ResultsListBox = New System.Windows.Forms.ListBox()
        Me.ResultsLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ClearResultsButton
        '
        Me.ClearResultsButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ClearResultsButton.Location = New System.Drawing.Point(78, 189)
        Me.ClearResultsButton.Name = "ClearResultsButton"
        Me.ClearResultsButton.Size = New System.Drawing.Size(191, 23)
        Me.ClearResultsButton.TabIndex = 13
        Me.ClearResultsButton.Text = "Clear Results"
        Me.ClearResultsButton.UseVisualStyleBackColor = True
        '
        'AnalysisResultsLabel
        '
        Me.AnalysisResultsLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.AnalysisResultsLabel.Location = New System.Drawing.Point(78, 136)
        Me.AnalysisResultsLabel.Name = "AnalysisResultsLabel"
        Me.AnalysisResultsLabel.Size = New System.Drawing.Size(188, 50)
        Me.AnalysisResultsLabel.TabIndex = 16
        Me.AnalysisResultsLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'AnalysisLabel
        '
        Me.AnalysisLabel.AutoSize = True
        Me.AnalysisLabel.Location = New System.Drawing.Point(75, 116)
        Me.AnalysisLabel.Name = "AnalysisLabel"
        Me.AnalysisLabel.Size = New System.Drawing.Size(53, 15)
        Me.AnalysisLabel.TabIndex = 15
        Me.AnalysisLabel.Text = "Analysis:"
        '
        'AnalyzeResultsButton
        '
        Me.AnalyzeResultsButton.Enabled = False
        Me.AnalyzeResultsButton.Location = New System.Drawing.Point(78, 86)
        Me.AnalyzeResultsButton.Name = "AnalyzeResultsButton"
        Me.AnalyzeResultsButton.Size = New System.Drawing.Size(189, 23)
        Me.AnalyzeResultsButton.TabIndex = 11
        Me.AnalyzeResultsButton.Text = "Analyze Results"
        Me.AnalyzeResultsButton.UseVisualStyleBackColor = True
        '
        'SubmitResultButton
        '
        Me.SubmitResultButton.Location = New System.Drawing.Point(78, 57)
        Me.SubmitResultButton.Name = "SubmitResultButton"
        Me.SubmitResultButton.Size = New System.Drawing.Size(189, 23)
        Me.SubmitResultButton.TabIndex = 10
        Me.SubmitResultButton.Text = "Submit Result"
        Me.SubmitResultButton.UseVisualStyleBackColor = True
        '
        'ResultTextBox
        '
        Me.ResultTextBox.Location = New System.Drawing.Point(232, 28)
        Me.ResultTextBox.Name = "ResultTextBox"
        Me.ResultTextBox.Size = New System.Drawing.Size(34, 23)
        Me.ResultTextBox.TabIndex = 8
        '
        'EnterResultLabel
        '
        Me.EnterResultLabel.AutoSize = True
        Me.EnterResultLabel.Location = New System.Drawing.Point(75, 31)
        Me.EnterResultLabel.Name = "EnterResultLabel"
        Me.EnterResultLabel.Size = New System.Drawing.Size(151, 15)
        Me.EnterResultLabel.TabIndex = 9
        Me.EnterResultLabel.Text = "Enter P for pass or F for fail:"
        '
        'ResultsListBox
        '
        Me.ResultsListBox.FormattingEnabled = True
        Me.ResultsListBox.ItemHeight = 15
        Me.ResultsListBox.Location = New System.Drawing.Point(15, 28)
        Me.ResultsListBox.Name = "ResultsListBox"
        Me.ResultsListBox.Size = New System.Drawing.Size(47, 184)
        Me.ResultsListBox.TabIndex = 12
        '
        'ResultsLabel
        '
        Me.ResultsLabel.AutoSize = True
        Me.ResultsLabel.Location = New System.Drawing.Point(12, 9)
        Me.ResultsLabel.Name = "ResultsLabel"
        Me.ResultsLabel.Size = New System.Drawing.Size(47, 15)
        Me.ResultsLabel.TabIndex = 14
        Me.ResultsLabel.Text = "Results:"
        '
        'AnalysisForm
        '
        Me.AcceptButton = Me.SubmitResultButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ClearResultsButton
        Me.ClientSize = New System.Drawing.Size(281, 226)
        Me.Controls.Add(Me.ClearResultsButton)
        Me.Controls.Add(Me.AnalysisResultsLabel)
        Me.Controls.Add(Me.AnalysisLabel)
        Me.Controls.Add(Me.AnalyzeResultsButton)
        Me.Controls.Add(Me.SubmitResultButton)
        Me.Controls.Add(Me.ResultTextBox)
        Me.Controls.Add(Me.EnterResultLabel)
        Me.Controls.Add(Me.ResultsListBox)
        Me.Controls.Add(Me.ResultsLabel)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "AnalysisForm"
        Me.Text = "Licensing-Exam Analysis"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ClearResultsButton As System.Windows.Forms.Button
    Friend WithEvents AnalysisResultsLabel As System.Windows.Forms.Label
    Friend WithEvents AnalysisLabel As System.Windows.Forms.Label
    Friend WithEvents AnalyzeResultsButton As System.Windows.Forms.Button
    Friend WithEvents SubmitResultButton As System.Windows.Forms.Button
    Friend WithEvents ResultTextBox As System.Windows.Forms.TextBox
    Friend WithEvents EnterResultLabel As System.Windows.Forms.Label
    Friend WithEvents ResultsListBox As System.Windows.Forms.ListBox
    Friend WithEvents ResultsLabel As System.Windows.Forms.Label

End Class
