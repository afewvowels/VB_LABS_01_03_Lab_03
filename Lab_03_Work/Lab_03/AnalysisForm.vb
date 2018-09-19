'Project: Lab #3
'Author: Anthony DePinto 
'Date: Fall 2014
'Description: Analysis of grades
'
Public Class AnalysisForm
    Private Sub SubmitResultButton_Click(sender As Object, e As EventArgs) Handles SubmitResultButton.Click
        ' Define variables
        Dim TempString As String

        ' Assign TextBox contents to temp string
        TempString = ResultTextBox.Text.ToUpper

        ' Conditional logic to determine if text input is valid
        Select Case TempString
            ' Define acceptable letter-grade entries
            Case "P", "F"
                ' Acceptable letter entered, add result to ListBox
                ResultsListBox.Items.Insert(ResultsListBox.Items.Count, TempString)
                ' Increment grade counter
                'GradesTotalCounterInteger += 1
            Case Else
                ' Invalid entry is detected, display error message
                MessageBox.Show("Invalid grade entry detected",
                                "Bad Data",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)
        End Select

        ' Regardless of correct or incorrect data entry, clear TextBox and reset Focus
        ResultTextBox.Clear()
        ResultTextBox.Focus()

        ' Check if ListBox has reached 10 items, which disables this button
        ' and enables the Analyze Results button
        If ResultsListBox.Items.Count >= 10 Then
            ResultTextBox.Enabled = False
            SubmitResultButton.Enabled = False
            AnalyzeResultsButton.Enabled = True
        End If
    End Sub

    Private Sub AnalyzeResultsButton_Click(sender As Object, e As EventArgs) Handles AnalyzeResultsButton.Click
        ' Declare variables
        Dim PInteger As Integer = 0
        Dim FInteger As Integer = 0

        ' Run For Next loop through contents of ResultsListBox and count total number
        ' of 'P's and 'F's
        For CounterInteger As Integer = 0 To (ResultsListBox.Items.Count - 1) Step 1
            Select Case ResultsListBox.Items.Item(CounterInteger).ToString
                Case "P"
                    PInteger += 1
                Case "F"
                    FInteger += 1
                Case Else
                    MessageBox.Show("Invalid grade entry detected",
                                "Bad Data",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)
            End Select
        Next CounterInteger
        ' End For Next loop

        ' Output results of two counter variables
        AnalysisResultsLabel.Text = "There were " & PInteger.ToString & " passing grades and " & FInteger.ToString & " failing grades."

        ' Once analysis is complete disable
        AnalyzeResultsButton.Enabled = False

    End Sub

    Private Sub ClearResultsButton_Click(sender As Object, e As EventArgs) Handles ClearResultsButton.Click
        ' Clear ListBox once comparison is done
        ResultsListBox.Items.Clear()

        ' Reverse button enabling/disabling to reset application
        ResultTextBox.Enabled = True
        SubmitResultButton.Enabled = True
        AnalysisResultsLabel.Text = String.Empty
        ResultTextBox.Focus()
    End Sub
End Class ' Analysis

