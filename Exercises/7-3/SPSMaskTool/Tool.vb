Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.SharePoint.Utilities
Imports Microsoft.SharePoint.WebPartPages
Imports System.Text.RegularExpressions

Public Class Tool
    Inherits ToolPart

    'Controls to appear in the tool part
    Protected WithEvents txtProperty As TextBox
    Protected WithEvents lblMessage As Label

    Protected Overrides Sub CreateChildControls()

        'Purpose: Add the child controls to the web part

        'Label for the errors
        lblMessage = New Label
        With lblMessage
            .Width = Unit.Percentage(100)
        End With
        Controls.Add(lblMessage)

        'Text Box for input
        txtProperty = New TextBox
        With txtProperty
            .Width = Unit.Percentage(100)
        End With
        Controls.Add(txtProperty)

    End Sub

    Public Overrides Sub ApplyChanges()
        'User pushes "OK" or "Apply"

        Try
            'Test the input value against the regular expression
            Dim objRegEx As New Regex("\b\d{3}-\d{3}-\d{4}\b")
            Dim objMatch As Match = objRegEx.Match(txtProperty.Text)
            If objMatch.Success Then
                'Move value from tool pane to web part
                Dim objWebPart As Part = _
    DirectCast(Me.ParentToolPane.SelectedWebPart, Part)
                objWebPart.Text = txtProperty.Text
                lblMessage.Text = ""
            Else
                lblMessage.Text = "Invalid phone number."
                txtProperty.Text = "###-###-####"
            End If
        Catch x As ArgumentException
        End Try
    End Sub

    Public Overrides Sub SyncChanges()
        'This is called after ApplyChanges to sync tool pane with web part
        Try
            'Test the input value against the regular expression
            Dim objRegEx As New Regex("\b\d{3}-\d{3}-\d{4}\b")
            Dim objMatch As Match = objRegEx.Match(txtProperty.Text)
            If objMatch.Success Then
                'Move value from tool pane to web part
                Dim objWebPart As Part = _
    DirectCast(Me.ParentToolPane.SelectedWebPart, Part)
                txtProperty.Text = objWebPart.Text
                lblMessage.Text = ""
            Else
                lblMessage.Text = "Invalid phone number."
                txtProperty.Text = "###-###-####"
            End If
        Catch x As ArgumentException
        End Try
    End Sub

    Public Overrides Sub CancelChanges()
        'User pushes "Cancel"
        Dim objWebPart As Part = DirectCast(Me.ParentToolPane.SelectedWebPart, Part)
        objWebPart.Text = ""
        txtProperty.Text = "###-###-####"
        lblMessage.Text = ""
    End Sub

    Protected Overrides Sub RenderToolPart(ByVal output As HtmlTextWriter)
        'Populate the existing property
        Dim objWebPart As Part = DirectCast(Me.ParentToolPane.SelectedWebPart, Part)
        txtProperty.Text = objWebPart.Text

        'Draw the tool part
        lblMessage.RenderControl(output)
        output.Write("<br>")
        txtProperty.RenderControl(output)
    End Sub



End Class
