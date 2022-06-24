Public NotInheritable Class AboutBox
    Private Sub AboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '---Set the title of the form---
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = String.Format("About {0}", ApplicationTitle)
        '---Initialize all of the text displayed on the About Box---
        Me.LabelProductName.Text = My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelDescription.Text = My.Application.Info.Description
        '---add icon to form---
        Dim bm As Bitmap = My.Resources.alert
        Dim icon As IntPtr = bm.GetHicon()
        Dim newIcon As Icon = System.Drawing.Icon.FromHandle(icon)
        Me.Icon = newIcon
        Me.Icon = My.Resources.sudoku
    End Sub
    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, _
   ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) _
   Handles LinkLabel1.LinkClicked
        Try
            VisitLink()
        Catch ex As Exception
            ' The error message
            MessageBox.Show("Unable to open link that was clicked.")
        End Try
    End Sub

    Sub VisitLink()
        LinkLabel1.LinkVisited = True
        System.Diagnostics.Process.Start(LinkLabel1.Text)
    End Sub



End Class
