Public Class frmMain
    Private bConnected As Boolean
    

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked
        
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If bConnected Then
            ToolStripComboBox1.Enabled = True
            ToolStripComboBox2.Enabled = True

            bConnected = False
            ToolStripButton1.Text = "Connect"

            ' Add Disconnect Code Here...
        Else

            ToolStripComboBox1.Enabled = False
            ToolStripComboBox2.Enabled = False

            bConnected = True
            ToolStripButton1.Text = "Disconnect"

            ' Add Connect Code Here....
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        TextBox1.Select(TextBox1.Text.Length, 0)
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            ' Send Data
            If TextBox2.Text.Length() = 0 Then
                e.Handled = True
                Return
            End If

            sendTerminalData(TextBox2.Text)
            TextBox2.Text = ""

            e.Handled = True
        End If
    End Sub

    Public Sub sendTerminalData(ByVal sText As String)
        TextBox1.Text = TextBox1.Text & sText & vbNewLine
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sendTerminalData(TextBox2.Text)
        TextBox2.Text = ""
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text.Length = 0 Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click

    End Sub
End Class
