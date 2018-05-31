Public Class Form1
    Dim A As Integer = 1
    Private Sub NewTabToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NewTabToolStripMenuItem.Click
        Dim RTB As New RichTextBox
        TabControl1.TabPages.Add(1, "Project " & A)
        TabControl1.SelectTab(A - 1)
        RTB.Name = "TE"
        RTB.Dock = DockStyle.Fill
        TabControl1.SelectedTab.Controls.Add(RTB)
        A = A + 1
    End Sub

    Private Sub CloseTabToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CloseTabToolStripMenuItem.Click
        If TabControl1.TabCount = 1 Then
            'Do Nothing
        Else
            TabControl1.TabPages.RemoveAt(TabControl1.SelectedIndex)
            TabControl1.SelectTab(TabControl1.TabPages.Count - 1)
            A = A - 1
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OPF As New OpenFileDialog
        Dim AllText As String = "", LineOfText As String = ""
        OPF.Filter = "Leopard Scripts|*.ls"
        OPF.ShowDialog()
        Try
            FileOpen(1, OPF.FileName, OpenMode.Input)
            Do Until EOF(1)
                LineOfText = LineInput(1)
                AllText = AllText & LineOfText & vbCrLf
            Loop
            CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Text = AllText
        Catch

        Finally
            FileClose(1)
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim SFD As New SaveFileDialog
        SFD.Filter = "Leopard Scripts|*.ls"
        SFD.ShowDialog()
        FileOpen(1, SFD.FileName, OpenMode.Output)
        PrintLine(1, CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Text)
        FileClose(1)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End

    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CutToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Paste()

    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Clear()

    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Redo()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Undo()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).SelectAll()
    End Sub

    Private Sub TimeDateToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TimeDateToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).SelectedText = TimeString & "/" & DateString
    End Sub

    Private Sub FindToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FindToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).FindForm()
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FontToolStripMenuItem.Click
        Dim FS As New FontDialog
        FS.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Font = FS.Font
    End Sub

    Private Sub ColorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ColorToolStripMenuItem.Click
        Dim FC As New ColorDialog
        FC.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).ForeColor = FC.Color
    End Sub

    Private Sub HighlightToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HighlightToolStripMenuItem.Click
        Dim HC As New ColorDialog
        HC.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).SelectionBackColor = HC.Color

    End Sub

    Private Sub PageColorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PageColorToolStripMenuItem.Click
        Dim BC As New ColorDialog
        BC.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).BackColor = BC.Color
    End Sub

    Private Sub DebugToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DebugToolStripMenuItem.Click
        For Each a As String In CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Text.Split(System.Environment.NewLine())
            If (a.Contains("Print=>")) Then
                Console.WriteLine(a.Substring(7))
            End If
            If (a.Contains("Cls")) Then
                Console.Clear()
            End If
            If (a.Contains("Title=>")) Then
                Console.Title = (a.Substring(7))
            End If
            If (a.Contains("Beep")) Then
                Beep()
            End If
            If (a.Contains("Msgbox=>")) Then
                MsgBox(a.Substring(8))
            End If
            If (a.Contains("Msgbox.Inputbox=>")) Then
                MsgBox(InputBox(a.Substring(17)) & (a.Substring(17)))
            End If
            If (a.Contains("Msgbox.Critical=>")) Then
                MsgBox(a.Substring(17), MsgBoxStyle.Critical)
            End If
            If (a.Contains("Msgbox.Information=>")) Then
                MsgBox(a.Substring(20), MsgBoxStyle.Information)
            End If
            If (a.Contains("Msgbox.Exclamation=>")) Then
                MsgBox(a.Substring(20), MsgBoxStyle.Exclamation)
            End If
            If (a.Contains("Msgbox.ApplicationModal=>")) Then
                MsgBox(a.Substring(17), MsgBoxStyle.ApplicationModal)
            End If
            If (a.Contains("Msgbox.OkOnly=>")) Then
                MsgBox(a.Substring(15), MsgBoxStyle.OkOnly)
            End If
            If (a.Contains("Msgbox.Yes&No=>")) Then
                MsgBox(a.Substring(15), MsgBoxStyle.YesNo)
            End If
            If (a.Contains("Msgbox.YesNo&Cancel=>")) Then
                MsgBox(a.Substring(21), MsgBoxStyle.YesNoCancel)
            End If
            If (a.Contains("Msgbox.Retry&Cancel=>")) Then
                MsgBox(a.Substring(21), MsgBoxStyle.RetryCancel)
            End If
            If (a.Contains("Msgbox.Ok&Cancel=>")) Then
                MsgBox(a.Substring(17), MsgBoxStyle.OkCancel)
            End If
            If (a.Contains("Msgbox.Question=>")) Then
                MsgBox(a.Substring(17), MsgBoxStyle.Question)
            End If
            If (a.Contains("Msgbox.AbortRetry&Ignore=>")) Then
                MsgBox(a.Substring(26), MsgBoxStyle.AbortRetryIgnore)
            End If
            If (a.Contains("Msgbox.Right=>")) Then
                MsgBox(a.Substring(26), MsgBoxStyle.MsgBoxRight)
            End If
            If (a.Contains("Msgbox.MsgHelp=>")) Then
                MsgBox(a.Substring(26), MsgBoxStyle.MsgBoxHelp)
            End If
            If (a.Contains("Msgbox.RtlReading=>")) Then
                MsgBox(a.Substring(26), MsgBoxStyle.MsgBoxRtlReading)
            End If
            If (a.Contains("Msgbox.Help=>")) Then
                MsgBox(a.Substring(13), MsgBoxStyle.MsgBoxHelp)
            End If
            If (a.Contains("Inputbox=>")) Then
                InputBox(a.Substring(10))
            End If
            If (a.Contains("Forecolor=>Cyan")) Then
                Console.ForegroundColor = ConsoleColor.Cyan
            End If
            If (a.Contains("Forecolor=> Blue")) Then
                Console.ForegroundColor = ConsoleColor.Black
            End If
            If (a.Contains("Forecolor=>Red")) Then
                Console.ForegroundColor = ConsoleColor.Red
            End If
            If (a.Contains("Forecolor=>Yellow")) Then
                Console.ForegroundColor = ConsoleColor.Yellow
            End If
            If (a.Contains("Forecolor=>Green")) Then
                Console.ForegroundColor = ConsoleColor.Green
            End If
            If (a.Contains("Forecolor=>White")) Then
                Console.ForegroundColor = ConsoleColor.White
            End If
            If (a.Contains("Forecolor=>Black")) Then
                Console.ForegroundColor = ConsoleColor.Black
            End If
            If (a.Contains("Forecolor=>DarkRed")) Then
                Console.ForegroundColor = ConsoleColor.DarkRed
            End If
            If (a.Contains("Forecolor=>DarkBlue")) Then
                Console.ForegroundColor = ConsoleColor.DarkBlue
            End If
            If (a.Contains("Forecolor=>DarkGreen")) Then
                Console.ForegroundColor = ConsoleColor.DarkGreen
            End If
            If (a.Contains("Forecolor=>DarkMagenta")) Then
                Console.ForegroundColor = ConsoleColor.DarkMagenta
            End If
            If (a.Contains("Forecolor=>Gray")) Then
                Console.ForegroundColor = ConsoleColor.Gray
            End If
            If (a.Contains("Forecolor=>DarkYellow")) Then
                Console.ForegroundColor = ConsoleColor.DarkYellow
            End If
            If (a.Contains("Backcolor=>Cyan")) Then
                Console.BackgroundColor = ConsoleColor.Cyan
            End If
            If (a.Contains("Backcolor=> Blue")) Then
                Console.BackgroundColor = ConsoleColor.Black
            End If
            If (a.Contains("Backcolor=>Red")) Then
                Console.BackgroundColor = ConsoleColor.Red
            End If
            If (a.Contains("Backcolor=>Yellow")) Then
                Console.BackgroundColor = ConsoleColor.Yellow
            End If
            If (a.Contains("Backcolor=>Green")) Then
                Console.BackgroundColor = ConsoleColor.Green
            End If
            If (a.Contains("Backcolor=>White")) Then
                Console.BackgroundColor = ConsoleColor.White
            End If
            If (a.Contains("Backcolor=>Black")) Then
                Console.BackgroundColor = ConsoleColor.Black
            End If
            If (a.Contains("Backcolor=>DarkRed")) Then
                Console.BackgroundColor = ConsoleColor.DarkRed
            End If
            If (a.Contains("Backcolor=>DarkBlue")) Then
                Console.BackgroundColor = ConsoleColor.DarkBlue
            End If
            If (a.Contains("Backcolor=>DarkGreen")) Then
                Console.BackgroundColor = ConsoleColor.DarkGreen
            End If
            If (a.Contains("Backcolor=>DarkMagenta")) Then
                Console.BackgroundColor = ConsoleColor.DarkMagenta
            End If
            If (a.Contains("Backcolor=>Gray")) Then
                Console.BackgroundColor = ConsoleColor.Gray
            End If
            If (a.Contains("Backcolor=>DarkYellow")) Then
                Console.BackgroundColor = ConsoleColor.DarkYellow
            End If
            If (a.Contains("About")) Then
                Console.Write("Leopard Script; Version: 0.1.0(ALPHA); All Rights Reserved")
                Console.ForegroundColor = ConsoleColor.Cyan
            End If
        Next

    End Sub
End Class
