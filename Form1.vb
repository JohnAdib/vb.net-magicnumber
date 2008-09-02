Public Class Form1

    Dim UserNumber As Long
    Dim GuessNumber As Long = -1
    Dim TimeCounter As Integer
    Dim CurrentLanguage As String = "En"

    Private Sub Btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCalc1.Click, _
                                                                BtnCalc2.Click, BtnCalc3.Click, BtnCalc4.Click, _
                                                                BtnCalc5.Click, BtnCalc6.Click, BtnCalc7.Click, _
                                                                BtnCalc8.Click, BtnCalc9.Click, BtnCalc0.Click
        txtCalcNumber.Text = txtCalcNumber.Text + sender.Text
        txtCalcNumber.Focus()
        txtCalcNumber.SelectionStart = txtCalcNumber.Text.Length
    End Sub

    Private Sub BtnCalcClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCalcClear.Click
        txtCalcNumber.Text = ""
        txtCalcNumber.Focus()
        txtCalcNumber.SelectionStart = txtCalcNumber.Text.Length
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click, NewToolStripMenuItem3.Click, NewToolStripMenuItem2.Click
        EndRange.Text = txtEnd.Text
        StartRange.Text = txtStart.Text
        UserNumber = 0

        If txtStart.Text.Length < 1 Then txtStart.Text = "1"
        If txtEnd.Text.Length < 1 Then txtEnd.Text = "100000"
        If (txtEnd.Text - txtStart.Text) < 100 Then txtEnd.Text = txtStart.Text + 100
        jStatusBar.BackColor = Color.SkyBlue
        Randomize()
        GuessNumber = Int(txtEnd.Text - txtStart.Text + 1) * Rnd() + txtStart.Text
        txtCalcNumber.Text = ""
        If CurrentLanguage = "En" Then

            If btnNew.Text = "Start" Then
                GameDetail.Text = "Detail of Game"
                jStatusBarLabel2.Text = "Guess Number"
                BtnCalcEnableStatus()
                btnNew.Text = "New"
                TimeCounter = txtEnd.Text - txtStart.Text
                If TimeCounter < 300 Then
                    TimeCounter = 300
                ElseIf TimeCounter > 3000 Then
                    TimeCounter = 3000
                End If
                txtCalcNumber.Focus()
                btnNew.TabStop = False
                TimeProgressBar.Maximum = TimeCounter + 1
                TimeProgressBar.Value = 0
                lblTime.Text = "Time"
                TryCounter.Text = 0
                GameTimer.Start()
            Else
                jStatusBarLabel2.Text = "Once again try your luck"
                BtnCalcEnableStatus(False)
                btnNew.TabStop = True
                btnNew.Focus()
                btnNew.Text = "Start"
                GameTimer.Stop()
                GuessNumber = -1
            End If

        ElseIf CurrentLanguage = "Fa" Then

            If btnNew.Text = "شروع" Then
                GameDetail.Text = "جزئیات بازی"
                jStatusBarLabel2.Text = "عدد جادویی را حدس بزن"
                BtnCalcEnableStatus()
                btnNew.Text = "جدید"
                TimeCounter = txtEnd.Text - txtStart.Text
                If TimeCounter < 300 Then
                    TimeCounter = 300
                ElseIf TimeCounter > 3000 Then
                    TimeCounter = 3000
                End If
                txtCalcNumber.Focus()
                btnNew.TabStop = False
                TimeProgressBar.Maximum = TimeCounter + 1
                TimeProgressBar.Value = 0
                lblTime.Text = "زمان"
                TryCounter.Text = 0
                GameTimer.Start()
            Else
                jStatusBarLabel2.Text = "یکبار دیگه شانس خودت رو امتحان کن!"
                BtnCalcEnableStatus(False)
                btnNew.TabStop = True
                btnNew.Focus()
                btnNew.Text = "شروع"
                GameTimer.Stop()
                GuessNumber = -1
            End If
        End If
    End Sub

    Private Sub BtnCalcOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCalcOK.Click
        Dim StatusMsg As String = ""

        If txtCalcNumber.Text.Length > 0 And txtCalcNumber.Text.Length < 10 Then _
        UserNumber = Convert.ToInt64(txtCalcNumber.Text) Else UserNumber = 0
        txtCalcNumber.Text = ""

        If UserNumber < StartRange.Text Then                                    ' do nothing
        ElseIf UserNumber > EndRange.Text Then                                  ' do nothing
        Else
            If UserNumber < GuessNumber Then StartRange.Text = UserNumber Else EndRange.Text = UserNumber
        End If

        If CurrentLanguage = "En" Then

            If GuessNumber = -1 Then
                StatusMsg = "Press Start"
            ElseIf UserNumber = 0 Then
                StatusMsg = "Input Number and then press OK"
            ElseIf UserNumber < txtStart.Text Then
                StatusMsg = "Invalid! you must input larger than " & txtStart.Text
                TryCounter.Text = TryCounter.Text + 1
            ElseIf UserNumber > txtEnd.Text Then
                StatusMsg = "Invalid! you must input less than " & txtEnd.Text
                TryCounter.Text = TryCounter.Text + 1
            ElseIf GuessNumber > UserNumber Then
                StatusMsg = "Greater than " & vbTab & UserNumber & vbTab & "( x> " & UserNumber & " )"
                TryCounter.Text = TryCounter.Text + 1
            ElseIf GuessNumber < UserNumber Then
                StatusMsg = "Less than " & vbTab & UserNumber & vbTab & "( x< " & UserNumber & " )"
                TryCounter.Text = TryCounter.Text + 1
            Else
                StatusMsg = "Congratulation, YOU WIN !!!"
                TryCounter.Text = TryCounter.Text + 1
                lblTime.Text = "You win with " & TryCounter.Text & " ties in " & TimeProgressBar.Value / 10 & " seconds "
                Call Me.btnNew_Click(Me, System.EventArgs.Empty)
                jStatusBar.BackColor = Color.GreenYellow
            End If

        ElseIf CurrentLanguage = "Fa" Then

            If GuessNumber = -1 Then
                StatusMsg = "بر روی شروع کلیک کن"
            ElseIf UserNumber = 0 Then
                StatusMsg = "عدد را وارد کرده و روی تایید کلیک کن"
            ElseIf UserNumber < txtStart.Text Then
                StatusMsg = "خطا! شما باید عددی بزرگتر از " & txtStart.Text & " وارد کنید."
                TryCounter.Text = TryCounter.Text + 1
            ElseIf UserNumber > txtEnd.Text Then
                StatusMsg = "خطا! شما باید عددی کوچکتر از " & txtEnd.Text & " وارد کنید."
                TryCounter.Text = TryCounter.Text + 1
            ElseIf GuessNumber > UserNumber Then
                StatusMsg = "عدد جادویی بزرگتر از " & UserNumber & " است.   ( x> " & UserNumber & " )"
                TryCounter.Text = TryCounter.Text + 1
            ElseIf GuessNumber < UserNumber Then
                StatusMsg = "عدد جادویی کوچکتر از " & UserNumber & " است.   ( x< " & UserNumber & " )"
                TryCounter.Text = TryCounter.Text + 1
            Else
                StatusMsg = "شما عدد جادویی را پیدا کردید، تبریک میگم ^_^"
                TryCounter.Text = TryCounter.Text + 1
                lblTime.Text = "شما با " & TryCounter.Text & " خطا در" & TimeProgressBar.Value / 10 & " ثانیه بازی را تمام کردید."
                Call Me.btnNew_Click(Me, System.EventArgs.Empty)
                jStatusBar.BackColor = Color.GreenYellow
            End If

        End If

        txtCalcNumber.Focus()
        jStatusBarLabel2.Text = StatusMsg
        txtCalcNumber.SelectionStart = txtCalcNumber.Text.Length
        GameDetail.Text = GameDetail.Text & Chr(13) & TryCounter.Text & ". " & StatusMsg
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BtnCalcEnableStatus(False)
    End Sub

    Private Sub GameTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GameTimer.Tick
        If CurrentLanguage = "En" Then

            If TimeProgressBar.Value <= TimeCounter Then
                lblTime.Text = "Time : " & TimeProgressBar.Value / 10 & " / " & TimeCounter / 10 & " s"
                TimeProgressBar.Value = TimeProgressBar.Value + 1
            Else
                Call Me.btnNew_Click(Me, System.EventArgs.Empty)
                jStatusBarLabel2.Text = "Game Over"
                GameTimer.Stop()
                GameDetail.Text = GameDetail.Text & Chr(13) & TryCounter.Text & ". " & jStatusBarLabel2.Text
                jStatusBar.BackColor = Color.OrangeRed
            End If

        ElseIf CurrentLanguage = "Fa" Then

            If TimeProgressBar.Value <= TimeCounter Then
                lblTime.Text = "زمان : " & TimeProgressBar.Value / 10 & " / " & TimeCounter / 10 & " ثانیه"
                TimeProgressBar.Value = TimeProgressBar.Value + 1
            Else
                Call Me.btnNew_Click(Me, System.EventArgs.Empty)
                jStatusBarLabel2.Text = "زمان شما به اتمام رسید>_<"
                GameTimer.Stop()
                GameDetail.Text = GameDetail.Text & Chr(13) & TryCounter.Text & ". " & jStatusBarLabel2.Text
                jStatusBar.BackColor = Color.OrangeRed
            End If

        End If

    End Sub

    Private Sub BtnCalcBackspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCalcBackspace.Click
        If txtCalcNumber.Text.Length > 0 Then txtCalcNumber.Text \= 10
        txtCalcNumber.Focus()
        txtCalcNumber.SelectionStart = txtCalcNumber.Text.Length
    End Sub

    Private Sub txtCalcNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCalcNumber.KeyPress
        If (Asc(e.KeyChar) < 48) Or (Asc(e.KeyChar) > 57) Then e.Handled = True
        If (Asc(e.KeyChar) = 8) Then e.Handled = False
        If e.KeyChar = ChrW(13) Then Call Me.BtnCalcOK_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub txtStart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
                                                                                            txtStart.KeyPress, _
                                                                                            txtEnd.KeyPress
        If (Asc(e.KeyChar) < 48) Or (Asc(e.KeyChar) > 57) Then e.Handled = True
        If (Asc(e.KeyChar) = 8) Then e.Handled = False
        If e.KeyChar = ChrW(13) Then Me.ProcessTabKey(True)
    End Sub

    Private Sub GameDetail_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GameDetail.MouseClick
        txtCalcNumber.Focus()
    End Sub


    Private Sub AboutToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem2.Click
        Dim tmp As String
        If GuessNumber = -1 Then tmp = "E-mail" Else tmp = GuessNumber
        MsgBox("Javad.Evazzadeh@gmail.com ", , "" & tmp)
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem3.Click, ExitToolStripMenuItem2.Click
        Close()
    End Sub

    Private Sub BtnCalcEnableStatus(Optional ByVal Status As Boolean = True)
        BtnCalc1.Enabled = Status
        BtnCalc2.Enabled = Status
        BtnCalc3.Enabled = Status
        BtnCalc4.Enabled = Status
        BtnCalc5.Enabled = Status
        BtnCalc6.Enabled = Status
        BtnCalc7.Enabled = Status
        BtnCalc8.Enabled = Status
        BtnCalc9.Enabled = Status
        BtnCalc0.Enabled = Status
        BtnCalcOK.Enabled = Status
        txtEnd.Enabled = Not Status
        txtStart.Enabled = Not Status
        BtnCalcClear.Enabled = Status
        txtCalcNumber.Enabled = Status
        BtnCalcBackspace.Enabled = Status
    End Sub

    Private Sub Changelanguage(Optional ByVal Language As String = "En")
        Dim xFont As Font
        Dim xControl As Control
        jStatusBarLabel2.Text = ""
        If Language = "En" Then
            xFont = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
            For Each xControl In Me.Controls
                xControl.Font = xFont
            Next
            FileToolStripMenuItem2.Font = xFont
            HelpToolStripMenuItem2.Font = xFont
            jStatusBarLabel2.Font = xFont
            TabControl1.Font = xFont
            xFont = New Font("Microsoft Sans Serif", 14.25, FontStyle.Bold)
            btnNew.Font = xFont
            txtStart.Font = xFont
            txtEnd.Font = xFont
            txtCalcNumber.Font = xFont



            Me.Text = "Magic Number"
            lblTime.Text = "Time"
            Label2.Text = "End Range"
            Label1.Text = "Start Range"
            TabPage1.Text = "Game"
            TabPage2.Text = "Setting"
            TabPage3.Text = "Detail"
            BtnCalcOK.Text = "OK"
            If GameDetail.Text.Length = 73 Then _
            GameDetail.Text = "What is attractive, is not easy, not too difficult, but is difficult to reach easily."
            BtnCalcClear.Text = "Clear"
            BtnCalcBackspace.Text = "Backspace"
            NewToolStripMenuItem2.Text = "New"
            NewToolStripMenuItem3.Text = "New"
            FileToolStripMenuItem2.Text = "File"
            ExitToolStripMenuItem2.Text = "Exit"
            ExitToolStripMenuItem3.Text = "Exit"
            HelpToolStripMenuItem2.Text = "Help"
            AboutToolStripMenuItem2.Text = "About"
            RadioButton1.Text = "English" & " (En)"
            RadioButton2.Text = "Farsi" & " (فارسی)"
            GroupBox1.Text = "Language"
            FarsiToolStripMenuItem3.Text = "Farsi" & " (فارسی)"
            FarsiToolStripMenuItem5.Text = "Farsi" & " (فارسی)"
            EnglishToolStripMenuItem3.Text = "English" & " (En)"
            EnglishToolStripMenuItem4.Text = "English" & " (En)"
            LanguageToolStripMenuItem2.Text = "Language"
            LanguageToolStripMenuItem3.Text = "Language"
            ContentsToolStripMenuItem3.Text = "Contents"
            If btnNew.Text = "شروع" Then
                btnNew.Text = "Start"
            ElseIf btnNew.Text = "جدید" Then
                btnNew.Text = "New"
            End If

            MenuStrip1.RightToLeft = Windows.Forms.RightToLeft.No
            For Each control In Me.Controls
                control.RightToLeft = False
            Next

            EnglishToolStripMenuItem3.Enabled = False
            EnglishToolStripMenuItem4.Enabled = False
            FarsiToolStripMenuItem3.Enabled = True
            FarsiToolStripMenuItem5.Enabled = True
            FarsiToolStripMenuItem3.Checked = False
            FarsiToolStripMenuItem5.Checked = False
            EnglishToolStripMenuItem3.Checked = True
            EnglishToolStripMenuItem4.Checked = True
            CurrentLanguage = "En"

        ElseIf Language = "Fa" Then
            xFont = New Font("b koodak", 9.25, FontStyle.Bold)
            For Each xControl In Me.Controls
                xControl.Font = xFont
            Next
            FileToolStripMenuItem2.Font = xFont
            HelpToolStripMenuItem2.Font = xFont
            jStatusBarLabel2.Font = xFont
            xFont = New Font("b koodak", 15, FontStyle.Bold)
            btnNew.Font = xFont

            Me.Text = "عدد جادویی را حدس بزن"
            lblTime.Text = "زمان"
            Label2.Text = "پایان محدوده"
            Label1.Text = "شروع محدوده"
            TabPage1.Text = "بازی"
            TabPage2.Text = "تنظیمات"
            TabPage3.Text = "جزئیات بازی"
            BtnCalcOK.Text = "تایید"
            If GameDetail.Text.Length = 85 Then _
            GameDetail.Text = "آنچه جذاب است سهولت نیست، دشواری هم نیست، بلکه دشواری رسیدن به سهولت است."
            BtnCalcClear.Text = "پاک"
            BtnCalcBackspace.Text = "پسبر!!!"
            NewToolStripMenuItem2.Text = "بازی جدید"
            NewToolStripMenuItem3.Text = "بازی جدید"
            FileToolStripMenuItem2.Text = "فایل"
            ExitToolStripMenuItem2.Text = "خروج"
            ExitToolStripMenuItem3.Text = "خروج"
            HelpToolStripMenuItem2.Text = "کمک!"
            AboutToolStripMenuItem2.Text = "درباره ی ..."
            RadioButton1.Text = "انگلیسی" & " (English)"
            RadioButton2.Text = "فارسی" & " (Farsi)"
            GroupBox1.Text = "زبان"
            FarsiToolStripMenuItem3.Text = "فارسی" & " (Farsi)"
            FarsiToolStripMenuItem5.Text = "فارسی" & " (Farsi)"
            EnglishToolStripMenuItem3.Text = "انگلیسی" & " (English)"
            EnglishToolStripMenuItem4.Text = "انگلیسی" & " (English)"
            LanguageToolStripMenuItem2.Text = "زبان"
            LanguageToolStripMenuItem3.Text = "زبان"
            ContentsToolStripMenuItem3.Text = "راهنما"
            If btnNew.Text = "Start" Then
                btnNew.Text = "شروع"
            ElseIf btnNew.Text = "New" Then
                btnNew.Text = "جدید"
            End If

            For Each control In Me.Controls
                control.RightToLeft = Windows.Forms.RightToLeft.Yes
            Next

            MenuStrip1.RightToLeft = Windows.Forms.RightToLeft.Yes
            EnglishToolStripMenuItem3.Enabled = True
            EnglishToolStripMenuItem4.Enabled = True
            FarsiToolStripMenuItem3.Enabled = False
            FarsiToolStripMenuItem5.Enabled = False
            FarsiToolStripMenuItem3.Checked = True
            FarsiToolStripMenuItem5.Checked = True
            EnglishToolStripMenuItem3.Checked = False
            EnglishToolStripMenuItem4.Checked = False
            CurrentLanguage = "Fa"
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Changelanguage("En")
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Changelanguage("Fa")
    End Sub

    Private Sub EnglishToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnglishToolStripMenuItem3.Click, EnglishToolStripMenuItem4.Click
        RadioButton1.Checked = True
    End Sub

    Private Sub FarsiToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FarsiToolStripMenuItem3.Click, FarsiToolStripMenuItem5.Click
        RadioButton2.Checked = True
    End Sub

    Private Sub ContentsToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContentsToolStripMenuItem3.Click
        MsgBox("Coming soon!", , "Help")
    End Sub

End Class