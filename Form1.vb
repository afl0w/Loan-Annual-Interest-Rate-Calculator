Public Class Form1

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles dblrate.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCalc.Click

        ' Input loan amount, period, and interest rate variables  

        Dim dblloan As Double = txtcredit.Text
        Dim intPeriod As Integer = txtyears.Text
        Dim dblrate As Double = txtinterestrate.Text

        ' Output Variables 

        Dim intPeriodmonths As Integer
        Dim intCount As Integer
        Dim dblpayment As Double
        Dim dblTolinterest As Double
        Dim dblToprincipal As Double
        Dim dblBalance As Double
        Dim strOutput As String

        'Loan Calculation/ Process 

        dblrate /= 1200
        intPeriodmonths = 12 * intPeriod
        dblpayment = Pmt(dblrate, intPeriodmonths, -dblloan)
        dblBalance = dblloan
        For intCount = 1 To intPeriodmonths
            dblTolinterest = (dblBalance * dblrate)
            dblToprincipal = (dblpayment - dblTolinterest)
            dblBalance = (dblBalance - dblToprincipal)

            strOutput = intCount & "              " & dblpayment.ToString("C") & "              " & dblTolinterest.ToString("C") & "             " & dblToprincipal.ToString("C") & "               " & dblBalance.ToString("C")
            lstOut.Items.Add(strOutput)

        Next


    End Sub

    Private Sub Txtyears_TextChanged(sender As Object, e As EventArgs) Handles txtinterestrate.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Clear textbox function

        txtcredit.Clear()
        txtinterestrate.Clear()
        txtyears.Clear()
        lstOut.Items.Clear()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'Exit the program

        Close()

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
