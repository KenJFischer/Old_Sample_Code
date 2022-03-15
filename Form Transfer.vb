
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmTransfer
    Dim strPath As String = Application.StartupPath

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        ' Create a transfer by generating transactions for the associated accounts
        ' Save transaction info to the corresponding text files
        Dim strFrom As String = lstFrom.SelectedItem.ToString().Substring(0, lstFrom.SelectedItem.ToString().IndexOf("("c)).Trim()
        Dim strTo As String = lstTo.SelectedItem.ToString().Substring(0, lstTo.SelectedItem.ToString().IndexOf("("c)).Trim()
        Dim dblTransferAmount As Double
        Dim strFormatDate As String
        Dim intHighestIDFrom As Integer
        Dim intHighestIDTo As Integer
        Dim outFile As IO.StreamWriter
        strFormatDate = Format(Now(), "MM/dd/yyyy")
        Double.TryParse(txtTransferAmount.Text, dblTransferAmount)

        If dblTransferAmount > 0 Then
            Try
                ' Set up string arrays to determine the highest transaction ID
                Dim strTransFrom() As String = IO.File.ReadAllLines(strPath & "\Accounts\" & strFrom & ".txt")
                Dim strTransTo() As String = IO.File.ReadAllLines(strPath & "\Accounts\" & strTo & ".txt")
                ' Generate and write new transactions to the appropriate account text files
                ' Transaction for the "from" account
                Integer.TryParse(strTransFrom(0), intHighestIDFrom)
                strTransFrom(0) = (intHighestIDFrom + 1).ToString()
                outFile = IO.File.CreateText(strPath & "\Accounts\" & strFrom & ".txt")
                For Each strTrans As String In strTransFrom
                    outFile.WriteLine(strTrans)
                Next strTrans
                outFile.WriteLine($"{strFormatDate}|{-dblTransferAmount}|Transfer from {strFrom} to {strTo}|Transfer|{intHighestIDFrom}")
                outFile.Close()

                ' Transaction for the "to" account
                Integer.TryParse(strTransTo(0), intHighestIDTo)
                strTransTo(0) = (intHighestIDTo + 1).ToString()
                outFile = IO.File.CreateText(strPath & "\Accounts\" & strTo & ".txt")
                For Each strTrans As String In strTransTo
                    outFile.WriteLine(strTrans)
                Next strTrans
                outFile.WriteLine($"{strFormatDate}|{dblTransferAmount}|Transfer from {strFrom} to {strTo}|Transfer|{intHighestIDTo}")
                outFile.Close()

                Dim Main As frmAccount = TryCast(Application.OpenForms("frmAccount"), frmAccount)

                ' Update totals in Account Listings.txt
                ' First determine old account totals
                Dim strAllAccts() As String = IO.File.ReadAllLines(strPath & "\Accounts\Account Listings.txt")
                Dim strFromAcct As String
                Dim strToAcct As String
                Dim dblTotal As Double
                ' Adjust totals for the "from" account
                strFromAcct = strAllAccts(lstFrom.SelectedIndex)
                Double.TryParse(strFromAcct.Substring(strFromAcct.IndexOf("("c) + 1, (strFromAcct.IndexOf(")"c) - 1) - strFromAcct.IndexOf("("c)), dblTotal)
                dblTotal += -dblTransferAmount
                Main.UpdateNameTotals(lstFrom.SelectedIndex, dblTotal)
                ' Adjust totals for the "to" account
                Dim intIndex As Integer
                If lstTo.SelectedIndex >= lstFrom.SelectedIndex Then
                    intIndex = lstTo.SelectedIndex + 1
                Else
                    intIndex = lstTo.SelectedIndex
                End If
                strToAcct = strAllAccts(intIndex)
                Double.TryParse(strToAcct.Substring(strToAcct.IndexOf("("c) + 1, (strToAcct.IndexOf(")"c) - 1) - strToAcct.IndexOf("("c)), dblTotal)
                dblTotal += dblTransferAmount
                Main.UpdateNameTotals(intIndex, dblTotal)


                ' Update the open account if it was one of the accounts involved in the transfer
                If lstFrom.SelectedItem.ToString().Substring(0, lstFrom.SelectedItem.ToString().IndexOf("("c)).Trim() =
                frmAccount.strAcct.Substring(0, frmAccount.strAcct.IndexOf("("c)).Trim() Then
                    ' If the account is being transferred from
                    frmAccount.strAcct = strFromAcct
                    Main.UpdateTransactions(-dblTransferAmount, $"Transfer from {strFrom} to {strTo}", "Transfer", strFormatDate, intHighestIDFrom, True,
                                            frmAccount.dgvTransaction.RowCount)
                    Main.UpdateTotals(-dblTransferAmount, 1, False)
                ElseIf lstTo.SelectedItem.ToString().Substring(0, lstTo.SelectedItem.ToString().IndexOf("("c)).Trim() =
                frmAccount.strAcct.Substring(0, frmAccount.strAcct.IndexOf("("c)).Trim() Then
                    ' If the account is being transferred to
                    frmAccount.strAcct = strToAcct
                    Main.UpdateTransactions(dblTransferAmount, $"Transfer from {strFrom} to {strTo}", "Transfer", strFormatDate, intHighestIDTo, True,
                                            frmAccount.dgvTransaction.RowCount)
                    Main.UpdateTotals(dblTransferAmount, 1, False)
                End If
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Transfer failed:" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Transfer amount must be greater than zero.", "Invalid Amount", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtTransferAmount.SelectAll()
        End If
    End Sub

    Private Sub frmTransfer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate the listbox with the accounts
        Dim inFile As IO.StreamReader
        Dim strAcct As String
        inFile = IO.File.OpenText(strPath & "\Accounts\Account Listings.txt")
        ' Read the file and populate the listbox
        Do Until inFile.Peek = -1
            strAcct = inFile.ReadLine()
            lstFrom.Items.Add(strAcct)
        Loop
        inFile.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub txtTransferAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTransferAmount.KeyPress
        ' Handle and format text entry
        ' Accept only numbers, backspace, and period (can't transfer negative amounts)
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso e.KeyChar <> ControlChars.Back _
           AndAlso e.KeyChar <> "." Then
            e.Handled = True
        ElseIf txtTransferAmount.Text.Contains(".") Then
            ' Allow only two digits to be entered after the period
            Dim strAfterDecimal() As String = txtTransferAmount.Text.Split("."c)
            If strAfterDecimal(1).Length >= 2 AndAlso e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If

            ' Allow only one period to be entered 
            If e.KeyChar = "." Then
                e.Handled = True
            End If
        End If
    End Sub


    Private Sub txtTransferAmount_TextChanged(sender As Object, e As EventArgs) Handles txtTransferAmount.TextChanged
        GenerateSummary()
    End Sub

    Private Sub lstFrom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstFrom.SelectedIndexChanged
        lstTo.Items.Clear()
        If lstFrom.SelectedIndex <> -1 Then
            Dim strSelectedFrom As String = lstFrom.SelectedItem.ToString()
            For Each strFrom As String In lstFrom.Items
                If strFrom = strSelectedFrom Then
                Else
                    lstTo.Items.Add(strFrom)
                End If
            Next strFrom
            GenerateSummary()
        End If
    End Sub

    Private Sub lstTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTo.SelectedIndexChanged
        GenerateSummary()
    End Sub

    ' Sub to generate the text in lblSummary
    Private Sub GenerateSummary()
        If lstFrom.SelectedIndex <> -1 AndAlso lstTo.SelectedIndex <> -1 AndAlso txtTransferAmount.Text <> String.Empty Then
            btnConfirm.Enabled = True
            lblSummary.Text = $"{txtTransferAmount.Text} will be transferred from {lstFrom.SelectedItem} to {lstTo.SelectedItem}"
        Else
            btnConfirm.Enabled = False
            lblSummary.Text = "Select the account to transfer from and the account to transfer to, and then enter the amount to transfer."
        End If
    End Sub
End Class