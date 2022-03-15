
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmAccount
    Dim strPath As String = Application.StartupPath
    Dim boolIsLoaded As Boolean = False
    ' Variables to determine how to sort dgvTransactions
    Dim intSortedColumn As Integer
    Dim SortedDirection As ComponentModel.ListSortDirection
    ' String to hold the account name with the account amount total information included
    Public Shared strAcct As String
    ' Integer to hold the index of the selected account in frmAccountSelector
    Public Shared intAcctIndex As Integer
    ' Holds transaction information
    Public Shared dblAmount As Double
    Public Shared strMemo As String
    Public Shared strVendor As String
    Public Shared strDate As String
    Public Shared intID As Integer
    Public Shared isNewTrans As Boolean ' true if new trans, false if editing trans

    ' Sub for sorting dgvTransactions by a specified column
    Public Sub SortTransactions()
        dgvTransaction.Sort(dgvTransaction.Columns(intSortedColumn), SortedDirection)
    End Sub

    ' Sub for loading transactions
    Public Sub LoadTransactions()
        ' Load and display transactions
        Dim strShortAcct As String = strAcct.Substring(0, strAcct.IndexOf("("c)).Trim()
        Dim inFile As IO.StreamReader
        Dim strTransInfo(4) As String
        Dim intTransID As Integer
        ' Variables for updating totals
        Dim dblTransAmount As Double
        Dim dblTransTotalDebits As Double
        Dim dblTransTotalDeposits As Double
        Dim dblTransTotal As Double
        Try
            If IO.File.Exists(strPath & "\Accounts\" & strShortAcct & ".txt") Then
                inFile = IO.File.OpenText(strPath & "\Accounts\" & strShortAcct & ".txt")
                If inFile.ReadLine <> "1" Then
                    ' Read the file for transactions and display any that are found
                    Do Until inFile.Peek = -1
                        strTransInfo = inFile.ReadLine().Split("|"c)
                        Double.TryParse(strTransInfo(1), dblTransAmount)
                        Integer.TryParse(strTransInfo(4), intTransID)
                        UpdateTransactions(dblTransAmount, strTransInfo(2), strTransInfo(3),
                                       strTransInfo(0), intTransID, True, dgvTransaction.RowCount)
                        ' Update totals information
                        If dblTransAmount < 0 Then
                            dblTransTotalDebits += dblTransAmount
                        Else
                            dblTransTotalDeposits += dblTransAmount
                        End If
                        dblTransTotal += dblTransAmount
                    Loop
                End If
                inFile.Close()
            Else
                MessageBox.Show(strShortAcct & ".txt not found. Unable to load transactions.", "File Not Found",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Account loading failed:" & Environment.NewLine & ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' Set totals
        lblTotalDebits.Text = dblTransTotalDebits.ToString("N2")
        lblTotalDeposits.Text = dblTransTotalDeposits.ToString("N2")
        lblAcctTotal.Text = dblTransTotal.ToString("N2")
        ' Set total labels color
        SetLblColors(dblTransTotalDebits, 0)
        SetLblColors(dblTransTotalDeposits, 1)
        SetLblColors(dblTransTotal, 2)
        dgvTransaction.Sort(dgvTransaction.Columns(0), ComponentModel.ListSortDirection.Descending)
        dgvTransaction.ClearSelection()
    End Sub

    Public Sub EditTotals(ByVal dblNewVal As Double, ByVal dblOldVal As Double)
        Dim dblTxtAmount As Double
        Dim dblOtherTxtAmount As Double
        ' Remove old values
        If dblOldVal < 0 Then
            Double.TryParse(lblTotalDebits.Text, dblTxtAmount)
            dblTxtAmount -= dblOldVal
            lblTotalDebits.Text = dblTxtAmount.ToString("N2")
            SetLblColors(dblTxtAmount, 0)
        Else
            Double.TryParse(lblTotalDeposits.Text, dblTxtAmount)
            dblTxtAmount -= dblOldVal
            lblTotalDeposits.Text = dblTxtAmount.ToString("N2")
            SetLblColors(dblTxtAmount, 1)
        End If
        ' Add new values
        If dblNewVal < 0 Then
            Double.TryParse(lblTotalDebits.Text, dblTxtAmount)
            dblTxtAmount += dblNewVal
            lblTotalDebits.Text = dblTxtAmount.ToString("N2")
            SetLblColors(dblTxtAmount, 0)
        Else
            Double.TryParse(lblTotalDeposits.Text, dblTxtAmount)
            dblTxtAmount += dblNewVal
            lblTotalDeposits.Text = dblTxtAmount.ToString("N2")
            SetLblColors(dblTxtAmount, 1)
        End If
        ' Set grand total label text and colors
        Double.TryParse(lblTotalDebits.Text, dblTxtAmount)
        Double.TryParse(lblTotalDeposits.Text, dblOtherTxtAmount)
        dblTxtAmount += dblOtherTxtAmount
        lblAcctTotal.Text = dblTxtAmount.ToString("N2")
        SetLblColors(dblTxtAmount, 2)
        UpdateNameTotals(intAcctIndex, dblTxtAmount)
    End Sub

    Public Sub SetLblColors(ByVal dblTotal As Double, ByVal intLblIndex As Integer) ' intLblIndex = 0 for debits, 1 for deposits, and 2 for grand total
        Select Case intLblIndex
            Case 0
                ' Set debit label color
                If dblTotal = 0 Then
                    lblTotalDebits.ForeColor = Color.Black
                Else
                    lblTotalDebits.ForeColor = Color.Red
                End If
            Case 1
                ' Set deposit label color
                If dblTotal = 0 Then
                    lblTotalDeposits.ForeColor = Color.Black
                Else
                    lblTotalDeposits.ForeColor = Color.FromArgb(0, 192, 0)
                End If
            Case 2
                ' Set total label color
                If dblTotal < 0 Then
                    lblAcctTotal.ForeColor = Color.Red
                ElseIf dblTotal = 0 Then
                    lblAcctTotal.ForeColor = Color.Black
                Else
                    lblAcctTotal.ForeColor = Color.FromArgb(0, 192, 0)
                End If
        End Select
    End Sub

    ' Sub for updating total labels6
    Public Sub UpdateTotals(ByVal dblChangeAmount As Double, ByVal intDeleter As Integer, ByVal boolUpdateNameTotals As Boolean)
        ' intDeleter is either 1 or -1; it is 1 normally and -1 when deleting transactions; it is used to adjust the totals correctly when deleting
        ' Adjust total values by a provided double value
        Dim dblTotal As Double
        If dblChangeAmount < 0 Then
            Double.TryParse(lblTotalDebits.Text, dblTotal)
            dblTotal += (dblChangeAmount * intDeleter)
            lblTotalDebits.Text = dblTotal.ToString("N2")
            ' Set label color
            SetLblColors(dblTotal, 0)
        Else
            Double.TryParse(lblTotalDeposits.Text, dblTotal)
            dblTotal += (dblChangeAmount * intDeleter)
            lblTotalDeposits.Text = dblTotal.ToString("N2")
            ' Set label color
            SetLblColors(dblTotal, 1)
        End If
        Double.TryParse(lblAcctTotal.Text, dblTotal)
        dblTotal += (dblChangeAmount * intDeleter)
        lblAcctTotal.Text = dblTotal.ToString("N2")

        ' Set lblAcctTotal color based on amount
        SetLblColors(dblTotal, 2)

        If boolUpdateNameTotals Then
            UpdateNameTotals(intAcctIndex, dblTotal)
        End If
    End Sub

    ' Sub for adjusting totals in the account names in Account Listings.txt
    Public Sub UpdateNameTotals(ByVal intLineIndex As Integer, ByVal dblTotal As Double)
        ' Update total information in Account Listings.txt
        Dim strSingleAcct As String
        Try
            Dim strAllAccts() As String = IO.File.ReadAllLines(strPath & "\Accounts\Account Listings.txt")
            strSingleAcct = strAllAccts(intLineIndex)
            strAllAccts(intLineIndex) = strSingleAcct.Substring(0, strSingleAcct.IndexOf("("c)).Trim() & "  (" & dblTotal.ToString("N2") & ")"
            IO.File.WriteAllLines(strPath & "\Accounts\Account Listings.txt", strAllAccts)
        Catch ex As Exception
            MessageBox.Show("Account totals could not be updated:" & Environment.NewLine & ex.Message,
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Sub for updating transaction info when a new transaction is created
    Public Sub UpdateTransactions(ByVal dblAmount As Double, ByVal strMemo As String, ByVal strVendor As String, ByVal strDate As String,
                                  ByVal intID As Integer, ByVal addRow As Boolean, ByVal intIndex As Integer)
        ' Add the transaction information to the DataGridView
        If addRow Then
            dgvTransaction.Rows.Add(strDate, dblAmount.ToString("N2"), strMemo, strVendor, intID)
        Else
            dgvTransaction.Rows.RemoveAt(intIndex)
            dgvTransaction.Rows.Insert(intIndex, strDate, dblAmount.ToString("N2"), strMemo, strVendor, intID)
            dgvTransaction.AutoResizeRows()
        End If

        ' Set color of text based on transaction amount
        If dblAmount < 0 Then
            dgvTransaction.Item(1, intIndex).Style.ForeColor = Color.Red
        Else
            dgvTransaction.Item(1, intIndex).Style.ForeColor = Color.Green
        End If
        dgvTransaction.Item(1, intIndex).Value = dblAmount.ToString("N2")

        ' Select the newly added row and then sort dgvTransaction based on date
        dgvTransaction.ClearSelection()
        dgvTransaction.Rows(intIndex).Selected = True
        SortTransactions()
    End Sub

    ' Sub procedure to remove orphaned account references (accounts listed in Account Listings.txt with no corresponding account file)
    Private Sub RemoveOrphan(ByVal intLine As Integer)
        ' Delete orphaned reference in Account Listings.txt
        Dim outFile As IO.StreamWriter
        Try
            outFile = IO.File.CreateText(strPath & "\Accounts\Account Listings.txt")
            Dim strAccounts() As String = IO.File.ReadAllLines(strPath & "\Accounts\Account Listings.txt")
            For intIndex As Integer = 0 To strAccounts.Length - 1
                If intIndex <> intLine Then
                    outFile.WriteLine(strAccounts(intIndex))
                End If
            Next intIndex
            outFile.Close()
        Catch ex As Exception
            MessageBox.Show("Account deletion failed:" & Environment.NewLine & ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        frmAccountSelector.boolOrphan = True
        Dim popUp As New frmAccountSelector
        popUp.Show()
        Me.Close()
    End Sub

    Private Sub frmAccount_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If frmAccountSelector.intOrphanIndex = -1 Then
            strAcct = frmAccountSelector.strSendAcct
            Dim strShortAcct As String = strAcct.Substring(0, strAcct.IndexOf("("c)).Trim()
            ' Set the form text and the account name label
            Me.Text = "Account Manager:  " & strShortAcct
            lblHeader.Text = "&Transactions for:  " & strShortAcct
            LoadTransactions()
            boolIsLoaded = True
        Else
            ' Remove orphaned account
            RemoveOrphan(frmAccountSelector.intOrphanIndex)
        End If
    End Sub

    Private Sub btnNewTrans_Click(sender As Object, e As EventArgs) Handles btnNewTrans.Click
        ' Create a new transaction and add it to the list of transactions
        ' Pop up the form for a new transaction
        isNewTrans = True
        Dim popUp As New frmTransaction
        popUp.ShowDialog()

        If strMemo <> "|" Then      ' strMemo only equals "|" when the user cancels the new transaction
            ' Add items to the transaction DataGridView
            UpdateTransactions(dblAmount, strMemo, strVendor, strDate, intID, True, dgvTransaction.RowCount)
            ' Adjust total values
            UpdateTotals(dblAmount, 1, True)
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If dgvTransaction.SelectedRows.Count <> 0 Then
            isNewTrans = False
            ' Get all data from selected row
            strDate = dgvTransaction.Item(0, dgvTransaction.SelectedRows(0).Index).Value.ToString()
            Double.TryParse(dgvTransaction.Item(1, dgvTransaction.SelectedRows(0).Index).Value.ToString(), dblAmount)
            strMemo = dgvTransaction.Item(2, dgvTransaction.SelectedRows(0).Index).Value.ToString()
            strVendor = dgvTransaction.Item(3, dgvTransaction.SelectedRows(0).Index).Value.ToString()
            Integer.TryParse(dgvTransaction.Item(4, dgvTransaction.SelectedRows(0).Index).Value.ToString(), intID)
            Dim popUp As New frmTransaction
            popUp.ShowDialog()
            UpdateTransactions(dblAmount, strMemo, strVendor, strDate, intID, False, dgvTransaction.SelectedRows(0).Index)
        Else
            MessageBox.Show("Select a transaction to edit.", "Selection Needed", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub btnDeleteTrans_Click(sender As Object, e As EventArgs) Handles btnDeleteTrans.Click
        ' Delete the transaction
        Dim strAcctPath As String = strPath & "\Accounts\" & strAcct.Substring(0, strAcct.IndexOf("("c)).Trim & ".txt"
        Dim intIndex As Integer = 0
        Dim outFile As IO.StreamWriter
        Dim result As DialogResult

        ' Ensure that rows have been selected
        If dgvTransaction.SelectedRows.Count = 0 Then
            MessageBox.Show("Select one or more transactions to delete.", "Select a Transaction",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            result = MessageBox.Show("Permanently delete the selected transaction(s)?", "Delete Transaction", MessageBoxButtons.YesNo,
                MessageBoxIcon.Information)
            If result = DialogResult.Yes Then
                SortTransactions()
                ' Determine which rows are selected
                Dim dblAmounts(dgvTransaction.SelectedRows.Count - 1) As Double
                Try
                    Dim strTrans() As String = IO.File.ReadAllLines(strAcctPath)
                    outFile = IO.File.CreateText(strAcctPath)
                    For Each row As DataGridViewRow In dgvTransaction.SelectedRows
                        strTrans(row.Index + 1) = String.Empty
                    Next row

                    For Each strTransaction As String In strTrans
                        If strTransaction = String.Empty Then ' do nothing
                        Else
                            outFile.WriteLine(strTransaction)
                        End If
                    Next strTransaction

                    ' Remove the selected rows from dgvTransaction and adjust totals
                    Dim dblDeletedAmount As Double
                    For Each row As DataGridViewRow In dgvTransaction.SelectedRows
                        Double.TryParse(row.Cells.Item("Transaction_Amount").Value.ToString(), dblDeletedAmount)
                        UpdateTotals(dblDeletedAmount, -1, True)
                        dgvTransaction.Rows.Remove(row)
                    Next row
                    outFile.Close()
                Catch ex As Exception
                    MessageBox.Show("Transaction deletion failed:" & Environment.NewLine & ex.Message,
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    Private Sub btnChangeAcct_Click(sender As Object, e As EventArgs) Handles btnChangeAcct.Click
        ' Open frmAccountSelector and close this form
        Dim popUp As New frmAccountSelector
        popUp.Show()
        Me.Close()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnTransfer_Click(sender As Object, e As EventArgs) Handles btnTransfer.Click
        ' Display the form for creating new transfers
        Dim popUp As New frmTransfer
        popUp.ShowDialog()
    End Sub

    Private Sub dgvTransaction_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvTransaction.ColumnHeaderMouseClick
        ' Assign values to variables used in sorting dgvTransaction
        If e.ColumnIndex <> 2 Then
            intSortedColumn = e.ColumnIndex
            If dgvTransaction.SortOrder = SortOrder.Ascending Then
                SortedDirection = ComponentModel.ListSortDirection.Ascending
            Else
                SortedDirection = ComponentModel.ListSortDirection.Descending
            End If
        End If
    End Sub

    Private Sub dgvTransaction_SortCompare(sender As Object, e As DataGridViewSortCompareEventArgs) Handles dgvTransaction.SortCompare
        ' Sort transaction amounts and IDs by their double values, rather than comparing their strings
        If boolIsLoaded Then
            If e.Column.Index = 1 OrElse e.Column.Index = 4 Then
                Dim dblRow1 As Double
                Dim dblRow2 As Double
                Double.TryParse(dgvTransaction.Item(e.Column.Index, e.RowIndex1).Value.ToString(), dblRow1)
                Double.TryParse(dgvTransaction.Item(e.Column.Index, e.RowIndex2).Value.ToString(), dblRow2)
                If dblRow1 < dblRow2 Then
                    e.SortResult = -1
                ElseIf dblRow1 = dblRow2 Then
                    e.SortResult = 0
                Else
                    e.SortResult = 1
                End If
                e.Handled = True
            ElseIf e.Column.Index = 0 Then ' Sort dates based on their datetime value
                Dim date1 As DateTime = DateTime.ParseExact(dgvTransaction.Item(e.Column.Index, e.RowIndex1).Value.ToString(), "MM/dd/yyyy",
                                                            Globalization.CultureInfo.InvariantCulture)
                Dim date2 As DateTime = DateTime.ParseExact(dgvTransaction.Item(e.Column.Index, e.RowIndex2).Value.ToString(), "MM/dd/yyyy",
                                                            Globalization.CultureInfo.InvariantCulture)
                If date1.CompareTo(date2) < 0 Then
                    e.SortResult = -1
                ElseIf date1.CompareTo(date2) = 0 Then
                    e.SortResult = 0
                Else
                    e.SortResult = 1
                End If
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub dgvTransaction_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvTransaction.CellMouseDoubleClick
        ' Double click to edit cell
        If e.RowIndex > -1 Then
            btnEdit.PerformClick()
        End If
    End Sub

    ' Display the corresponding labels' text in tooltips.  This is useful if a label's text 
    ' overflows the screen by being longer than 46 characters (although this is extremely unlikely).
    Private Sub lblTotalDebits_MouseHover(sender As Object, e As EventArgs) Handles lblTotalDebits.MouseHover
        ToolTip1.SetToolTip(lblTotalDebits, lblTotalDebits.Text)
    End Sub

    Private Sub lblTotalDeposits_MouseHover(sender As Object, e As EventArgs) Handles lblTotalDeposits.MouseHover
        ToolTip1.SetToolTip(lblTotalDeposits, lblTotalDeposits.Text)
    End Sub

    Private Sub lblAcctTotal_MouseHover(sender As Object, e As EventArgs) Handles lblAcctTotal.MouseHover
        ToolTip1.SetToolTip(lblAcctTotal, lblAcctTotal.Text)
    End Sub

    Private Sub frmAccount_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        dgvTransaction.ClearSelection()
    End Sub

    Private Sub dgvTransaction_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dgvTransaction.KeyPress
        ' acts as the accept button handler
        If e.KeyChar = ChrW(Keys.Return) Then
            ' Move selected row up by one, since pressing enter advances selected row by one
            If dgvTransaction.SelectedRows.Count > 0 AndAlso dgvTransaction.SelectedRows(0).Index > 0 Then
                dgvTransaction.Rows(dgvTransaction.SelectedRows(0).Index - 1).Selected = True
            End If
            btnNewTrans.PerformClick()
            e.Handled = True
        End If
    End Sub

    ' Zoom in and out 
    Private Sub btnZoomIn_Click(sender As Object, e As EventArgs) Handles btnZoomIn.Click
        If dgvTransaction.Font.Size < 21 Then
            dgvTransaction.Font = New Font("Segoe UI", dgvTransaction.Font.Size + 1, FontStyle.Regular)
            dgvTransaction.AutoResizeColumns()
            dgvTransaction.AutoResizeRows()
        End If
    End Sub

    Private Sub btnZoomOut_Click(sender As Object, e As EventArgs) Handles btnZoomOut.Click
        If dgvTransaction.Font.Size > 7 Then
            dgvTransaction.Font = New Font("Segoe UI", dgvTransaction.Font.Size - 1, FontStyle.Regular)
            dgvTransaction.AutoResizeColumns()
            dgvTransaction.AutoResizeRows()
        End If
    End Sub
End Class