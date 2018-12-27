Public Class Form1
    Dim vTotal As Double = 0 'always keeps the display total
    Dim vCalc As Double = 0 'always keeps the arithmatic total
    Dim vChain As String = ""
    Dim vOpera As String = "None"
    Dim vSenderType As String 'good
    Dim vVirgin As Boolean = True '?
    Dim vOpt As Boolean = False 'good
    Dim vTB As Double = 0
    Dim vMemory As Double = 0
    Dim vReg1 As Double = 0
    Dim vReg2 As Double = 0
    Dim vAns As Double = 0 'good
    Dim vLoop As Integer = 0
    Dim ListDisplay As New List(Of String)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click,
        Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button10.Click, Button12.Click,
        btnPlus.Click, btnMinus.Click, btnTimes.Click, btnDivide.Click, btnPercent.Click, btnEqual.Click
        Dim array As String()
        Dim vSender As String
        Dim vDisplay As String

        'Part I:  create the chains
        If sender.text = "." And vChain = "" Then
            vSender = "0."
        Else
            vSender = sender.text
        End If
        ListDisplay.Add(vSender) 'list for all the keystrokes
        Select Case vSender
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0.", "."
                If vOpt = True Then  'determines if a comma is inserted before the entry
                    vChain = vChain & "," & vSender
                    vOpt = False
                Else
                    vChain = vChain & vSender
                End If
            Case Else 'an operator was added
                vChain = vChain & "," & vSender
                vVirgin = False
                vOpt = True
        End Select

        'Part II:  split the chain
        array = vChain.Split(",")
        vAns = array(0)
        If array.Count > 3 Then
            For i As Integer = 3 To array.Count - 1
                Select Case array(i)
                    Case "+", "-", "x", "/", "%", "="
                        Select Case array(i - 2)
                            Case "+"
                                vAns = vAns + CDbl(array(i - 1))
                            Case "-"
                                vAns = vAns - CDbl(array(i - 1))
                            Case "x"
                                vAns = vAns * CDbl(array(i - 1))
                            Case "/"
                                vAns = vAns / CDbl(array(i - 1))
                            Case "%"
                                vAns = vAns / CDbl(array(i - 1)) 'work on this formula
                            Case "="
                                Select Case array(i - 3)
                                    Case "+"
                                        vAns = vAns + CDbl(array(i - 1))
                                    Case "-"
                                        vAns = vAns - CDbl(array(i - 1))
                                    Case "x"
                                        vAns = vAns * CDbl(array(i - 1))
                                    Case "/"
                                        vAns = vAns / CDbl(array(i - 1))
                                    Case "%"
                                        vAns = vAns / CDbl(array(i - 1)) 'work on this formula 
                                End Select
                        End Select
                        Clipboard.Clear()
                        Clipboard.SetText(tbDisplay.Text)
                End Select
            Next
        End If

        'Part III:  what to show in the Display
        Select Case vSender 'what is the current entry
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0.", "."
                vSenderType = "Digit"
            Case Else
                vSenderType = "Operator"
        End Select
        If ListDisplay.Count = 1 Then
            If vSenderType = "Digit" Then
                vDisplay = "aNothingDigit"
            Else
                vDisplay = "aNothingOperator"
            End If
        Else 'ListDisplay.count > 1
            Select Case ListDisplay(ListDisplay.Count - 2) 'what was the entry just before the current entry; -2 for zero-based list
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0.", "."
                    If vSenderType = "Digit" Then
                        vDisplay = "aDigitDigit"
                    Else
                        vDisplay = "aDigitOperator"
                    End If
                Case Else
                    If vSenderType = "Digit" Then
                        vDisplay = "aOperatorDigit"
                    Else
                        vDisplay = "aOperatorOperator"
                    End If
            End Select
        End If
        Select Case vDisplay
            Case "aNothingDigit"
                tbDisplay.Text = vSender
            Case "aNothingOperator"
                MsgBox("Start with a Digit")
            Case "aDigitDigit"
                tbDisplay.Text = InsertCommas(tbDisplay.Text & vSender)
            Case "aDigitOperator"
                tbDisplay.Text = InsertCommas(vAns.ToString)
            Case "aOperatorDigit"
                tbDisplay.Text = vSender
        End Select


    End Sub
    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click, btnClearEntry.Click
        tbDisplay.Text = "0"
        vTotal = 0
        vCalc = 0
        vVirgin = True
        vOpera = "None"
        vChain = ""
        vAns = 0
        vOpt = False
    End Sub




    Private Sub btnPercent_Click(sender As Object, e As EventArgs) Handles btnPercent.Click
        If vTotal = 0 Then
            vTotal = tbDisplay.Text
        Else
            ' vTotal = tbDisplay.Text / vTotal
        End If
        vCalc = tbDisplay.Text / vCalc
        vVirgin = True
        'vChain = vChain & "," & "%"
        ' vChain = vChain & sender.Text & ","
        If vOpera = "Percent" Then
            tbDisplay.Text = Format(tbDisplay.Text / vTotal.ToString, "#,##0.00%")
        Else
            vOpera = "Percent"
            'tbDisplay.Text = 0
        End If
    End Sub

    Private Function InsertCommas(fComma) As String
        Dim vLen As Integer
        Dim vDot As Integer
        fComma = fComma.Replace(",", "") 'remove any commas from the string
        vDot = InStr(fComma, ".")
        If vDot = 0 Then
            vLen = Len(fComma)
        Else
            vLen = vDot - 1
        End If

        Select Case vLen
            Case 4, 5, 6
                InsertCommas = fComma.insert(vLen - 3, ",")
            Case 7, 8, 9
                fComma = fComma.insert(vLen - 6, ",")
                InsertCommas = fComma.insert(vLen - 2, ",")
            Case 10, 11, 12
                fComma = fComma.insert(vLen - 6, ",")
                fComma = fComma.insert(vLen - 2, ",")
                InsertCommas = fComma.insert(vLen - 9, ",")
            Case Else
                InsertCommas = fComma
        End Select
    End Function

    Private Sub btnMemory_Click(sender As Object, e As EventArgs) Handles btnMplus.Click, btnMminus.Click, btnMrecall.Click, btnMclear.Click
        Dim s As String = DirectCast(sender, Button).Text
        Select Case s
            Case "M+"
                vMemory = vMemory + tbDisplay.Text
            Case "M-"
                vMemory = vMemory - tbDisplay.Text
            Case "MC"
                ' MsgBox("Are you Sure you want to Clear Memory?")
                vMemory = 0
                tbDisplay.Text = "0"
                vTotal = 0
                vOpera = "None"
                vVirgin = True
            Case "MR"
                tbDisplay.Text = InsertCommas(vMemory.ToString)
        End Select

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        btnDivide.Text = Chr(247)
    End Sub
End Class
