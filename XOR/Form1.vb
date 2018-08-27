Public Class Form1

    ' Created in 2014 by Marwix.

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox2.Text = XOREncryption("Password", TextBox1.Text)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox3.Text = XORDecryption("Password", TextBox4.Text)
    End Sub

    Public Function XORDecryption(ByVal CodeKey As String, ByVal DataIn As String) As String

        Dim lonDataPtr As Long
        Dim strDataOut As String
        Dim intXOrValue1 As Integer
        Dim intXOrValue2 As Integer

        For lonDataPtr = 1 To (Len(DataIn) / 2)
            intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
            intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
            strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
        Next lonDataPtr

        XORDecryption = strDataOut
    End Function

    Public Function XOREncryption(ByVal CodeKey As String, ByVal DataIn As String) As String

        Dim lonDataPtr As Long
        Dim strDataOut As String
        Dim temp As Integer
        Dim tempstring As String
        Dim intXOrValue1 As Integer
        Dim intXOrValue2 As Integer

        For lonDataPtr = 1 To Len(DataIn)
            intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
            intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
            temp = (intXOrValue1 Xor intXOrValue2)
            tempstring = Hex(temp)
            If Len(tempstring) = 1 Then tempstring = "0" & tempstring
            strDataOut = strDataOut + tempstring
        Next lonDataPtr
        XOREncryption = strDataOut

    End Function
End Class
