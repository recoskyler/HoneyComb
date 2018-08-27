' Author  : Adil Atalay Hamamcioglu
' Created : 10/02/2017
' Thanks to Burak Erinc Cetin and Arman Ghazanchyan

Imports Microsoft.Office.Interop

Public Class Form1

    Dim results As New List(Of Integer)
    Dim finalRes As New List(Of String)

#Region " HANDLERS "

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Debug.WriteLine("--------[ HoneyComb 1.0 ]--------")
        My.Settings.Reload()
        Label8.Text = My.Settings.path
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text IsNot "" Then
            If Convert.ToInt32(TextBox1.Text) >= 3 Then
                calculate(Convert.ToInt32(TextBox1.Text))
            Else
                MsgBox("Lütfen 2'den büyük bir sayı giriniz.")
            End If
        Else
            MsgBox("Lütfen 2'den büyük bir sayı giriniz.")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        populateExcel(My.Settings.path)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If My.Settings.path IsNot "NONE" Or My.Settings.path IsNot "" Then
                If IO.File.Exists(My.Settings.path) Then
                    Process.Start(My.Settings.path)
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            Label8.Text = FolderBrowserDialog1.SelectedPath & "\HoneyComb_Output.xls"

            My.Settings.path = FolderBrowserDialog1.SelectedPath & "\HoneyComb_Output.xls"
            My.Settings.Save()
            My.Settings.Reload()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

        If e.KeyChar = Chr(13) Then
            calculate(Convert.ToInt32(TextBox1.Text))
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

#End Region

#Region " CODE "

    Public Sub calculate(ByVal num As Integer)
        results.Clear()

        sumGen(filter(comb(getVals(num), 3)))

        Label3.Text = compare(results, num, 3)

        Debug.WriteLine("+++++++++++++++++++++")
        Debug.WriteLine(compare(results, num, 3))
        Debug.WriteLine("*********************")
    End Sub

    Public Function getVals(ByVal num As Integer) As Object()
        Dim arr(num) As Object

        For i As Integer = 1 To num
            arr(i - 1) = (i.ToString())
        Next

        Return arr
    End Function

    Public Function comb(ByVal arr As Object(), ByVal depth As Integer) As List(Of String)
        Dim res As New List(Of String)

        Dim mySubsets As New List(Of List(Of Object))

        mySubsets = Combination.GetSubsets(arr, depth)

        Dim str As String

        For Each i As List(Of Object) In mySubsets
            str = ""

            For Each j As Integer In i
                str &= (j & "+")
            Next j

            res.Add(str)
        Next i

        Return res
    End Function

    Public Function filter(ByVal arr As List(Of String)) As List(Of String)
        Dim res As New List(Of String)

        Debug.Print(arr.Count)

        For Each s As String In arr
            s = s.Substring(0, s.Length - 1)

            If Not s.Contains("+0") Then
                res.Add(s)
            End If
        Next

        Debug.Print(res.Count)

        res.Sort()

        Debug.Print(res.Count)

        Return res
    End Function

    Public Function C(ByVal n As Integer, ByVal r As Integer) As Integer
        Return Combination.Count(n, r)
    End Function

    Public Function factorial(ByVal n As Integer) As Long
        Dim res As ULong = 1

        If n >= 1 Then
            For i As Integer = 1 To n
                Debug.WriteLine(res)
                res = res * i
            Next
        End If

        Return res
    End Function

    Public Function isPrime(ByVal n1 As Integer, ByVal n2 As Integer) As List(Of Integer)
        Dim a1 As New List(Of Integer)
        Dim a2 As New List(Of Integer)
        Dim a3 As New List(Of Integer)

        For i As Integer = 1 To n1
            If (n1 Mod i) = 0 Then
                a1.Add(i)
            End If
        Next

        For i As Integer = 1 To n2
            If (n2 Mod i) = 0 Then
                a2.Add(i)
            End If
        Next

        a3 = a1.Intersect(a2).ToList()
        a3.Sort()

        Return a3
    End Function

    Public Sub sumGen(ByVal arr As List(Of String))
        For Each s As String In arr
            Dim temp() As String = s.Split("+")
            Dim sum As Integer = 0

            For Each n As String In temp
                If n <> "" Then
                    sum = sum + Convert.ToInt32(n)
                End If
            Next

            results.Add(sum)
        Next
    End Sub

    Public Function compare(ByVal arr As List(Of Integer), ByVal num As Integer, ByVal depth As Integer) As String
        Dim dividend As Integer = 0
        Dim divisor As String = C(num, depth)
        Dim pnum As Integer = 0
        Dim res As String = ""

        For Each s As String In arr
            If isPrime(CInt(s), num).Count = 1 Then
                dividend += 1
            End If
        Next

        res = dividend & "/" & divisor

        Return res
    End Function

    Public Sub populateExcel(ByVal path As String)
        Debug.WriteLine("Populating Excel File")
        Button3.Enabled = False
        MsgBox("Bu işlem birkaç dakika sürebilir.")

        Try
            If path = "NONE" Or path = "" Then
                path = "HoneyComb_Output.xls"

                My.Settings.path = path
                My.Settings.Save()
                My.Settings.Reload()
            End If

            If IO.File.Exists(path) Then
                Try
                    IO.File.Delete(path)
                Catch ex As Exception
                    Debug.WriteLine(ex.ToString)
                End Try
            End If

            If TextBox2.Text IsNot "" And TextBox3.Text IsNot "" Then
                Dim misValue As Object = System.Reflection.Missing.Value
                Dim oExcel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application()

                If oExcel Is Nothing Then
                    MessageBox.Show("Excel kurulu degil.")
                    Return
                End If

                Dim oBook As Excel.Workbook = oExcel.Workbooks.Add(misValue)
                Dim oSheet As Object = oBook.Worksheets(1)
                Dim row As Integer = 2

                oSheet.Range("A1").Value = "Sayı"
                oSheet.Range("B1").Value = "Sonuç"

                For i As Integer = CInt(TextBox2.Text) To CInt(TextBox3.Text)
                    results.Clear()
                    sumGen(filter(comb(getVals(i), 3)))

                    oSheet.Range(("A" & CStr(row))).Value = CStr(i)
                    oSheet.Range(("B" & CStr(row))).Value = "_" & compare(results, i, 3)

                    Dim pos(2) As Integer
                    Dim pt() As String = compare(results, i, 3).Split("/")

                    pos(0) = CInt(pt(0))
                    pos(1) = CInt(pt(1))

                    oSheet.Cells((pos(0) + 3), (pos(1) + 3)) = "*"

                    If i Mod 3 = 0 Then
                        oSheet.Cells((pos(0) + 3), (pos(1) + 3)).Interior.ColorIndex = 3
                    Else
                        oSheet.Cells((pos(0) + 3), (pos(1) + 3)).Interior.ColorIndex = 42
                    End If

                    row += 1
                Next

                results.Clear()
                sumGen(filter(comb(getVals(11), 3)))

                Debug.Print(filter(comb(getVals(11), 3)).Count)

                For row1 As Integer = 1 To C(11, 3)
                    oSheet.Range(("C" & CStr(row1))).Value = results.Item(row1 - 2)
                Next

                oBook.SaveAs(path)
                oBook.Close()
                oExcel.Quit()

                Debug.WriteLine("DONE")
                Button3.Enabled = True
                MsgBox("Excel dosyası dolduruldu.")
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
    End Sub

    Public Function isZero(ByVal n As Integer) As Boolean
        Dim res As Boolean = False

        If n = 0 Then
            res = True
        End If

        Return res
    End Function

#End Region

End Class
