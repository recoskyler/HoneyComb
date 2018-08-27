'Author: Arman Ghazanchyan
'Created date: 10/10/2006
'Updated date: 11/13/2007

''' <summary>
''' Provides methods to generate and calculate unique combinations.
''' </summary>
Public Class Combination : Implements IComparer(Of List(Of Object))

#Region " CLASS METHODS "

    ''' <summary>
    ''' Gets a nested(2D) list of unique combinations.
    ''' </summary>
    ''' <param name="items">A list(set) of objects from which 
    ''' the unique combinations of subsets should be returned.</param>
    ''' <param name="k">The subset size.</param>
    <DebuggerHidden()>
    Public Shared Function GetSubsets(ByVal items() As Object, ByVal k As Integer) As List(Of List(Of Object))

        If items Is Nothing Then
            'The items parameter is Null. Throw an exception.
            Throw New ArgumentNullException(
            "items", "The parameter can’t be Null.")
        End If
        If items.Length = 0 Then
            'The list is empty. Throw an exception.
            Throw New ArgumentException(
            "The list must contain at least one object.", "items")
        End If
        Dim i As Integer = k
        Dim n As Integer = items.Length
        If (i - 1) < 0 OrElse k > n Then
            'Invalid subset size. Throw an exception.
            Throw New ArgumentOutOfRangeException(
            "k", k, "The value must be in the range {1 to " & CStr(n) & "}.")
        End If
        Dim finalList As New List(Of List(Of Object))
        i -= 1
        'Add the objects to the first sub list.
        Dim indexs(k - 1) As Integer
        Dim firstSubList As New List(Of Object)
        For j As Integer = 0 To k - 1
            indexs(j) = j
            firstSubList.Add(items(j))
        Next
        'Add the first sub list to the parent list.
        finalList.Add(firstSubList)
        'Continue adding lasts.
        While indexs(0) <> n - k AndAlso finalList.Count < 2147483647
            If indexs(i) < i + (n - k) Then
                indexs(i) += 1
                Dim subList As New List(Of Object)
                For Each j As Integer In indexs
                    'Add the objects to the sub list.
                    subList.Add(items(j))
                Next
                'Add the sub list to the parent list.
                finalList.Add(subList)
            Else
                Do
                    i -= 1
                Loop While indexs(i) = i + (n - k)
                indexs(i) += 1
                For j As Integer = i + 1 To k - 1
                    indexs(j) = indexs(j - 1) + 1
                Next
                Dim subList As New List(Of Object)
                For Each j As Integer In indexs
                    'Add the objects to the sub list.
                    subList.Add(items(j))
                Next
                'Add the sub list to the parent list.
                finalList.Add(subList)
                i = k - 1
            End If
        End While
        Return finalList
    End Function

    ''' <summary>
    ''' Calculates the number of unique combinations. 
    ''' </summary>
    ''' <param name="n">The total number(set size) of objects.</param>
    ''' <param name="k">The subset size.</param>
    <DebuggerHidden()>
    Public Shared Function Count(ByVal n As ULong, ByVal k As ULong) As ULong
        If n < k Then Return 0
        If n = k Then Return 1
        Dim delta, iMax As ULong
        If (k < n - k) Then
            delta = n - k
            iMax = k
        Else
            delta = k
            iMax = n - k
        End If
        Dim ans As ULong = CULng(delta + 1)
        For i As ULong = 2 To iMax
            ans = CULng((ans * (delta + i)) / i)
        Next
        Return ans
    End Function

    ''' <summary>
    ''' Sorts the elements in a List(Of List(Of Object)).
    ''' </summary>
    ''' <param name="combLists">A list of combinations list to sort.</param>
    <DebuggerHidden()>
    Public Overloads Shared Sub Sort(ByVal combLists As List(Of List(Of Object)))
        If combLists.Count = 1 Then
            combLists(0).Sort()
        ElseIf combLists.Count > 1 Then
            combLists.Sort(New Combination)
        End If
    End Sub

    ''' <summary>
    ''' Compares two specified List(Of Object).
    ''' </summary>
    ''' <param name="x">The first List(Of Object).</param>
    ''' <param name="y">The second List(Of Object).</param>
    <DebuggerHidden()>
    Protected Overridable Function Compare(ByVal x As List(Of Object), ByVal y As List(Of Object)) As Integer _
    Implements System.Collections.Generic.IComparer(Of List(Of Object)).Compare
        x.Sort()
        y.Sort()
        Dim t As Integer
        'Numeric type compare
        If IsNumeric(x(0)) Then
            For i As Integer = 0 To x.Count - 1
                t = Val(x(i)).CompareTo(Val(y(i)))
                If t <> 0 Then
                    Return t
                End If
            Next i
            Return 0
        End If
        'String type compare
        If x(0).GetType Is GetType(String) Then
            For i As Integer = 0 To x.Count - 1
                t = CStr(x(i)).CompareTo(y(i))
                If t <> 0 Then
                    Return t
                End If
            Next i
            Return 0
        End If
        'Date type compare
        If x(0).GetType Is GetType(Date) Then
            For i As Integer = 0 To x.Count - 1
                t = CDate(x(i)).CompareTo(y(i))
                If t <> 0 Then
                    Return t
                End If
            Next i
            Return 0
        End If
    End Function

#End Region

End Class
