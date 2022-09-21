Imports System.Collections

' This class implements a simple dictionary using an array of DictionaryEntry objects (key/value pairs).
Public Class SimpleDictionary
    Implements IDictionary

    ' The array of items
    Dim items() As DictionaryEntry

    Dim ItemsInUse As Integer = 0

    ' Construct the SimpleDictionary with the desired number of items.
    ' The number of items cannot change for the life time of this SimpleDictionary.
    Public Sub New(ByVal numItems As Integer)
        items = New DictionaryEntry(numItems - 1) {}
    End Sub

    ' IDictionary Members
    Public ReadOnly Property IsReadOnly() As Boolean Implements IDictionary.IsReadOnly
        Get
            Return False
        End Get
    End Property

    Public Function Contains(ByVal key As Object) As Boolean Implements IDictionary.Contains
        Dim index As Integer
        Return TryGetIndexOfKey(key, index)
    End Function

    Public ReadOnly Property IsFixedSize() As Boolean Implements IDictionary.IsFixedSize
        Get
            Return False
        End Get
    End Property

    Public Sub Remove(ByVal key As Object) Implements IDictionary.Remove
        If key = Nothing Then
            Throw New ArgumentNullException("key")
        End If
        ' Try to find the key in the DictionaryEntry array
        Dim index As Integer
        If TryGetIndexOfKey(key, index) Then

            ' If the key is found, slide all the items up.
            Array.Copy(items, index + 1, items, index, (ItemsInUse - index) - 1)
            ItemsInUse = ItemsInUse - 1
        Else

            ' If the key is not in the dictionary, just return.
        End If
    End Sub

    Public Sub Clear() Implements IDictionary.Clear
        ItemsInUse = 0
    End Sub

    Public Sub Add(ByVal key As Object, ByVal value As Object) Implements IDictionary.Add

        ' Add the new key/value pair even if this key already exists in the dictionary.
        If ItemsInUse = items.Length Then
            Throw New InvalidOperationException("The dictionary cannot hold any more items.")
        End If
        items(ItemsInUse) = New DictionaryEntry(key, value)
        ItemsInUse = ItemsInUse + 1
    End Sub

    Public ReadOnly Property Keys() As ICollection Implements IDictionary.Keys
        Get

            ' Return an array where each item is a key.
            ' Note: Declaring keyArray() to have a size of ItemsInUse - 1
            '       ensures that the array is properly sized, in VB.NET
            '       declaring an array of size N creates an array with
            '       0 through N elements, including N, as opposed to N - 1
            '       which is the default behavior in C# and C++.
            Dim keyArray() As Object = New Object(ItemsInUse - 1) {}
            Dim n As Integer
            For n = 0 To ItemsInUse - 1
                keyArray(n) = items(n).Key
            Next n

            Return keyArray
        End Get
    End Property

    Public ReadOnly Property Values() As ICollection Implements IDictionary.Values
        Get
            ' Return an array where each item is a value.
            Dim valueArray() As Object = New Object(ItemsInUse - 1) {}
            Dim n As Integer
            For n = 0 To ItemsInUse - 1
                valueArray(n) = items(n).Value
            Next n

            Return valueArray
        End Get
    End Property

    Public Property Item(ByVal key As Object) As Object Implements IDictionary.Item
        Get

            ' If this key is in the dictionary, return its value.
            Dim index As Integer
            If TryGetIndexOfKey(key, index) Then

                ' The key was found return its value.
                Return items(index).Value
            Else

                ' The key was not found return null.
                Return Nothing
            End If
        End Get

        Set(ByVal value As Object)
            ' If this key is in the dictionary, change its value.
            Dim index As Integer
            If TryGetIndexOfKey(key, index) Then

                ' The key was found change its value.
                items(index).Value = value
            Else

                ' This key is not in the dictionary add this key/value pair.
                Add(key, value)
            End If
        End Set
    End Property

    Private Function TryGetIndexOfKey(ByVal key As Object, ByRef index As Integer) As Boolean
        For index = 0 To ItemsInUse - 1
            ' If the key is found, return true (the index is also returned).
            If items(index).Key.Equals(key) Then
                Return True
            End If
        Next index

        ' Key not found, return false (index should be ignored by the caller).
        Return False
    End Function

    Private Class SimpleDictionaryEnumerator
        Implements IDictionaryEnumerator

        ' A copy of the SimpleDictionary object's key/value pairs.
        Dim items() As DictionaryEntry

        Dim index As Integer = -1

        Public Sub New(ByVal sd As SimpleDictionary)
            ' Make a copy of the dictionary entries currently in the SimpleDictionary object.
            items = New DictionaryEntry(sd.Count - 1) {}
            Array.Copy(sd.items, 0, items, 0, sd.Count)
        End Sub

        ' Return the current item.
        Public ReadOnly Property Current() As Object Implements IDictionaryEnumerator.Current
            Get
                ValidateIndex()
                Return items(index)
            End Get
        End Property

        ' Return the current dictionary entry.
        Public ReadOnly Property Entry() As DictionaryEntry Implements IDictionaryEnumerator.Entry
            Get
                Return Current
            End Get
        End Property

        ' Return the key of the current item.
        Public ReadOnly Property Key() As Object Implements IDictionaryEnumerator.Key
            Get
                ValidateIndex()
                Return items(index).Key
            End Get
        End Property

        ' Return the value of the current item.
        Public ReadOnly Property Value() As Object Implements IDictionaryEnumerator.Value
            Get
                ValidateIndex()
                Return items(index).Value
            End Get
        End Property

        ' Advance to the next item.
        Public Function MoveNext() As Boolean Implements IDictionaryEnumerator.MoveNext
            If index < items.Length - 1 Then
                index = index + 1
                Return True
            End If

            Return False
        End Function

        ' Validate the enumeration index and throw an exception if the index is out of range.
        Private Sub ValidateIndex()
            If index < 0 Or index >= items.Length Then
                Throw New InvalidOperationException("Enumerator is before or after the collection.")
            End If
        End Sub

        ' Reset the index to restart the enumeration.
        Public Sub Reset() Implements IDictionaryEnumerator.Reset
            index = -1
        End Sub

    End Class

    Public Function GetEnumerator() As IDictionaryEnumerator Implements IDictionary.GetEnumerator

        'Construct and return an enumerator.
        Return New SimpleDictionaryEnumerator(Me)
    End Function

    ' ICollection Members
    Public ReadOnly Property IsSynchronized() As Boolean Implements IDictionary.IsSynchronized
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property SyncRoot() As Object Implements IDictionary.SyncRoot
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property Count() As Integer Implements IDictionary.Count
        Get
            Return ItemsInUse
        End Get
    End Property

    Public Sub CopyTo(ByVal array As Array, ByVal index As Integer) Implements IDictionary.CopyTo
        Throw New NotImplementedException()
    End Sub

    ' IEnumerable Members
    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator

        ' Construct and return an enumerator.
        Return Me.GetEnumerator()
    End Function

End Class