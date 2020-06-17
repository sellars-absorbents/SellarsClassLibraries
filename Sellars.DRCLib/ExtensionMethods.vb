Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection

Module ExtensionMethods
    ' *********************************************************************************************
    ' * This extension method will ensure that the passed in string is returned as a fixed length *
    ' * based on the passed in length parameter.  If the original string is too long, it will be  *
    ' * truncated, if too short, it will be padded on the end with spaces.                        *
    ' *********************************************************************************************
    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFixedLengthString(ByVal input As String, ByVal length As Integer) As String
        Dim result As String = String.Empty

        If String.IsNullOrEmpty(input) Then
            result = New String(" "c, length)
        ElseIf input.Length > length Then
            result = input.Substring(0, length)
        Else
            result = input.PadRight(length)
        End If

        Return result
    End Function

    ''' <summary>
    ''' Copies all the properties of the "from" object to this object if they exist.
    ''' </summary>
    ''' <param name="to">The object in which the properties are copied</param>
    ''' <param name="from">The object which is used as a source</param>
    ''' <param name="excludedProperties">Exclude these properties from the copy</param>
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub CopyFrom(ByVal [to] As Object, ByVal from As Object, ByVal excludedProperties As String())
        prvCopyFrom([to], from, excludedProperties)
    End Sub

    Private Sub prvCopyFrom(ByVal [to] As Object, ByVal from As Object, ByVal excludedProperties As String())
        Dim targetType As Type = [to].[GetType]()
        Dim sourceType As Type = from.[GetType]()

        ' Get all the properties for the source and to objects
        'Dim sourceProps As PropertyInfo() = sourceType.GetProperties()
        'Dim toProps As PropertyInfo() = targetType.GetProperties()
        Dim matchingProps As List(Of PropertyInfo) = GetMatchingProperties(from, [to])

        For Each propInfo As PropertyInfo In matchingProps
            'filter the properties
            If excludedProperties IsNot Nothing AndAlso excludedProperties.Contains(propInfo.Name) Then
                Continue For
            End If

            'Get the matching property from the target
            Dim toProp As PropertyInfo = If((targetType Is sourceType), propInfo, targetType.GetProperty(propInfo.Name))

            'If it exists and it's writeable
            If toProp IsNot Nothing AndAlso toProp.CanWrite Then
                'Copy the value from the source to the target
                Dim value As [Object] = propInfo.GetValue(from, Nothing)
                toProp.SetValue([to], value, Nothing)
            End If
        Next
    End Sub

    ' This function uses LINQ to get all matching properties between the two objects.
    ' It is considered a match if both the type and name of the property are identical
    ' on both objects.
    Private Function GetMatchingProperties(ByVal source As Object, ByVal target As Object) As IList(Of PropertyInfo)
        If source Is Nothing Then
            Throw New ArgumentNullException("source")
        End If

        If target Is Nothing Then
            Throw New ArgumentNullException("target")
        End If

        Dim sourceType As System.Type = source.GetType()
        Dim sourceProperties As System.Reflection.PropertyInfo() = sourceType.GetProperties()
        Dim targetType As System.Type = target.GetType()
        Dim targetProperties As System.Reflection.PropertyInfo() = targetType.GetProperties()

        Dim properties = (From s In sourceProperties _
            From t In targetProperties _
            Where s.Name = t.Name AndAlso s.PropertyType Is t.PropertyType _
            Select s).ToList()
        Return properties
    End Function

    ''' <summary>
    ''' Copies all the properties of the "from" object to this object if they exist.
    ''' </summary>
    ''' <param name="to">The object in which the properties are copied</param>
    ''' <param name="from">The object which is used as a source</param>
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub CopyFrom(ByVal [to] As Object, ByVal from As Object)
        prvCopyFrom([to], from, Nothing)
    End Sub

    ''' <summary>
    ''' Copies all the properties of this object to the "to" object
    ''' </summary>
    ''' <param name="to">The object in which the properties are copied</param>
    ''' <param name="from">The object which is used as a source</param>
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub CopyTo(ByVal from As Object, ByVal [to] As Object)
        prvCopyFrom([to], from, Nothing)
    End Sub

    ''' <summary>
    ''' Copies all the properties of this object to the "to" object
    ''' </summary>
    ''' <param name="to">The object in which the properties are copied</param>
    ''' <param name="from">The object which is used as a source</param>
    ''' <param name="excludedProperties">Exclude these properties from the copy</param>
    <System.Runtime.CompilerServices.Extension()> _
    Public Sub CopyTo(ByVal from As Object, ByVal [to] As Object, ByVal excludedProperties As String())
        prvCopyFrom([to], from, excludedProperties)
    End Sub

End Module
