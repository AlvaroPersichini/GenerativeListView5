Option Explicit On
Option Strict On

Public Class ProductData

    Public Property PartNumber As String = ""
    Public Property Nomenclature As String = ""
    Public Property DescriptionRef As String = ""
    Public Property Definition As String = ""
    Public Property FileName As String = ""
    Public Property FullPath As String = ""
    Public Property ImageFilePath As String = ""
    Public Property Quantity As Integer = 1
    Public Property Level As Integer = 0
    Public Property ProductType As String = ""
    Public Property Source As Integer = 0
    Public Property Dimensions As (X As Double, Y As Double, Z As Double)

End Class
