Option Explicit On
Option Strict On
Module Program

    Sub Main()

        Console.WriteLine("> Starting Process...")

        ' Catia
        Dim CATIAsession As New CatiaSession
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False


        ' Directorios y nombres
        Dim baseDir As String = "C:\Temp"
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim folderPath As String = IO.Path.Combine(baseDir, "Export_" & timestamp)
        Dim excelFileName As String = IO.Path.Combine(folderPath, "Reporte_" & timestamp & ".xlsx")
        If Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If



        ' Extracción CATIA
        Dim oCatiaDataextractor As New CatiaDataExtractor
        Dim oDataTable As DataTable = oCatiaDataextractor.ExtractData(oProduct, folderPath)


        ' Imprimir el resumen con columnas alineadas
        Console.WriteLine(Environment.NewLine & "--- Listado de Productos, Niveles y Cantidades ---")

        ' Cabecera de la tabla
        Console.WriteLine($"{"PartNumber",-30} | {"Nivel",-6} | {"Cantidad",-8}")
        Console.WriteLine(New String("-"c, 52))

        ' Filas de datos
        For Each row As DataRow In oDataTable.Rows
            Dim partNumber As String = row("PartNumber").ToString()
            Dim level As Integer = CInt(row("Level"))
            Dim quantity As Integer = CInt(row("Quantity"))

            Console.WriteLine($"{partNumber,-30} | {level,-6} | {quantity,-8}")
        Next



    End Sub

End Module
