Option Strict On
Option Explicit On


' El manejo de los "components": los detecta y los salta: si encuentra un "Component", no los pone en el oDictionary.
' Aunque salte el Component, entra a mirar qué tiene dentro, si adentro hay piezas reales, las trata normalmente.

' Links rotos: si el link de una pieza está roto, lo detecta porque al intentar acceder al documento de la referencia, lanza un error.
' Tiene un bloque try-catch para detectar si el link está roto. Si lo está, avisa por consola y omite ese elemento.

' En los diccionarios no hay que usar tipo de datos "ProductStructureTypeLib.CatProductSource" porque no es serializable.
' Por eso se guarda como Integer u otros tipos nativos.



Public Class CatiaDataExtractor


    Public Function ExtractData(oRootProduct As ProductStructureTypeLib.Product,
                                folderPath As String) As DataTable


        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] - Extracting data from CATIA...")


        If Not String.IsNullOrWhiteSpace(folderPath) AndAlso Not IO.Directory.Exists(folderPath) Then
            IO.Directory.CreateDirectory(folderPath)
        End If


        Dim table As DataTable = CreateProductDataTable()


        ' Cargar fila para el root product
        AddProductRow(oRootProduct, table, 0, folderPath)

        ' Recorrer la estructura del ensamble recursivamente
        ' ProcesarHijosRecursivo(oRootProduct, table, 1, folderPath)

        Return table

    End Function



    Private Sub AddProductRow(oProduct As ProductStructureTypeLib.Product,
                             table As DataTable,
                             level As Integer,
                             folderPath As String)

        Dim rootDoc As INFITF.Document = CType(oProduct.ReferenceProduct.Parent, INFITF.Document)

        Dim row As DataRow = table.NewRow()

        Dim snapshot = TakeSnapshot(oProduct, folderPath, level)


        ' Bounding Box
        row("DimX") = snapshot.dimX
        row("DimY") = snapshot.dimY
        row("DimZ") = snapshot.dimZ


        ' Identificación
        row("PartNumber") = oProduct.PartNumber
        row("InstanceName") = oProduct.Name
        row("Revision") = oProduct.Revision
        row("Nomenclature") = oProduct.Nomenclature
        row("DescriptionRef") = oProduct.DescriptionRef
        row("Definition") = oProduct.Definition
        row("Quantity") = 1
        row("Level") = level
        row("IsAssembly") = (oProduct.Products.Count > 0)
        row("ProductType") = TypeName(rootDoc)
        row("Source") = CInt(oProduct.Source)


        ' Archivos
        row("FileName") = rootDoc.Name
        row("FullPath") = GetJustDirectory(rootDoc.FullName)
        row("ImageFilePath") = snapshot.finalFileName


        ' Físicas
        ' row("Material") = materialName
        'row("Mass_kg") = inertia.Mass
        'row("Volume_m3") = inertia.Volume
        'row("SurfaceArea_m2") = inertia.Area

        '' Posición
        'row("PosX") = pos.X
        'row("PosY") = pos.Y
        'row("PosZ") = pos.Z

        table.Rows.Add(row)

    End Sub


    Private Sub ProcesarHijosRecursivo(oProduct As ProductStructureTypeLib.Product,
                                   table As DataTable,
                                   level As Integer,
                                   folderPath As String)

        For Each oProduct In oProduct.Products

            Dim childProduct As ProductStructureTypeLib.Product = oProduct

            ' Si la pieza ya existe en la lista, acumulamos la cantidad

            Dim existingRow As DataRow = table.Rows.Find(childProduct.PartNumber)

            If existingRow IsNot Nothing Then
                existingRow("Quantity") = CInt(existingRow("Quantity")) + 1
            Else
                Dim childDoc As INFITF.Document = CType(childProduct.ReferenceProduct.Parent, INFITF.Document)
                AddProductRow(childProduct, table, level, folderPath)
            End If

            ' Si tiene subcomponentes, continuamos la recursión
            If childProduct.Products.Count > 0 Then

                ProcesarHijosRecursivo(childProduct, table, level + 1, folderPath)

            End If

        Next

    End Sub












    Private Function TakeSnapshot(oProd As ProductStructureTypeLib.Product, folder As String, level As Integer) As (finalFileName As String, dimX As Double, dimY As Double, dimZ As Double)

        Dim safePartNumber As String = CleanFileName(oProd.PartNumber)

        Dim finalFileName As String = IO.Path.Combine(folder, safePartNumber & ".jpg")

        Dim oApp As INFITF.Application = oProd.Application

        Dim docPrincipal As INFITF.Document = oApp.ActiveDocument

        Dim oNewWindow As INFITF.Window

        Dim oOriginalWindow As INFITF.Window

        ' Acá maneja el root, porque si es root hace "NewWindow()" y si no es root hace "open in new window"

        If level <> 0 Then

            Dim oSelection As INFITF.Selection = docPrincipal.Selection
            oSelection.Clear()
            oSelection.Add(oProd)
            oApp.StartCommand("Open in New Window")
            oApp.RefreshDisplay = True

            ' Verificar si la nueva ventana realmente se abrió o no
            ' Si no se abrió, significa que el componente es un "Component" o un "Reference Product" y no tiene un documento asociado.
            If oApp.ActiveDocument Is docPrincipal Then
                oSelection.Clear()
                Return (String.Empty, 0.0, 0.0, 0.0)
            End If

            ' en caso de que sí sea root, hace "NewWindow()"
        Else
            oOriginalWindow = oApp.ActiveWindow
            oNewWindow = oOriginalWindow.NewWindow()



        End If

        ' creo que esta linea es redundante, porque cuando se abre una ventana nueva,
        ' ya sea con "NewWindow()" o con "Open in New Window",
        ' la ventana activa pasa a ser la nueva ventana. Pero por las dudas, la dejo.
        ' aunque en la linea siguiente uso oCurrentWindow para referenciar el oSpecWindow.



        Dim oSpecsWindow As INFITF.SpecsAndGeomWindow = CType(oApp.ActiveWindow, INFITF.SpecsAndGeomWindow)

        Dim oViewer As INFITF.Viewer3D = CType(oSpecsWindow.Viewers.Item(1), INFITF.Viewer3D)


        Dim oldColor(2), white(2) As Object
        white(0) = 1 : white(1) = 1 : white(2) = 1
        oViewer.GetBackgroundColor(oldColor)
        oViewer.PutBackgroundColor(white)


        ' cuidado que el metodo "StartCommand" es 
        oSpecsWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly
        oApp.StartCommand("Compass")
        Threading.Thread.Sleep(1000)
        oApp.ActiveWindow.Height = 300
        oApp.ActiveWindow.Width = 300


        ' Ocultar los sistemas de ejes de las piezas
        Dim oDocActual As INFITF.Document = oApp.ActiveDocument
        Dim oSelAxes As INFITF.Selection = oDocActual.Selection
        oSelAxes.Clear()
        oSelAxes.Search("CATPrtSearch.AxisSystem.Visibility=Shown,all")
        Threading.Thread.Sleep(1000)
        If oSelAxes.Count > 0 Then
            oSelAxes.VisProperties.SetShow(INFITF.CatVisPropertyShow.catVisPropertyNoShowAttr)
            oSelAxes.Clear()
        End If


        oViewer.Viewpoint3D.ProjectionMode = INFITF.CatProjectionMode.catProjectionCylindric
        oViewer.Viewpoint3D = CType(oApp.ActiveDocument.Cameras.Item(1), INFITF.Camera3D).Viewpoint3D
        oViewer.Reframe()
        oViewer.Update()
        oApp.RefreshDisplay = True


        ' acá toma la captura.
        oViewer.CaptureToFile(INFITF.CatCaptureFormat.catCaptureFormatJPEG, finalFileName)


        oViewer.PutBackgroundColor(oldColor)
        oApp.StartCommand("Compass")
        Threading.Thread.Sleep(1000)
        oSpecsWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom



        '---------------------------------------------------------------------
        Dim oInertia As SPATypeLib.Inertia = CType(oProd.GetTechnologicalObject("Inertia"), SPATypeLib.Inertia)

        Dim bBox As Array = Array.CreateInstance(GetType(Object), 9)

        oInertia.GetPrincipalAxes(bBox)


        Dim dimX As Double = CDbl(bBox.GetValue(1)) - CDbl(bBox.GetValue(0))
        Dim dimY As Double = CDbl(bBox.GetValue(3)) - CDbl(bBox.GetValue(2))
        Dim dimZ As Double = CDbl(bBox.GetValue(5)) - CDbl(bBox.GetValue(4))


        Dim oReference As INFITF.Reference = oProd.CreateReferenceFromName("")


        Return (finalFileName, dimX, dimY, dimZ)


    End Function











    Private Function CreateProductDataTable() As DataTable

        Dim dt As New DataTable("ProductData")

        ' --- Identificación y Administración ---
        dt.Columns.Add("PartNumber", GetType(String))
        dt.Columns.Add("InstanceName", GetType(String))
        dt.Columns.Add("Revision", GetType(String))
        dt.Columns.Add("Nomenclature", GetType(String))
        dt.Columns.Add("DescriptionRef", GetType(String))
        dt.Columns.Add("Definition", GetType(String))
        dt.Columns.Add("Quantity", GetType(Integer))
        dt.Columns.Add("Level", GetType(Integer))
        dt.Columns.Add("IsAssembly", GetType(Boolean))
        dt.Columns.Add("ProductType", GetType(String))
        dt.Columns.Add("Source", GetType(Integer))

        ' --- Archivos y Rutas ---
        dt.Columns.Add("FileName", GetType(String))
        dt.Columns.Add("FullPath", GetType(String))
        dt.Columns.Add("ImageFilePath", GetType(String))

        ' --- Dimensiones de Caja Envolvente (Bounding Box) ---
        dt.Columns.Add("DimX", GetType(Double))
        dt.Columns.Add("DimY", GetType(Double))
        dt.Columns.Add("DimZ", GetType(Double))

        ' --- Propiedades Físicas e Inerciales ---
        dt.Columns.Add("Mass_kg", GetType(Double))
        dt.Columns.Add("Volume_m3", GetType(Double))
        dt.Columns.Add("SurfaceArea_m2", GetType(Double))

        ' --- Posición Espacial Absoluta/Relativa (Traslación) ---
        dt.Columns.Add("PosX", GetType(Double))
        dt.Columns.Add("PosY", GetType(Double))
        dt.Columns.Add("PosZ", GetType(Double))


        ' esta linea es para agregar la primary key a la tabla, que es el PartNumber.
        ' Esta abreviada. Ver documentacion de DataColumnCollection.Add para más detalles.

        dt.PrimaryKey = New DataColumn() {dt.Columns("PartNumber")}

        Return dt

    End Function

    Private Function CleanFileName(name As String) As String
        Dim invalidChars As New String(IO.Path.GetInvalidFileNameChars())
        Dim cleaned As String = name
        For Each c As Char In invalidChars
            cleaned = cleaned.Replace(c, "_"c)
        Next
        Return cleaned
    End Function



    Private Function GetJustDirectory(fullPath As String) As String
        If String.IsNullOrEmpty(fullPath) Then Return ""
        Dim lastSlash As Integer = Math.Max(fullPath.LastIndexOf("\"), fullPath.LastIndexOf("/"))
        If lastSlash > 0 Then
            Return fullPath.Substring(0, lastSlash)
        End If
        Return fullPath
    End Function


End Class








