Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Linq
Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Imports System.Runtime.InteropServices
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Table = DocumentFormat.OpenXml.Wordprocessing.Table
Imports Run = DocumentFormat.OpenXml.Wordprocessing.Run
Imports Text = DocumentFormat.OpenXml.Wordprocessing.Text

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button1.Click

        Dim PlantillasDir As String = txtCarpetaPlantillas.Text 'Directorio de las plantillas
        Dim ProductotDir As String = txtCarpetaProducto.Text 'Directorio de destino de los formularios para un solo producto/lote
        Dim RegistrosDir As String = txtCarpetaRegistros.Text 'Directorio de destino de los formularios de registro
        Dim PlantillaPath As String 'Path de la plantilla
        Dim ArchivoPath As String 'Path del archivo de destino
        Dim FechaLib As String = txtFechaLib.Text 'Fecha de liberación del lote
        Dim Producto As String = txtProducto.Text 'Nombre del producto
        Dim NumFab As String = txtNumFab.Text
        Dim Cantidad As String = txtCantidad.Text
        Dim Lote As String = txtLote.Text
        Dim FechaCad As String = txtFechaCad.Text
        Dim Mes As String = txtMes.Text
        Dim FechaEnv As String = txtFechaEnv.Text
        Dim NumEnv1 As String = txtNumEnv1.Text
        Dim NumEnv2 As String = txtNumEnv2.Text
        Dim TipoEnv1 As String = txtTipoEnv1.Text
        Dim TipoEnv2 As String = txtTipoEnv2.Text
        Dim CodEnv1 As String = txtCodEnv1.Text
        Dim CodEnv2 As String = txtCodEnv2.Text
        Dim CodTapa1 As String = txtCodTapa1.Text
        Dim CodTapa2 As String = txtCodTapa2.Text
        Dim TipoEnvasado As String = comboTipoEnv.Text
        Dim FechaFab As String = txtFechaFab.Text
        Dim YearFab As String = txtYearFab.Text
        Dim RelEq As String = txtRelEq.Text
        Dim RefEnv1 As String = txtRefEnv1.Text
        Dim RefEnv2 As String = txtRefEnv2.Text
        Dim Observaciones As String = comboObs.Text
        Dim result As DialogResult = MessageBox.Show("¿Estás seguro?", "Confirmar", MessageBoxButtons.YesNo)
        Dim i, nrows As Integer
        If checkConsole.Checked Then
            Win32.AllocConsole()
        Else
            Win32.FreeConsole()
        End If
        If result = DialogResult.Yes Then
            'FICHA DE LIBERACIÓN DE LOTE
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Liberacion_Lote.docx"
            ArchivoPath = ProductotDir + "\" + "1-" + FechaFab + "_" + Producto + "_Registro_Liberacion_Lote.docx"
            My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
            SearchAndReplace(ArchivoPath, "LIBERACIÓN", "LIBERACIÓN: " + FechaLib)
            SearchAndReplace(ArchivoPath, "PRODUCTO:", "PRODUCTO: " + Producto)
            SearchAndReplace(ArchivoPath, "Fabricación:", "Fabricación: " + NumFab)
            SearchAndReplace(ArchivoPath, "CANTIDAD:", "CANTIDAD: " + Cantidad)
            SearchAndReplace(ArchivoPath, "LOTE:", "LOTE: " + Lote)
            If Strings.StrComp(Observaciones, "Aceite") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
parte realizado por Bindu 2013 (Caracteres organolépticos) y otra parte por Laboratorios DR GOYA (Índice de acidez e Índice de yodo).")
            ElseIf Strings.StrComp(Observaciones, "Crema") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
realizado por Bindu 2013.")
            ElseIf Strings.StrComp(Observaciones, "Solución acuosa") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
realizado por Bindu 2013.")
            ElseIf Strings.StrComp(Observaciones, "Sólido") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
parte realizado por Bindu 2013 (Caracteres organolépticos) y otra parte por Laboratorios DR GOYA (pH).")
            End If
            SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: ")
            'FICHA DE ENVASADO
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Envasado.docx"
            ArchivoPath = ProductotDir + "\" + "2-" + FechaFab + "_" + Producto + "_Registro_Envasado.docx"
            My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)

            AddTextInCell(ArchivoPath, FechaEnv, 0, 0)
            ChangeTextInCell(ArchivoPath, Producto + Environment.NewLine + "Lote:" + Lote, 1, 1)
            If NumEnv2 Is "" Or NumEnv2 Is " " Then
                ChangeTextInCell(ArchivoPath, TipoEnv1 + ": " + CodEnv1, 2, 1)
                ChangeTextInCell(ArchivoPath, TipoEnv1 + ": " + CodTapa1, 3, 1)
                ChangeTextInCell(ArchivoPath, NumEnv1 + "x" + TipoEnv1, 4, 1)
            Else
                ChangeTextInCell(ArchivoPath, TipoEnv1 + ": " + CodEnv1 + "                          " + TipoEnv2 + ": " + CodEnv2, 2, 1)
                ChangeTextInCell(ArchivoPath, TipoEnv1 + ": " + CodTapa1 + "                          " + TipoEnv2 + ": " + CodTapa2, 3, 1)
                ChangeTextInCell(ArchivoPath, NumEnv1 + "x" + TipoEnv1 + "                           " + NumEnv2 + "x" + TipoEnv2, 4, 1)
            End If
            'FICHA DE PRODUCTO PARA ANÁLISIS
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Producto_Analisis.docx"
            ArchivoPath = ProductotDir + "\" + "3-" + FechaFab + "_" + Producto + "_Registro_Producto_Analisis.docx"
            My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
            SearchAndReplace(ArchivoPath, "MES:", "MES: " + Mes)
            ChangeTextInCell(ArchivoPath, NumFab, 1, 1)
            ChangeTextInCell(ArchivoPath, Producto, 1, 3)
            ChangeTextInCell(ArchivoPath, Lote, 2, 1)
            ChangeTextInCell(ArchivoPath, FechaFab, 5, 1)
            ChangeTextInCell(ArchivoPath, Cantidad, 5, 3)
            If NumEnv2 Is "" Or NumEnv2 Is " " Then
                ChangeTextInCell(ArchivoPath, NumEnv1 + "x" + TipoEnv1, 5, 5)
            Else
                ChangeTextInCell(ArchivoPath, NumEnv1 + "x" + TipoEnv1 + "                               " + NumEnv2 + "x" + TipoEnv2, 5, 5)
            End If

            ChangeTextInCell(ArchivoPath, FechaCad, 6, 1)
            AddTextInCell(ArchivoPath, vbLf + FechaLib + vbLf, 7, 1)
            If Strings.StrComp(Observaciones, "Aceite") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
parte realizado por Bindu 2013 (Caracteres organolépticos) y otra parte por Laboratorios DR GOYA (Índice de acidez e Índice de yodo).")
            ElseIf Strings.StrComp(Observaciones, "Crema") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
realizado por Bindu 2013.")
            ElseIf Strings.StrComp(Observaciones, "Solución acuosa") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
realizado por Bindu 2013.")
            ElseIf Strings.StrComp(Observaciones, "Sólido") = 0 Then
                SearchAndReplace(ArchivoPath, "Observaciones:", "Observaciones: Análisis microbiológico realizado por Laboratorios DR GOYA y análisis físico-químico
parte realizado por Bindu 2013 (Caracteres organolépticos) y otra parte por Laboratorios DR GOYA (pH).")
            End If

            'AÑADIR A REGISTRO DE ENVASADO MENSUAL
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Envasado_Mensual.docx"
            ArchivoPath = RegistrosDir + "\" + Mes + "_Registro_Envasado_Mensual.docx"
            If My.Computer.FileSystem.FileExists(ArchivoPath) = False Then
                Console.WriteLine(ArchivoPath + " does not exist, create new one")
                My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
                AddTextInCell(ArchivoPath, Mes, 0, 0)
            Else
                Console.WriteLine(ArchivoPath + " already exists, do not create")
            End If

            'OBTENER LA SIGUIENTE FILA VACÍA, SI NO HAY, INSERTAR UNA NUEVA
            i = 0
            nrows = NumRows(ArchivoPath)
            Console.WriteLine("number of rows:")
            Console.WriteLine(nrows)
            Do While RowIsEmpty(ArchivoPath, i) = False
                i = i + 1
                If i = nrows - 1 Then
                    'INSERTA UNA NUEVA FILA Y BORRA EL CONTENIDO
                    insertRow(ArchivoPath, nrows - 3, nrows - 2)
                    Console.WriteLine("row inserted, nrows:")
                    nrows = NumRows(ArchivoPath)
                    Console.WriteLine(nrows)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 0)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 1)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 2)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 3)
                End If
                Console.WriteLine(i)
            Loop

            'RELLENAR CON LOS DATOS DEL PRODUCTO
            ChangeTextInCell(ArchivoPath, FechaEnv, i, 0)
            ChangeTextInCell(ArchivoPath, Lote, i, 1)
            ChangeTextInCell(ArchivoPath, Producto, i, 2)
            If NumEnv2 Is "" Or NumEnv2 Is " " Then
                ChangeTextInCell(ArchivoPath, NumEnv1 + "x" + TipoEnv1, i, 3)
            Else
                ChangeTextInCell(ArchivoPath, NumEnv1 + "x" + TipoEnv1 + "                           " + NumEnv2 + "x" + TipoEnv2, i, 3)
            End If
            'AÑADIR A REGISTRO DE MUESTROTECA
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Muestroteca.docx"
            ArchivoPath = RegistrosDir + "\Registro_Muestroteca.docx"
            If My.Computer.FileSystem.FileExists(ArchivoPath) = False Then
                Console.WriteLine(ArchivoPath + " does not exist, create new one")
                My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
            Else
                Console.WriteLine(ArchivoPath + " already exists, do not create")
            End If
            'OBTENER LA SIGUIENTE FILA VACÍA, SI NO HAY, INSERTAR UNA NUEVA
            i = 0
            nrows = NumRows(ArchivoPath)
            Console.WriteLine("number of rows:")
            Console.WriteLine(nrows)
            Do While RowIsEmpty(ArchivoPath, i) = False
                i = i + 1
                If i = nrows - 1 Then
                    'INSERTA UNA NUEVA FILA Y BORRA EL CONTENIDO
                    insertRow(ArchivoPath, nrows - 2, nrows - 1)
                    Console.WriteLine("row inserted, nrows:")
                    nrows = NumRows(ArchivoPath)
                    Console.WriteLine(nrows)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 0)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 1)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 2)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 3)
                End If
                Console.WriteLine(i)
            Loop
            ChangeTextInCell(ArchivoPath, FechaFab, i, 0)
            ChangeTextInCell(ArchivoPath, Lote, i, 1)
            ChangeTextInCell(ArchivoPath, Producto, i, 2)
            ChangeTextInCell(ArchivoPath, FechaLib, i, 3)

            'AÑADIR AL REGISTRO DE TOMA DE MUESTRAS
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_PN_L_M_1_Toma_Muestras.docx"
            ArchivoPath = RegistrosDir + "\" + YearFab + "_Registro_PN_L_M_1_Toma_Muestras.docx"
            If My.Computer.FileSystem.FileExists(ArchivoPath) = False Then
                Console.WriteLine(ArchivoPath + " does not exist, create new one")
                My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
            Else
                Console.WriteLine(ArchivoPath + " already exists, do not create")
            End If

            'OBTENER LA SIGUIENTE FILA VACÍA, SI NO HAY, INSERTAR UNA NUEVA
            i = 0
            nrows = NumRows(ArchivoPath)
            Console.WriteLine("number of rows:")
            Console.WriteLine(nrows)
            Do While RowIsEmpty(ArchivoPath, i) = False
                i = i + 1
                If i = nrows Then
                    'INSERTA UNA NUEVA FILA Y BORRA EL CONTENIDO
                    insertRow(ArchivoPath, nrows - 2, nrows - 1)
                    Console.WriteLine("row inserted, nrows:")
                    nrows = NumRows(ArchivoPath)
                    Console.WriteLine(nrows)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 0)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 1)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 2)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 3)
                End If
                Console.WriteLine(i)
            Loop
            ChangeTextInCell(ArchivoPath, RefEnv1 + "                                " + RefEnv2, i, 0)
            ChangeTextInCell(ArchivoPath, Lote, i, 1)
            If Strings.StrComp(TipoEnvasado, "Automático") = 0 Then
                ChangeTextInCell(ArchivoPath, "Toma de muestra en el proceso de envasado, de la envasadora", i, 2)
            ElseIf Strings.StrComp(TipoEnvasado, "Manual") = 0 Then
                ChangeTextInCell(ArchivoPath, "Toma de muestra en el proceso de envasado, de la jarra volumétrica", i, 2)
            End If
            ChangeTextInCell(ArchivoPath, FechaFab, i, 3)


            'AÑADIR AL REGISTRO DE RELACIÓN DE EQUIPOS

            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Revision_Equipos.docx"
            ArchivoPath = RegistrosDir + "\" + Mes + "_Registro_Revision_Equipos.docx"
            If My.Computer.FileSystem.FileExists(ArchivoPath) = False Then
                Console.WriteLine(ArchivoPath + " does not exist, create new one")
                My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
            Else
                Console.WriteLine(ArchivoPath + " already exists, do not create")
            End If

            'OBTENER LA SIGUIENTE FILA VACÍA, SI NO HAY, INSERTAR UNA NUEVA
            i = 0
            nrows = NumRows(ArchivoPath)
            Console.WriteLine("number of rows:")
            Console.WriteLine(nrows)
            Do While RowIsEmpty(ArchivoPath, i) = False
                i = i + 1
                If i = nrows Then
                    'INSERTA UNA NUEVA FILA Y BORRA EL CONTENIDO
                    insertRow(ArchivoPath, nrows - 2, nrows - 1)
                    Console.WriteLine("row inserted, nrows:")
                    nrows = NumRows(ArchivoPath)
                    Console.WriteLine(nrows)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 0)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 1)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 2)
                    ChangeTextInCell(ArchivoPath, "", nrows - 1, 3)
                End If
                Console.WriteLine(i)
            Loop
            ChangeTextInCell(ArchivoPath, FechaFab, i, 0)
            ChangeTextInCell(ArchivoPath, NumFab, i, 1)
            ChangeTextInCell(ArchivoPath, RelEq, i, 2)
            ChangeTextInCell(ArchivoPath, "CONFORME", i, 3)

            'AÑADIR A REGISTRO ANUAL DE LIBERACIÓN DE LOTES
            PlantillaPath = PlantillasDir + "\Plantilla_Registro_Anual_Liberacion_de_Lotes.docx"
            ArchivoPath = RegistrosDir + "\" + YearFab + "_Registro_Anual_Liberacion_de_Lotes.docx"
            If My.Computer.FileSystem.FileExists(ArchivoPath) = False Then
                Console.WriteLine(ArchivoPath + " does not exist, create new one")
                My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
                AddTextInCell(ArchivoPath, YearFab, 0, 0)
            Else
                Console.WriteLine(ArchivoPath + " already exists, do not create")
            End If

            'OBTENER LA SIGUIENTE FILA VACÍA, SI NO HAY, INSERTAR UNA NUEVA
            i = 0
            nrows = NumRows(ArchivoPath)
            Console.WriteLine("number of rows:")
            Console.WriteLine(nrows)
            Do While RowIsEmpty(ArchivoPath, i) = False
                i = i + 1
                If i = nrows - 1 Then
                    'INSERTA UNA NUEVA FILA Y BORRA EL CONTENIDO
                    insertRow(ArchivoPath, nrows - 3, nrows - 2)
                    Console.WriteLine("row inserted, nrows:")
                    nrows = NumRows(ArchivoPath)
                    Console.WriteLine(nrows)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 0)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 1)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 2)
                    ChangeTextInCell(ArchivoPath, "", nrows - 2, 3)
                End If
                Console.WriteLine(i)
            Loop

            'RELLENAR CON LOS DATOS DEL PRODUCTO
            ChangeTextInCell(ArchivoPath, FechaFab, i, 0)
            ChangeTextInCell(ArchivoPath, Lote, i, 1)
            ChangeTextInCell(ArchivoPath, Producto, i, 2)
            ChangeTextInCell(ArchivoPath, FechaLib, i, 3)

            'AÑADIR AL REGISTRO DE LOTES ANUALES
            PlantillaPath = PlantillasDir + "\Registro_Asignacion_Lotes_Anuales.xlsx"
            ArchivoPath = RegistrosDir + "\" + YearFab + "_Registro_Asignacion_Lotes_Anuales.xlsx"
            If My.Computer.FileSystem.FileExists(ArchivoPath) = False Then
                Console.WriteLine(ArchivoPath + " does not exist, create new one")
                My.Computer.FileSystem.CopyFile(PlantillaPath, ArchivoPath, overwrite:=True)
            Else
                Console.WriteLine(ArchivoPath + " already exists, do not create")
            End If
            i = 6
            Do While RowIsEmptyExcel(ArchivoPath, i) = False
                i = i + 2
            Loop
            InsertTextExcel(ArchivoPath, FechaFab, "A" + i.ToString)
            InsertTextExcel(ArchivoPath, NumFab, "B" + i.ToString)
            InsertTextExcel(ArchivoPath, Producto, "C" + i.ToString)
            InsertTextExcel(ArchivoPath, TipoEnv1, "D" + i.ToString)
            InsertTextExcel(ArchivoPath, TipoEnv2, "D" + (i + 1).ToString)
            InsertTextExcel(ArchivoPath, NumEnv1, "E" + i.ToString)
            InsertTextExcel(ArchivoPath, NumEnv2, "E" + (i + 1).ToString)
            InsertTextExcel(ArchivoPath, Lote, "F" + i.ToString)
            MessageBox.Show("Documentos rellenados!")
        Else
            MessageBox.Show("Calcelado!")
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button2.Click
        If (fbd1.ShowDialog() = DialogResult.OK) Then
            txtCarpetaPlantillas.Text = fbd1.SelectedPath
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button3.Click
        If (fbd2.ShowDialog() = DialogResult.OK) Then
            txtCarpetaProducto.Text = fbd2.SelectedPath
        End If
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object,
      ByVal e As System.EventArgs) Handles Button4.Click
        If (fbd3.ShowDialog() = DialogResult.OK) Then
            txtCarpetaRegistros.Text = fbd3.SelectedPath
        End If
    End Sub
    Public Sub SearchAndReplace(ByVal document As String, oldString As String, newString As String)
        'BUSCA Y REEMPLAZA UNA CADENA DE CARACTERES EN EL DOCUMENTO WORD
        Dim wordDoc As WordprocessingDocument = WordprocessingDocument.Open(document, True)
        Using (wordDoc)
            Dim docText As String = Nothing
            Dim sr As StreamReader = New StreamReader(wordDoc.MainDocumentPart.GetStream)

            Using (sr)
                docText = sr.ReadToEnd
            End Using

            Dim regexText As Regex = New Regex(oldString)
            docText = regexText.Replace(docText, newString)
            Dim sw As StreamWriter = New StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create))

            Using (sw)
                sw.Write(docText)
            End Using
        End Using
    End Sub
    Public Function AddressCell(worksheetpart As WorksheetPart, addressName As String) As Cell
        Dim cell As Cell = worksheetpart.Worksheet.Descendants(Of Cell).Where(Function(c) c.CellReference = addressName).FirstOrDefault
        Return cell
    End Function
    Public Function GetRow(worksheet As Worksheet, rowindex As Integer) As Row
        Dim row As Row = worksheet.GetFirstChild(Of SheetData)().Elements(Of Row)().ElementAt(rowindex)
        If row Is Nothing Then
            Throw New ArgumentException(String.Format("No row with index {0} found in spreadsheet", rowindex))
        End If
        Return row
    End Function
    Public Sub ChangeTextInCell(ByVal filepath As String, ByVal txt As String, ByVal iRow As Integer, ByVal iColumn As Integer)
        ' ACCEDE A UNA CELDA EN UNA TABLA DEL DOCUMENTO WORD Y CAMBIA EL CONTENIDO
        ' iRow fila, iColumn columna. Busca la primera tabla en el archivo
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().FirstOrDefault()

            ' Find the selected row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(iRow)

            ' Find the selected cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(iColumn)

            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().FirstOrDefault()

            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().FirstOrDefault()
            If r Is Nothing Then
                r = p.AppendChild(New Run())
            End If
            ' Set the text for the run.
            Dim t As Text = r.Elements(Of Text)().FirstOrDefault()
            If t Is Nothing Then
                t = r.AppendChild(New Text())
            End If
            t.Text = txt
        End Using
    End Sub
    Public Sub AddTextInCell(ByVal filepath As String, ByVal txt As String, ByVal iRow As Integer, ByVal iColumn As Integer)
        ' ACCEDE A UNA CELDA EN UNA TABLA DEL DOCUMENTO WORD Y CAMBIA EL CONTENIDO
        ' iRow fila, iColumn columna. Busca la primera tabla en el archivo
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().FirstOrDefault()

            ' Find the selected row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(iRow)

            ' Find the selected cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(iColumn)

            ' Find the first paragraph in the table cell.
            Dim p As Paragraph = cell.Elements(Of Paragraph)().LastOrDefault()

            ' Find the first run in the paragraph.
            Dim r As Run = p.Elements(Of Run)().LastOrDefault()
            If r Is Nothing Then
                r = p.AppendChild(New Run())
            End If
            ' Add text for the run.
            Dim t As Text = r.AppendChild(New Text())
            t.Text = txt
        End Using
    End Sub
    Public Function NumRows(ByVal filepath As String) As Integer
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().FirstOrDefault()

            NumRows = table.Descendants(Of TableRow)().ToList().Count()
            Return NumRows
        End Using
    End Function
    Public Sub insertRow(ByVal filepath As String, beforeRow As Integer, nextRow As Integer)
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().FirstOrDefault()
            Dim trb As TableRow = table.Elements(Of TableRow)().ElementAt(beforeRow)
            Dim trn As TableRow = table.Elements(Of TableRow)().ElementAt(nextRow)
            Console.WriteLine("---INSERT NEW ROW---")
            Console.WriteLine(beforeRow)
            Console.WriteLine(trb.InnerText)
            Console.WriteLine(nextRow)
            Console.WriteLine(trn.InnerText)
            trn.InsertBeforeSelf(trn.CloneNode(True))
        End Using
    End Sub
    Public Function RowIsEmpty(ByVal filepath As String, ByVal iRow As Integer) As Boolean
        ' ACCEDE A UNA FILA EN UNA TABLA DEL DOCUMENTO WORD Y DEVUELVE TRUE SI LA PRIMERA CELDA ESTÁ VACÍA
        ' iRow fila
        ' Use the file name and path passed in as an argument to 
        ' Open an existing document. 
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(filepath, True)
            ' Find the first table in the document.
            Dim table As Table = doc.MainDocumentPart.Document.Body.Elements(Of Table)().FirstOrDefault()

            ' Find the selected row in the table.
            Dim row As TableRow = table.Elements(Of TableRow)().ElementAt(iRow)

            ' Find the selected cell in the row.
            Dim cell As TableCell = row.Elements(Of TableCell)().ElementAt(0)
            If Len(cell.InnerText) = 0 Then
                Console.WriteLine("cell is empty")
                Return True
            Else
                Console.WriteLine("cell is not empty")
                Return False
            End If
        End Using
    End Function
    Public Function RowIsEmptyExcel(ByVal filepath As String, ByVal iRow As Integer) As Boolean
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(filepath, True)
            Dim WorkbookPart As WorkbookPart = spreadSheet.WorkbookPart
            Dim Sheet As Sheet = WorkbookPart.Workbook.Descendants(Of Sheet).FirstOrDefault()
            Dim WorksheetPart As WorksheetPart = WorkbookPart.GetPartById(Sheet.Id.Value)
            Dim SheetData As SheetData = WorksheetPart.Worksheet.GetFirstChild(Of SheetData)
            Dim addressName As String = "A" + iRow.ToString
            Dim cell As Cell = AddressCell(WorksheetPart, addressName)
            If cell.CellValue Is Nothing Then
                Return True
            Else
                If Len(cell.CellValue.Text) = 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Using
    End Function
    Public Sub InsertTextExcel(ByVal docName As String, ByVal text As String, addressName As String)
        ' Open the document for editing.
        Using spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
            Dim WorkbookPart As WorkbookPart = spreadSheet.WorkbookPart
            Dim Sheet As Sheet = WorkbookPart.Workbook.Descendants(Of Sheet).FirstOrDefault()
            Dim WorksheetPart As WorksheetPart = WorkbookPart.GetPartById(Sheet.Id.Value)
            Dim SheetData As SheetData = WorksheetPart.Worksheet.GetFirstChild(Of SheetData)
            Dim cell As Cell = AddressCell(WorksheetPart, addressName)
            If cell Is Nothing Then
                Console.WriteLine("cell is nothing")
            Else
                If cell.CellValue Is Nothing Then
                    cell.CellValue = New CellValue
                End If
                cell.CellValue.Text = text
            End If
        End Using
    End Sub
    Private Sub Form1_FormClosing(
    ByVal sender As System.Object,
    ByVal e As System.Windows.Forms.FormClosingEventArgs) _
    Handles MyBase.FormClosing

        Dim message As String =
                "¿Seguro que deseas salir de la aplicación?"
        Dim caption As String = "Salir"
        Dim result = MessageBox.Show(message, caption,
                                     MessageBoxButtons.YesNo,
                                     MessageBoxIcon.Question)

        ' If the no button was pressed ...
        If (result = DialogResult.No) Then
            ' cancel the closure of the form.
            e.Cancel = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim ctrl As Windows.Forms.Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                If ctrl IsNot txtCarpetaPlantillas And ctrl IsNot txtCarpetaProducto And ctrl IsNot txtCarpetaRegistros Then
                    ctrl.Text = String.Empty
                End If
            End If
        Next
    End Sub

End Class

Public Class Win32
    <DllImport("kernel32.dll")> Public Shared Function AllocConsole() As Boolean

    End Function
    <DllImport("kernel32.dll")> Public Shared Function FreeConsole() As Boolean

    End Function

End Class