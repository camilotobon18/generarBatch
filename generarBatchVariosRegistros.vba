Sub GenerarBatch()
Dim numEncabezado As Integer
Dim transaccion As String
Dim fechaDocumento As String
Dim sociedad As String
Dim fechaContabilizacion As String
Dim mes As String
Dim moneda As String
Dim referencia As String
Dim txtEncabezado As String
Const division As Integer = 3206
Const finLinea As String = "/"
Const norelsal As String = "NORELSAL"
Const rp As String = "RP"
Const bbseg As String = "BBSEG"
Const primerColumna As Integer = 2
Dim largoArray As Integer
Dim registroGasto As Integer
Dim registroProvision As Integer
Dim claseMovGasto As Integer
Dim claseMovProvision As Integer
Dim indicadorIva As String
Dim ivaTele As String
Dim ivaTelc As String
Dim ivaTela As String
Dim cuentaProvision As String
Dim cuentaTele As String
Dim cuentaTelc As String
Dim cuentaTela As String
Dim cuentaTelpUSD As String
Dim cuentaTelpPEN As String
Dim claseDocumento As String
Dim ubicacionArchivo As String
Dim nombreArchivo As String
Dim ivaTelp As String

'capturar los datos del Iva y Cuenta
ivaTele = Sheets("Maestro").Range("C2").Value
ivaTelc = Sheets("Maestro").Range("C3").Value
ivaTela = Sheets("Maestro").Range("c4").Value
ivaTelp = Sheets("Maestro").Range("c6").Value
cuentaTele = Sheets("Maestro").Range("B2").Value
cuentaTeleITCO = Sheets("Maestro").Range("B5").Value
cuentaTelc = Sheets("Maestro").Range("B3").Value
cuentaTela = Sheets("Maestro").Range("B4").Value
cuentaTelpUSD = Sheets("Maestro").Range("B8").Value
cuentaTelpPEN = Sheets("Maestro").Range("B7").Value
ubicacionArchivo = ActiveWorkbook.Path
nombreArchivo = Range("B14").Value

'Declaracion de un Array
Dim listaCuenta() As String

'If para definir el Db y Cr
If Range("B1").Value = "PROVISION" Then
    claseMovGasto = 40
    claseMovProvision = 50
Else
    claseMovGasto = 50
    claseMovProvision = 40
End If

'If para asignar el indicador IVA, cuenta provision
Dim sociedadProvision As String
sociedadProvision = Range("B8").Value

Select Case sociedadProvision
Case "TELE"
    indicadorIva = ivaTele
    
    If Range("B15").Value = "ITCO" Then
        cuentaProvision = cuentaTeleITCO
    Else
        cuentaProvision = cuentaTele
    End If
Case "TELP"
    indicadorIva = ivaTelp
    If Range("B9").Value = "PEN" Then
        cuentaProvision = cuentaTelpPEN
    Else
        cuentaProvision = cuentaTelpUSD
    End If
Case "TELC"
    indicadorIva = ivaTelc
    cuentaProvision = cuentaTelc
Case "TELA"
    indicadorIva = ivaTela
    cuentaProvision = cuentaTela
Case Else
    indicadorIva = ""
End Select


'capturar los datos para el encabezado
numEncabezado = Range("B2").Value
transaccion = Range("B3").Value
fechaDocumento = Range("B4").Value
sociedad = Range("B8").Value
fechaContabilizacion = Range("B5").Value
mes = Range("B6").Value
claseDocumento = Range("B7").Value
moneda = Range("B9").Value
referencia = Range("B10").Value
txtEncabezado = Range("B11").Value


'contar el numero de filas con datos en la columna de cuentaContable
largoArray = Application.CountA(Columns("D")) - 1
'el primer dato son la cantidad de columnas, el segundo la cantidad de filas
ReDim listaCuenta(11, largoArray)

Dim fila As Integer
Dim columna As Integer

'For anidado, para llenar el array
For fila = 1 To largoArray
    For columna = 1 To 11
        listaCuenta(columna, fila) = ActiveSheet.Cells(fila + 1, columna + 3).Value
    Next columna
Next fila


'Crear el libro nuevo y asignarle los datos
Workbooks.Add



Const cuentaArray As Integer = 1
Const asignacionArray As Integer = 5
Const cecoArray As Integer = 6
Const valorArray As Integer = 7
Const sociedadGLArray As Integer = 8
Const docComprasArray As Integer = 3
Const textoArray As Integer = 11
Const cuentaBatch As Integer = 114
Const valorBatch As Integer = 7
Const asignacionBatch As Integer = 34
Const asignacionBatchDos As Integer = 146
Const cecoBatch As Integer = 16
Const sociedadGLBatch As Integer = 208
Const textoBatch As Integer = 37
Const fechaBatch As Integer = 32
Const finLineaBatch As Integer = 280
Const norelsalBatch As Integer = 124
Const rpBatch As Integer = 142
Const bbsegBatch As Integer = 2
Const claseMovBatch As Integer = 3
Const divisionBatch As Integer = 15
Const indicadorIvaBatch As Integer = 11
Const primerColumnaBatch As Integer = 1
Const docComprasBatch As Integer = 19

Const numEncabezadoBatch As Integer = 1
Const transaccionBatch As Integer = 2
Const fechaDocumentoBatch As Integer = 3
Const claseDocumentoBatch As Integer = 4
Const sociedadBatch As Integer = 5
Const fechaContabilizacionBatch As Integer = 6
Const mesBatch As Integer = 7
Const monedaBatch As Integer = 8
Const referenciaBatch As Integer = 12
Const txtEncabezadoBatch As Integer = 14
Const divisionEncabezadoBatch As Integer = 15
Const finLineaEncabezadoBatch As Integer = 36

Dim filaTres As Integer
filaTres = 1

For fila = 1 To largoArray
'crear encabezado
    Worksheets("Hoja1").Cells(filaTres, numEncabezadoBatch).Value = numEncabezado
    Worksheets("Hoja1").Cells(filaTres, transaccionBatch).Value = transaccion
    Worksheets("Hoja1").Cells(filaTres, fechaDocumentoBatch).Value = fechaDocumento
    Worksheets("Hoja1").Cells(filaTres, claseDocumentoBatch).Value = claseDocumento
    Worksheets("Hoja1").Cells(filaTres, sociedadBatch).Value = sociedad
    Worksheets("Hoja1").Cells(filaTres, fechaContabilizacionBatch).Value = fechaContabilizacion
    Worksheets("Hoja1").Cells(filaTres, mesBatch).Value = mes
    Worksheets("Hoja1").Cells(filaTres, monedaBatch).Value = moneda
    Worksheets("Hoja1").Cells(filaTres, referenciaBatch).Value = referencia
    Worksheets("Hoja1").Cells(filaTres, txtEncabezadoBatch).Value = txtEncabezado
    Worksheets("Hoja1").Cells(filaTres, divisionEncabezadoBatch).Value = division
    Worksheets("Hoja1").Cells(filaTres, finLineaEncabezadoBatch).Value = finLinea
'asignar los valores a las celdas del gasto
    Worksheets("Hoja1").Cells(filaTres + 1, valorBatch).Value = Round(listaCuenta(valorArray, fila), 2)
    Worksheets("Hoja1").Cells(filaTres + 1, asignacionBatch).Value = listaCuenta(asignacionArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 1, asignacionBatchDos).Value = listaCuenta(asignacionArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 1, cuentaBatch).Value = listaCuenta(cuentaArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 1, cecoBatch).Value = listaCuenta(cecoArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 1, sociedadGLBatch).Value = listaCuenta(sociedadGLArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 1, textoBatch).Value = listaCuenta(textoArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 1, norelsalBatch).Value = norelsal
    Worksheets("Hoja1").Cells(filaTres + 1, rpBatch).Value = rp
    Worksheets("Hoja1").Cells(filaTres + 1, divisionBatch).Value = division
    Worksheets("Hoja1").Cells(filaTres + 1, bbsegBatch).Value = bbseg
    Worksheets("Hoja1").Cells(filaTres + 1, claseMovBatch).Value = claseMovGasto
    Worksheets("Hoja1").Cells(filaTres + 1, indicadorIvaBatch).Value = indicadorIva
    Worksheets("Hoja1").Cells(filaTres + 1, finLineaBatch).Value = finLinea
    Worksheets("Hoja1").Cells(filaTres + 1, primerColumnaBatch).Value = primerColumna
    Worksheets("Hoja1").Cells(filaTres + 1, docComprasBatch).Value = listaCuenta(docComprasArray, fila)

'Asignar los valores a las celdas de la provision
    Worksheets("Hoja1").Cells(filaTres + 2, valorBatch).Value = Round(listaCuenta(valorArray, fila), 2)
    Worksheets("Hoja1").Cells(filaTres + 2, asignacionBatch).Value = listaCuenta(asignacionArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 2, asignacionBatchDos).Value = listaCuenta(asignacionArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 2, cuentaBatch).Value = cuentaProvision
    Worksheets("Hoja1").Cells(filaTres + 2, fechaBatch).Value = fechaDocumento
    Worksheets("Hoja1").Cells(filaTres + 2, textoBatch).Value = listaCuenta(textoArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 2, norelsalBatch).Value = norelsal
    Worksheets("Hoja1").Cells(filaTres + 2, rpBatch).Value = rp
    Worksheets("Hoja1").Cells(filaTres + 2, divisionBatch).Value = division
    Worksheets("Hoja1").Cells(filaTres + 2, bbsegBatch).Value = bbseg
    Worksheets("Hoja1").Cells(filaTres + 2, claseMovBatch).Value = claseMovProvision
    Worksheets("Hoja1").Cells(filaTres + 2, finLineaBatch).Value = finLinea
    Worksheets("Hoja1").Cells(filaTres + 2, primerColumnaBatch).Value = primerColumna
    Worksheets("Hoja1").Cells(filaTres + 2, sociedadGLBatch).Value = listaCuenta(sociedadGLArray, fila)
    Worksheets("Hoja1").Cells(filaTres + 2, docComprasBatch).Value = listaCuenta(docComprasArray, fila)
    
filaTres = filaTres + 3
Next fila

'Guardar el archivo nuevo, en la misma ubicacion del generador
ActiveWorkbook.SaveAs Filename:=ubicacionArchivo & "/" & nombreArchivo & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.SaveAs Filename:=ubicacionArchivo & "/" & nombreArchivo & ".csv", FileFormat:=xlCSV, CreateBackup:=False
      

End Sub
