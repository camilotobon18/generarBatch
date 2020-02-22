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
Const ivaTele As String = "V0"
Const ivaTelc As String = "C0"
Const ivaTela As String = "C0"
Dim cuentaProvision As String
Const cuentaTele As String = "2790909013"
Const cuentaTelc As String = "1"
Const cuentaTela As String = "2"
Dim claseDocumento As String
Const claseTele As String = "SA"
Const claseTelc As String = "aa"
Const claseTela As String = "aa"

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

'If para asignar el indicador IVA, cuenta provision y clase documento
If Range("B8").Value = "TELE" Then
indicadorIva = ivaTele
cuentaProvision = cuentaTele
claseDocumento = claseTele
ElseIf Range("B8").Value = "TELC" Then
indicadorIva = ivaTelc
cuentaProvision = cuentaTelc
claseDocumento = claseTelc
Else
indicadorIva = ivaTela
cuentaProvision = cuentaTela
claseDocumento = claseTela
End If

'capturar los datos para el encabezado
numEncabezado = Range("B2").Value
transaccion = Range("B3").Value
fechaDocumento = Range("B4").Value
sociedad = Range("B8").Value
fechaContabilizacion = Range("B5").Value
mes = Range("B6").Value
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
listaCuenta(columna, fila) = Worksheets("Hoja1").Cells(fila + 1, columna + 3).Value
Next columna
Next fila


'Crear el libro nuevo y asignarle los datos
Workbooks.Add
Range("A1").Value = numEncabezado
Range("B1").Value = transaccion
Range("C1").Value = fechaDocumento
Range("D1").Value = claseDocumento
Range("E1").Value = sociedad
Range("F1").Value = fechaContabilizacion
Range("G1").Value = mes
Range("H1").Value = moneda
Range("L1").Value = referencia
Range("N1").Value = txtEncabezado
Range("O1").Value = division
Range("AJ1").Value = finLinea
Range("O1").Value = division


Const cuentaArray As Integer = 1
Const asignacionArray As Integer = 5
Const cecoArray As Integer = 6
Const valorArray As Integer = 7
Const sociedadGLArray As Integer = 8
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


'For para asignar los valores a las celdas del gasto
For fila = 1 To largoArray
Worksheets("Hoja1").Cells(fila + 1, valorBatch).Value = listaCuenta(valorArray, fila)
Worksheets("Hoja1").Cells(fila + 1, asignacionBatch).Value = listaCuenta(asignacionArray, fila)
Worksheets("Hoja1").Cells(fila + 1, asignacionBatchDos).Value = listaCuenta(asignacionArray, fila)
Worksheets("Hoja1").Cells(fila + 1, cuentaBatch).Value = listaCuenta(cuentaArray, fila)
Worksheets("Hoja1").Cells(fila + 1, cecoBatch).Value = listaCuenta(cecoArray, fila)
Worksheets("Hoja1").Cells(fila + 1, sociedadGLBatch).Value = listaCuenta(sociedadGLArray, fila)
Worksheets("Hoja1").Cells(fila + 1, textoBatch).Value = listaCuenta(textoArray, fila)
Worksheets("Hoja1").Cells(fila + 1, norelsalBatch).Value = norelsal
Worksheets("Hoja1").Cells(fila + 1, rpBatch).Value = rp
Worksheets("Hoja1").Cells(fila + 1, divisionBatch).Value = division
Worksheets("Hoja1").Cells(fila + 1, bbsegBatch).Value = bbseg
Worksheets("Hoja1").Cells(fila + 1, claseMovBatch).Value = claseMovGasto
Worksheets("Hoja1").Cells(fila + 1, indicadorIvaBatch).Value = indicadorIva
Worksheets("Hoja1").Cells(fila + 1, finLineaBatch).Value = finLinea
Worksheets("Hoja1").Cells(fila + 1, primerColumnaBatch).Value = primerColumna
Next fila

'For para asignar los valores a las celdas de las provisiones
For fila = 1 To largoArray
Worksheets("Hoja1").Cells(fila + 1 + largoArray, valorBatch).Value = listaCuenta(valorArray, fila)
Worksheets("Hoja1").Cells(fila + 1 + largoArray, asignacionBatch).Value = listaCuenta(asignacionArray, fila)
Worksheets("Hoja1").Cells(fila + 1 + largoArray, asignacionBatchDos).Value = listaCuenta(asignacionArray, fila)
Worksheets("Hoja1").Cells(fila + 1 + largoArray, cuentaBatch).Value = cuentaProvision
Worksheets("Hoja1").Cells(fila + 1 + largoArray, fechaBatch).Value = fechaDocumento
Worksheets("Hoja1").Cells(fila + 1 + largoArray, textoBatch).Value = listaCuenta(textoArray, fila)
Worksheets("Hoja1").Cells(fila + 1 + largoArray, norelsalBatch).Value = norelsal
Worksheets("Hoja1").Cells(fila + 1 + largoArray, rpBatch).Value = rp
Worksheets("Hoja1").Cells(fila + 1 + largoArray, divisionBatch).Value = division
Worksheets("Hoja1").Cells(fila + 1 + largoArray, bbsegBatch).Value = bbseg
Worksheets("Hoja1").Cells(fila + 1 + largoArray, claseMovBatch).Value = claseMovProvision
Worksheets("Hoja1").Cells(fila + 1 + largoArray, finLineaBatch).Value = finLinea
Worksheets("Hoja1").Cells(fila + 1 + largoArray, primerColumnaBatch).Value = primerColumna
Next fila

End Sub


