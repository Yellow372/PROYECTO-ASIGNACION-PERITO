VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_asignacion 
   Caption         =   "Asignación de peritos"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frm_asignacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_asignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public RutaPDF_Global As String
Public wbDatos As Workbook ' Referencia

Private Sub UserForm_Initialize()
    With Me.cmb_ListaUsuarios
        .Clear
        .AddItem "Fransheska Rodriguez"
        .AddItem "Prueba Prueba"
        .ListIndex = 0
    End With
End Sub

Private Sub btn_enviar_Click()
    Dim olApp As Object 'Aplicacion Outlook
    Dim olMailItm As Object 'Aplicacion Outlook
    Dim archivoPDF As String
    Dim i As Long
    Dim anexo As String
    Dim usuario As String
    Dim garantia As String
    Dim totalEnviados As Long
    Dim ws As Worksheet
    Dim reporteWs As Worksheet
    Dim repRow As Long

    If Me.cmb_ListaUsuarios.ListIndex = -1 Then
        MsgBox "Selecciona un firmante antes de continuar.", vbExclamation
        Exit Sub
    End If

    If wbDatos Is Nothing Then 'En caso de no tener un excel de datos
        MsgBox "Debes seleccionar primero el archivo Excel de origen.", vbExclamation
        Exit Sub
    End If

    Set ws = wbDatos.Sheets(1) 'Vamos a escribir las garantías sin pdf que no se envian
    On Error Resume Next
    Set reporteWs = ThisWorkbook.Sheets("REPORTES")
    If reporteWs Is Nothing Then
        Set reporteWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        reporteWs.Name = "REPORTES"
    End If
    On Error GoTo 0
    reporteWs.Cells.ClearContents
    reporteWs.Range("A1").Value = "Garantías sin PDF adjunto"
    repRow = 2

    usuario = Me.cmb_ListaUsuarios.Value 'para los anexos
    Select Case usuario
        Case "Fransheska Rodriguez": anexo = "2516"
        Case "Prueba Prueba": anexo = "0011"
        Case Else: anexo = "0000"
    End Select

    Set olApp = CreateObject("Outlook.Application")
    totalEnviados = 0

    For i = 2 To ws.Range("B" & ws.Rows.Count).End(xlUp).Row 'Buscaremos en toda la columna D los valores
        garantia = ws.Cells(i, "D").Value

        ' Busqueda de PDF
        archivoPDF = ""
        Dim f As String
        f = Dir(RutaPDF_Global & "\*.pdf")
        Do While f <> ""
            If InStr(1, f, garantia, vbTextCompare) > 0 Then
                archivoPDF = RutaPDF_Global & "\" & f
                Exit Do
            End If
            f = Dir
        Loop

        If archivoPDF = "" Then 'evita enviar el correo si no hay pdf
            reporteWs.Cells(repRow, 1).Value = garantia
            repRow = repRow + 1
            GoTo SkipEnvio
        End If

        Set olMailItm = olApp.CreateItem(0)

        Dim htmlBody As String 'Solución formato
        
        htmlBody = "<body style='font-family:Cambria;font-size:9pt;'>"
        htmlBody = htmlBody & "<p><b>Estimado(a) Perito: " & ws.Cells(i, "V").Value & "</b></p>"
        htmlBody = htmlBody & "<p>La presente es para solicitar sus servicios, a fin de que procedan con la siguiente retasación, <span style='color:red;'>recordándoles alcanzar toda la información solicitada en los formatos establecidos por el BCP.</span><br>"
        htmlBody = htmlBody & "<span style='color:red; text-decoration:underline;'>Esta información es obligatoria en los informes para todos los bienes inmuebles, maquinarias de todo tipo, contenidos y existencias.</span></p>"
        htmlBody = htmlBody & "<p>Los datos de la garantía son:</p>"

        htmlBody = htmlBody & "<table border='1' cellspacing='0' cellpadding='3' style='font-family:Cambria;font-size:9pt;border-collapse:collapse;'>"
        htmlBody = htmlBody & "<tr><td><b>GARANTIA</b></td><td>" & ws.Cells(i, "D").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>CLIENTE</b></td><td>" & ws.Cells(i, "E").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>BANCA</b></td><td><b>" & ws.Cells(i, "B").Value & "</b></td></tr>"
        htmlBody = htmlBody & "<tr><td><b>CONTACTO</b></td><td>" & ws.Cells(i, "E").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>TELEFONO</b></td><td>" & ws.Cells(i, "F").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>CORREO</b></td><td>" & ws.Cells(i, "G").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>DIRECCION</b></td><td>" & ws.Cells(i, "N").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>LOCALIDAD</b></td><td>" & ws.Cells(i, "O").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>DESCRIPCION DEL BIEN</b></td><td>" & ws.Cells(i, "W").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>FECHA ULTIMA VALUACION</b></td><td>" & ws.Cells(i, "Q").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>AREA DE TERRENO</b></td><td>" & ws.Cells(i, "U").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>AREA CONSTRUIDA</b></td><td>" & ws.Cells(i, "T").Value & "</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>NÚMERO DE PARTIDA</b></td><td>" & ws.Cells(i, "P").Value & "</td></tr>"
        htmlBody = htmlBody & "</table><br>"

        htmlBody = htmlBody & "<table border='1' cellspacing='0' cellpadding='3' style='font-family:Cambria;font-size:9pt;border-collapse:collapse;'>"
        htmlBody = htmlBody & "<tr><td><b>PLAZO DE ENTREGA</b></td><td>[POR VER]</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>ENTREGA DE FISICO</b></td><td>CONSEJEROS Y CORREDORES DE SEGUROS | ¦ Av. Javier Prado Este 488 – Torre Orquídeas – Piso 6 – Of. 601, San Isidro, Lima.</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>OBSERVACIONES</b></td><td>[POR VER]</td></tr>"
        htmlBody = htmlBody & "<tr><td><b>FACTURAR AL</b></td><td>[POR VER]</td></tr>"
        htmlBody = htmlBody & "</table><br>"

        htmlBody = htmlBody & "<p><span style='font-weight:bold;background:yellow;'>Por favor tenga a bien informar la ACEPTACIÓN o la NO ACEPTACIÓN dentro del plazo de 24hrs. De lo contrario lo tomaremos como ACEPTADO.</span></p>"
        htmlBody = htmlBody & "<p><span style='color:red;'>INFORMAR si cuenta con correos electrónicos nuevos para actualizarlos inmediatamente.</span></p>"
        htmlBody = htmlBody & "<br><b>" & usuario & "<br>Centro de Seguros<br>+51(1) 200-4-222 - Anexo " & anexo & "</b>"
        htmlBody = htmlBody & "</body>"

        With olMailItm 'Destinataria del correo
            .To = "ealgendones@consejeros.com.pe"
            .Subject = "Solicitud de retasación - " & garantia
            .htmlBody = htmlBody
            .Attachments.Add archivoPDF
            .Send
        End With

        totalEnviados = totalEnviados + 1 'Mide la cantidad de correos enviados
SkipEnvio:
    Next i

    MsgBox totalEnviados & " correo(s) enviados correctamente. Revise la hoja REPORTES para los no enviados.", vbInformation
    wbDatos.Close SaveChanges:=False
    Set wbDatos = Nothing
    Unload Me
End Sub
'Boton "Buscar pdf"
Private Sub btn_pdf_Click()
    Dim fd As FileDialog
    Dim folderPath As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Selecciona carpeta donde están los PDF"
        .AllowMultiSelect = False
        If .Show = -1 Then
            RutaPDF_Global = .SelectedItems(1)
            MsgBox "Ruta seleccionada: " & RutaPDF_Global, vbInformation
        End If
    End With
End Sub
'Boton "Buscar base de datos"
Private Sub btn_bd_Click()
    Dim fd As FileDialog 'Sirve como interfaz
    Dim rutaBase As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecciona archivo Excel de origen"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm" 'Tipo de archhivos
        .AllowMultiSelect = False
        If .Show = -1 Then
            rutaBase = .SelectedItems(1) 'Abrimos el archivo para sacar los datos
            Set wbDatos = Workbooks.Open(rutaBase, ReadOnly:=True)
            MsgBox "Archivo abierto: " & wbDatos.Name, vbInformation
        End If
    End With
End Sub

