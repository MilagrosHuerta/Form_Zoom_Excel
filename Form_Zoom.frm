VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Zoom 
   Caption         =   "ZOOM"
   ClientHeight    =   7695
   ClientLeft      =   2115
   ClientTop       =   2460
   ClientWidth     =   18510
   OleObjectBlob   =   "Form_Zoom.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'Centrar en pantalla
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form_Zoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------ '
' ---              Formulario creado por                   --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' --- Puedes usarlo libremente en tus aplicaciones,        --- '
' --- pero no asignarte la autoría.                        --- '
' ------------------------------------------------------------ '
' --- Las 3 siguientes líneas son las que hay que escribir --- '
' --- para la macro que llame al formulario ZOOM           --- '
' ------------------------------------------------------------ '
' ---  Sub Zoom_Celda()                                    --- '
' ---    Form_Zoom.Show                                    --- '
' ---  End Sub                                             --- '
' ------------------------------------------------------------ '
Option Explicit
Dim N_Celda, T_Celda As String
Dim N_Fila, N_Colu As Integer
Dim N_Fila_min, N_Colu_min As Integer
Dim N_Fila_max, N_Colu_max As Integer
Dim N_Suma, i, j As Integer
Private Sub UserForm_Initialize()
    If Formula.Value = True Then
        Me.Text_Celda = ActiveCell.FormulaLocal
    Else
        Me.Text_Celda = ActiveCell.Value
    End If
    Form_Zoom.Caption = "ZOOM - " & ActiveSheet.Name
    Call Titulo_Nombre_Celda
    j = 1000                    ' Número de veces que se repite el bucle de la etiqueta SIGIENTE antes de cerrar
End Sub
Private Sub Titulo_Nombre_Celda()
'---------------------------------------------------------------------------------------- '
'--- Escribe el título del nombre de la celda, cada uno lo configura como lo necesite --- '
'--- Se ponen algunos ejemplos, para que se vea cómo se puede configurar, en función  --- '
'--- de la hoja en la que estemos trabajando.                                         --- '
'---------------------------------------------------------------------------------------- '
    If ActiveSheet.Name = "Tabla 1" Then
        N_Celda = Cells(1, ActiveCell.Column) & " - " & Cells(ActiveCell.Row, 1) & " - " & _
                        Replace(ActiveCell.AddressLocal, "$", "")                                   ' Nombre Columna - Nombre Fila - Nombre Celda
    ElseIf ActiveSheet.Name = "Tabla 2" Then
        N_Celda = Cells(1, ActiveCell.Column) & " - " & Replace(ActiveCell.AddressLocal, "$", "")   ' Nombre Columna - Nombre Celda
    ElseIf ActiveSheet.Name = "Formulario" Then
        N_Celda = Cells(ActiveCell.Row, ActiveCell.Column - 1)                                      ' Nombre Formulario - Nombre Celda
    Else
        N_Celda = Replace(ActiveCell.AddressLocal, "$", "")                                         ' Nombre Celda sin el símbolo "&"
    End If
    Me.Text_N_Celda = N_Celda
End Sub
Private Sub Rangos_Celdas()
Dim Protegida As Integer
' ----------------------------------------------------------------------------- '
' --- Asignamos los datos mínimos o máximos por los que se puede desplazar. --- '
' --- Estos datos podrían variar en función de la hoja en la que está.      --- '
' --- Para ello, habría que poner los condicionales necesarios.             --- '
' ----------------------------------------------------------------------------- '
    Application.ScreenUpdating = False
' --- Se puede indicar la primera fila y la primera columna por la que se quiere que empiece a desplazarse
    N_Fila_min = 1
    N_Colu_min = 1
' --- Si la hoja está protegida, las celdas especiales no sirven, por lo que habría que desproteger temporalmente la hoja
   If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect
        Protegida = 0
    Else
        Protegida = 1
    End If
    N_Fila_max = ActiveCell.SpecialCells(xlLastCell).Row        ' Busca la última Fila con datos
    N_Colu_max = ActiveCell.SpecialCells(xlLastCell).Column     ' Busca la última Columna con datos
    
    If N_Fila_min > N_Fila_max Then N_Fila_min = N_Fila_max     ' Si el valor mínimo definido es mayor que el máximo, los iguala
    If N_Colu_min > N_Colu_max Then N_Colu_min = N_Colu_max
    
    If Protegida = 0 Then ActiveSheet.Protect
End Sub
Private Sub B_Formula_Click()
    If Formula.Value = True Then
        Formula.Value = False
        Me.Text_Celda = ActiveCell.Value
    Else
        Formula.Value = True
        Me.Text_Celda = ActiveCell.FormulaLocal
    End If
End Sub
Private Sub B_Actualiza_Click()
    Call UserForm_Initialize
End Sub
Private Sub B_Arriba_Click()
' --------------------------------------------------------------------------------------------------------------- '
' --- Asignamos los datos mínimos y máximos. Estos datos podrían variar en función de la hoja en la que está. --- '
' --- Para ello, habría que poner los condicionales necesarios                                                --- '
' --------------------------------------------------------------------------------------------------------------- '
    N_Suma = -1         ' Negativo porque sube filas
    Call Rangos_Celdas
' --- EMPIEZA EL PROCESO --------------------------------------------------------
SIGUIENTE:
    If ActiveCell.Row = N_Fila_min Then
        N_Fila = N_Fila_max
        N_Colu = ActiveCell.Column + N_Suma
        If N_Colu < N_Colu_min Then N_Colu = N_Colu_max
        Cells(N_Fila, N_Colu).Select
    Else
        ActiveCell.Offset(N_Suma, 0).Range("A1").Select         ' Se desplaza tantas filas como N_Suma se haya definido
    End If
    i = i + 1
    If i > j Then End                                           ' Para que salga del bucle si hay algún error
    If ActiveCell.Value = "" Or ActiveCell.Width = 0 Or ActiveCell.Height = 0 Then GoTo SIGUIENTE
' ------------------------------------------------------------------------------------------------------------ '
' --- Este condicional solo se pone si hay celdas que no estén bloqueadas, de otra forma entraría en bucle --- '
' ------------------------------------------------------------------------------------------------------------ '
    If ActiveSheet.ProtectContents = True Then
        If ActiveCell.Locked = True Then GoTo SIGUIENTE
    End If
    
    Call UserForm_Initialize
End Sub
Private Sub B_Abajo_Click()
' --------------------------------------------------------------------------
' Asignamos los datos mínimos y máximos, podrían definirse en función de la hoja en la que esté con CONDICIONAL
' --------------------------------------------------------------------------
    N_Suma = 1
    Call Rangos_Celdas
' --- EMPIEZA EL PROCESO --------------------------------------------------------
SIGUIENTE:
    If ActiveCell.Row = N_Fila_max Then
        N_Fila = N_Fila_min
        N_Colu = ActiveCell.Column + N_Suma
        If N_Colu > N_Colu_max Then N_Colu = N_Colu_min
        Cells(N_Fila, N_Colu).Select
    Else
        ActiveCell.Offset(N_Suma, 0).Range("A1").Select         ' Se desplaza tantas filas como N_Suma se haya definido
    End If
    
    i = i + 1
    If i > j Then End         ' Para que salga del bucle si hay algún error
    If ActiveCell.Value = "" Or ActiveCell.Width = 0 Or ActiveCell.Height = 0 Then GoTo SIGUIENTE
' ------------------------------------------------------------------------------------------------------------ '
' --- Este condicional solo se pone si hay celdas que no estén bloqueadas, de otra forma entraría en bucle --- '
' ------------------------------------------------------------------------------------------------------------ '
    If ActiveSheet.ProtectContents = True Then
        If ActiveCell.Locked = True Then GoTo SIGUIENTE
    End If

    Call UserForm_Initialize
End Sub
Private Sub B_Derecha_Click()
' --------------------------------------------------------------------------
' Asignamos los datos mínimos y máximos, podrían definirse en función de la hoja en la que esté con CONDICIONAL
' --------------------------------------------------------------------------
    N_Suma = 1
    Call Rangos_Celdas
' --- EMPIEZA EL PROCESO --------------------------------------------------------
SIGUIENTE:
    If ActiveCell.Column = N_Colu_max Then
        N_Fila = ActiveCell.Row + N_Suma
        N_Colu = N_Colu_min
        If N_Fila > N_Fila_max Then N_Fila = N_Fila_min
        Cells(N_Fila, N_Colu).Select
    Else
        ActiveCell.Offset(0, N_Suma).Range("A1").Select    ' Se desplaza tantas columnas como N_Suma se haya definido
    End If

    i = i + 1
    If i > j Then End         ' Para que salga del bucle si hay algún error
    If ActiveCell.Value = "" Or ActiveCell.Width = 0 Or ActiveCell.Height = 0 Then GoTo SIGUIENTE
' ------------------------------------------------------------------------------------------------------------ '
' --- Este condicional solo se pone si hay celdas que no estén bloqueadas, de otra forma entraría en bucle --- '
' ------------------------------------------------------------------------------------------------------------ '
    If ActiveSheet.ProtectContents = True Then
        If ActiveCell.Locked = True Then GoTo SIGUIENTE
    End If

    Call UserForm_Initialize
 End Sub
Private Sub B_Izquierda_Click()
' --------------------------------------------------------------------------
' Asignamos los datos mínimos y máximos, podrían definirse en función de la hoja en la que esté con CONDICIONAL
' --------------------------------------------------------------------------
    N_Suma = -1
    Call Rangos_Celdas
' --- EMPIEZA EL PROCESO --------------------------------------------------------
SIGUIENTE:
    If ActiveCell.Column = N_Colu_min Or ActiveCell.Column = 1 Then
        N_Fila = ActiveCell.Row + N_Suma
        N_Colu = N_Colu_max
        If N_Fila < N_Fila_min Then N_Fila = N_Fila_max
        Cells(N_Fila, N_Colu).Select
    Else
        ActiveCell.Offset(0, N_Suma).Range("A1").Select    ' Se desplaza tantas columnas como N_Suma se haya definido
    End If

    i = i + 1
    If i > j Then End         ' Para que salga del bucle si hay algún error
    If ActiveCell.Value = "" Or ActiveCell.Width = 0 Or ActiveCell.Height = 0 Then GoTo SIGUIENTE
' ------------------------------------------------------------------------------------------------------------ '
' --- Este condicional solo se pone si hay celdas que no estén bloqueadas, de otra forma entraría en bucle --- '
' ------------------------------------------------------------------------------------------------------------ '
    If ActiveSheet.ProtectContents = True Then
        If ActiveCell.Locked = True Then GoTo SIGUIENTE
    End If

    Call UserForm_Initialize
End Sub
Private Sub B_Guardar_Click()
'----------------------------------- '
'--- Guarda los cambios la celda --- '
'----------------------------------- '
    T_Celda = Me.Text_Celda
    N_Celda = Replace(ActiveCell.AddressLocal, "$", "")
    
    If T_Celda = ActiveCell.FormulaLocal Then
        Me.Text_N_Celda = "NO hay cambios en la celda " & N_Celda
    Else
    On Error GoTo Mensaje_ERROR
        ActiveCell = T_Celda
        Me.Text_N_Celda = "Guardando cambios en la celda " & N_Celda
    End If
    Application.Wait Now + TimeValue("00:00:02")
    Call Titulo_Nombre_Celda
    
    Exit Sub

Mensaje_ERROR:
    Me.Text_N_Celda = "No se pueden guardar los cambios en la celda " & N_Celda
    Application.Wait Now + TimeValue("00:00:02")
End Sub
Private Sub B_Cerrar_Click()
    Unload Me
End Sub
Private Sub B_GuardarCerrar_Click()
    B_Guardar_Click
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ------------------------------------------------------------------------------------ '
' --- Para evitar que se pueda cerrar el formulario en la X de arriba a la derecha --- '
' ------------------------------------------------------------------------------------ '
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
