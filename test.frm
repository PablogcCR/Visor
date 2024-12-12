VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Visor 
   Caption         =   "Búsqueda de Información de Pago"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   Icon            =   "test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cb_cancelar_mass 
      Caption         =   "Ca&ncelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox sle_file_actual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cb_cerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   10200
      TabIndex        =   3
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   9000
      Picture         =   "test.frx":030A
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   6600
      Width           =   5655
      Begin VB.CommandButton cb_imprimir_todas 
         Caption         =   "Imprimir &Directorio"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cb_imprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   1800
         X2              =   1800
         Y1              =   120
         Y2              =   720
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
      ExtentX         =   7223
      ExtentY         =   1296
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   5535
      Begin VB.Label lbl_progreso 
         Caption         =   "Archivo xx de xxx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo Actual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Label lbl_mensaje 
      Caption         =   "Ningún archivo seleccionado, presione ""Buscar"" para visualizar información del pago."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   8055
   End
End
Attribute VB_Name = "Visor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      'Definición de variables
      Dim fso As New FileSystemObject
      Dim fld As Folder
      Dim istr_path As String
      Dim ibln_listo As Boolean
      Dim ib_seguir As Boolean
      
      

Private Sub cb_cancelar_mass_Click()
    ib_seguir = False
End Sub

Private Sub cb_cerrar_Click()
    On Error Resume Next
    Unload Visor
End Sub

Private Sub cb_imprimir_Click()
    On Error Resume Next
    MsgBox "Verfique que la impresora esté conectada, encendida y con papel tamaño carta colocado. Haga click en Aceptar para continuar.", vbExclamation + vbOKOnly, "Verificación de Impresora"
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
    
End Sub

Private Sub cb_imprimir_todas_Click()
    Dim fs, f, f1, fc
    Dim ls_path, ls_temp As String
    Dim li_actual As Integer
    
    If MsgBox("El siguiente proceso imprime todos los archivos ubicados en el directorio en el cual se encuentra ubicado el archivo actual en pantalla.  Desea continuar?", vbYesNo + vbQuestion, "Impresión Masiva de Archivos") = vbYes Then
        MsgBox "Verfique que la impresora esté conectada, encendida y con papel tamaño carta colocado. Haga click en Aceptar para continuar.", vbExclamation + vbOKOnly, "Verificación de Impresora"
        MousePointer = vbHourglass
        WebBrowser2.Visible = True
        lbl_progreso.Visible = True
        sle_file_actual.Visible = True
        Label1.Visible = True
        cb_cancelar_mass.Visible = True
        
        WebBrowser2.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
        ls_path = istr_path
        ls_temp = ls_path
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(ls_path)
        Set fc = f.Files
        li_actual = 1
        For Each f1 In fc
            If ib_seguir Then
                ls_path = ls_path + f1.Name
                
                'Se refrescan los elementos de retroalimentación al usuario
                lbl_progreso.Caption = "Archivo " + CStr(li_actual) + " de " + CStr(fc.Count)
                sle_file_actual.Text = f1.Name
                
                ibln_listo = False
                WebBrowser2.Navigate2 (ls_path)
                Do While (Not ibln_listo)
                    DoEvents
                Loop
                
                WebBrowser2.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
                ls_path = ls_temp
                li_actual = li_actual + 1
            Else
                MousePointer = vbDefault
                WebBrowser2.Visible = False
                lbl_progreso.Visible = False
                sle_file_actual.Visible = False
                Label1.Visible = False
                cb_cancelar_mass.Visible = False
                MsgBox "Proceso de impresión Cancelado por el Usuario.", vbOKOnly + vbInformation, "Resultado del Proceso"
                Exit Sub
            End If
        Next
        MousePointer = vbDefault
        WebBrowser2.Visible = False
        lbl_progreso.Visible = False
        sle_file_actual.Visible = False
        Label1.Visible = False
        cb_cancelar_mass.Visible = False
        MsgBox "Proceso de impresión finalizado.", vbOKOnly + vbInformation, "Resultado del Proceso"
    End If
End Sub

Private Sub Command1_Click()
'*****************************************************************/
'* Objetivo:                                                     */
'*    Llamar la funcón de búsqueda del un archivo y si se encuentra*/
'*    cargarlo en el control web browser.                        */
'*****************************************************************/
'* Estrategia General de Solución:                               */
'*    Se obtiene el nombre del archivo a buscar, se llama la f(x)*/
'*    y si se encuentra el archivo se manda a cargar por medio de*/
'*    la funcion navigate2 del control web browser               */
'*****************************************************************/
'* Programado Por:              David Antonio Rodríguez Alfaro.  */
'*  Fecha de Creación:          2003-06-13                       */
'*  Fecha de Liberación:                                         */
'*****************************************************************/
'* Ult. Modificador:            David Antonio Rodríguez Alfaro.  */
'*  Ult. Fecha Modificación:    2003-06-16                       */
'*  Cambios Realizados:       0                                  */
'*****************************************************************/
    Dim sDir As String, sSrchString As String
    Dim ls_rutacompleta As String
    Dim ls_par As String
    Dim lb_seguir As Boolean
    
    'se posiciona en la ruta actual
    sDir = CurDir
    lb_seguir = True
    sSrchString = InputBox("Ingrese el Número de Cédula del Empleado", _
                    "Número de Cédula", "000000000")
    MousePointer = vbHourglass
    sSrchString = sSrchString + ".html "
    ls_par = FindFile(sDir, sSrchString, ls_rutacompleta, lb_seguir)
    MousePointer = vbDefault
    
    
    
    If Len(Trim(ls_rutacompleta)) > 0 Then
        istr_path = Left(ls_rutacompleta, (Len(ls_rutacompleta) - Len(sSrchString)) + 1)
        WebBrowser1.Visible = True
        WebBrowser1.Navigate2 (ls_rutacompleta)
        cb_imprimir.Enabled = True
        cb_imprimir_todas.Enabled = True
    Else
        MsgBox "Cédula no encontrada, intente nuevamente.", vbOKOnly + vbExclamation, "Archivo no encontrado"
        WebBrowser1.Visible = False
        cb_imprimir.Enabled = False
        cb_imprimir_todas.Enabled = False
    End If
    
    
End Sub

Private Function FindFile(ByVal sFol As String, _
                         sFile As String, _
                         ByRef filenamex As String, _
                         ByRef ab_seguir As Boolean) As String
'*****************************************************************/
'* Objetivo:                                                     */
'*      Realizar la búsqueda recursiva de un archivo a partir del*/
'*    directorio actual.                                         */
'*****************************************************************/
'* Estrategia General de Solución:                               */
'*                                                               */
'*      Se utilizan funciones propias del objeto filesystemobject*/
'*    para recorrer recursivamente el sistema de archivos y buscar/
'*    así un archivo específico.                                 */
'*****************************************************************/
'* Programado Por:              David Antonio Rodríguez Alfaro.  */
'*  Fecha de Creación:          2003-06-13                       */
'*  Fecha de Liberación:                                         */
'*****************************************************************/
'* Ult. Modificador:            David Antonio Rodríguez Alfaro.  */
'*  Ult. Fecha Modificación:    2003-06-16                       */
'*  Cambios Realizados:       0                                  */
'*****************************************************************/
    Dim tFld As Folder, tFil As File, FileName As String
    
    Set fld = fso.GetFolder(sFol)
    filenamex = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                    vbHidden Or vbSystem Or vbReadOnly)
    If Len(filenamex) <> 0 Then
        filenamex = fso.BuildPath(fld.Path, filenamex)
        ab_seguir = False
        FindFile = fso.BuildPath(fld.Path, filenamex)
    End If
    If fld.SubFolders.Count > 0 And Len(filenamex) = 0 Then
        For Each tFld In fld.SubFolders
            If ab_seguir Then
                 FindFile tFld.Path, sFile, filenamex, ab_seguir
            End If
        Next
    End If
End Function
 

Private Sub Form_Load()
    ib_seguir = True
End Sub

Private Sub WebBrowser2_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    ibln_listo = True
End Sub

