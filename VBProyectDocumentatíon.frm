VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00A85828&
   Caption         =   "Documentatión Tools"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "VBProyectDocumentatíon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Tag             =   "Form Principal Para la documentación de Proyectos."
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Tag             =   "Texto no Visible utilizado para paso de parametrizaciones."
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D16727&
      Height          =   495
      Left            =   2949
      Picture         =   "VBProyectDocumentatíon.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "Boton de Carga de Datos a Pantalla y Liberación de Impresión de Informe"
      ToolTipText     =   "Ejecutar Selección"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D16727&
      Enabled         =   0   'False
      Height          =   495
      Left            =   6246
      Picture         =   "VBProyectDocumentatíon.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Boton de Impresión de Informe"
      ToolTipText     =   "Ejecutar Selección"
      Top             =   6480
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   10271
      Tag             =   "Objeto de Dialogo utilizado para Captura del Archivo a Seleccionar."
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4935
      Left            =   878
      TabIndex        =   1
      Tag             =   "Arbol de Nodos para Despliegue de la Información"
      Top             =   1245
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8705
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Tag             =   "Imagenes utilizadas en el proceso de Visualización en pantalla"
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":1A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":2738
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":3412
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":3CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":413E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":4A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":52F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":560C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VBProyectDocumentatíon.frx":5EE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   983
      Picture         =   "VBProyectDocumentatíon.frx":6D38
      Tag             =   "Imagen aderida a la Visualización del Diseño del Form"
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documentation Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1583
      TabIndex        =   2
      Tag             =   "Titulo del Diseño de Presentación de Información"
      Top             =   465
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5415
      Left            =   518
      Shape           =   4  'Rounded Rectangle
      Tag             =   "Segundo Fondo del Diseño de la forma."
      Top             =   1005
      Width           =   9975
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00F2D5CC&
      FillStyle       =   0  'Solid
      Height          =   6975
      Left            =   398
      Shape           =   4  'Rounded Rectangle
      Tag             =   "Primer Fondo del Diseño de la forma"
      Top             =   225
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum xlConstants
'// Definición de Constantes utilizadas en la llamada a EXCEL para definir diversas propiedades.
  xlSolid = 1
  xlCenter = &HFFFFEFF4
  xlNormal = &HFFFFEFD1
End Enum



Dim MyAppID99 As Object
Dim MyWorkBook99
Dim MyWorkShet99

Dim XLS As Boolean

Private Sub ProceesToFile(ByRef MyAppID2 As Object, ByRef MyWorkBook2, ByRef MyWorkShet2)
'// Segundo segmento de Generación de Archivo de Informe en EXCEL.
'// El Proceso genera básicamente  el ultimo Sheet de Documentación que genera el scaneo del código de cada parte del proyecto.
Dim ContadorFuente As Integer
Dim PathProyect As String
Dim Con1 As ADODB.Connection
Dim TextoImpreso As String
Dim FileOpenRead As Integer
Dim StringLectura As String
Dim PrincipalKey2 As String
Dim DescriptionKey2 As String
Dim LocateDate2 As Boolean
'**************************************************************************************
'*********                    Tercer Hoja de Documentación                  **********
'**************************************************************************************
MyWorkShet2(3).Name = "Doc. Process"
MyWorkShet2(3).Select
MyAppID2.Range("A1:D2").Select
MyAppID2.Selection.Interior.ColorIndex = 46
MyAppID2.Selection.Interior.Pattern = xlSolid
MyAppID2.Selection.MergeCells = True
MyAppID2.Selection.HorizontalAlignment = xlCenter
MyAppID2.Selection.Font.Size = 12
MyAppID2.Range("A1").FormulaR1C1 = "Corporación GYT Continental" & vbLf & "Soluciones Tecnológicas"
MyAppID2.Columns("A:C").Select
MyAppID2.Selection.ColumnWidth = 30.71
MyAppID2.Rows("1:2").Select
MyAppID2.Selection.RowHeight = 20
MyAppID2.Range("A4").Select
MyAppID2.Selection.FormulaR1C1 = "Formato de Descripción de Sistema"
MyAppID2.Range("A4:D4").Select
MyAppID2.Selection.Interior.ColorIndex = 6
MyAppID2.Selection.Interior.Pattern = xlSolid
MyAppID2.Selection.MergeCells = True
MyAppID2.Selection.Font.Bold = True

MyAppID2.Range("A6").Select
MyAppID2.Selection.Interior.ColorIndex = 15
MyAppID2.Selection.Interior.Pattern = xlSolid
MyAppID2.Selection.Font.Bold = True
MyAppID2.Selection.FormulaR1C1 = "Proyecto:"
MyAppID2.Range("B6").Select
MyAppID2.Selection.FormulaR1C1 = TreeView1.Nodes(1).Text
'******************************************************
' Extraccion de Propiedades del Informe
'******************************************************
y = 1
If TreeView1.Nodes(2).Children > 0 Then

      i = TreeView1.Nodes(2).Child.Index
      
      If i = TreeView1.Nodes(2).Child.LastSibling.Index Then
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 15
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
            MyAppID2.Range("B" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            y = y + 1
      End If
      
      While i <> TreeView1.Nodes(2).Child.LastSibling.Index
         DoEvents
         MyAppID2.Range("A" & (7 + y)).Select
         MyAppID2.Selection.Interior.ColorIndex = 15
         MyAppID2.Selection.Interior.Pattern = xlSolid
         MyAppID2.Selection.Font.Bold = True
         MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
         MyAppID2.Range("B" & (7 + y)).Select
         MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)

         i = TreeView1.Nodes(i).Next.Index
         y = y + 1
         If i = TreeView1.Nodes(2).Child.LastSibling.Index Then
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 15
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
            MyAppID2.Range("B" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            y = y + 1
         End If
      Wend
      
End If

If TreeView1.Nodes(3).Children > 0 Then
y = y + 3
ContadorFuente = 0
MyAppID2.Range("A" & (7 + y - 2)).Select
MyAppID2.Selection.FormulaR1C1 = "Componentes del Proyecto:"
MyAppID2.Range("A" & (7 + y - 2) & ":D" & (7 + y - 2)).Select
MyAppID2.Selection.Interior.ColorIndex = 6
MyAppID2.Selection.Interior.Pattern = xlSolid
MyAppID2.Selection.MergeCells = True
MyAppID2.Selection.Font.Bold = True


i = TreeView1.Nodes(3).Child.Index

        If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
          If TreeView1.Nodes(i).Image <> 9 Then
            
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID2.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 47
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.Font.ColorIndex = 2
            MyAppID2.Selection.MergeCells = True

            y = y + 1
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = "No. Componente"
            MyAppID2.Range("B" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = "Codigo"
            
            MyAppID2.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 50
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.Font.ColorIndex = 2
            MyAppID2.Selection.HorizontalAlignment = xlCenter
            MyAppID2.Range("B" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.MergeCells = True
            y = y + 2
            
            TextoImpreso = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            PathProyect = Mid(dlgOpenFile.FileName, 1, InStr(1, dlgOpenFile.FileName, dlgOpenFile.FileTitle) - 1)
            
            If InStr(1, TextoImpreso, ";") <> 0 _
               And IsNull(InStr(1, TextoImpreso, ";")) = False _
               And InStr(1, TextoImpreso, ";") <> 1 Then
                PathProyect = PathProyect & Trim(Mid(TextoImpreso, InStr(1, TextoImpreso, ";") + 1, Len(TextoImpreso) - InStr(1, TextoImpreso, ";")))
            Else
                PathProyect = PathProyect & Trim(TextoImpreso)
            End If
            ContadorFuente = 0
            FileOpenRead = FreeFile
            Open PathProyect For Input As #FileOpenRead
            Do While Not EOF(FileOpenRead)
            DoEvents
            Line Input #FileOpenRead, StringLectura
            PrincipalKey2 = Trim(Replace(Replace(Trim(StringLectura), Chr(34), " "), "'", " "))
                If Trim(PrincipalKey2) <> "" And Len(Trim(PrincipalKey2)) > 1 Then
                    Do While Not Mid(PrincipalKey2, Len(PrincipalKey2), 1) <> "_"
                    Line Input #FileOpenRead, StringLectura
                    PrincipalKey2 = PrincipalKey2 & vbLf & Trim(Replace(Replace(Trim(StringLectura), Chr(34), " "), "'", " "))
                    Loop
                    
                     OpenConnectionMDB Con1, App.Path & "\ParamDocumentation.mdb"
                       VerifycationOtherOBJ PrincipalKey2, DescriptionKey2, LocateDate2, Con1
                     CloseConnectionMDB Con1
                     If LocateDate2 = True Then
                       ContadorFuente = ContadorFuente + 1
                       MyAppID2.Range("A" & (7 + y)).Select
                       MyAppID2.Selection.FormulaR1C1 = ContadorFuente
                       MyAppID2.Range("B" & (7 + y)).Select
                       MyAppID2.Selection.FormulaR1C1 = DescriptionKey2 & ": " & Chr(10) & PrincipalKey2
                       MyAppID2.Range("B" & (7 + y) & ":D" & (7 + y)).Select
                       MyAppID2.Selection.MergeCells = True
                       
                       y = y + 1
                       
                     End If
                End If
            '*****************************************************
            Loop
            Close #FileOpenRead
            
            
            
            End If
        End If

While i <> TreeView1.Nodes(3).Child.LastSibling.Index
 DoEvents
    If TreeView1.Nodes(i).Image <> 9 Then
            ContadorFuente = 0
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID2.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 47
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.Font.ColorIndex = 2
            MyAppID2.Selection.MergeCells = True
            
        y = y + 1
        MyAppID2.Range("A" & (7 + y)).Select
        MyAppID2.Selection.FormulaR1C1 = "No. Componente"
        MyAppID2.Range("B" & (7 + y)).Select
        MyAppID2.Selection.FormulaR1C1 = "Codigo"
        
        MyAppID2.Range("A" & (7 + y) & ":D" & (7 + y)).Select
        MyAppID2.Selection.Interior.ColorIndex = 50
        MyAppID2.Selection.Interior.Pattern = xlSolid
        MyAppID2.Selection.Font.Bold = True
        MyAppID2.Selection.Font.ColorIndex = 2
        MyAppID2.Selection.HorizontalAlignment = xlCenter
        MyAppID2.Range("B" & (7 + y) & ":D" & (7 + y)).Select
        MyAppID2.Selection.MergeCells = True
        y = y + 2
                    TextoImpreso = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            PathProyect = Mid(dlgOpenFile.FileName, 1, InStr(1, dlgOpenFile.FileName, dlgOpenFile.FileTitle) - 1)
            
            If InStr(1, TextoImpreso, ";") <> 0 _
               And IsNull(InStr(1, TextoImpreso, ";")) = False _
               And InStr(1, TextoImpreso, ";") <> 1 Then
                PathProyect = PathProyect & Trim(Mid(TextoImpreso, InStr(1, TextoImpreso, ";") + 1, Len(TextoImpreso) - InStr(1, TextoImpreso, ";")))
            Else
                PathProyect = PathProyect & Trim(TextoImpreso)
                
            End If
            ContadorFuente = 0
            FileOpenRead = FreeFile
            Open PathProyect For Input As #FileOpenRead
            Do While Not EOF(FileOpenRead)
            DoEvents
            Line Input #FileOpenRead, StringLectura
            PrincipalKey2 = Trim(Replace(Replace(Trim(StringLectura), Chr(34), " "), "'", " "))
                If Trim(PrincipalKey2) <> "" And Len(Trim(PrincipalKey2)) > 1 Then
                    Do While Not Mid(PrincipalKey2, Len(PrincipalKey2), 1) <> "_"
                    Line Input #FileOpenRead, StringLectura
                    PrincipalKey2 = PrincipalKey2 & vbLf & Trim(Replace(Replace(Trim(StringLectura), Chr(34), " "), "'", " "))
                    Loop
                    
                     OpenConnectionMDB Con1, App.Path & "\ParamDocumentation.mdb"
                       VerifycationOtherOBJ PrincipalKey2, DescriptionKey2, LocateDate2, Con1
                     CloseConnectionMDB Con1
                     If LocateDate2 = True Then
                       ContadorFuente = ContadorFuente + 1
                       MyAppID2.Range("A" & (7 + y)).Select
                       MyAppID2.Selection.FormulaR1C1 = ContadorFuente
                       MyAppID2.Range("B" & (7 + y)).Select
                       MyAppID2.Selection.FormulaR1C1 = DescriptionKey2 & ": " & Chr(10) & PrincipalKey2
                       MyAppID2.Range("B" & (7 + y) & ":D" & (7 + y)).Select
                       MyAppID2.Selection.MergeCells = True
                       
                       y = y + 1
                       
                     End If
                End If
            '*****************************************************
            Loop
            Close #FileOpenRead

    End If
    
    i = TreeView1.Nodes(i).Next.Index
    
    
    If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
        If TreeView1.Nodes(i).Image <> 9 Then
            ContadorFuente = 0
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID2.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 47
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.Font.ColorIndex = 2
            MyAppID2.Selection.MergeCells = True
            y = y + 1
            MyAppID2.Range("A" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = "No. Componente"
            MyAppID2.Range("B" & (7 + y)).Select
            MyAppID2.Selection.FormulaR1C1 = "Codigo"
            MyAppID2.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.Interior.ColorIndex = 50
            MyAppID2.Selection.Interior.Pattern = xlSolid
            MyAppID2.Selection.Font.Bold = True
            MyAppID2.Selection.Font.ColorIndex = 2
            MyAppID2.Selection.HorizontalAlignment = xlCenter
            MyAppID2.Range("B" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID2.Selection.MergeCells = True
            y = y + 2
                        TextoImpreso = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            PathProyect = Mid(dlgOpenFile.FileName, 1, InStr(1, dlgOpenFile.FileName, dlgOpenFile.FileTitle) - 1)
            
            If InStr(1, TextoImpreso, ";") <> 0 _
               And IsNull(InStr(1, TextoImpreso, ";")) = False _
               And InStr(1, TextoImpreso, ";") <> 1 Then
                PathProyect = PathProyect & Trim(Mid(TextoImpreso, InStr(1, TextoImpreso, ";") + 1, Len(TextoImpreso) - InStr(1, TextoImpreso, ";")))
            Else
                PathProyect = PathProyect & Trim(TextoImpreso)
            End If
            ContadorFuente = 0
            FileOpenRead = FreeFile
            Open PathProyect For Input As #FileOpenRead
            Do While Not EOF(FileOpenRead)
            DoEvents
            Line Input #FileOpenRead, StringLectura
            PrincipalKey2 = Trim(Replace(Replace(Trim(StringLectura), Chr(34), " "), "'", " "))
                If Trim(PrincipalKey2) <> "" And Len(Trim(PrincipalKey2)) > 1 Then
                    Do While Not Mid(PrincipalKey2, Len(PrincipalKey2), 1) <> "_"
                    Line Input #FileOpenRead, StringLectura
                    PrincipalKey2 = PrincipalKey2 & vbLf & Trim(Replace(Replace(Trim(StringLectura), Chr(34), " "), "'", " "))
                    Loop
                    
                     OpenConnectionMDB Con1, App.Path & "\ParamDocumentation.mdb"
                       VerifycationOtherOBJ PrincipalKey2, DescriptionKey2, LocateDate2, Con1
                     CloseConnectionMDB Con1
                     If LocateDate2 = True Then
                       ContadorFuente = ContadorFuente + 1
                       MyAppID2.Range("A" & (7 + y)).Select
                       MyAppID2.Selection.FormulaR1C1 = ContadorFuente
                       MyAppID2.Range("B" & (7 + y)).Select
                       MyAppID2.Selection.FormulaR1C1 = DescriptionKey2 & ": " & Chr(10) & PrincipalKey2
                       MyAppID2.Range("B" & (7 + y) & ":D" & (7 + y)).Select
                       MyAppID2.Selection.MergeCells = True
                       
                       y = y + 1
                       
                     End If
                End If
            '*****************************************************
            Loop
            Close #FileOpenRead

        End If
    End If
Wend

End If
XLS = True



End Sub

Private Sub SetVarsPublics(ByRef MyAppID As Object, ByRef MyWorkBook, ByRef MyWorkShet)
'// Seteo y traslado de valores a Variables tipo objeto que aperturan documento EXCEL
Set MyAppID99 = MyAppID
Set MyWorkBook99 = MyWorkBook
Set MyWorkShet99 = MyWorkShet

End Sub

Private Sub Command1_Click()
'// Proceso Inicial de Generación de Archivo de Informe en EXCEL.
'// Este proceso genera el Sheet 1 y 2 del proyecto a documentar.
Dim MyAppID As Object
Dim MyWorkBook
Dim MyWorkShet
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim y As Integer
Command1.Enabled = False
Command2.Enabled = False
'******************************************************
'Construccion de EXCEL File
'******************************************************
Set MyAppID = CreateObject("EXCEL.Application")
Set MyWorkBook = MyAppID.WorkBooks.Add
Set MyWorkShet = MyWorkBook.WorkSheets()

MyWorkShet(1).Name = "Doc. Visual Basic Proyect"
'******************************************************
' Encabezado del Informe
'******************************************************
MyWorkShet(1).Select
MyAppID.Range("A1:D2").Select
MyAppID.Selection.Interior.ColorIndex = 46
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.HorizontalAlignment = xlCenter
MyAppID.Selection.Font.Size = 12
MyAppID.Range("A1").FormulaR1C1 = "Corporación GYT Continental" & vbLf & "Soluciones Tecnológicas"
MyAppID.Columns("A:C").Select
MyAppID.Selection.ColumnWidth = 30.71
MyAppID.Rows("1:2").Select
MyAppID.Selection.RowHeight = 20
MyAppID.Range("A4").Select
MyAppID.Selection.FormulaR1C1 = "Formato de Descripción de Sistema"
MyAppID.Range("A4:D4").Select
MyAppID.Selection.Interior.ColorIndex = 6
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.Font.Bold = True

MyAppID.Range("A6").Select
MyAppID.Selection.Interior.ColorIndex = 15
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.Font.Bold = True
MyAppID.Selection.FormulaR1C1 = "Proyecto:"
MyAppID.Range("B6").Select
MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(1).Text
'******************************************************
' Extraccion de Propiedades del Informe
'******************************************************
y = 1
If TreeView1.Nodes(2).Children > 0 Then

      i = TreeView1.Nodes(2).Child.Index
      
      If i = TreeView1.Nodes(2).Child.LastSibling.Index Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 15
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            y = y + 1
      End If
      
      While i <> TreeView1.Nodes(2).Child.LastSibling.Index
         DoEvents
         MyAppID.Range("A" & (7 + y)).Select
         MyAppID.Selection.Interior.ColorIndex = 15
         MyAppID.Selection.Interior.Pattern = xlSolid
         MyAppID.Selection.Font.Bold = True
         MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
         MyAppID.Range("B" & (7 + y)).Select
         MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)

         i = TreeView1.Nodes(i).Next.Index
         y = y + 1
         If i = TreeView1.Nodes(2).Child.LastSibling.Index Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 15
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            y = y + 1
         End If
      Wend
      
End If
'******************************************************
' Extraccion de Modulos y Pantallas
'******************************************************
If TreeView1.Nodes(3).Children > 0 Then
y = y + 3

MyAppID.Range("A" & (7 + y - 2)).Select
MyAppID.Selection.FormulaR1C1 = "Componentes del Proyecto:"
MyAppID.Range("A" & (7 + y - 2) & ":D" & (7 + y - 2)).Select
MyAppID.Selection.Interior.ColorIndex = 6
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.Font.Bold = True

MyAppID.Range("A" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "No. Componente"
MyAppID.Range("B" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "Tipo de OBJ"
MyAppID.Range("C" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "Nombre OBJ"
MyAppID.Range("D" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "Lenguaje"

MyAppID.Range("A" & (7 + y - 1) & ":D" & (7 + y - 1)).Select
MyAppID.Selection.Interior.ColorIndex = 15
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.Font.Bold = True
MyAppID.Selection.HorizontalAlignment = xlCenter


i = TreeView1.Nodes(3).Child.Index

        If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
          If TreeView1.Nodes(i).Image <> 9 Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "1"
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":") - 1)
            MyAppID.Range("C" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID.Range("D" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Visual Basic"
            MyAppID.Range("A" & (7 + y) & ":A" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            MyAppID.Range("D" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter


            y = y + 1
            End If
        End If

While i <> TreeView1.Nodes(3).Child.LastSibling.Index
    DoEvents
    If TreeView1.Nodes(i).Image <> 9 Then
        MyAppID.Range("A" & (7 + y)).Select
        If i = TreeView1.Nodes(3).Child.Index Then
        MyAppID.Selection.FormulaR1C1 = "1"
        Else
        MyAppID.Selection.Formula = "=if(isnumber(A" & (7 + y - 1) & "),A" & (7 + y - 1) & " + 1, 1)"
        End If
        MyAppID.Range("B" & (7 + y)).Select
        MyAppID.Selection.Font.Bold = True
        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":") - 1)
        MyAppID.Range("C" & (7 + y)).Select
        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
        MyAppID.Range("D" & (7 + y)).Select
        MyAppID.Selection.FormulaR1C1 = "Visual Basic"
        MyAppID.Range("A" & (7 + y) & ":A" & (7 + y)).Select
        MyAppID.Selection.HorizontalAlignment = xlCenter
        MyAppID.Range("D" & (7 + y) & ":D" & (7 + y)).Select
        MyAppID.Selection.HorizontalAlignment = xlCenter
        y = y + 1
    End If
    
    i = TreeView1.Nodes(i).Next.Index
    
    
    If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
        If TreeView1.Nodes(i).Image <> 9 Then
            MyAppID.Range("A" & (7 + y)).Select
            If i = TreeView1.Nodes(3).Child.Index Then
            MyAppID.Selection.FormulaR1C1 = "1"
            Else
            MyAppID.Selection.Formula = "=if(isnumber(A" & (7 + y - 1) & "),A" & (7 + y - 1) & " + 1, 1)"
            End If
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":") - 1)
            MyAppID.Range("C" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID.Range("D" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Visual Basic"
            MyAppID.Range("A" & (7 + y) & ":A" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            MyAppID.Range("D" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            y = y + 1
        End If
    End If
Wend

End If
      
      
If TreeView1.Nodes(3).Children > 0 Then
y = y + 4

MyAppID.Range("A" & (7 + y - 2)).Select
MyAppID.Selection.FormulaR1C1 = "Archivos Relacionados:"
MyAppID.Range("A" & (7 + y - 2) & ":D" & (7 + y - 2)).Select
MyAppID.Selection.Interior.ColorIndex = 6
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.Font.Bold = True

MyAppID.Range("A" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "No. Archivo"
MyAppID.Range("B" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "Nombre de Archivo"
MyAppID.Range("C" & (7 + y - 1)).Select
MyAppID.Selection.FormulaR1C1 = "Extención"

MyAppID.Range("A" & (7 + y - 1) & ":C" & (7 + y - 1)).Select
MyAppID.Selection.Interior.ColorIndex = 15
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.Font.Bold = True
MyAppID.Selection.HorizontalAlignment = xlCenter


i = TreeView1.Nodes(3).Child.Index

        If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
          If TreeView1.Nodes(i).Image = 9 Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "1"
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID.Range("C" & (7 + y)).Select
            If InStr(1, Mid(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1), Len(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)) - 3, 4), ".") = 0 Then
            MyAppID.Selection.FormulaR1C1 = "Not Extencion"
            Else
            MyAppID.Selection.FormulaR1C1 = Mid(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1), Len(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)) - 3, 4)
            End If
            MyAppID.Range("A" & (7 + y) & ":A" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            MyAppID.Range("C" & (7 + y) & ":C" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            y = y + 1
            End If
        End If

While i <> TreeView1.Nodes(3).Child.LastSibling.Index
    DoEvents
    If TreeView1.Nodes(i).Image = 9 Then
        MyAppID.Range("A" & (7 + y)).Select
        If i = TreeView1.Nodes(3).Child.Index Then
        MyAppID.Selection.FormulaR1C1 = "1"
        Else
        MyAppID.Selection.Formula = "=if(isnumber(A" & (7 + y - 1) & "),A" & (7 + y - 1) & " + 1, 1)"
        End If
        MyAppID.Range("B" & (7 + y)).Select
        MyAppID.Selection.Font.Bold = True
        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
        MyAppID.Range("C" & (7 + y)).Select
        If InStr(1, Mid(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1), Len(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)) - 3, 4), ".") = 0 Then
        MyAppID.Selection.FormulaR1C1 = "Not Extencion"
        Else
        MyAppID.Selection.FormulaR1C1 = Mid(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1), Len(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)) - 3, 4)
        End If
        MyAppID.Range("A" & (7 + y) & ":A" & (7 + y)).Select
        MyAppID.Selection.HorizontalAlignment = xlCenter
        MyAppID.Range("C" & (7 + y) & ":C" & (7 + y)).Select
        MyAppID.Selection.HorizontalAlignment = xlCenter
        y = y + 1
    End If
    
    i = TreeView1.Nodes(i).Next.Index
    
    
    If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
        If TreeView1.Nodes(i).Image = 9 Then
            MyAppID.Range("A" & (7 + y)).Select
            If i = TreeView1.Nodes(3).Child.Index Then
            MyAppID.Selection.FormulaR1C1 = "1"
            Else
            MyAppID.Selection.Formula = "=if(isnumber(A" & (7 + y - 1) & "),A" & (7 + y - 1) & " + 1, 1)"
            End If
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            MyAppID.Range("C" & (7 + y)).Select
            If InStr(1, Mid(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1), Len(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)) - 3, 4), ".") = 0 Then
            MyAppID.Selection.FormulaR1C1 = "Not Extencion"
            Else
            MyAppID.Selection.FormulaR1C1 = Mid(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1), Len(Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)) - 3, 4)
            End If
            MyAppID.Range("A" & (7 + y) & ":A" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            MyAppID.Range("C" & (7 + y) & ":C" & (7 + y)).Select
            MyAppID.Selection.HorizontalAlignment = xlCenter
            y = y + 1
        End If
    End If
Wend

End If
'**************************************************************************************
'*********                    Segunda Hora de Documentación                  **********
'**************************************************************************************
MyWorkShet(2).Name = "Doc. Components"
MyWorkShet(2).Select
MyAppID.Range("A1:D2").Select
MyAppID.Selection.Interior.ColorIndex = 46
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.HorizontalAlignment = xlCenter
MyAppID.Selection.Font.Size = 12
MyAppID.Range("A1").FormulaR1C1 = "Corporación GYT Continental" & vbLf & "Soluciones Tecnológicas"
MyAppID.Columns("A:C").Select
MyAppID.Selection.ColumnWidth = 30.71
MyAppID.Rows("1:2").Select
MyAppID.Selection.RowHeight = 20
MyAppID.Range("A4").Select
MyAppID.Selection.FormulaR1C1 = "Formato de Descripción de Sistema"
MyAppID.Range("A4:D4").Select
MyAppID.Selection.Interior.ColorIndex = 6
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.Font.Bold = True

MyAppID.Range("A6").Select
MyAppID.Selection.Interior.ColorIndex = 15
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.Font.Bold = True
MyAppID.Selection.FormulaR1C1 = "Proyecto:"
MyAppID.Range("B6").Select
MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(1).Text
'******************************************************
' Extraccion de Propiedades del Informe
'******************************************************
y = 1
If TreeView1.Nodes(2).Children > 0 Then

      i = TreeView1.Nodes(2).Child.Index
      
      If i = TreeView1.Nodes(2).Child.LastSibling.Index Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 15
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            y = y + 1
      End If
      
      While i <> TreeView1.Nodes(2).Child.LastSibling.Index
         DoEvents
         MyAppID.Range("A" & (7 + y)).Select
         MyAppID.Selection.Interior.ColorIndex = 15
         MyAppID.Selection.Interior.Pattern = xlSolid
         MyAppID.Selection.Font.Bold = True
         MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
         MyAppID.Range("B" & (7 + y)).Select
         MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)

         i = TreeView1.Nodes(i).Next.Index
         y = y + 1
         If i = TreeView1.Nodes(2).Child.LastSibling.Index Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 15
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, 1, InStr(1, TreeView1.Nodes(i).Text, ":"))
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(i).Text, InStr(1, TreeView1.Nodes(i).Text, ":") + 1, Len(TreeView1.Nodes(i).Text) - InStr(1, TreeView1.Nodes(i).Text, ":") + 1)
            y = y + 1
         End If
      Wend
      
End If
      
      
If TreeView1.Nodes(3).Children > 0 Then
y = y + 3

MyAppID.Range("A" & (7 + y - 2)).Select
MyAppID.Selection.FormulaR1C1 = "Sub Componentes del Proyecto:"
MyAppID.Range("A" & (7 + y - 2) & ":D" & (7 + y - 2)).Select
MyAppID.Selection.Interior.ColorIndex = 6
MyAppID.Selection.Interior.Pattern = xlSolid
MyAppID.Selection.MergeCells = True
MyAppID.Selection.Font.Bold = True



i = TreeView1.Nodes(3).Child.Index

        If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
          If TreeView1.Nodes(i).Image <> 9 Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(i).Text
            MyAppID.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 37
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.MergeCells = True
            MyAppID.Selection.Font.Bold = True
            y = y + 1
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Descripción"
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Nombre de Componente"
            MyAppID.Range("C" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Propiedades Especificadas"
            MyAppID.Range("A" & (7 + y) & ":C" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 50
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.Font.ColorIndex = 2
            MyAppID.Selection.HorizontalAlignment = xlCenter
            y = y + 1
                        
                If TreeView1.Nodes(i).Children = 0 Then
                 y = y + 1
                Else
                  j = TreeView1.Nodes(i).Child.Index
                  If j = TreeView1.Nodes(i).Child.LastSibling.Index Then
                     MyAppID.Range("A" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                     MyAppID.Range("B" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                     MyAppID.Range("C" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                     MyAppID.Selection.HorizontalAlignment = xlCenter
                     y = y + 1
                     If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                  End If
                  
                  While j <> TreeView1.Nodes(i).Child.LastSibling.Index
                        DoEvents
                        MyAppID.Range("A" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                        MyAppID.Selection.HorizontalAlignment = xlCenter

                        y = y + 1
                        If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                             DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                        j = TreeView1.Nodes(j).Next.Index
                        If j = TreeView1.Nodes(i).Child.LastSibling.Index Then
                            MyAppID.Range("A" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                            MyAppID.Selection.HorizontalAlignment = xlCenter

                            y = y + 1
                            If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                        End If
                  Wend
                
                End If
            End If
        End If

While i <> TreeView1.Nodes(3).Child.LastSibling.Index
    If TreeView1.Nodes(i).Image <> 9 Then
        MyAppID.Range("A" & (7 + y)).Select
        MyAppID.Selection.Font.Bold = True
        MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(i).Text
        MyAppID.Range("A" & (7 + y) & ":D" & (7 + y)).Select
        MyAppID.Selection.Interior.ColorIndex = 37
        MyAppID.Selection.Interior.Pattern = xlSolid
        MyAppID.Selection.MergeCells = True
        MyAppID.Selection.Font.Bold = True

        y = y + 1
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Descripción"
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Nombre de Componente"
            MyAppID.Range("C" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Propiedades Especificadas"
            MyAppID.Range("A" & (7 + y) & ":C" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 50
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.Font.ColorIndex = 2
            MyAppID.Selection.HorizontalAlignment = xlCenter
            y = y + 1
               If TreeView1.Nodes(i).Children = 0 Then
                 y = y + 1
               Else
                  j = TreeView1.Nodes(i).Child.Index
                  If j = TreeView1.Nodes(i).Child.LastSibling.Index Then
                     MyAppID.Range("A" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                     MyAppID.Range("B" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                     MyAppID.Range("C" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                     MyAppID.Selection.HorizontalAlignment = xlCenter

                     y = y + 1
                     If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                  End If
                  While j <> TreeView1.Nodes(i).Child.LastSibling.Index
                       DoEvents
                        MyAppID.Range("A" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                        j = TreeView1.Nodes(j).Next.Index
                        If j = TreeView1.Nodes(i).Child.LastSibling.Index Then
                            MyAppID.Range("A" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                            MyAppID.Selection.HorizontalAlignment = xlCenter
 
                            y = y + 1
                            If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                        End If
                  Wend
                
                End If

    End If
    
    i = TreeView1.Nodes(i).Next.Index
    
    
    If i = TreeView1.Nodes(3).Child.LastSibling.Index Then
        If TreeView1.Nodes(i).Image <> 9 Then
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(i).Text
            MyAppID.Range("A" & (7 + y) & ":D" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 37
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.MergeCells = True
            MyAppID.Selection.Font.Bold = True

            y = y + 1
            MyAppID.Range("A" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Descripción"
            MyAppID.Range("B" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Nombre de Componente"
            MyAppID.Range("C" & (7 + y)).Select
            MyAppID.Selection.FormulaR1C1 = "Propiedades Especificadas"
            MyAppID.Range("A" & (7 + y) & ":C" & (7 + y)).Select
            MyAppID.Selection.Interior.ColorIndex = 50
            MyAppID.Selection.Interior.Pattern = xlSolid
            MyAppID.Selection.Font.Bold = True
            MyAppID.Selection.Font.ColorIndex = 2
            MyAppID.Selection.HorizontalAlignment = xlCenter
            y = y + 1
            
                If TreeView1.Nodes(i).Children = 0 Then
                y = y + 1
                Else
                  j = TreeView1.Nodes(i).Child.Index
                  If j = TreeView1.Nodes(i).Child.LastSibling.Index Then
                     MyAppID.Range("A" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                     MyAppID.Range("B" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                     MyAppID.Range("C" & (7 + y)).Select
                     MyAppID.Selection.Font.Bold = True
                     MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                     MyAppID.Selection.HorizontalAlignment = xlCenter

                     y = y + 1
                     If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                  End If
                  While j <> TreeView1.Nodes(i).Child.LastSibling.Index
                        DoEvents
                        MyAppID.Range("A" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                        MyAppID.Selection.HorizontalAlignment = xlCenter

                        y = y + 1
                        If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                        j = TreeView1.Nodes(j).Next.Index
                        If j = TreeView1.Nodes(i).Child.LastSibling.Index Then
                            MyAppID.Range("A" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, 1, InStr(1, TreeView1.Nodes(j).Text, ":") - 1)
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(j).Text, InStr(1, TreeView1.Nodes(j).Text, ":") + 1, Len(TreeView1.Nodes(j).Text) - InStr(1, TreeView1.Nodes(j).Text, ":"))
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = True
                            MyAppID.Selection.FormulaR1C1 = TreeView1.Nodes(j).Children
                            MyAppID.Selection.HorizontalAlignment = xlCenter

                            y = y + 1
                            If TreeView1.Nodes(j).Children = 0 Then
                       y = y + 1
                     Else
                        MyAppID.Range("B" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Propiedad"
                        MyAppID.Range("C" & (7 + y)).Select
                        MyAppID.Selection.FormulaR1C1 = "Valor"
                        MyAppID.Range("B" & (7 + y) & ":C" & (7 + y)).Select
                        MyAppID.Selection.Interior.ColorIndex = 11
                        MyAppID.Selection.Interior.Pattern = xlSolid
                        MyAppID.Selection.Font.Bold = True
                        MyAppID.Selection.Font.ColorIndex = 2
                        MyAppID.Selection.HorizontalAlignment = xlCenter
                        y = y + 1
                        k = TreeView1.Nodes(j).Child.Index
                        If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                        End If
                        
                        While k <> TreeView1.Nodes(j).Child.LastSibling.Index
                            DoEvents
                            MyAppID.Range("B" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                            MyAppID.Range("C" & (7 + y)).Select
                            MyAppID.Selection.Font.Bold = False
                            MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                            y = y + 1
                            k = TreeView1.Nodes(k).Next.Index
                            If k = TreeView1.Nodes(j).Child.LastSibling.Index Then
                                MyAppID.Range("B" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, 1, InStr(1, TreeView1.Nodes(k).Text, ":") - 1)
                                MyAppID.Range("C" & (7 + y)).Select
                                MyAppID.Selection.Font.Bold = False
                                MyAppID.Selection.FormulaR1C1 = Mid(TreeView1.Nodes(k).Text, InStr(1, TreeView1.Nodes(k).Text, ":") + 1, Len(TreeView1.Nodes(k).Text) - InStr(1, TreeView1.Nodes(k).Text, ":"))
                                y = y + 1
                            End If
                        Wend
                        
                     End If
                        End If
                  Wend
                
                End If

        End If
    End If
Wend

End If

Call ProceesToFile(MyAppID, MyWorkBook, MyWorkShet)
MyAppID.Visible = True
Call SetVarsPublics(MyAppID, MyWorkBook, MyWorkShet)
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
'// Proceso de Carga de información de proyecto VB a Pantalla en Arbol de Nodos.
'// El arbol de nodos no es más que un arreglo de memoria dinámica en el cual se muestra graficamente toda la estructura fisica del proyecto, colocando la información previamente parametrizada
On Error Resume Next
Dim OpenFileVBP As Integer
Dim Nod1
Dim StringLeido As String
Dim StringLeidoSubObject As String
Dim LlaveConsulta As String
Dim Con1 As ADODB.Connection
Dim DescriptObjectResult1 As String
Dim CargaDetalle1 As Boolean
Dim DetectoBegin As Boolean
Dim IsProperty1 As Boolean
Dim ImageAsociate1 As Integer
Dim LocateDate1 As Boolean
Dim CountObject As Integer
Dim NameFilesToRead As String
Dim ReadToObject As Integer
Dim ContadordeBegins As Integer
Dim PrincipalKeySubObject As String
Dim DescripcionObjeto As String
Dim EncontroRegistro As Boolean
Dim UltimaLlave As String
Dim DescrypcionPropiedad As String
Dim EncontroPropiedad As Boolean
Dim ErrCount As Integer
Dim ErrCount2 As Integer

    Command2.Enabled = False
    Command1.Enabled = False
         CountObject = 0
         TreeView1.Enabled = False
         dlgOpenFile.CancelError = True ' Causes a trappable error to occur when the user hits the 'Cancel' button
         dlgOpenFile.DialogTitle = "Seleccione archivo a Procesar"
         dlgOpenFile.FileName = ""
         dlgOpenFile.Filter = "Visual Basic Proyect (*.vbp)|*.vbp"
         dlgOpenFile.FilterIndex = 1
         dlgOpenFile.Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly
         dlgOpenFile.ShowOpen
         If Err = cdlCancel Then  ' 'Cancel' button was hit
        'Do nothing
          MsgBox "Proceso Cancelado", vbInformation, "Carga de Proyecto Cancelada"
          TreeView1.Enabled = False
          On Error GoTo 0
          Exit Sub
         Else
            On Error GoTo 0
            TreeView1.Nodes.Clear
            TreeView1.LineStyle = tvwRootLines
            TreeView1.Checkboxes = False
            TreeView1.ImageList = ImageList1
            TreeView1.Sorted = True
         
            Nod1 = TreeView1.Nodes.Add(, , "Root", dlgOpenFile.FileTitle, 1, 1)
            Nod1 = TreeView1.Nodes.Add("Root", tvwChild, "RootProperty", "Propiedades del Proyecto", 5, 5)
            Nod1 = TreeView1.Nodes.Add("Root", tvwChild, "RootObjects", "Objetos del Proyecto", 6, 6)
            
            Nod1 = TreeView1.Nodes.Add("RootProperty", tvwChild, "DateModProperty", "Ultima Modificación: " & DateToFile(dlgOpenFile.FileName, Text1, FechaModificacion), 4, 4)
            Nod1 = TreeView1.Nodes.Add("RootProperty", tvwChild, "DateCreateProperty", "Fecha de Creación: " & DateToFile(dlgOpenFile.FileName, Text1, FechaCreacion), 4, 4)
            Nod1 = TreeView1.Nodes.Add("RootProperty", tvwChild, "DateAccesProperty", "Ultimo Acceso: " & DateToFile(dlgOpenFile.FileName, Text1, FechaUltimoAcceso), 4, 4)
            OpenFileVBP = FreeFile
            
            Open dlgOpenFile.FileName For Input As #OpenFileVBP
            
            Do While Not EOF(OpenFileVBP)
              Line Input #OpenFileVBP, StringLeido
              StringLeido = Trim(StringLeido)
              DoEvents
              If InStr(1, StringLeido, "=") <> 0 And _
                 IsNull(InStr(1, StringLeido, "=")) = False And _
                 InStr(1, StringLeido, "=") <> 1 Then
              
              LlaveConsulta = Mid(StringLeido, 1, InStr(1, StringLeido, "="))
              'ParamDocumentation.mdb
              OpenConnectionMDB Con1, App.Path & "\ParamDocumentation.mdb"
                  VerificationKey LlaveConsulta, DescriptObjectResult1, CargaDetalle1, IsProperty1, ImageAsociate1, LocateDate1, Con1
              CloseConnectionMDB Con1
              
              
              If LocateDate1 = True Then
                       If IsProperty1 = True Then
                         Nod1 = TreeView1.Nodes.Add("RootProperty", tvwChild, LlaveConsulta, Trim(DescriptObjectResult1) & " " & Mid(StringLeido, Len(LlaveConsulta) + 1, Len(StringLeido) - Len(LlaveConsulta)), ImageAsociate1, ImageAsociate1)
                         
                       Else
                         CountObject = CountObject + 1
                         Nod1 = TreeView1.Nodes.Add("RootObjects", tvwChild, LlaveConsulta & CountObject, Trim(DescriptObjectResult1) & " " & Mid(StringLeido, Len(LlaveConsulta) + 1, Len(StringLeido) - Len(LlaveConsulta)), ImageAsociate1, ImageAsociate1)
                         
                         If CargaDetalle1 = True Then
                           'Buscar nombre del form modulo o clase que pueda tener detalle
                               NameFilesToRead = ""
                                    If InStr(1, StringLeido, ";") <> 0 And _
                                       IsNull(InStr(1, StringLeido, ";")) = False And _
                                       InStr(1, StringLeido, ";") <> 1 Then
                                        NameFilesToRead = Trim(Mid(StringLeido, InStr(1, StringLeido, ";") + 1, Len(StringLeido) - InStr(1, StringLeido, ";")))
                                    Else
                                        NameFilesToRead = Trim(Mid(StringLeido, InStr(1, StringLeido, "=") + 1, Len(StringLeido) - InStr(1, StringLeido, "=")))
                                    End If
                                    If NameFilesToRead <> "" Then
                                       NameFilesToRead = Mid(dlgOpenFile.FileName, 1, InStr(1, dlgOpenFile.FileName, dlgOpenFile.FileTitle) - 1) & NameFilesToRead
                                       StringLeidoSubObject = ""
                                       DetectoBegin = False
                                       ContadordeBegins = 0
                                       ErrCount = 1
                                       ReadToObject = FreeFile
                                       Open NameFilesToRead For Input As #ReadToObject
                                       Do While Not EOF(ReadToObject)
                                        DoEvents
                                        Line Input #ReadToObject, StringLeidoSubObject
                                        StringLeidoSubObject = Trim(StringLeidoSubObject)
                                        
                                        If StringLeidoSubObject <> "" And Len(StringLeidoSubObject) >= 3 Then
                                          If Mid(StringLeidoSubObject, 1, 3) = "End" And Len(Trim(StringLeidoSubObject)) = 3 Then
                                            ContadordeBegins = ContadordeBegins - 1
                                            UltimaLlave = ""
                                          Else
                                            If Mid(StringLeidoSubObject, 1, 6) = "Begin " Then
                                              ContadordeBegins = ContadordeBegins + 1
                                              If ContadordeBegins = 1 Then
                                                 DetectoBegin = True
                                              End If
                                              PrincipalKeySubObject = Mid(StringLeidoSubObject, 7, InStr(7, StringLeidoSubObject, " ") - 7)
                                              PrincipalKeySubObject = Trim(PrincipalKeySubObject)
                                              EncontroRegistro = False
                                              OpenConnectionMDB Con1, App.Path & "\ParamDocumentation.mdb"
                                              VerifycationSubKey PrincipalKeySubObject, LlaveConsulta, DescripcionObjeto, EncontroRegistro, Con1
                                              CloseConnectionMDB Con1
                                              If EncontroRegistro = True Then
                                                 UltimaLlave = LlaveConsulta & CountObject & "-" & PrincipalKeySubObject & "-" & Trim(Mid(StringLeidoSubObject, 6 + Len(PrincipalKeySubObject) + 1, Len(StringLeidoSubObject) - 6 + Len(PrincipalKeySubObject)))
                                                 On Error Resume Next
                                                 Nod1 = TreeView1.Nodes.Add(LlaveConsulta & CountObject, tvwChild, UltimaLlave, Trim(DescripcionObjeto) & " " & Mid(StringLeidoSubObject, 6 + Len(PrincipalKeySubObject) + 1, Len(StringLeidoSubObject) - 6 + Len(PrincipalKeySubObject)), 4, 4)
                                                 If Err.Number = 35602 Then
                                                    ErrCount = ErrCount + 1
                                                    UltimaLlave = UltimaLlave & ErrCount
                                                    Nod1 = TreeView1.Nodes.Add(LlaveConsulta & CountObject, tvwChild, UltimaLlave, Trim(DescripcionObjeto) & " " & Mid(StringLeidoSubObject, 6 + Len(PrincipalKeySubObject) + 1, Len(StringLeidoSubObject) - 6 + Len(PrincipalKeySubObject)), 4, 4)
                                                 End If
                                                 On Error GoTo 0
                                                 ErrCount2 = 1
                                              End If
                                            Else
                                                   
                                             If UltimaLlave <> "" Then
                                                 If InStr(1, StringLeidoSubObject, "=") <> 0 And _
                                                    IsNull(InStr(1, StringLeidoSubObject, "=")) = False And _
                                                    InStr(1, StringLeidoSubObject, "=") <> 1 Then
                                                    EncontroPropiedad = False
                                                    OpenConnectionMDB Con1, App.Path & "\ParamDocumentation.mdb"
                                                    VerifycationPropertyKey Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "=") - 1)), LlaveConsulta, PrincipalKeySubObject, DescrypcionPropiedad, EncontroPropiedad, Con1
                                                    CloseConnectionMDB Con1
                                                    If EncontroPropiedad = True Then
                                                    On Error Resume Next
                                                    Nod1 = TreeView1.Nodes.Add(UltimaLlave, tvwChild, UltimaLlave & Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "=") - 1)), _
                                                           DescrypcionPropiedad & ": " & Mid(StringLeidoSubObject, Len(Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "=")))) + 1, _
                                                           Len(StringLeidoSubObject) - Len(Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "="))))), 7, 7)
                                                    If Err.Number = 35602 Then
                                                        ErrCount2 = ErrCount2 + 1
                                                        Nod1 = TreeView1.Nodes.Add(UltimaLlave, tvwChild, UltimaLlave & Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "=") - 1)) & ErrCount2, _
                                                               DescrypcionPropiedad & ": " & Mid(StringLeidoSubObject, Len(Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "=")))) + 1, _
                                                               Len(StringLeidoSubObject) - Len(Trim(Mid(StringLeidoSubObject, 1, InStr(1, StringLeidoSubObject, "="))))), 7, 7)
                                                    
                                                    End If
                                                    On Error GoTo 0
                                                    
                                                    End If
                                                 End If
                                             End If
                                                   
                                            End If
                                          End If
                                          
                                          
                                          
                                          
                                          If ContadordeBegins = 0 And DetectoBegin = True Then
                                            Exit Do
                                          End If
                                        End If
                                       Loop
                                       Close #ReadToObject
                                    End If
                         End If
                       End If
                       
              End If
              
              End If
            Loop
            
            Close #OpenFileVBP
         TreeView1.Refresh
         TreeView1.Enabled = True
         
             If XLS = True Then
                On Error Resume Next
                MyAppID99.Application.Quit
                If Err = 91 Then
                  XLS = False
                End If
                On Error GoTo 0
                Set MyAppID99 = Nothing
                Set MyWorkBook99 = Nothing
                Set MyWorkShet99 = Nothing
                XLS = False
            End If
            XLS = False
         Command1.Enabled = True
         End If
    Command2.Enabled = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
'// Proceso de Cierre de Aplicativo
'// Se realiza una verificación Simple de si existe aperturado algún Proceso Excel de tal forma que no se quede abierto dicho proceso.
        If XLS = True Then
            On Error Resume Next
            MyAppID99.Application.Quit
            If Err = 91 Then
              Exit Sub
            End If
            On Error GoTo 0
            Set MyAppID99 = Nothing
            Set MyWorkBook99 = Nothing
            Set MyWorkShet99 = Nothing
            XLS = False
        End If

End Sub
