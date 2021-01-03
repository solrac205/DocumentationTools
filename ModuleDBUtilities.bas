Attribute VB_Name = "ModuleDBUtilities"


Public Enum DatosDeArchivo
 '// Colección para identificación del dato de Fecha de Archivo que queremos consultar
 '// del archivo consultado en funcion DateToFile
     FechaCreacion = 1
     FechaModificacion = 2
     FechaUltimoAcceso = 3
End Enum

Public Sub OpenConnectionMDB(ConnectionToOpen As ADODB.Connection, DBName As String)
'// Establecer conexión a base de datos en Access Cualquiera que esta sea
'// Solicitando unicamente un Objeto de Conexión y el nombre del Archivo Extencion <.MDB>
Set ConnectionToOpen = New ADODB.Connection
ConnectionToOpen.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ";Persist Security Info=False"
ConnectionToOpen.Open

End Sub


Public Function DateToFile(ByRef FileToEvaluate As String, ByRef TextObject As TextBox, ByRef DateToConsult As DatosDeArchivo) As String
'// Extracción de datos de fechas de un archivo cualquiera que este sea, utilizando tres parametros de entrada y dandonos como resultado un string.
'// Los Parametros de entrada son: Archivo a Evaluar, Objeto Tipo TextBox y tipo de Fecha de consulta
On Error GoTo Err_DateToFile
Dim FileToAcces
Dim File
Dim Drive
Dim Folder
Dim SubFol As TextBox

TextObject.Text = ""
Set SubFol = TextObject

Set FileToAcces = CreateObject("Scripting.FileSystemObject")
Set Drive = FileToAcces.Drives(Mid(FileToEvaluate, 1, 1))
Set Folder = Drive.RootFolder
For i = 1 To Len(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5))
  If Mid(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5), i, 1) = "\" Then
    If Dir(Folder & "\" & SubFol.Text, vbDirectory) = "" Then
     SubFol.Text = SubFol.Text & "."
    End If
    Set Folder = Folder.SubFolders(SubFol.Text)
    SubFol.Text = ""
  Else
    SubFol.Text = SubFol.Text & Mid(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5), i, 1)
  End If
Next i
If Dir(Folder & "\" & SubFol.Text, vbDirectory) = "" Then
 SubFol.Text = SubFol.Text & "."
End If
Set Folder = Folder.SubFolders(SubFol.Text)
Set File = Folder.Files(FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate))
Select Case DateToConsult
Case 1
DateToFile = File.DateCreated
Case 2
DateToFile = File.DateLastModified
Case 3
DateToFile = File.DateLastAccessed
End Select

Exit_Err_DateToFile:
Exit Function

Err_DateToFile:
DateToFile = "File Or Path Not Found"
Resume Exit_Err_DateToFile
End Function
Public Sub CloseConnectionMDB(ConnectionOpen As ADODB.Connection)
'// Cierre de Conexión a Base de Datos Acces, Unico parametro de entrada = Conexión Abierta
ConnectionOpen.Close
Set ConnectionOpen = Nothing
End Sub
Public Sub VerifycationSubKey(ByRef PrincipalKey As String, ByRef SubKey As String, ByRef DescriptionKey As String, ByRef LocateDate As Boolean, ByRef ConnectionToBase As ADODB.Connection)
'// Proceso de Verificación de palabra Clave para determinar que Sub Objetos Estariamos utilizando al documentar.
Dim RSConsultaBase As ADODB.Recordset
Dim SQLString As String
'// Query de Ejecución Conformado por: SELECT DISTINCT * FROM SubObjectsorProperty WHERE NombreObjeto = '" & SubKey & "' AND ObjectName = '" & PrincipalKey & "';"
SQLString = "SELECT DISTINCT * FROM SubObjectsorProperty WHERE NombreObjeto = '" & SubKey & "' AND ObjectName = '" & PrincipalKey & "';"
Set RSConsultaBase = New ADODB.Recordset
With RSConsultaBase
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Open SQLString, ConnectionToBase, , , adCmdText

If .EOF = True And .BOF = True Then
   DescriptionKey = ""
   LocateDate = False
Else
   .MoveFirst
   DescriptionKey = !DescriptionObject
   LocateDate = True

End If
.Close
End With

Set RSConsultaBase = Nothing

End Sub

Public Sub VerifycationOtherOBJ(ByRef PrincipalKey As String, ByRef DescriptionKey As String, ByRef LocateDate As Boolean, ByRef ConnectionToBase As ADODB.Connection)
'// Verificación para determinar que fragmentos de codigo se estarian tomando en cuenta para documentar.
Dim RSConsultaBase As ADODB.Recordset
Dim SQLString As String
'// Query de Ejecución Conformado por:"SELECT DISTINCT TOP 1 * FROM OtherObject WHERE  '" & PrincipalKey & "' LIKE KeyToFind & " & "'%'" & ";"
SQLString = "SELECT DISTINCT TOP 1 * FROM OtherObject WHERE  '" & PrincipalKey & "' LIKE KeyToFind & " & "'%'" & ";"

Set RSConsultaBase = New ADODB.Recordset
With RSConsultaBase
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Open SQLString, ConnectionToBase, , , adCmdText

If .EOF = True And .BOF = True Then
   DescriptionKey = ""
   LocateDate = False
Else
   .MoveFirst
   DescriptionKey = !Descripcion
   LocateDate = True

End If
.Close
End With

Set RSConsultaBase = Nothing

End Sub

Public Sub VerifycationPropertyKey(ByRef PropertyKey As String, ByRef ObjectKey As String, ByRef SubKey As String, ByRef DescriptionKey As String, ByRef LocateDate As Boolean, ByRef ConnectionToBase As ADODB.Connection)
'// Verificación de Propiedades dentro de código para documentar.
Dim RSConsultaBase As ADODB.Recordset
Dim SQLString As String
'// Query de Ejecución Conformado por: "SELECT DISTINCT * " & "FROM DetailPropertyToEvaluate " & "WHERE NombreObjeto = '" & ObjectKey & "' " & "AND ObjectName = '" & SubKey & "' " & "AND PropertyToEvaluate = '" & PropertyKey & "';"
SQLString = "SELECT DISTINCT * " & _
            "FROM DetailPropertyToEvaluate " & _
            "WHERE NombreObjeto = '" & ObjectKey & "' " & _
            "AND ObjectName = '" & SubKey & "' " & _
            "AND PropertyToEvaluate = '" & PropertyKey & "';"
Set RSConsultaBase = New ADODB.Recordset
With RSConsultaBase
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Open SQLString, ConnectionToBase, , , adCmdText

If .EOF = True And .BOF = True Then
   DescriptionKey = ""
   LocateDate = False
Else
   .MoveFirst
   DescriptionKey = !DescrypcionToProperty
   LocateDate = True

End If
.Close
End With

Set RSConsultaBase = Nothing

End Sub
Public Sub VerificationKey(ByRef KeyToVerification As String, ByRef DescriptObjectResult As String, ByRef CargaDetalle As Boolean, ByRef IsProperty As Boolean, ByRef ImageAsociate As Integer, ByRef LocateDate As Boolean, ByRef ConnectionToBase As ADODB.Connection)
'// Verificación de Tipos de Archivos a tomar en cuenta para Documentar
Dim RSConsultaBase As ADODB.Recordset
Dim SQLString As String

'// Query de Ejecución Conformado por: "SELECT DISTINCT * FROM ObjectsOfStructProyect WHERE NombreObjeto = '" & KeyToVerification & "';"

SQLString = "SELECT DISTINCT * FROM ObjectsOfStructProyect WHERE NombreObjeto = '" & KeyToVerification & "';"
Set RSConsultaBase = New ADODB.Recordset
With RSConsultaBase
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Open SQLString, ConnectionToBase, , , adCmdText

If .EOF = True And .BOF = True Then
   DescriptObjectResult = ""
   CargaDetalle = False
   IsProperty = False
   ImageAsociate = 0
   LocateDate = False
Else
   .MoveFirst
   DescriptObjectResult = !DescripcionObjeto
   CargaDetalle = CBool(!RequiereDetalle)
   IsProperty = CBool(!PropertyUObjectExternal)
   ImageAsociate = !ImagenAsociada
   LocateDate = True

End If
.Close
End With

Set RSConsultaBase = Nothing


End Sub
Public Sub ExecQuery(StrQuery As String, Connection As ADODB.Connection, Transaccional As Boolean)
'// Proceso de Ejecución de Querys, permitiendo ejecutarlos Transaccionales o No transaccionales en un momento dado.
On Error GoTo err_execquery
Dim ExecuteQuery  As ADODB.Command

Set ExecuteQuery = New ADODB.Command
Set ExecuteQuery.ActiveConnection = Connection
ExecuteQuery.CommandText = StrQuery

If Transaccional = True Then
    Connection.BeginTrans
    ExecuteQuery.Execute
    Connection.CommitTrans
Else
    ExecuteQuery.Execute
End If

exit_execquery:
Exit Sub

err_execquery:
  If Transaccional = True Then
  Connection.RollbackTrans
  End If
Resume exit_execquery
End Sub

