Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports System.Xml
Imports Excel = Microsoft.Office.Interop.Excel
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports System.Net.Security
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System
Imports MS.WindowsAPICodePack.Internal
Imports System.Net.Mime.MediaTypeNames
Imports System.Collections
Imports Microsoft.Office.Tools
Imports Marca_4._0.ExtendedApplication

Public Class Utils
    Public Shared Property SeparadorDeLista As String = CultureInfo.CurrentCulture.TextInfo.ListSeparator
    Public Shared Function CloseStrBuilder(param As StringBuilder, delim As Char) As String()
        If param.Length < 1 Then Return Nothing
        param.Remove(param.Length - 1, 1)
        Return param.ToString().Split(New Char() {delim})
    End Function
    Public Shared Function CalcularHashSHA256(texto As String) As String

        Using sha256 As SHA256 = SHA256.Create()
            ' Convertir el texto a un array de bytes y calcular el hash
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(texto)
            Dim hashBytes As Byte() = sha256.ComputeHash(bytes)

            ' Convertir el array de bytes del hash a una cadena hexadecimal
            Dim stringBuilder As New StringBuilder()
            For Each b As Byte In hashBytes
                stringBuilder.Append(b.ToString("x2"))
            Next
            Return stringBuilder.ToString()
        End Using
    End Function
    Public Shared Sub AsignarValidacionDeLista(ByRef ExcelRangeProp As Excel.Range, valores As List(Of Object))
        ' Limpiar cualquier validación existente
        ExcelRangeProp.Validation.Delete()
        ' Unir los valores en una cadena separada por comas
        If valores.Count = 0 Then Exit Sub

        'Modificar los elementos de la lista para tratar los textos como tales
        Dim valoresModificados As New List(Of Object)
        For Each valor As Object In valores
            If valor Is Nothing Then Continue For


            If TypeOf valor Is String AndAlso IsNumeric(valor) Then
                ' Agregar una comilla simple al principio para forzar el tratamiento como texto
                valoresModificados.Add("'" & valor.ToString())
            Else
                valoresModificados.Add(valor)
            End If

        Next
        Dim listaValores As String
        If valoresModificados.Count > 0 Then
            listaValores = String.Join(Utils.SeparadorDeLista, valoresModificados)

            With ExcelRangeProp.Validation
                .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                 Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:=listaValores)
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        End If
    End Sub
    Public Shared Function SeekOrMakeSheet(ByRef Wb As Excel.Workbook, Name As String) As Excel.Worksheet
        Dim Sh As Excel.Worksheet = Nothing

        On Error Resume Next
        Sh = Wb.Worksheets(Name)
        Err.Clear()

        If Sh Is Nothing Then

            Sh = Wb.Worksheets.Add()
            On Error Resume Next
            Sh.Name = Name
            Sh.Visible = Excel.XlSheetVisibility.xlSheetHidden
            Err.Clear()

        End If

        If Sh IsNot Nothing AndAlso Sh.Name.Equals(Name) Then
            Return Sh
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function SeekOrMakeListObject(Sh As Excel.Worksheet, TableName As String, ColumnNames As String()) As Excel.ListObject

        If Sh Is Nothing Then Return Nothing

        ' Crear la tabla si no existe
        Dim Tbl As Excel.ListObject = Nothing
        If Sh.ListObjects.Count > 0 Then
            On Error Resume Next
            Tbl = Sh.ListObjects(TableName)
            Err.Clear()
        End If

        If Tbl Is Nothing Then

            Dim OtherTbl As Excel.ListObject = Nothing
            On Error Resume Next
            OtherTbl = Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, ColumnNames.Count)).ListObject
            Err.Clear()
            If OtherTbl IsNot Nothing Then OtherTbl.Delete()


            Tbl = Sh.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
                                              Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, ColumnNames.Count)), Type.Missing,
                                              Excel.XlYesNoGuess.xlYes, Type.Missing)
            Tbl.Name = TableName
            Tbl.HeaderRowRange.Value2 = ColumnNames
        End If

        Return Tbl
    End Function
End Class

Public Class ExtendedApplication
    Public Structure StateStructure
        Public Counter As Integer
        Public Enable As Boolean
    End Structure
    Public Enum StateTransition
        Start
        Enable
        Disable
        Finish
    End Enum
    Public Enum StateOption
        ForceEncript
        AppEvents
        AppAlerts
    End Enum
    Public Property States As Dictionary(Of StateOption, StateStructure)
    Public ReadOnly Property State(Et As StateOption) As Boolean
        Get
            Dim Struc As StateStructure = States(Et)
            Return Struc.Enable
        End Get
    End Property


    Public WithEvents ExcelApplicationProp As Excel.Application
    Public Property ExtendedBooksProp As ExtendedBooks
    Public Property MarkedFilesProp As MarkeddFiles

    Public ReadOnly Property ActiveExtendedWorkBook As ExtendedBook
        Get
            If ExcelApplicationProp.ActiveWorkbook Is Nothing Then Return Nothing
            Return ExtendedBooksProp.ExtendedBookProp(ExcelApplicationProp.ActiveWorkbook.Name)
        End Get
    End Property

    Public UserId As Dictionary(Of String, String)

    Sub New()
        States = New Dictionary(Of StateOption, StateStructure) From {
            {StateOption.ForceEncript, New StateStructure() With {.Counter = -1, .Enable = False}},
            {StateOption.AppEvents, New StateStructure() With {.Counter = 0, .Enable = True}},
            {StateOption.AppAlerts, New StateStructure() With {.Counter = 0, .Enable = True}}
        }
        ExtendedBooksProp = New ExtendedBooks With {
            .ExtendedAppProp = Me
        }
        MarkedFilesProp = New MarkeddFiles With {
            .ExtendedAppProp = Me
        }

    End Sub

    Public Function AddNewWorkBook(Wb As Excel.Workbook) As ExtendedBook
        'registra un nuevo libro
        Dim ExtBk As New ExtendedBook() With {
                .ExtendedBooksProp = ExtendedBooksProp,
                .ExcelWorkbookProp = Wb,
                .MarkedFileProp = Nothing
            }
        ExtendedBooksProp.Add(ExtBk)
        If State(StateOption.ForceEncript) Then
            If Not ExtBk.AddMarkedWorkbook() Then
                ExtendedBooksProp.Remove(ExtBk)
                Return Nothing
            End If
        End If
        Return ExtBk
    End Function

    Sub UnloadEventHandling()
        RemoveHandler ExcelApplicationProp.WorkbookNewSheet, AddressOf HandleWorkbookNewSheet
        RemoveHandler ExcelApplicationProp.NewWorkbook, AddressOf Application_WorkbookNew
        RemoveHandler ExcelApplicationProp.WorkbookOpen, AddressOf Application_WorkbookOpen
        RemoveHandler ExcelApplicationProp.WorkbookBeforeClose, AddressOf Application_WorkbookBeforeClose
    End Sub

    Public Sub Undo()
        ExcelApplicationProp.Undo()
    End Sub

    Public Sub UpdateState(Et As StateOption, Trans As StateTransition)
        Dim Struc As StateStructure = States(Et)
        Select Case Trans
            Case StateTransition.Disable
                Struc.Counter -= 1
                If Struc.Counter < 0 Then
                    Struc.Enable = False
                End If
            Case StateTransition.Enable
                Struc.Counter += 1
                If Struc.Counter >= 0 Then
                    Struc.Enable = True
                End If
            Case StateTransition.Finish
                Struc.Counter = -1
                Struc.Enable = False
            Case StateTransition.Start
                Struc.Counter = 0
                Struc.Enable = True
        End Select
        If Et = StateOption.AppEvents Then ExcelApplicationProp.EnableEvents = Struc.Enable
        If Et = StateOption.AppAlerts Then ExcelApplicationProp.DisplayAlerts = Struc.Enable
        States(Et) = Struc
    End Sub

    Public Sub ResetState()
        ThisApp.UpdateState(StateOption.ForceEncript, StateTransition.Finish)
    End Sub

    Public Sub SuspendProcess()
        UpdateState(StateOption.AppEvents, StateTransition.Disable)
        UpdateState(StateOption.AppAlerts, StateTransition.Disable)
    End Sub

    Public Sub ResumeProces()
        UpdateState(StateOption.AppEvents, StateTransition.Enable)
        UpdateState(StateOption.AppAlerts, StateTransition.Enable)
    End Sub

    Public Function GetUserId() As Dictionary(Of String, String)
        Dim outlookApp As Outlook.Application = Nothing
        Try
            outlookApp = New Outlook.Application
            Dim addrEntry As Outlook.AddressEntry = outlookApp.Session.CurrentUser.AddressEntry
            Dim currentUser As Outlook.ExchangeUser
            If addrEntry.Type = "EX" Then
                currentUser = outlookApp.Session.CurrentUser.AddressEntry.GetExchangeUser()
                If currentUser IsNot Nothing Then
                    Dim result As New Dictionary(Of String, String) From {
                            {"UsuarioExcel", Me.ExcelApplicationProp.UserName},
                            {"UsuarioOutlook", currentUser.Name},
                            {"Correo", currentUser.PrimarySmtpAddress},
                            {"Movil", currentUser.MobileTelephoneNumber}
                        }
                    outlookApp.Quit()
                    Return result
                End If
                outlookApp.Quit()
                Return Nothing
            Else
                outlookApp.Quit()
                Return Nothing
            End If
        Catch ex As Exception
            If outlookApp IsNot Nothing Then outlookApp.Quit()
            Return Nothing
        End Try
    End Function

    Private Sub HandleWorkbookNewSheet(ByVal Wb As Excel.Workbook, ByVal Sh As Object) Handles ExcelApplicationProp.WorkbookNewSheet
        Dim tupla As ExtendedBook = ExtendedBooksProp.ExtendedBookProp(Wb.Name)
        If tupla Is Nothing Then Exit Sub
        SuspendProcess()
        If State(StateOption.ForceEncript) Then 'marcado obligatorio
            If tupla.MarkedFileProp Is Nothing Then 'Nueva hoja en libro no marcado
                MessageBox.Show($"Se marcara libro entero {Wb.Name}")
                tupla.AddMarkedWorkbook()
            End If
            tupla.MarkedFileProp.MarkWorksheet(Sh)
        End If
        ResumeProces()
    End Sub
    Private Sub HandlerWorkbookBeforeSave(ByVal Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles ExcelApplicationProp.WorkbookBeforeSave
        Dim ExtWb As ExtendedBook = ExtendedBooksProp.ExtendedBookProp(Wb.Name)
        If ExtWb Is Nothing Then
            SuspendProcess()
            Wb.Close(SaveChanges:=False)
            ResumeProces()
            Cancel = True
            Exit Sub
        End If
        Cancel = False
        If State(StateOption.ForceEncript) Then 'marcado obligatorio
            If ExtWb.MarkedFileProp Is Nothing Then 'se marca libro no marcado
                MessageBox.Show($"Se marcara libro entero {Wb.Name}")
                ExtWb.AddMarkedWorkbook()
            End If
        End If
        SuspendProcess()
        ExtWb.PreSave()
        ResumeProces()
    End Sub
    Public Sub Application_WorkbookOpen(ByVal Wb As Excel.Workbook) Handles ExcelApplicationProp.WorkbookOpen
        SuspendProcess()
        Dim ExtBk As ExtendedBook = AddNewWorkBook(Wb)
        If ExtBk Is Nothing Then Wb.Close(SaveChanges:=False)
        ResumeProces()
    End Sub
    Public Sub Application_WorkbookNew(ByVal Wb As Excel.Workbook) Handles ExcelApplicationProp.NewWorkbook
        SuspendProcess()
        Dim ExtBk As ExtendedBook = AddNewWorkBook(Wb)
        If ExtBk Is Nothing Then Wb.Close(SaveChanges:=False)
        ResumeProces()
    End Sub
    Public Sub Application_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, ByRef Cancel As Boolean) Handles ExcelApplicationProp.WorkbookBeforeClose
        Cancel = True
        Dim ExtWb As ExtendedBook = ExtendedBooksProp.ExtendedBookProp(Wb.Name)
        If ExtWb Is Nothing Then
            SuspendProcess()
            Wb.Close(SaveChanges:=False)
            ResumeProces()
            Exit Sub
        End If
        Dim resultado As DialogResult
        If Not Wb.Saved Then
            resultado = MessageBox.Show("Los Cambios se Perderan", "Continuar", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        End If
        If Wb.Saved Or resultado = DialogResult.Yes Then
            SuspendProcess()
            ExtendedBooksProp.Remove(ExtWb)
            Wb.Close(SaveChanges:=False)
            ResumeProces()
        End If
    End Sub
End Class
Public Class ExtendedBook
    Public Const ClaveLibroMarcado As String = "ClaveLibroMarcado"
    Public Const HojaCaratula As String = "Caratula"
    Public Property ExtendedBooksProp As ExtendedBooks
    Public Property MarkedFileProp As MarkeddFile
    Public WithEvents ExcelWorkbookProp As Excel.Workbook
    Public ReadOnly Property IsEncripted As Boolean
        Get
            If ExcelWorkbookProp Is Nothing Then
                Return False
            Else
                Return ExcelWorkbookProp.ProtectStructure
            End If
        End Get
    End Property
    Public Overrides Function Equals(obj As Object) As Boolean
        ' Define aquí cómo comparar dos objetos
        Dim x As ExtendedBook = TryCast(obj, ExtendedBook)
        If x Is Nothing Then Return False
        Return x.ExcelWorkbookProp.Name.Equals(Me.ExcelWorkbookProp.Name)
    End Function
    Public Overrides Function GetHashCode() As Integer
        ' Devuelve un código hash para un objeto
        Return Utils.CalcularHashSHA256(Me.ExcelWorkbookProp.Name)
    End Function
    Sub New()
        MarkedFileProp = Nothing
    End Sub
    Public Function PreSave() As Boolean
        'Encripta y borra datos en caratula visible
        ExcelWorkbookProp.Activate()
        If ExcelWorkbookProp.Saved Then Return True
        Try
            If MarkedFileProp IsNot Nothing Then
                EncriptWorkbook()
                Dim ws As Excel.Worksheet
                Try
                    ws = ExcelWorkbookProp.Worksheets(HojaCaratula)
                Catch ex As Exception
                    ws = Nothing
                End Try
                If ws IsNot Nothing Then ws.Cells.Clear()
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function AddMarkedWorkbook() As Boolean
        'Registra como marcado y marca libro 
        Try
            If MarkedFileProp Is Nothing Then
                MarkedFileProp = New MarkeddFile() With {
                    .ExtendedBookProp = Me,
                    .MarkedFilesProp = ExtendedBooksProp.ExtendedAppProp.MarkedFilesProp
                }
                MarkedFileProp.MarkedFilesProp.Add(MarkedFileProp)
                If Not IsEncripted Then
                    Dim ws As Excel.Worksheet = Utils.SeekOrMakeSheet(ExcelWorkbookProp, HojaCaratula)
                    ws.Visible = Excel.XlSheetVisibility.xlSheetVisible
                End If
                MarkedFileProp.MarkWoorkbook()
            Else
                MessageBox.Show($"Libro ya registrado {ExcelWorkbookProp.Name}")
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show($"Error al registrar o marcar libro {ExcelWorkbookProp.Name}")
            Return False
        End Try
    End Function
    Public Function EncriptWorkbook() As Boolean
        If ExcelWorkbookProp Is Nothing Then Return False
        Try
            If IsEncripted Then ExcelWorkbookProp.Unprotect(ClaveLibroMarcado)
            For Each ws As Excel.Worksheet In ExcelWorkbookProp.Worksheets
                If Not ws.Name.Equals(HojaCaratula) Then ws.Visible = Excel.XlSheetVisibility.xlSheetHidden
            Next
            ExcelWorkbookProp.Protect(ClaveLibroMarcado)
            Return True
        Catch es As Exception
            Return False
        End Try
    End Function
    Public Function DesEncriptWorkbook() As Boolean
        If ExcelWorkbookProp Is Nothing Then Return False
        If IsEncripted Then
            Try
                ExcelWorkbookProp.Unprotect(Password:=ClaveLibroMarcado)
                For Each ws As Excel.Worksheet In ExcelWorkbookProp.Worksheets
                    ws.Visible = Excel.XlSheetVisibility.xlSheetVisible
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function
End Class
Public Class ExtendedBooks
    Inherits List(Of ExtendedBook)
    Public Property ExtendedAppProp As ExtendedApplication
    Public ReadOnly Property ExtendedBookProp(Name As String) As ExtendedBook
        Get
            'itera por si los libros han cambiado de nombre respecto de la clave
            For Each Ewb As ExtendedBook In Me
                If Ewb.ExcelWorkbookProp.Name = Name Then Return Ewb
            Next
            Return Nothing
        End Get
    End Property
    Public Overloads Function Remove(ExtWb As ExtendedBook) As Boolean
        Try
            If ExtWb.MarkedFileProp IsNot Nothing Then ExtendedAppProp.MarkedFilesProp.Remove(ExtWb.MarkedFileProp)
            If ExtendedAppProp.MarkedFilesProp.Count = 0 Then ExtendedAppProp.UpdateState(StateOption.ForceEncript, StateTransition.Finish)
            MyBase.Remove(ExtWb)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Sub MarkWorkbooks()
        For Each ExtWb As ExtendedBook In Me
            If Not ExtWb.AddMarkedWorkbook() Then MessageBox.Show($"Error al marcar libro {ExtWb.ExcelWorkbookProp.Name}")
        Next
    End Sub
End Class
Public Class MarkeddFile
    Public Property ExtendedBookProp As ExtendedBook
    Public Property MarkedFilesProp As MarkeddFiles
    Private Property InnerImagen As Image
    Public ReadOnly Property Imagen As Image
        Get
            Return InnerImagen
        End Get
    End Property
    Public Overrides Function Equals(obj As Object) As Boolean
        ' Define aquí cómo comparar dos objetos
        Dim x As MarkeddFile = TryCast(obj, MarkeddFile)
        If x Is Nothing Then Return False
        Return x.ExtendedBookProp.Equals(Me.ExtendedBookProp)
    End Function
    Public Overrides Function GetHashCode() As Integer
        ' Devuelve un código hash para un objeto
        Return Me.ExtendedBookProp.GetHashCode()
    End Function
    Public Function GenerateMark() As Boolean
        Dim thisApp As ExtendedApplication = ExtendedBookProp.ExtendedBooksProp.ExtendedAppProp
        Dim IdUsuario As Dictionary(Of String, String) = thisApp.UserId
        Try
            InnerImagen = New Image(IdUsuario("UsuarioExcel"), IdUsuario("UsuarioOutlook"), IdUsuario("Correo"), IdUsuario("Movil"))
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function MarkWoorkbook() As Boolean
        Dim Result As Boolean = True
        GenerateMark()
        For Each Ws As Excel.Worksheet In ExtendedBookProp.ExcelWorkbookProp.Worksheets
            If Not MarkWorksheet(Ws) Then Result = False
        Next
        Return Result
    End Function
    Public Function MarkWorksheet(ByRef sheet As Excel.Worksheet) As Boolean
        ' Establecer imagen como encabezado central
        Try
            With sheet
                .Cells.Interior.ColorIndex = 0 ' Excel.XlColorIndex.xlColorIndexNone
                With .PageSetup
                    .CenterHeaderPicture.Filename = Imagen.Header
                    .CenterFooterPicture.Filename = Imagen.Footer
                    .CenterHeader = Chr(10) & "&G"
                    .CenterFooter = "&G" & Chr(10)
                End With
                ' Establecer imagen como fondo
                .SetBackgroundPicture(Imagen.Screen)
            End With
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
End Class
Public Class MarkeddFiles
    Inherits List(Of MarkeddFile)
    Public Property ExtendedAppProp As ExtendedApplication
End Class
Public Class Image
    Private Property InnerScreen As String
    Public ReadOnly Property Screen As String
        Get
            Return InnerScreen
        End Get
    End Property
    Private Property InnerHeader As String
    Public ReadOnly Property Header As String
        Get
            Return InnerHeader
        End Get
    End Property
    Private Property InnerFooter As String
    Public ReadOnly Property Footer As String
        Get
            Return InnerFooter
        End Get
    End Property
    Public Sub New(Id1 As String, Id2 As String, Id3 As String, Id4 As String)
        InnerScreen = System.IO.Path.GetTempFileName() & ".png"
        InnerHeader = System.IO.Path.GetTempFileName() & ".png"
        InnerFooter = System.IO.Path.GetTempFileName() & ".png"
        GenerateGraphic(InnerScreen, Brushes.LightPink, Id1, Id2, Id3, Id4)
        GenerateGraphic(InnerHeader, Brushes.Red, Id1, Id2, Id3, Id4)
        GenerateGraphic(InnerFooter, Brushes.Red, Id1, Id2, Id3, Id4)
    End Sub
    Public Sub GenerateGraphic(imageFilePath As String, brush As Brush, Id1 As String, Id2 As String, Id3 As String, Id4 As String)
        Dim width As Integer = 600
        Dim height As Integer = 325
        Using bitmap As New Bitmap(width, height)
            Using g = Graphics.FromImage(bitmap)
                g.Clear(Color.Transparent)
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
                Dim stringFormat As New StringFormat() With
            {
            .Alignment = StringAlignment.Center,
            .LineAlignment = StringAlignment.Center,
            .FormatFlags = StringFormatFlags.LineLimit, ' Limita las líneas para ajustarlas al recuadro
            .Trimming = StringTrimming.Word ' Asegura que las palabras no se corten al ajustar
            }
                Dim font As New System.Drawing.Font("Arial", 38)
                Dim text As String = "CONFIDENCIAL PARA: " & Id1 & Chr(10) & Id2 & Chr(10) & Id3 & Chr(10) & Id4
                Dim rect As New System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height)
                g.DrawString(text, font, brush, rect, stringFormat)
            End Using
            bitmap.Save(imageFilePath, Imaging.ImageFormat.Png)
        End Using
    End Sub
End Class
