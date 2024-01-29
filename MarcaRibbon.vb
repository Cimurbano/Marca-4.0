'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New MarcaRibbon()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports System.Drawing
Imports stdole
Imports Microsoft.Office.Tools.Ribbon
Imports System.Runtime.CompilerServices
Imports Marca_4._0.ExtendedApplication

<Runtime.InteropServices.ComVisible(True)>
Public Class MarcaRibbon
    Implements Office.IRibbonExtensibility
    Private ribbon As Office.IRibbonUI
    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("Marca_4._0.MarcaRibbon.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub
    Public Sub BlockCommand(control As Office.IRibbonControl, ByRef cancelDefault As Boolean)
        If ThisApp.State(StateOption.ForceEncript) Then
            MessageBox.Show("Comando Bloqueado.")
            cancelDefault = True
        Else
            cancelDefault = False
        End If
    End Sub
    Public Function ConfigEncriptVisible(control As Office.IRibbonControl) As Boolean
        If ThisApp.State(StateOption.ForceEncript) Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Sub MarkOne(control As Office.IRibbonControl)
        If ThisApp.ActiveExtendedWorkBook Is Nothing Then Exit Sub
        ThisApp.SuspendProcess()
        ThisApp.ActiveExtendedWorkBook.AddMarkedWorkbook()
        ThisApp.ResumeProces()
    End Sub
    Public Sub MarkAll(control As Office.IRibbonControl)
        If ThisApp.ExtendedBooksProp Is Nothing Then Exit Sub
        ThisApp.SuspendProcess()
        ThisApp.ExtendedBooksProp.MarkWorkbooks()
        ThisApp.ResumeProces()
    End Sub
    Public Sub DesEncriptMarked(control As Office.IRibbonControl)
        If ThisApp.ActiveExtendedWorkBook Is Nothing Then Exit Sub
        If ThisApp.ActiveExtendedWorkBook.DesEncriptWorkbook() Then
            ThisApp.SuspendProcess()
            ThisApp.UpdateState(StateOption.ForceEncript, StateTransition.Start)
            ThisApp.ExtendedBooksProp.MarkWorkbooks()
            ThisApp.ResumeProces()
            Me.ribbon.Invalidate()
        Else
            MessageBox.Show("Error archivo no encriptado")
        End If
    End Sub
    Public Function GetGroupWorkbookViewsVisible(control As Office.IRibbonControl) As Boolean
        If ThisApp.State(StateOption.ForceEncript) Then
            Return False ' desHabilita el grupo
        Else
            Return True ' habilita el grupo
        End If
    End Function

    Public Sub ResetRibbon(control As Office.IRibbonControl)
        Me.ribbon.Invalidate()
    End Sub
#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
