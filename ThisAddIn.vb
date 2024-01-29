Imports System.Drawing.Text
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Tools
Imports Microsoft.VisualBasic.ApplicationServices

Module MyGlobals
    Public ThisApp As ExtendedApplication
End Module

Public Class ThisAddIn

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New MarcaRibbon()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ThisApp = New ExtendedApplication With {
            .ExcelApplicationProp = Globals.ThisAddIn.Application
        }
        ThisApp.UserId = ThisApp.GetUserId()
        If ThisApp.UserId Is Nothing Then
            MessageBox.Show("Error al obtener credenciales del usuario")
            ThisApp.ExcelApplicationProp.Quit()
        Else
            ThisApp.ResetState()
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ThisApp.UnloadEventHandling()
    End Sub
End Class
