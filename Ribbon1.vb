'TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

'1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
'   zu behandeln, zum Beispiel das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem
'   Menüband-Designer exportiert haben, verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und
'   ändern Sie den Code für die Verwendung mit dem Programmiermodell für die Menübanderweiterung (RibbonX).

'3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.

'Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("OutlookAddIn5.Ribbon1.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub GetButtonID(control As Office.IRibbonControl)
        Dim myOutlook As Outlook.Application = Globals.ThisAddIn.Application

        System.Diagnostics.Debug.WriteLine(CStr(myOutlook.ActiveExplorer.Selection.Count) + "(Count)")
        If myOutlook.ActiveExplorer.Selection.Count > 0 Then

            Dim mailItem As Outlook.MailItem
            mailItem = DirectCast(myOutlook.ActiveExplorer.Selection(1), Outlook.MailItem)
            Select Case control.Id
                Case "FolderDisplay"
                    Dim messageBoxText As String = "The item is in folder:" & vbLf & vbLf _
                                                   & DirectCast(mailItem.Parent, Outlook.Folder).FolderPath
                    MsgBox(messageBoxText, 0, "In which folder?")

                Case "FolderGoto"
                    System.Diagnostics.Debug.WriteLine("Current: " & TryCast(myOutlook.ActiveExplorer.CurrentFolder.Name, String) _
                                                                   & myOutlook.ActiveExplorer.Selection.Count)  ' yields <1>
                    ' Change the current folder:
                    myOutlook.ActiveExplorer.CurrentFolder = DirectCast(mailItem.Parent, Outlook.Folder)
                    System.Diagnostics.Debug.WriteLine("Current: " & TryCast(myOutlook.ActiveExplorer.CurrentFolder.Name, String) _
                                                                   & myOutlook.ActiveExplorer.Selection.Count)  ' yields <0>
                    System.Windows.Forms.Application.DoEvents()         ' needed, otherwise the mailItem will not be selectable

                    ' Additionally, set the initial mailItem as selected:
                    ' But this doesn't work for unknown reasons.
                    ' .AddToSelection() does not work for unknown reasons.
                    If myOutlook.ActiveExplorer.IsItemSelectableInView(mailItem) Then
                        ' https://social.msdn.microsoft.com/Forums/en-US/aedcbda9-5304-4969-82ac-dbd41e0879b0/select-item-in-activeexplorer?forum=outlookdev
                        System.Diagnostics.Debug.WriteLine("Selectable!")

                        System.Diagnostics.Debug.WriteLine("Adding0: " & myOutlook.ActiveExplorer.Selection.Count)  ' <0> is OK
                        System.Diagnostics.Debug.WriteLine(CStr("Item: " & mailItem.Subject))                       ' verify, that the element is the desired one
                        myOutlook.ActiveExplorer.AddToSelection(mailItem)                                           ' This FAILS !
                        System.Diagnostics.Debug.WriteLine("Adding1: " & myOutlook.ActiveExplorer.Selection.Count)  ' should be <1> but is still <0> :-(
                        System.Windows.Forms.Application.DoEvents()                                                 ' thought in vain that this would help
                        System.Diagnostics.Debug.WriteLine("Adding2: " & myOutlook.ActiveExplorer.Selection.Count)  ' should be <1> but is still <0> :-(
                        System.Diagnostics.Debug.WriteLine("Which Folder?    = " & myOutlook.ActiveExplorer.CurrentFolder.Name)

                        'Dim oneitem As Object
                        'Dim oneMsg As Outlook.MailItem
                        'For Each oneitem In Globals.ThisAddIn.Application.ActiveExplorer.CurrentFolder.Items
                        '    If Globals.ThisAddIn.Application.ActiveExplorer.IsItemSelectableInView(oneitem) Then
                        '        If TypeOf oneitem Is Outlook.MailItem Then
                        '            oneMsg = DirectCast(oneitem, Outlook.MailItem)
                        '            'System.Diagnostics.Debug.WriteLine("MailItem = " & oneitem.Subject & oneitem.CreationTime)
                        '            System.Diagnostics.Debug.WriteLine("MailItem =" & oneMsg.Subject & " (" & oneMsg.CreationTime & ")")
                        '            Globals.ThisAddIn.Application.ActiveExplorer.AddToSelection(oneMsg)
                        '            System.Diagnostics.Debug.WriteLine("Selection.Countx = " & Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count)
                        '            Globals.ThisAddIn.Application.ActiveExplorer.Activate()
                        '            System.Windows.Forms.Application.DoEvents()
                        '        End If
                        '    End If
                        'Next
                    Else
                        System.Diagnostics.Debug.WriteLine("Not selectable!")
                        'Globals.ThisAddIn.Application.ActiveExplorer.Display()
                        'Globals.ThisAddIn.Application.ActiveExplorer.AddToSelection(mailItem)
                    End If
            End Select
        End If
    End Sub

#End Region

#Region "Hilfsprogramme"

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
