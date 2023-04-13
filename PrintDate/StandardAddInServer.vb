Imports Inventor
Imports System.Runtime.InteropServices
'Imports Microsoft.Win32

Namespace PrintDate
  <ProgId("PrintDate.StandardAddInServer"),
  Guid("845aa20b-673d-444f-a9d1-97d2d0597aa1")>
  Public Class StandardAddInServer
    Implements ApplicationAddInServer

#Region "Data"
    ' Inventor application object.
    Private m_inventorApplication As Application

    'For adding Toolbar Buttons
    Private WithEvents M_UserInterfaceEvents As UserInterfaceEvents
    Private WithEvents PrintButtonDef As ButtonDefinition
    Private WithEvents PrintPreviewButtonDef As ButtonDefinition
    Private IVPrintButtonDef As ButtonDefinition
    Private IVPrintPreviewButtonDef As ButtonDefinition

#End Region

#Region "ApplicationAddInServer Members"

    Public Sub Activate(ByVal AddInSiteObject As ApplicationAddInSite, ByVal FirstTime As Boolean
                        ) Implements ApplicationAddInServer.Activate

      Try
        ' This method is called by Inventor when it loads the AddIn.
        ' The AddInSiteObject provides access to the Inventor Application object.
        ' The FirstTime flag indicates if the AddIn is loaded for the first time.
        ' Initialize AddIn members.
        m_inventorApplication = AddInSiteObject.Application

        ' TODO:  Add ApplicationAddInServer.Activate implementation.
        ' e.g. event initialization, command creation etc.

        'If firstTime = True Then
        ReplaceButton(m_inventorApplication)
        'End If

      Catch ex As Exception
        Windows.Forms.MessageBox.Show(ex.ToString())
      End Try

    End Sub

    Public Sub ReplaceButton(app As Application)

      Dim controlDefs As ControlDefinitions
      Dim controlDef As ControlDefinition
      Dim parentCtrl As CommandControl
      Dim targetCtrl As CommandControl
      Dim cmdCtrl As CommandControl
      Dim oQAT As CommandControls
      Dim FoundCount As Integer = 0
      Dim OldShortCut As String

      'MsgBox("Start Replace Button")
      controlDefs = app.CommandManager.ControlDefinitions
      For Each controlDef In app.CommandManager.ControlDefinitions
        If (controlDef.InternalName = "AppFilePrintCmd") Then
          IVPrintButtonDef = controlDef
          FoundCount += 1
          'MsgBox("Found Print")
        ElseIf (controlDef.InternalName = "AppFilePrintPreviewCmd") Then
          IVPrintPreviewButtonDef = controlDef
          FoundCount += 1
          'MsgBox("Found Print Preview")
        End If
        If FoundCount = 2 Then
          Exit For
        End If
      Next
      If FoundCount < 2 Then
        'MsgBox("No button found")
      End If

      'MsgBox("Create PrintDate Button")
      PrintButtonDef = controlDefs.AddButtonDefinition(IVPrintButtonDef.DisplayName,
          "AppFilePrintDateCmd",
          CommandTypesEnum.kFilePropertyEditCmdType,
          IVPrintButtonDef.ClientId,
          IVPrintButtonDef.DescriptionText,
          IVPrintButtonDef.ToolTipText,
          IVPrintButtonDef.StandardIcon,
          IVPrintButtonDef.LargeIcon)

      'MsgBox("Clear Old Override")
      OldShortCut = IVPrintButtonDef.OverrideShortcut
      IVPrintButtonDef.OverrideShortcut = ""

      'MsgBox("Set Override")
      PrintButtonDef.OverrideShortcut = OldShortCut

      'MsgBox("Create PrintDatePreview Button")
      PrintPreviewButtonDef = controlDefs.AddButtonDefinition(IVPrintPreviewButtonDef.DisplayName,
          "AppFilePrintDatePreviewCmd",
          CommandTypesEnum.kFilePropertyEditCmdType,
          IVPrintPreviewButtonDef.ClientId,
          IVPrintPreviewButtonDef.DescriptionText,
          IVPrintPreviewButtonDef.ToolTipText,
          IVPrintPreviewButtonDef.StandardIcon,
          IVPrintPreviewButtonDef.LargeIcon)

      'Add Button to FileBrowerControl, hide original Button
      'MsgBox("Adding to File Brower")
      parentCtrl = app.UserInterfaceManager.FileBrowserControls("AppFilePrintCmd")
      targetCtrl = parentCtrl.ChildControls("AppFilePrintCmd")
      cmdCtrl = parentCtrl.ChildControls.AddButton(PrintButtonDef, False, True,
                                                   IVPrintButtonDef.InternalName, True)
      targetCtrl.Visible = False
      targetCtrl = parentCtrl.ChildControls("AppFilePrintPreviewCmd")
      cmdCtrl = parentCtrl.ChildControls.AddButton(PrintPreviewButtonDef, False, True,
                                                   IVPrintPreviewButtonDef.InternalName, True)
      targetCtrl.Visible = False
      'MsgBox("Added to File Brower")

      'Add Button to Drawing.QuickAccessToolbar, hide original Button
      'MsgBox("Adding to QAT")
      oQAT = app.UserInterfaceManager.Ribbons.Item("Drawing").QuickAccessControls
      targetCtrl = oQAT.Item(IVPrintButtonDef.InternalName)
      cmdCtrl = oQAT.AddButton(PrintButtonDef, False, True, IVPrintButtonDef.InternalName, True)
      targetCtrl.Visible = False
      'MsgBox("Added to QAT")

    End Sub

    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate

      ' This method is called by Inventor when the AddIn is unloaded.
      ' The AddIn will be unloaded either manually by the user or
      ' when the Inventor session is terminated.

      ' TODO:  Add ApplicationAddInServer.Deactivate implementation
      'Try

      ' Release objects.
      'Marshal.ReleaseComObject(m_inventorApplication)
      m_inventorApplication = Nothing
      GC.WaitForPendingFinalizers()
      GC.Collect()
      'Catch ex As Exception
      'Windows.Forms.MessageBox.Show(ex.ToString())
      'End Try
    End Sub

    Public ReadOnly Property Automation() As Object Implements ApplicationAddInServer.Automation

      ' This property is provided to allow the AddIn to expose an API 
      ' of its own to other programs. Typically, this  would be done by
      ' implementing the AddIn's API interface in a class and returning 
      ' that class object through this property.

      Get
        Return Nothing
      End Get

    End Property

    Public Sub ExecuteCommand(ByVal CommandID As Integer) Implements ApplicationAddInServer.ExecuteCommand

      ' Note:this method is now obsolete, you should use the 
      ' ControlDefinition functionality for implementing commands.

    End Sub

    Private Sub M_UserInterfaceEvents_OnResetRibbonInterface(Context As NameValueMap) Handles M_UserInterfaceEvents.OnResetRibbonInterface
      ReplaceButton(m_inventorApplication)
    End Sub

    Private Sub SetProperties()
      'Get Current Date Time and add to Drawing custom Properties
      ' Set a reference to the drawing document.
      ' This assumes a drawing document is active.
      If m_inventorApplication.ActiveDocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
        Dim oDrawDoc As DrawingDocument = m_inventorApplication.ActiveDocument
        Dim sDateTime As String = Now.ToString
        Dim oPropSet As PropertySet = oDrawDoc.PropertySets.Item("User Defined Properties")

        Try
          oPropSet.Item("LastPrintDate").Value = sDateTime
        Catch
          oPropSet.Add(sDateTime, "LastPrintDate")
        End Try

        Try
          oPropSet.Item("LastPrintUser").Value = m_inventorApplication.UserName
        Catch
          oPropSet.Add(m_inventorApplication.UserName, "LastPrintUser")
        End Try

        oDrawDoc.Update()
        'oDrawDoc.Save()
      End If
    End Sub

    Private Sub PrintButtonDef_OnExecute(Context As NameValueMap) Handles PrintButtonDef.OnExecute
      SetProperties()
      IVPrintButtonDef.Execute()
    End Sub

    Private Sub PrintPreviewButtonDef_OnExecute(Context As NameValueMap) Handles PrintPreviewButtonDef.OnExecute
      SetProperties()
      IVPrintPreviewButtonDef.Execute()
    End Sub

#End Region

  End Class
End Namespace