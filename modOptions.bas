Attribute VB_Name = "Modoptions"
' *************************************************************************************************
' * This Module Contains The main SOurce for Setting and Retrieving Options                       *
' * For NextPad , It Basically Sets options By BOOLEAN True (-1) , FALSE (0)                      *
' * by Using this Method we can save Space On the Usrs Harddisk and Make the Final Executable     *
' * Smaller than expected , In the Past Options Were Confronted With Barberic If Then             *
' * Statements for one thing , Now the Option Can be Retreived , Set and Makes everything Else    *
' * Run Smoother ......................                                                           *
' *************************************************************************************************




Type ToolbarBOOL
visible As Integer
End Type
Public isToolbarvisible As ToolbarBOOL

Type ExternalEditor
use As Integer
End Type
Public UseExternalEditor As ExternalEditor

Public Usewordwrap As Boolean

Public AskIfTooBig As Boolean

Public Check_Associations_at_Startup As Boolean
'********************************************************************

Sub SaveSetting_Toolbar(BOOLobject As Boolean)
    
    Form1.Toolbar.visible = BOOLobject
    
    Form1.mnuhidetoolbaritem.Checked = Form1.Toolbar.visible
    
    Resizenotewithtoolbar ' resize the form
        
    SaveRegistryString "Toolbar", "Visible", Abs(CInt(CBool(Form1.Toolbar.visible)))
 
End Sub

Sub saveSetting_UseExternalEditor(BOOLobject As Boolean)
   Dim retval As Long 'declare Variables
    retval = IIf(BOOLobject, True, False)
    Select Case retval
      Case True
       SaveRegistryString "UseExternalEditor", "Use", 1
     Case False
       SaveRegistryString "UseExternalEditor", "Use", 0
    End Select
End Sub
 

Sub SaveSetting_Wordwrap(BOOLobject As Boolean)

 
 
  Usewordwrap = BOOLobject
           
  Togglewordwrap (BOOLobject) ' togglewordwrap (Retval as BOOLEAN)

  SaveRegistryString "Wordwrap", "Wordwrap", Abs(CInt(CBool(BOOLobject)))
  
     
End Sub
Sub SaveSetting_chckassociations(BOOLobject As Boolean)
  Dim retval As Long 'declare Variables
    retval = IIf(BOOLobject, True, False)
      Select Case retval
        Case True
          SaveRegistryString "chckassociations", "show", 1
        Case False
          SaveRegistryString "chckassociations", "show", 0
      End Select
End Sub
Sub SaveSetting_Prioritylevel(nLevel As Integer)
    
      SaveRegistryString "Priority", "Level", nLevel
    
End Sub
Sub SaveSetting_AskIfToobig(BOOLobject As Boolean)
  Dim retval As Long 'declare Variables
    retval = IIf(BOOLobject, True, False)
      
      Select Case retval
       Case True
        SaveRegistryString "Misc", "AskifTooBig", 1
         AskIfTooBig = True
       Case False
        SaveRegistryString "Misc", "AskifTooBig", 0
         AskIfTooBig = False
      End Select
End Sub
Sub RetrieveALLSettings()
Dim externaleditorval As String 'declare Variables
Dim toolbarval As String 'declare Variables
Dim wordwrapval As String 'declare Variables
Dim Chckassociationsval As String 'declare Variables
Dim AskIftoobigVal As String 'declare Variables
Dim nPriorityLevel As Integer 'declare Variables
Dim rtval As Integer 'declare Variables
Dim Response 'declare Variables
   On Error GoTo RegistrySettingsError
' Toolbar
   
   toolbarval = GetSetting("NextPad", "Toolbar", "Visible")
     Select Case toolbarval
       Case 1
         isToolbarvisible.visible = True
         Form1.Toolbar.visible = True
         Form1.mnuhidetoolbaritem.Checked = True
       Case 0
        isToolbarvisible.visible = False
        Form1.Toolbar.visible = False
        Form1.mnuhidetoolbaritem.Checked = False
     End Select

' External Editor
    externaleditorval = GetSetting("NextPad", "UseExternalEditor", "Use")
      Select Case externaleditorval
        Case 1
         UseExternalEditor.use = True
        Case 0
         UseExternalEditor.use = False
      End Select

'Word wrap
    wordwrapval = GetSetting("NextPad", "Wordwrap", "Wordwrap")
      Select Case wordwrapval
        Case 1
         Usewordwrap = True
        Case 0
         Usewordwrap = False
      End Select

' check Associations
   Chckassociationsval = GetSetting("NextPad", "chckassociations", "show")
     Select Case Chckassociationsval
      Case 1
       Check_Associations_at_Startup = True
      Case 0
       Check_Associations_at_Startup = False
     End Select

'Ask If too big
   AskIftoobigVal = GetSetting("nextPad", "Misc", "AskIftoobig")
     Select Case AskIftoobigVal
      Case 1
       AskIfTooBig = True
      Case 0
       AskIfTooBig = False
      End Select
      
    nPriorityLevel = GetSetting("NextPad", "Priority", "Level")
     SetPriority (nPriorityLevel)
      
RegistrySettingsError:
   If Err.Number <> 0 Then
       Response = MsgBox("An Error Has Occured While Accessing One or More Registry Entrys , Settings May Be Missing Or Corrupt." _
       & vbNewLine & vbNewLine & "Repair Settings ?" & vbNewLine & vbNewLine & "Click Yes Too Repair Settings", vbYesNo + vbExclamation, "Error,Options")
     Select Case Response
       Case vbYes
         SaveSetting_Toolbar True
         SaveSetting_chckassociations False
         saveSetting_UseExternalEditor True
         SaveSetting_Prioritylevel 10
         SaveSetting_Wordwrap True
         SaveSetting_AskIfToobig True
         SaveRegistryString "associations", "isassociated", "0"
         MsgBox "All Settings Were Saved Successfully.", vbInformation, "Successful"
         Exit Sub
       Case vbNo
         MsgBox "Settings Were Chosen Not To be Repaired This Message Will Continue Too appear Until you do so.", vbCritical, "Settings Not Saved"
         Exit Sub
     End Select
End If



End Sub
