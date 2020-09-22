VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save"
   ClientHeight    =   75
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   75
   ScaleWidth      =   1200
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'This code was taken from:-
'API-Guide 3.7
'Copyright © The KPD-Team, 1998-2002

    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
'--------------------------------------------------------
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub Form_Load()
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = Me.hWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    OFName.lpstrFilter = strFileToOpen$ + Chr$(0) + strTypeToOpen$ + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the title
    OFName.lpstrTitle = "Open File - KPD-Team 1998"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        strFileToOpen$ = Trim$(OFName.lpstrFile)
        frmBASSMK2.Enabled = True
        frmBASSMK2.Hide
        frmBASSMK2.Show
        Unload Me
    Else
        'if cancel is clicked then exit
        strFileToOpen$ = "////"
        frmBASSMK2.Enabled = True
        frmBASSMK2.Hide
        frmBASSMK2.Show
        Unload Me
    End If
End Sub
'--------------------------------------------------------

