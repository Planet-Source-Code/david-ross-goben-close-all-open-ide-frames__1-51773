VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'NOTE: In the Connect.dsr file, in the AddinInstance_OnConnection() event,
'be sure to chage the default caption in the AddToAddInCommandBar()
'function to the name that you want to see in the Add-Ins menu in the IDE
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'****************************************************************************
'COMPILATION NOTES:
'
' 1. Compile the VBCloseALL.DLL either to the project folder, or to your
'    \Windows\System32 (Windows\System for Win95\98\ME) folder. If compiled
'    to the project folder, copy it to your System32 (or System) folder.
'
' 2. Exit VB, then re-enter it. Fram the "Add-Ins" menu,choose
'    "Add-In Manager...". Find the "VB Development Close All Windows" entry
'    and insure that the "Loaded/Unloaded" and "Load on Startup" are checked,
'    then hit OK.
'
' You should now see "VB Dev Close All Windows" in the Add-Ins menu. Select it
' anytime you want to close all open forms and code modules.
'----------------------------------------------------------------------------
' IMPORTANT NOTE:
' If you are updating an Add-in, BE SURE to first unclock the Loaded/Unloaded
' open in the Add-In Manager (it doesn't hurt to also uncheck Load on Startup.
' This way you can write the new DLL without it yelling at you about access
' being denied because it is in use.
'
' Also, I've noticed that when you exit VB after compiling an Add-in, it
' suffers a small (but not harmful) conniption and issues a warning. Don't
' sweat it. You can cheat by opening up a different project and then exiting.
'----------------------------------------------------------------------------
' COOL TIP:
' BY THE WAY. If you want a particular project to always open up empty, or
' with certain file frames opened (except graphical forms, which never open on
' startup), you can force this by setting up your display (including code
' fram positioning), exiting VB, and editing the properties on the project's
' *.VBW file (associated with the *.VBP file). Change its Attribute to
' Read-Only. This way VB will always load its settings, and will not update
' it on app exit, even if you exited with all frames closed. VB will not
' complain if this file is read-only (it can't, since Microsoft wants VB to
' be able to access source code on CD and DVD discs).
'****************************************************************************
Option Explicit

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : close either code or form design windows that are visible
'*******************************************************************************
Private Sub Form_Load()
  Dim w As Window
  
  For Each w In VBInstance.Windows
    If (w.Type = vbext_wt_CodeWindow Or _
        w.Type = vbext_wt_Designer) And w.Visible Then w.Close
  Next w
  Unload Me
End Sub

