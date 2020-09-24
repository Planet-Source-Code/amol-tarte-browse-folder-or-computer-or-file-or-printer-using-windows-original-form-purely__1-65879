VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Browser"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   90
      TabIndex        =   10
      Top             =   90
      Width           =   6015
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   945
         Left            =   3990
         Picture         =   "Form1.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1380
         Width           =   1785
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   945
         Left            =   2070
         Picture         =   "Form1.frx":1594
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1380
         Width           =   1785
      End
      Begin VB.CommandButton cmdFiles 
         Caption         =   "Browse For Folder Including Files"
         Height          =   945
         Left            =   150
         Picture         =   "Form1.frx":180A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1380
         Width           =   1785
      End
      Begin VB.CheckBox chkDisplayEditBox 
         Caption         =   "Display Edit Box"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   3390
         Width           =   2175
      End
      Begin VB.CheckBox chkDontRootDir 
         Caption         =   "Dont Use Root Directories"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CheckBox chkDontBelowDomain 
         Caption         =   "Dont Go Below Domain"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   2850
         Width           =   2055
      End
      Begin VB.CheckBox chkNewLooks 
         Caption         =   "Use New Looks"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   2580
         Width           =   1755
      End
      Begin VB.CommandButton cmdComputer 
         Caption         =   "Browse For &Computer"
         Height          =   945
         Left            =   150
         Picture         =   "Form1.frx":20D4
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1785
      End
      Begin VB.CommandButton cmdFolder 
         Caption         =   "Browse For &Folder"
         Height          =   945
         Left            =   3990
         Picture         =   "Form1.frx":299E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1785
      End
      Begin VB.CommandButton cmdPrinter 
         Caption         =   "Browse For &Printer"
         Height          =   945
         Left            =   2070
         Picture         =   "Form1.frx":3668
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblMessage 
         Caption         =   "Remember: If you Check Use New Looks And / Or Display Edit Box, The Ok button Of Windows Form will always be enabled."
         Height          =   735
         Left            =   2730
         TabIndex        =   11
         Top             =   2460
         Width           =   2985
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sOutPut As String
Dim sMsg    As String

Private Sub chkDisplayEditBox_Click()
    
    If chkDisplayEditBox.Value = 1 Then chkNewLooks.Value = 1
End Sub

Private Sub cmdAbout_Click()

    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdComputer_Click()
        
    sOutPut = ""
    sMsg = "Browse For Computer?"
    sMsg = sMsg & " The Ok button will not be enabled until you select a valid Computer."
    
    Call BrowseForComputer(Me.hWnd, "C:\", sMsg, chkDontBelowDomain.Value)
End Sub

Private Sub cmdExit_Click()

    Unload Me
    Set frmMain = Nothing
End Sub

Private Sub cmdFiles_Click()
    sOutPut = ""
    sMsg = "Want to browse for Folder?"
    If chkNewLooks.Value = 1 Or chkDisplayEditBox.Value = 1 Then
        sMsg = sMsg _
        & " The Ok button will always be enabled even if you dont" _
        & " select a valid folder because you have selected New Looks or Display Edit Box."
    Else
        sMsg = sMsg & " The Ok button will not be enabled until you select a valid folder or file."
    End If
    sOutPut = BrowseForFolder(Me.hWnd, CurDir, sMsg, True, chkNewLooks.Value, chkDontBelowDomain.Value, chkDontRootDir.Value, chkDisplayEditBox.Value)
    If sOutPut <> "" Then
        MsgBox "You Selected: " & sOutPut, vbInformation
    Else
        MsgBox "You did not select anything?", vbQuestion
    End If
End Sub

Private Sub cmdFolder_Click()
        
    sOutPut = ""
    sMsg = "Want to browse for Folder?"
    If chkNewLooks.Value = 1 Or chkDisplayEditBox.Value = 1 Then
        sMsg = sMsg _
        & " The Ok button will always be enabled even if you dont" _
        & " select a valid folder because you have selected New Looks or Display Edit Box."
    Else
        sMsg = sMsg & " The Ok button will not be enabled until you select a folder."
    End If
     
    sOutPut = BrowseForFolder(Me.hWnd, CurDir, sMsg, False, chkNewLooks.Value, chkDontBelowDomain.Value, chkDontRootDir.Value, chkDisplayEditBox.Value)
    If sOutPut <> "" Then
        MsgBox "Your Selected Folder: " & sOutPut, vbInformation
    Else
        MsgBox "You did not select anything?", vbQuestion
    End If
End Sub

Private Sub cmdPrinter_Click()
    
    sOutPut = ""
    If chkNewLooks.Value = 1 Or chkDisplayEditBox.Value = 1 Then
        sMsg = sMsg _
        & " The Ok button will always be enabled even if you dont" _
        & " select a valid printer because you have selected New Looks or Display Edit Box."
    Else
        sMsg = sMsg & " The Ok button will not be enabled until you select a printer."
    End If

    sOutPut = BrowseForPrinter(Me.hWnd, "", "Want to search for printers?" & vbNewLine & "The Ok button will be disabled until you select a printer.", chkDontBelowDomain.Value)
    If sOutPut <> "" Then
        MsgBox "Your Selected Printer: " & sOutPut, vbInformation
    Else
        MsgBox "You did not select anything?", vbQuestion
    End If
End Sub
