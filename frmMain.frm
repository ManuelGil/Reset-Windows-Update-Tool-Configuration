VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WURESET Config"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3975
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Program Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
      Begin VB.CommandButton btnColor 
         Caption         =   "..."
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtFont 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblLanguage 
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblFont 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.Label lblSystemArchitecture 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSystemArchitecture"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblSystemVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSystemVersion"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblSystemName 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSystemName"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    This file is part of WURESET Configuration Project
'    WURESET Config Free GNU Application
'    Copyright (C) 2018 Manuel Gil.
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http:'www.gnu.org/licenses/>.
'

' -----------------------------------------------------------------
' $Id$
' $Title: WURESET Config Free GNU Application. $
' $Description: Configuration for Reset Windows Update Tool. $
' $Copyright: GPL product. $
'
' $Author: Manuel Gil. $
' $version: 1.0.0.3. $
' -----------------------------------------------------------------

Option Explicit

' -----------------------------------------------------------------
' Methods
' -----------------------------------------------------------------

' This subroutine load the Form
' void Form_Load()
Private Sub Form_Load()
    LoadSettings
    SetLanguageFile
    
    lblSystemName.Caption = "Name: " & systemName
    lblSystemVersion.Caption = "Version: " & systemVersion
    lblSystemArchitecture.Caption = "Architecture: " & systemArchitecture & "-bits"
    
    ' "Colors:" & vbCrLf & _
    ' "    2 = Green       9 = Light Blue" & vbCrLf & _
    ' "    3 = Aqua       10 = Light Green" & vbCrLf & _
    ' "    4 = Red        11 = Light Aqua" & vbCrLf & _
    ' "    5 = Purple     12 = Light Red" & vbCrLf & _
    ' "    6 = Yellow     13 = Light Purple" & vbCrLf & _
    ' "    7 = White      14 = Light Yellow" & vbCrLf & _
    ' "    8 = Gray       15 = Bright White"
    
    Dim i As Integer
    
    For i = 0 To cmbLanguage.ListCount - 1
        If cmbLanguage.List(i) = programLanguage Then
            cmbLanguage.ListIndex = i
        End If
    Next i
    
    txtFont.Text = programFont
End Sub

' Subroutine Click on Button btnColor
' void cmdColor_Click()
Private Sub btnColor_Click()
    frmColor.Show vbModal
End Sub

' Subroutine Click on Button btnSave
' void btnSave_Click()
Private Sub btnSave_Click()
    ' If (Text on txtFont is number) And_
    ' (Text on txtFont major to 2 Or minor to 15) save settings
    If Not IsNumeric(txtFont.Text) Then
        MsgBox "Invalid Color", vbOKOnly
        txtFont.Text = "7"
        txtFont.SetFocus
    ElseIf CInt(txtFont.Text) < 2 Or CInt(txtFont.Text) > 15 Then
        MsgBox "Invalid Color", vbOKOnly
        txtFont.Text = "7"
        txtFont.SetFocus
    Else
        ' If ComboBox cmbLanguage isn't Empty
        ' Program Language = Text on ComboBox cmbLanguage
        If cmbLanguage <> "" Then
            programLanguage = cmbLanguage.Text
        Else
            programLanguage = ""
        End If
        
        ' Program Font = Text on TextBox txtFont
        programFont = txtFont.Text
        
        ' Save Program Language And Program Font
        SaveSettings
        MsgBox "The operation completed successfully."
        Unload Me
    End If
End Sub
