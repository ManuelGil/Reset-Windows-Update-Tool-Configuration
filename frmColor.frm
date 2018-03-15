VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Color"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2070
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   840
         Picture         =   "frmColor.frx":000C
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   120
         Picture         =   "frmColor.frx":05EE
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   840
         Picture         =   "frmColor.frx":0BD0
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   840
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   840
         Picture         =   "frmColor.frx":11B2
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1320
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   840
         Picture         =   "frmColor.frx":1794
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1560
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   840
         Picture         =   "frmColor.frx":1D76
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   840
         Picture         =   "frmColor.frx":2358
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   840
         Picture         =   "frmColor.frx":293A
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   120
         Picture         =   "frmColor.frx":2F1C
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   120
         Picture         =   "frmColor.frx":34FE
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   120
         Picture         =   "frmColor.frx":3AE0
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   120
         Picture         =   "frmColor.frx":40C2
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   120
         Picture         =   "frmColor.frx":46A4
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   600
      End
      Begin VB.PictureBox picBackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   120
         Picture         =   "frmColor.frx":4C86
         ScaleHeight     =   180
         ScaleWidth      =   600
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmColor"
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

' Subroutine Click on Picture picBackcolor
' void picBackcolor_Click()
Private Sub picBackcolor_Click(Index As Integer)
    fMainForm.txtFont.Text = Index
    Unload Me
End Sub
