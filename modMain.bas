Attribute VB_Name = "modMain"
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
' Constants
' -----------------------------------------------------------------
Public Const program = "WURESET Config"

' -----------------------------------------------------------------
' Attributes
' -----------------------------------------------------------------
Public settingsFile As String

' -----------------------------------------------------------------
' Relations
' -----------------------------------------------------------------
Public fMainForm As frmMain

' -----------------------------------------------------------------
' Methods
' -----------------------------------------------------------------

' This subroutine load the program
' void Main()
Public Sub Main()
    ' Declare the variables
    Dim fMainShow As String
    Dim appData As String
    
    ' Single Instance
    If App.PrevInstance Then End

    appData = Environ("APPDATA")
    
    ' Set the path of the Application
    If Right(appData, 1) = "\" Then
        settingsFile = appData & "wureset\settings.ini"
    Else
        settingsFile = appData & "\wureset\settings.ini"
    End If
    
    ' Show frmMain
    Set fMainForm = New frmMain
    Load fMainForm
    fMainShow = GetSetting(program, "Restore", "Start.Main", fMainShow)
    If fMainShow <> "0" Then fMainForm.Show
End Sub
