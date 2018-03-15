Attribute VB_Name = "modLAnguage"
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

' This subroutine load the language
' void SetLanguageFile()
Public Sub SetLanguageFile()
    ' Declare the variables
    Dim strPath As String
    Dim strFound As String
    
    ' Set the path of the Application
    If Right(App.Path, 1) <> "\" Then
        strPath = App.Path & "\wureset\lang\*.txt"
    Else
        strPath = App.Path & "wureset\lang\*.txt"
    End If
    
    ' Search for ALL avaibles Languages
    strFound = Dir(strPath)
    
    Do Until strFound = ""
        ' Add the language to the List and cut the ".txt"
        fMainForm.cmbLanguage.AddItem strFound
        strFound = Dir
    Loop
End Sub

