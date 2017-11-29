Attribute VB_Name = "modINIFile"
'    This file is part of WURESET Configuration Project
'    WURESET Config Free GNU Application
'    Copyright (C) 2017 Manuel Gil.
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
' $version: 1.0.0.2. $
' -----------------------------------------------------------------

Option Explicit

' Declare Function GetPrivateProfileString% Lib "Kernel"
'            (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$,
'            ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
'
' @param lpAppName$         Name of a Windows-based application that appears in
'                           the initialization file.
'
' @param lpKeyName$         Key name that appears in the initialization file.
'
' @param nDefault$          Specifies the default value for the given key if the
'                           key cannot be found in the initialization file.
'
' @param lpFileName$        Points to a string that names the initialization
'                           file. If lpFileName does not contain a path to the
'                           file, Windows searches for the file in the Windows
'                           directory.
'
' @param lpDefault$         Specifies the default value for the given key if the
'                           key cannot be found in the initialization file.
'
' @param lpReturnedString$  Specifies the buffer that receives the character
'                           string.
'
' @param nSize%             Specifies the maximum number of characters (including
'                           the last null character) to be copied to the buffer.
'
' @param lpString$          Specifies the string that contains the new key value.
'
' GetPrivateProfileString(lpAppName -> pSection, lpKeyValue -> pKey, nDefault -> pDefaultValue, _
'     lpReturnString -> strReturnString, nSize -> pSize, lpFileName -> pFileName)
'
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpsettingsFile As String) As Long

' Declare Function WritePrivateProfileString% Lib "Kernel"
'            (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$,
'            ByVal lpFileName$)
'
' @param lpAppName$         Name of a Windows-based application that appears in
'                           the initialization file.
'
' @param lpKeyName$         Key name that appears in the initialization file.
'
' @param nDefault$          Specifies the default value for the given key if the
'                           key cannot be found in the initialization file.
'
' @param lpFileName$        Points to a string that names the initialization
'                           file. If lpFileName does not contain a path to the
'                           file, Windows searches for the file in the Windows
'                           directory.
'
' @param lpDefault$         Specifies the default value for the given key if the
'                           key cannot be found in the initialization file.
'
' @param lpReturnedString$  Specifies the buffer that receives the character
'                           string.
'
' @param nSize%             Specifies the maximum number of characters (including
'                           the last null character) to be copied to the buffer.
'
' @param lpString$          Specifies the string that contains the new key value.
'
' WritePrivateProfileString(lpAppName -> pSection, lpKeyName -> pKey, lpString -> pValue, lpFileName -> pFile)
'
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpsettingsFile As String) As Long

' -----------------------------------------------------------------
' Attributes
' -----------------------------------------------------------------
Public systemName As String
Public systemVersion As String
Public systemArchitecture As String
Public programLanguage As String
Public programFont As String

' -----------------------------------------------------------------
' Methods
' -----------------------------------------------------------------

' This subroutine load the settings
' void LoadSettings()
Public Sub LoadSettings()
    ' Declare the variables
    Dim line As String
    
    ' Generally, API do not expect string buffers longer than 255 character
    line = String(255, vbNullChar)
    
    ' Get String for System Name
    GetPrivateProfileString "system", "name", "", line, Len(line), settingsFile
    systemName = Trim(Left(line, InStr(line, vbNullChar) - 1))
    
    ' Get String for System Version
    GetPrivateProfileString "system", "version", "", line, Len(line), settingsFile
    systemVersion = Trim(Left(line, InStr(line, vbNullChar) - 1))
    
    ' Get String for System Architecture
    GetPrivateProfileString "system", "architecture", "", line, Len(line), settingsFile
    systemArchitecture = Trim(Left(line, InStr(line, vbNullChar) - 1))
    
    ' Get String for Program Language
    GetPrivateProfileString "program", "language", "", line, Len(line), settingsFile
    programLanguage = Trim(Left(line, InStr(line, vbNullChar) - 1))
    
    ' Get String for Program Font
    GetPrivateProfileString "program", "font", "", line, Len(line), settingsFile
    programFont = Trim(Left(line, InStr(line, vbNullChar) - 1))
End Sub

' This subroutine save the settings
' void SaveSettings()
Public Sub SaveSettings()
    ' Set String of Program Language
    WritePrivateProfileString "program", "language", programLanguage, settingsFile
    
    ' Set String of Program Font
    WritePrivateProfileString "program", "font", programFont, settingsFile
End Sub
