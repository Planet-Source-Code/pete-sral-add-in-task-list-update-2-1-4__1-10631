Attribute VB_Name = "modGlobals"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const TASK_FILE_NAME As String = "Tasks.dat"

Global gstrTaskFile As String
Global gobjVBInstance  As VBIDE.VBE
Global gwinWindow   As VBIDE.Window
Global gblMouseClick As Integer



Public Function GetCurrentVersion() As String

On Error GoTo GetCurrentVersion_Error

Dim FileHandle As Integer
Dim sText As String, sTemp As String
Dim bEnd As Boolean
Dim sMajor As String, sRev As String, sMinor As String
    
    'Force Save to get recent version for VBP file
    If Len(gobjVBInstance.ActiveVBProject.FileName) <> 0 Then
        If gobjVBInstance.ActiveVBProject.Saved = False Then
            If MsgBox("Save " & gobjVBInstance.ActiveVBProject.Name & "?", vbYesNo, "Save Project") = vbYes Then
                gobjVBInstance.ActiveVBProject.SaveAs (gobjVBInstance.ActiveVBProject.FileName)
            End If
        End If
    Else
        MsgBox "Can not get version until project is saved!", vbInformation, "Getting Current Version"
        Exit Function
    End If
    
    FileHandle = FreeFile
    Open gobjVBInstance.ActiveVBProject.FileName For Input As FileHandle
    Do Until bEnd
        Line Input #FileHandle, sText
        
        sTemp = GetToken(sText, "=")
        
        'MajorVer = 1
        'MinorVer = 0
        'RevisionVer = 0
        
        Select Case UCase(sTemp)
            Case Is = "MAJORVER"
                sMajor = sText
            Case Is = "MINORVER"
                sMinor = sText
            Case Is = "REVISIONVER"
                sRev = sText
                bEnd = True
        End Select
    Loop
    Close FileHandle
    'MsgBox Trim$(sMajor) & "." & Trim$(sMinor) & "." & Trim$(sRev)
    GetCurrentVersion = Trim$(sMajor) & "." & Trim$(sMinor) & "." & Trim$(sRev)
    
    
GetCurrentVersion_Error:
    Close FileHandle
    DoError "modGlobals", "GetCurrentVersion", Err
End Function

Public Sub SetListViewColor(pCtrlListView As ListView, pCtrlPictureBox As PictureBox)

On Error GoTo SetListViewColor_Error
    
    Dim iLineHeight As Long
    Dim iBarHeight As Long
    Dim lBarWidth As Long
    Dim lColor1 As Long
    Dim lColor2 As Long
    
    ' Creates a color bar background for a ListView when in
    ' report mode. Passing the listview and picturebox allows
    ' you to use this with more than one control. You can also
    ' change the colors used for each by passing new RGB color
    ' values in the optional color parameters.
    
    
    'set picture to none and exit sub if not in report mode
    
    lColor1 = IIf(Len(GetFromIni("GridColor", "Color1", gstrTaskFile)) <> 0, GetFromIni("GridColor", "Color1", gstrTaskFile), vbWhite)
    lColor2 = IIf(Len(GetFromIni("GridColor", "Color2", gstrTaskFile)) <> 0, GetFromIni("GridColor", "Color2", gstrTaskFile), vbWhite)
    
    If pCtrlListView.View = lvwReport Then
        pCtrlListView.Picture = LoadPicture("")
        pCtrlListView.Refresh
        
        pCtrlPictureBox.Cls
        
        'these can be commented out if the pCtrlPictureBox control
        'is set correctly.
        pCtrlPictureBox.AutoRedraw = True
        pCtrlPictureBox.BorderStyle = vbBSNone
        pCtrlPictureBox.ScaleMode = vbTwips
        pCtrlPictureBox.Visible = False
        
        'set the alignment to "Tile" and you only need
        'two bars of color.
        pCtrlListView.PictureAlignment = lvwTile
    
        'needed because ListView does not have "TextHeight"
        pCtrlPictureBox.Font = pCtrlListView.Font
    
        
        
        pCtrlPictureBox.Font = pCtrlListView.Font
        With pCtrlPictureBox.Font
            .Size = pCtrlListView.Font.Size
            .Bold = pCtrlListView.Font.Bold
            .Charset = pCtrlListView.Font.Charset
            .Italic = pCtrlListView.Font.Italic
            .Name = pCtrlListView.Font.Name
            .Strikethrough = pCtrlListView.Font.Strikethrough
            .Underline = pCtrlListView.Font.Underline
            .Weight = pCtrlListView.Font.Weight
        End With
        pCtrlPictureBox.Refresh
        
        'set height to a single line of text plus a
        'one pixel spacer.
        iLineHeight = pCtrlPictureBox.TextHeight("W") + Screen.TwipsPerPixelY
    
        'set color bars to 3-line wide.
        iBarHeight = (iLineHeight * 1)
        lBarWidth = pCtrlListView.Width
    
        'resize the pCtrlPictureBox picturebox
        pCtrlPictureBox.Height = iBarHeight * 2
        pCtrlPictureBox.Width = lBarWidth
    
        'paint the two bars of color
        pCtrlPictureBox.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
        pCtrlPictureBox.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
        
        pCtrlPictureBox.AutoSize = True
        'set the pCtrlListView picture to the
        'pCtrlPictureBox image
        pCtrlListView.Picture = pCtrlPictureBox.Image
    Else
        pCtrlListView.Picture = LoadPicture("")
    End If
    
    pCtrlListView.Refresh
    Exit Sub
SetListViewColor_Error:
    'clear pCtrlListView's picture and then exit
    pCtrlListView.Picture = LoadPicture("")
    pCtrlListView.Refresh
End Sub
Public Function DoError(psModule As String, psProc As String, plErr As Long)

   On Error Resume Next
   
   If plErr <> 0 Then
      MsgBox "Error: " & Str(plErr) & " - " & Error(plErr) & " in Module: " & psModule & " in Procedure: " & psProc, vbCritical
   End If
   
End Function

Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String

   On Error GoTo GetFromIni_Error
   
   Dim strReturn As String
   
   strReturn = String(255, Chr(0))
   GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
   
   Exit Function
   
GetFromIni_Error:
   DoError "modGlobals", "GetFromIni", Err
    
End Function

Function WriteToIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer

   On Error GoTo WriteToIni_Error
   WriteToIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)

   Exit Function
   
WriteToIni_Error:
   DoError "modGlobals", "WriteToIni", Err
   
End Function

Public Function FileExists(strFile As String) As Boolean
   
   On Error Resume Next 'Doesn't raise error - FileExists will be False
   
   FileExists = Dir(strFile, vbHidden) <> ""

End Function

Public Function GetToken(sSrc As String, sDelimit As String)
    
   Dim ilast As Integer, iLoop As Integer, iPos As Integer
   Dim sToken As String
   
   If sDelimit = "" Then sDelimit = ","
   
   ilast = 32767
   
   For iLoop = 1 To Len(sDelimit)
      iPos = InStr(sSrc, Mid$(sDelimit, iLoop, 1))
      If iPos <> 0 And iPos < ilast Then ilast = iPos
   Next
   
   If ilast <> 32767 Then
      If ilast = 1 Then
         sToken = ""
      Else
         sToken = Mid$(sSrc, 1, ilast - 1)
      End If
      sSrc = Mid$(sSrc, ilast + 1)
   Else
      sToken = sSrc
      sSrc = ""
   End If
   
   GetToken = sToken

End Function
