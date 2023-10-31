#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         CHITRANJAN KUMAR

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <File.au3>
#include <GDIPlus.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>

_GDIPlus_Startup()
Global Const $hGUI = GUICreate("Test", 384, 120)
Global Const $iBtn_Path = GUICtrlCreateButton("Path", 10, 10, 70, 40)
Global Const $iInput_Path = GUICtrlCreateInput("", 90, 20, 160, 20)
GUICtrlSetState(-1, $GUI_DISABLE)
Global Const $iCombo_Ext = GUICtrlCreateCombo("JPG", 270, 20, 100, 21)
GUICtrlSetData($iCombo_Ext, "BMP|PNG", "JPG")
Global Const $iLabel_X = GUICtrlCreateLabel("X:", 20, 80, 20, 10)
Global Const $iInput_X = GUICtrlCreateInput("0", 40, 76, 30, 20, $ES_NUMBER)
Global Const $iLabel_Y = GUICtrlCreateLabel("Y:", 80, 80, 20, 10)
Global Const $iInput_Y = GUICtrlCreateInput("0", 100, 76, 30, 20, $ES_NUMBER)
Global Const $iLabel_W = GUICtrlCreateLabel("W:", 140, 80, 20, 10)
Global Const $iInput_W = GUICtrlCreateInput("0", 160, 76, 30, 20, $ES_NUMBER)
Global Const $iLabel_H = GUICtrlCreateLabel("H:", 200, 80, 20, 10)
Global Const $iInput_H = GUICtrlCreateInput("0", 220, 76, 30, 20, $ES_NUMBER)
Global Const $iBtn_Start = GUICtrlCreateButton("Start", 290, 60, 70, 40)
GUICtrlSetState(-1, $GUI_DISABLE)
GUISetState()

Global $sPath, $aFiles, $i, $iX, $iY, $iW, $iH

While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            _GDIPlus_Shutdown()
            GUIDelete()
            Exit
        Case $iBtn_Path
            $sPath = FileSelectFolder("Select a path with images", "", 0, $hGUI)
            If @error Then
                MsgBox($MB_ICONWARNING, "Warning", "Operation has been aborted", 20, $hGUI)
                ContinueLoop
            EndIf
            GUICtrlSetState($iBtn_Start, $GUI_ENABLE)
            GUICtrlSetData($iInput_Path, $sPath)
        Case $iBtn_Start
            $aFiles = _FileListToArray($sPath, "*." & GUICtrlRead($iCombo_Ext), $FLTA_FILES, True)
            If @error Then
                MsgBox($MB_ICONWARNING, "Warning", "Operation has been aborted", 20, $hGUI)
                ContinueLoop
            EndIf
            $iX = GUICtrlRead($iInput_X)
            $iY = GUICtrlRead($iInput_Y)
            $iW = GUICtrlRead($iInput_W)
            $iH = GUICtrlRead($iInput_H)
            If BitOR($iW = 0, $iH = 0) Then
                MsgBox($MB_ICONWARNING, "Warning", "Please check width / height", 20, $hGUI)
                ContinueLoop
            EndIf
            For $i = 1 To $aFiles[0]
                Crop_Images($iX, $iY, $iW, $iH, $aFiles[$i])
            Next
            MsgBox($MB_ICONINFORMATION, "Done", "Please check out provided folder for the results", 0, $hGUI)
    EndSwitch
WEnd

Func Crop_Images($iX, $iY, $iW, $iH, $sFilename)
    If FileExists($sFilename) Then
        Local $hImage = _GDIPlus_ImageLoadFromFile($sFilename)
        Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
        Local $hGfx = _GDIPlus_ImageGetGraphicsContext($hBitmap)
        _GDIPlus_GraphicsDrawImageRectRect($hGfx, $hImage, $iX, $iY, $iW, $iH, 0, 0, $iW, $iH)
        _GDIPlus_ImageSaveToFile($hBitmap, StringTrimRight($sFilename, 4) & "_P.jpg")
        _GDIPlus_GraphicsDispose($hGfx)
        _GDIPlus_ImageDispose($hBitmap)
        _GDIPlus_ImageDispose($hImage)
        Return True
    EndIf
    Return False
EndFunc