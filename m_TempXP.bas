Attribute VB_Name = "m_TempXP"
Private Function CPopupMenu(ParentID As Long, Optional hPopMenu As Long = 0) As Long
  
  Dim hMenu As Long
  Dim hMenu1 As Long
  Dim sItem As String      'popisok MenuItem
  Dim bDummy As Boolean
  Dim lFlags As Long       'flags pro menuitem
  Dim sPicture As String   'meno obrazku
  Dim lIdx As Long         'do for cyklu
  Dim Ret As Long

    hMenu = IIf(hPopMenu = 0, CreateMenu(), hPopMenu)

    For lIdx = 1 To lArr
    
        If Caps(lIdx, 3) = ParentID Then

            sItem = Caps(lIdx, 1)
            If Caps(lIdx, 4) = "A" Then

                hMenu1 = CPopupMenu(lIdx)
                bDummy = AppendMenu(hMenu, MF_POPUP + MF_OWNERDRAW + MF_STRING, hMenu1, ByVal sItem)
                Caps(lIdx, 5) = hMenu1
                
            Else
                
                lFlags = MF_OWNERDRAW + MF_STRING
                If sItem = "-" Then
                    lFlags = MF_SEPARATOR + lFlags
                End If

                bDummy = AppendMenu(hMenu, lFlags, lIdx, ByVal sItem)
                Caps(lIdx, 5) = lIdx
                
            End If
            
        End If
        
    Next lIdx

    CPopupMenu = hMenu

End Function

Public Sub CInitMenu()

  'naplnime informace o menu

    Caps(1, 1) = "&File"
    Caps(1, 2) = ""
    Caps(1, 3) = "0"
    Caps(1, 4) = "A"
    Caps(1, 6) = "mnuFile"
    Caps(1, 7) = ""

    Caps(2, 1) = "&Open"
    Caps(2, 2) = "open"
    Caps(2, 3) = "1"
    Caps(2, 4) = "N"
    Caps(2, 6) = "mnuOpen"
    Caps(2, 7) = "Otvoriù s˙bor ..."

    Caps(3, 1) = "&Save"
    Caps(3, 2) = "save"
    Caps(3, 3) = "1"
    Caps(3, 4) = "N"
    Caps(3, 6) = "mnuSave"
    Caps(3, 7) = "Uloûiù s˙bor ..."

    Caps(4, 1) = "-"
    Caps(4, 2) = ""
    Caps(4, 3) = "1"
    Caps(4, 4) = "N"
    Caps(4, 6) = "mnuLine1"
    Caps(4, 7) = ""


    Caps(5, 1) = "&Konec"
    Caps(5, 2) = ""
    Caps(5, 3) = "1"
    Caps(5, 4) = "N"
    Caps(5, 6) = "mnuEnd"
    Caps(5, 7) = "UkonËiù program"

    Caps(6, 1) = "&Popup"
    Caps(6, 2) = "3"
    Caps(6, 3) = "0"
    Caps(6, 4) = "A"
    Caps(6, 6) = "mnuPopup"
    Caps(6, 7) = ""

    Caps(7, 1) = "&Pokus1"
    Caps(7, 2) = ""
    Caps(7, 3) = "6"
    Caps(7, 4) = "A"
    Caps(7, 6) = "mnuPokus1"
    Caps(7, 7) = ""


    Caps(8, 1) = "&PodPokus1/2"
    Caps(8, 2) = "open"
    Caps(8, 3) = "7"
    Caps(8, 4) = "N"
    Caps(8, 6) = "mnuPokus2"
    Caps(8, 7) = "PokusnÈ menu 1"


    Caps(9, 1) = "P&odPokus2"
    Caps(9, 2) = "save"
    Caps(9, 3) = "7"
    Caps(9, 4) = "N"
    Caps(9, 6) = "mnuPokus3"
    Caps(9, 7) = "PokusnÈ menu 2"

    lArr = 9 'm·me celkom 9 MenuItems

End Sub

Public Sub CSetupMenu(hwnd As Long)

  Dim hMenu As Long
  Dim lIndex As Long
  Dim bDummy As Boolean

    hMainMenu = CreateMenu()

    For lIndex = 1 To lArr
        If Caps(lIndex, 3) = "0" Then
            hMenu = CPopupMenu(lIndex)
            'bDummy = AppendMenu(hMainMenu, MF_POPUP + MF_STRING + MF_OWNERDRAW, hmenu, ByVal Caps(lIndex, 1))
            bDummy = AppendMenu(hMainMenu, MF_POPUP + MF_STRING, hMenu, ByVal Caps(lIndex, 1))
            Caps(lIndex, 5) = hMenu
        End If
    Next lIndex

    bDummy = SetMenu(hwnd, hMainMenu)

End Sub

Public Function FillRectTmp(hwnd As Long, m_DC As Long) As Boolean

    Dim Rec As RECT, nRec As RECT

            GetWindowRect hwnd, Rec
            
            nRec = Rec
            nRec.Right = nRec.Right - nRec.Left
            nRec.Bottom = nRec.Bottom - nRec.Top
            nRec.Left = 0: nRec.Top = 0
            
            'hColorFill = CreateSolidBrush(RGB(255, 251, 247))
            'SelectObject m_DC, hColorFill
            'FillRect m_DC, nRec, hColorFill
            'DeleteObject hColorFill
            
            'nastavenie per a ötetca pre kreslenie obdlûnika
            hBrFill = CreateSolidBrush(RGB(246, 246, 246))  'farba v˝beru
            hPenFill = CreatePen(0, 1, RGB(102, 102, 102))     'farba okraja
            
            'uloûÌme info o starom pere a ötetci
            hOldBrFill = SelectObject(m_DC, hBrFill)
            hOldPenFill = SelectObject(m_DC, hPenFill)
            
            Rectangle m_DC, 0, 0, nRec.Right - nRec.Left - 4, nRec.Bottom - nRec.Top - 4
            
            'nastavenie pÙvodnÈho pera a ötetca
            Call SelectObject(m_DC, hOldBrFill)
            Call SelectObject(m_DC, hOldPenFill)
    
            'zmazanie nami vytvorenÈho brush a pen
            Call DeleteObject(hBrFill)
            Call DeleteObject(hPenFill)
            
End Function

Private Sub PrintGlyph(hdc As Long, Glyph As String, Color As Long, rt As RECT, ByVal wFormat As Long)
'glyph pre öÌpku je 3 alebo 4

  'Create the Marlett font if it doesn't exist already
  If m_MarlettFont& = 0& Then
    'Dim tLF As LOGFONT
  
    tLF.lfFaceName = "Marlett" + Chr(0)
    tLF.lfCharSet = SYMBOL_CHARSET
    tLF.lfHeight = 13 'the value could be changed in relation to the real MenuFont to draw proportional boxes
'    tLF.lfWeight = 500
'    tLF.lfWidth = 31

    'm_MarlettFont& = CreateFontIndirect(tLF)
  End If

  'write text with transparent background
  Call SetBkMode(hdc&, TRANSPARENT)
    
  Dim hOldFont As Long
  
  'Select the font for the device context
  hOldFont& = SelectObject(hdc&, m_MarlettFont&)
  
  'select the color for the glyph
  Call SetTextColor(hdc&, Color&)
  
  Call DrawText(hdc&, Glyph, 1, rt, wFormat&)

End Sub
