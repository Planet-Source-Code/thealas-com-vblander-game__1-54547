VERSION 5.00
Begin VB.Form frmLander 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Lander"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8880
   FillStyle       =   0  'Solid
   Icon            =   "frmLander.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New game"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameStart 
         Caption         =   "&Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGamePause 
         Caption         =   "&Pause"
         Shortcut        =   ^P
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameAbout 
         Caption         =   "&About game"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmLander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'*  (C) Sala Bojan 2004, alas@eunet.yu
'*  You may use any file as you like, if you liked it, you can vote.
'*  This is not finished game, just a simple 'Lander' demo.
'*  It can be played on any computer with Win32, with the same quality
'**********************************************************************

Option Explicit

' Some holders for DC's we will use in the game
Dim hLunar              As HHOLDER ' Lunar module
Dim hlSide              As HHOLDER ' Side fire, left and right image...
Dim hlDown              As HHOLDER ' Down fire image
Dim hlCrashed           As HHOLDER ' Crashed lunar module image

Dim hLevel              As HHOLDER ' Moon surface

' Location data
Dim hlLeft              As POINTAPI ' This is a location of the left lander's leg
Dim hlRight             As POINTAPI ' Right leg
Dim hlMiddle            As POINTAPI ' The middle one

Dim pMouse              As POINTAPI ' Location of the mouse

' Some other stuffs now
Dim lPen                As Long ' Pen that will be used for a moon surface, it is be HPEN in c++

Dim BitmapLoaded        As BITMAP ' Bitmap that we have just loaded, we will store the data here
Dim BitmapInfoHeader    As HBITMAPINFOHEADER ' Bitmap INFOHEADRER, a bitmap file informations header

Dim IsGameRunning       As Boolean
Dim IsGamePaused        As Boolean
' When the lander safely lands, both IsCrashed and IsLanded will be TRUE!
Dim IsCrashed           As Boolean ' If the lander has landed and/or crashed, this is TRUE
Dim IsLanded            As Boolean ' But if it is landed, this option will be TRUE too

Dim lTmrCounter         As Long ' Just a counter for performance

Dim lmX                 As Single ' lunar module position X
Dim lmY                 As Single ' lm position Y

Dim lmAccX              As Single ' Acceleration increment
Dim lmAccY              As Single

Dim gTime               As Long ' Count the socore, the faster you land, better

'**********************************************************************
'*  This function will create a new empty DC
'*
'*  W      - Width of the new DC
'*  H      - Height of the new DC
'*  BPP    - Bit count, 1, 4, 8, 24...
'*  hDC    - A buffer that will recieve the created DC
'**********************************************************************
Public Sub CreateBlankDC(W&, H&, BPP&, hdc&, hBMP&)
    Dim hDIB& ' DIB location
    
    ' Write some basic bitmap informations by using INFOHEADER
    With BitmapInfoHeader
        .biSize = Len(BitmapInfoHeader) ' size of the structure in bytes
        .biBitCount = BPP
        .biHeight = H
        .biWidth = W
        .biPlanes = 1 ' Must be one :)
        .biSizeImage = GetImageSize(W, H) ' Image size, use the function I've found somewhere
    End With
   
    hdc = CreateCompatibleDC(0) ' Create empty DC (device context)
    hDIB = CreateDIBSection(hdc, BitmapInfoHeader, DIB_RGB_COLORS, 0, 0, 0) ' Create DIB section
   
    If hDIB Then ' If it was created
        hBMP = SelectObject(hdc, hDIB) ' Select our DIB to DC, BMP that is
        BitBlt hdc, 0, 0, W, H, hdc, 0, 0, vbBlackness ' Just fill the image with blackness
    Else
        MsgBox "Error in creation of DIB"
        Exit Sub
    End If

End Sub

'**********************************************************************
'*  It will copy desired bitmap to a specifed DC, simply draws the bitmap to it
'*
'*  BitmapFileName      - File name of the bitmap
'*  BPP                 - Bit count
'*  hDC                 - Where to paint the BMP
'*  hBMP                - ID for the bitmap, to have it in memory :)
'**********************************************************************
Public Sub LoadBitmapIntoDC(BitmapFileName$, BPP&, hdc&, hBMP&)
    Dim VBImage As StdPicture ' VBA's object for some image manipulations, VERY nice stuff!
    Dim hDCT&, hBMPT& ' Temporary data to copy the bitmap
    
    
    Set VBImage = LoadPicture(BitmapFileName) ' Load the image, you can even open JPG, GIF, WMF...
    GetObjectA VBImage.handle, Len(BitmapLoaded), BitmapLoaded ' Get the handle of BITMAP structure

    ' This function is made to CREATE DC and COPY the bitmap to it, to simplify the process
    CreateBlankDC BitmapLoaded.bmWidth, BitmapLoaded.bmHeight, BPP, hdc, hBMP ' Now we will create simple blank DC, so you wouldn bother with it
    
    hDCT = CreateCompatibleDC(hdc) ' Create a temporary dc, to store the bitmap
    hBMPT = SelectObject(hDCT, VBImage.handle) ' Temp bmp
    
    ' Now we will simply copy the bitmap from temp to specified DC bit by bit, safer way to "put" the bitmap to a DC
    BitBlt hdc, 0, 0, BitmapInfoHeader.biWidth, BitmapInfoHeader.biHeight, hDCT, 0, 0, vbSrcCopy
    
    ' Select the temp, delete it, not needed anymore
    SelectObject hDCT, hBMPT
    
    DeleteDC hDCT
    DeleteObject hBMPT
    
End Sub

'**********************************************************************
'*  This func will create a black mask of the source HDC
'*
'*  SrcDC       - Source DC, to create mask for
'*  DstDC       - Destination, where to put the mask
'*  DstBMP      - Destination bitmap for the mask (just long integer data)
'*  W           - Width of the dst bitmap
'*  H           - height
'**********************************************************************
Public Sub CreateMaskDC(SrcDC&, DstDC&, DstBMP&, W&, H&)
    Dim x&, y&
    
    CreateBlankDC W, H, 24, DstDC&, DstBMP ' Create new DC
    
    ' Now cycle pixels
    For x = 0 To W
        For y = 0 To H
            If GetPixel(SrcDC&, x, y) <> 0 Then ' If we have some 'non-black' color
                SetPixel DstDC&, x, y, RGB(255, 255, 255) ' Put the WHITE color
            End If
        Next y
    Next x
    
    ' We have used white for mask, black for background, invert it to be correct
    BitBlt DstDC&, 0, 0, W, H, DstDC&, 0, 0, vbDstInvert
End Sub

'**********************************************************************
'*  Simply draws a moon surface to a hLevel's hdc
'*
'**********************************************************************
Public Sub DrawLevel()
    Dim i&
    ' pt is array of points for line(s)
    Dim pt(10) As POINTAPI, p As POINTAPI
    
    ' First fill memory with black color
    BitBlt hLevel.hdc, 0, 0, Me.ScaleWidth, 60, 0, 0, hLevel.hdc, vbBlackness
    
    ' Create a PEN  which we will use for drawing a surface (see included module for API explanations)
    lPen = CreatePen(0, 1, RGB(255, 255, 255))
    SelectObject hLevel.hdc, lPen ' Select created pen to surf DC
    
    ' Now put some random points to a memory
    For i = 0 To 10
        Randomize i * GetTickCount
        pt(i).x = i * CInt((Me.ScaleWidth / 10))
        pt(i).y = CInt(Rnd * 50) + 1
    Next i
    pt(10).x = Me.ScaleWidth + 1 ' Last one must be >= then the form's width
    
    ' There must be at least two points with the same Y, so lander could land
    Randomize
    i = CInt(Rnd * 10) + 1: If i > 10 Then i = 10 ' Check if i is bigger then 10 (dunno why, it just happends)
    pt(i).y = pt(i - 1).y ' See...
    
    ' Now just draw the level
    For i = 1 To 10
        ' We use 'p' just because stupid API Text Viewer have only Ex version of this api
        MoveToEx hLevel.hdc, pt(i - 1).x, pt(i - 1).y, p
        LineTo hLevel.hdc, pt(i).x, pt(i).y
    Next i
    
    ' Fill it with the white color, there must not be a "leak"!
    FillColor = RGB(255, 255, 255)
    ExtFloodFill hLevel.hdc, 0, 59, 0, 1 ' You see... only 'Ex' or 'Ext' versions are avivable :P...
End Sub

'**********************************************************************
'*  This func will calculate image size, not written by me...
'*  W   - Width
'*  H   - Height
'**********************************************************************
Public Function GetImageSize(W&, H&) As Long ' from c++ macro
    GetImageSize = ((W * 3 + 3) And &HFFFFFFFC) * H
End Function

'**********************************************************************
'*  Form's 'load' message handler
'**********************************************************************
Private Sub Form_Load()
    ' The game has started
    IsGameRunning = True
    
    ' Show the form just before the loop, 'DoEvents' message processor will not do this
    Me.Show
    
    ' Set lander's position to center
    lmX = Me.ScaleWidth / 2 - 10
    lmY = 50
    
    ' Create some sprites
    LoadBitmapIntoDC App.Path & "\lander.bmp", 24, hLunar.hdc, hLunar.hBMP
    CreateMaskDC hLunar.hdc, hLunar.mhDC, hLunar.mBMP, 30, 25
    
    LoadBitmapIntoDC App.Path & "\lside.bmp", 24, hlSide.hdc, hlSide.hBMP
    
    LoadBitmapIntoDC App.Path & "\ldown.bmp", 24, hlDown.hdc, hlDown.hBMP
    
    LoadBitmapIntoDC App.Path & "\lcrashed.bmp", 24, hlCrashed.hdc, hlCrashed.hBMP
    
    CreateBlankDC Me.ScaleWidth, 60, 24, hLevel.hdc, hLevel.hBMP
    
    DrawLevel
    
    lmdMaxFuel = 500 ' Starting fuel level
    lmdFuel = lmdMaxFuel ' Current fuel :)
    
    GameLoop ' Start the loop
End Sub

'**********************************************************************
'*  If player has moved the mouse on this form, activate the game
'*  (if it is not paused before by him)
'**********************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not IsGameRunning Then
        If Not IsGamePaused Then IsGameRunning = True
        GameLoop
    End If
End Sub

'**********************************************************************
'*  Delete holders from memory
'**********************************************************************
Private Sub Form_Unload(Cancel As Integer)
    DeleteHolder hLunar
    DeleteHolder hLevel
    DeleteHolder hlDown
    DeleteHolder hlSide
    DeleteHolder hlCrashed
End Sub

'**********************************************************************
'*  Now the game starts
'**********************************************************************
Public Sub GameLoop()
    Dim kLeft As Boolean ' If LEFT was pressed
    Dim kRight As Boolean ' RIGHT
    Dim kDown As Boolean ' DOWN
    Dim kLegs(1 To 3) As Boolean ' To see if all the legs have touched the surface at the same Y level
        
    Dim i% ' :), '%' is integer
    
    Dim sBuf$ ' Buffer to store some strings

    If Not IsGamePaused Then SetWindowCaption "" ' Set the window title to default
    
    SetBkMode hdc, 0 ' Set background mode, TRANSPARENT in this case (hdc is the 'default' hdc, window's in this case)
    
    ' Set window font now
    Font.Bold = True
    Font.Size = 10

    ' Here we go
    While IsGameRunning
        DoEvents ' Process some messages
            
        GetCursorPos pMouse ' Get the mouse position
        
        ' See if user has left the window with his mouse, pause the game then
        With pMouse
            ' VB Form position is in Twips, so we must divide it with 15
            If .x < Left / 15 Then GoTo pause
            If .x > Left / 15 + ScaleWidth Then GoTo pause
        
            If .y < Top / 15 + 48 Then GoTo pause
            If .y > Top / 15 + ScaleHeight Then GoTo pause
        End With
        
        ' If TMR_INTERVAL time has passed, then process next frame...
        If lTmrCounter + TMR_INTERVAL <= GetTickCount Then
            
            lTmrCounter = GetTickCount ' Reset counter
            
            gTime = gTime + 1 ' Count the players time, for scoring
            
            ' Reset key indicators
            kDown = False
            kRight = False
            kLeft = False
                
            ' As I said, IsCrashed indicates if player has crashed OR/AND landed ('soft' crashing... :)
            If Not IsCrashed = True Then
            
                lmAccY = lmAccY + 0.05 ' Send the lander down, gravity
            
                If lmdFuel > 0 Then ' If we have some fuel
                    ' Use GetKeyState api to check if the key is pressed, slow but better
                    ' than VB's key events
                    If GetKeyState(vbKeyDown) < 0 Then
                        kDown = True
                        lmAccY = lmAccY - 0.1
                    End If
                    
                    If GetKeyState(vbKeyRight) < 0 Then
                        lmAccX = lmAccX - 0.05
                        kLeft = True
                    End If
                    
                    If GetKeyState(vbKeyLeft) < 0 Then
                        lmAccX = lmAccX + 0.05
                        kRight = True
                    End If
                Else
                    lmdFuel = 0 ' If it is <0 for some reason, like -1
                End If
            
                lmX = lmX + lmAccX ' Now move the ship
                lmY = lmY + lmAccY
            
            End If
            
            ' Set position of the module's legs
            hlLeft.x = lmX
            hlLeft.y = lmY + 24
            hlRight.x = lmX + 29
            hlRight.y = lmY + 24
            hlMiddle.x = lmX + 15
            hlMiddle.y = lmY + 24
            
            ' Reset legs, all of them must touch the land at the same frame
            For i = 1 To 3: kLegs(i) = False: Next i
            
            ' Check for lander's collision, first check if he is in the needed range
            If lmY + 24 > Me.ScaleHeight - 60 Then
                ' One by one, check all the legs by using GetPixel
                If GetPixel(hLevel.hdc, hlLeft.x, hlLeft.y - (Me.ScaleHeight - 59)) = vbWhite Then
                    IsCrashed = True
                    kLegs(1) = True
                End If
                If GetPixel(hLevel.hdc, hlRight.x, hlRight.y - (Me.ScaleHeight - 59)) = vbWhite Then
                    IsCrashed = True
                    kLegs(2) = True
                End If
                If GetPixel(hLevel.hdc, hlMiddle.x, hlMiddle.y - (Me.ScaleHeight - 59)) = vbWhite Then
                    IsCrashed = True
                    kLegs(3) = True
                End If
                ' Lander has landed, all the legs have collided
                If kLegs(1) And kLegs(2) And kLegs(3) Then
                    If lmAccY < 1 Then
                        IsLanded = True
                    End If
                End If
            End If
            
            ' Drawing part, just to draw the game
            BackColor = 0 ' Set the back color to BLACK
            Cls ' Clear the window
            
            ' Draw the level
            BitBlt hdc, 0, Me.ScaleHeight - 60, Me.ScaleWidth, 60, hLevel.hdc, 0, 0, vbSrcCopy
            
            ' Now draw lander (if it has landed it looks a bit different)...
            If IsCrashed And Not IsLanded Then
                BitBlt hdc, lmX, lmY + 10, 30, 25, hlCrashed.hdc, 0, 0, vbSrcCopy
            ElseIf IsLanded Then
                BitBlt hdc, lmX, lmY - 1, 30, 25, hLunar.mhDC, 0, 0, vbSrcAnd
                BitBlt hdc, lmX, lmY - 1, 30, 25, hLunar.hdc, 0, 0, vbSrcPaint
            Else
                BitBlt hdc, lmX, lmY, 30, 25, hLunar.mhDC, 0, 0, vbSrcAnd
                BitBlt hdc, lmX, lmY, 30, 25, hLunar.hdc, 0, 0, vbSrcPaint
            End If

            ' Draw acceleration flame
            If Not IsCrashed Then
                If kRight Then
                    BitBlt hdc, lmX - 12, lmY + 4, 18, 9, hlSide.hdc, 18, 0, vbSrcInvert
                    
                    lmdFuel = lmdFuel - 1
                End If
                If kLeft Then
                    BitBlt hdc, lmX + 24, lmY + 4, 18, 9, hlSide.hdc, 0, 0, vbSrcInvert
                
                    lmdFuel = lmdFuel - 1
                End If
                If kDown Then
                    BitBlt hdc, lmX + 8, lmY + 17, 14, 31, hlDown.hdc, 0, 0, vbSrcInvert
                
                    lmdFuel = lmdFuel - 1.5
                End If
            End If
            
            
            ForeColor = RGB(255, 255, 255) ' Set fore color for text
            
            ' Print some text
            sBuf = "Speed: " & Round(lmAccY, 1)
            TextOut hdc, 16, 16, sBuf, Len(sBuf)
            
            sBuf = "Fuel: " & Round(lmdFuel, 0)
            TextOut hdc, 16, 32, sBuf, Len(sBuf)
        
            sBuf = "Time: " & gTime
            TextOut hdc, 16, 48, sBuf, Len(sBuf)
        
            ' Handle 'landed', and 'crashed' modes
            If IsCrashed Then
                If IsLanded Then
                    With frmResult
                        .picRes(0).Visible = True ' Change the picture
                        .Caption = "You have landed..."
                        .lblRes = "You have landed, now how to go back... one-way ticket I guess..."
                        .Show vbModal, Me
                    End With
                Else
                    With frmResult
                        .picRes(1).Visible = True
                        .Caption = "You have crashed..."
                        .lblRes = "Very nice hole... well, you know what? I bet you can make a bigger one!"
                        .Show vbModal, Me
                    End With
                End If
                IsGameRunning = False
                IsGamePaused = True
            End If
                
        End If
        
        ' Finito
    Wend
    
Exit Sub
pause:
    SetWindowCaption "- PAUSED"
    IsGameRunning = False
crashed:
    IsGameRunning = False
End Sub

'**********************************************************************
'*  Set the title of the main window
'**********************************************************************
Public Sub SetWindowCaption(Text As String)
    Caption = "Lunar Lander " & Text
End Sub

Private Sub mnuGameAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'**********************************************************************
'*  Exit game
'**********************************************************************
Private Sub mnuGameExit_Click()
    Unload Me
End Sub

'**********************************************************************
'*  Restart game
'**********************************************************************
Private Sub mnuGameNew_Click()
    
    lmX = Me.ScaleWidth / 2 - 10
    lmY = 50
    
    lmdMaxFuel = 500
    lmdFuel = lmdMaxFuel
    
    IsLanded = False
    IsCrashed = False
    IsGameRunning = True
    IsGamePaused = False
    
    gTime = 0
    
    DrawLevel
    
    GameLoop
    
End Sub

'**********************************************************************
'*  Pause game
'**********************************************************************
Private Sub mnuGamePause_Click()
    IsGameRunning = False
    IsGamePaused = True
End Sub

'**********************************************************************
'*  Unpause game
'**********************************************************************
Private Sub mnuGameStart_Click()
    IsGameRunning = True
    IsGamePaused = False
    GameLoop
End Sub

