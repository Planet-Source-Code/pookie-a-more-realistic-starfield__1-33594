VERSION 5.00
Begin VB.Form frmStars 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Starfield"
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrKeyBoard 
      Interval        =   50
      Left            =   2460
      Top             =   780
   End
   Begin VB.Label lblStars 
      BackStyle       =   0  'Transparent
      Caption         =   "Stars:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "frmStars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------------
' Pookie's Starfield Generator V 1.2
' If you wish to use this code in your game, then give me some credit. ;)
' Email: arbml999@hotmail.com
' ICQ: 147085490
' --------------------------------------------------------------------------------------

' Set up the stars
Private Type UDTStars
  Radius  As Single
  Angle   As Single
  Speed   As Single
  Red     As Single
  Grn     As Single
  Blu     As Single
  RadiusStart As Single
  XStar As Long
  YStar As Long
End Type

Dim Stars(9999)   As UDTStars ' 10000 max stars
Dim Finished      As Boolean  ' Exit program if true
Dim XMiddle       As Long     ' X middle of the screen
Dim YMiddle       As Long     ' Y middle of the screen
Dim MaxRadius     As Long     ' Max Radius stars can go before off the screen
Dim MaxStars      As Long     ' Counter for how many stars on screen now
Dim Ticker        As Long     ' Counter for how much time has gone in a loop
Dim RayLength     As Single   ' Length of the rays (Brightness of stars)
Dim StarRoll      As Single   ' Amount of spinning of the stars
Dim StarSpeed     As Single   ' Speed of flying though space

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long

' Memorize Pi*2 (360d of Radians)
Const PI2 = 6.283185307

Private Sub Form_Load()

Dim nloop As Long
  
  ' display keys
  MsgBox "To use:" & vbNewLine & "Cursor Up/Down: Flight speed" & vbNewLine & _
         "Cursor Left/Right: Flight rolling" & vbNewLine & _
         "Page Up/Down: Amount of stars in space" & vbNewLine & _
         "Escape: Exit program", vbInformation, "--- Pookie's StarField ---"
  
  Me.Show
  DoEvents
  
  Randomize
  
  ' Place the stars label at the bottom of the screen
  lblStars.Top = Me.ScaleHeight - 16
  
  ' Find the middle of the screen
  XMiddle = Me.ScaleWidth / 2
  YMiddle = Me.ScaleHeight / 2
  ' Calc the MaxRadius before dots can't be seen on screen
  MaxRadius = XMiddle * 1.33333
  ' Calc the Raylength so that the brighness should be similar regardless of resolution
  RayLength = 768# / MaxRadius
  ' Amount of rolling in space
  StarRoll = 0.002
  ' Amount of speed through space
  StarSpeed = 1.00001
  ' Amount of stars in space
  MaxStars = 1000
  
  ' Place all the stars
  For nloop = 0 To MaxStars - 1
    NewStar nloop, True
  Next

  ' Jump into the mainloop until done
  MainLoop
  
End Sub

Private Sub MainLoop()

Dim nloop         As Long ' Loop counter
Dim Brightness    As Long ' Brightness of current star
Dim Red           As Long ' Red value of star
Dim Grn           As Long ' Green value of star
Dim Blu           As Long ' Blue value of star

' If true then exit program
Finished = False

Do
  Ticker = GetTickCount
  For nloop = 0 To MaxStars - 1
    With Stars(nloop)
      ' Clear the star from the screen
      SetPixelV Me.hdc, .XStar, .YStar, 0
      
      ' Move the star
      .Radius = .Radius * .Speed
      .Speed = .Speed * StarSpeed
      .Angle = .Angle + StarRoll
      ' If star is offscreen then create a new one
      If .Radius > MaxRadius Then
        NewStar nloop, False
      End If
      
      ' Calc the brightness of the star
      Brightness = (.Radius - .RadiusStart) * RayLength
      Red = .Red + Brightness
      Grn = .Grn + Brightness
      Blu = .Blu + Brightness
      If Red < 0 Then Red = 0
      If Grn < 0 Then Grn = 0
      If Blu < 0 Then Blu = 0
      If Red > 255 Then Red = 255
      If Grn > 255 Then Grn = 255
      If Blu > 255 Then Blu = 255
      ' Memorize the X/Y of the star
      .XStar = XMiddle + .Radius * Cos(.Angle)
      .YStar = YMiddle + .Radius * Sin(.Angle)
      ' Draw the star
      SetPixelV Me.hdc, .XStar, .YStar, RGB(Red, Grn, Blu)
    End With
  Next
  ' Refresh the screen
  Me.Refresh
  
  ' Slow the routine down if going too quick
  Do
    DoEvents
  Loop Until GetTickCount > Ticker + 5
  
Loop Until Finished

  ' Unload the form and exit
  Unload Me
  
End Sub

Private Sub NewStar(nStar As Long, fFirst As Boolean)

  With Stars(nStar)
    .Red = Rnd * 32
    .Grn = Rnd * 48
    .Blu = Rnd * 64
    If Rnd > 0.5 Then .Red = -1000
    If Rnd > 0.75 Then .Grn = -1000
    .Angle = Rnd * PI2                    ' random 360 degrees angle
    .RadiusStart = Rnd * MaxRadius / 1.5  ' Some stars can be far from the centre
    .Radius = .RadiusStart                ' Make sure radius and radius start are the same
    .Speed = 1#                           ' Set the initial speed of the star to 1
    ' make some stars move fast when program stars for variation
    If fFirst Then
      .Speed = .Speed + Rnd / 100 - Rnd / 100
    End If
  End With

End Sub

Private Sub tmrkeyboard_Timer()

  Dim nloop As Long
  
  ' Speed up
  If GetAsyncKeyState(vbKeyUp) And StarSpeed < 1.0001 Then
    StarSpeed = StarSpeed + 0.00001
  End If
  ' Speed down
  If GetAsyncKeyState(vbKeyDown) And StarSpeed > 1.00001 Then
    StarSpeed = StarSpeed - 0.00001
  End If
  
  ' Increase Stars
  If GetAsyncKeyState(vbKeyPageDown) And MaxStars > 100 Then
    For nloop = MaxStars - 100 To MaxStars - 1
      With Stars(nloop)
        SetPixelV Me.hdc, .XStar, .YStar, 0
      End With
    Next
    MaxStars = MaxStars - 100
  End If
  ' Decrease stars
  If GetAsyncKeyState(vbKeyPageUp) And MaxStars < 9976 Then
    MaxStars = MaxStars + 100
    For nloop = MaxStars - 100 To MaxStars - 1
      NewStar nloop, False
    Next
  End If
  
  ' Roll left
  If GetAsyncKeyState(vbKeyLeft) And StarRoll < 0.01 Then
    StarRoll = StarRoll + 0.001
  End If
  ' Roll Right
  If GetAsyncKeyState(vbKeyRight) And StarRoll > -0.01 Then
    StarRoll = StarRoll - 0.001
  End If
   
  ' Exit program
  If GetAsyncKeyState(vbKeyEscape) Then
    Finished = True
  End If
  
  ' Show how many stars on screen
  lblStars.Caption = "Stars: " & MaxStars

End Sub
