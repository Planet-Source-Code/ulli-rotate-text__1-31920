VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF80FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   403
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resize me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   2310
      TabIndex        =   0
      Top             =   2475
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFacename          As String * 33
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private The_Text        As String
Private F               As LOGFONT
Private hPrevFont       As Long
Private hFont           As Long
Private Sleeptime       As Long
Private cx              As Single
Private cy              As Single
Private Theta           As Single
Private dTheta          As Single
Private xRadius         As Single
Private yRadius         As Single
Private i               As Integer
Private pi              As Single
Private Char            As String
Private Const SizeFact  As Single = 3  'play with this and sizefont to create the full ark
Private Const SizeFont  As Long = 10

Private Sub Form_Initialize()

    Sleeptime = 66
    FontSize = SizeFont
    
End Sub

Private Sub Form_Resize()

    The_Text = "PLANET SOURCE CODE ROTATING TEXT EXAMPLE...   "
    Cls
    Label1.Move (ScaleWidth - Label1.Width) / 2, (ScaleHeight - Label1.Height) / 2
    pi = 4 * Atn(1)
    cx = ScaleWidth / 2
    cy = ScaleHeight / 2
    Theta = pi
    dTheta = pi / TextWidth(The_Text)
    xRadius = ScaleWidth * 0.4
    yRadius = ScaleHeight * 0.4
    FontBold = True
    For i = 1 To Len(The_Text)
        F.lfEscapement = Theta * 3600 / pi / 2 - 950   'rotation angle, in tenths°. the additional 5 degrees tilt is to adjust the slight x/y offset
        F.lfFacename = FontName + Chr$(0)  'null terminated
        F.lfHeight = FontSize * SizeFact
        hFont = CreateFontIndirect(F)
        hPrevFont = SelectObject(hdc, hFont)
        Char = Mid$(The_Text, i, 1)
        CurrentX = cx + xRadius * Cos(Theta) - Cos(Theta) * TextHeight(Char) / SizeFact
        CurrentY = cy - yRadius * Sin(Theta) - Sin(Theta) * TextWidth(Char) / SizeFact
        Print Char
        Sleep Sleeptime
        Theta = Theta - dTheta * TextWidth(Char)
        hFont = SelectObject(hdc, hPrevFont)
        DeleteObject hFont
    Next i
    Sleeptime = 0
    Label1.Visible = True
    
End Sub

':) Ulli's VB Code Formatter V2.10.7 (19.02.2002 08:59:08) 35 + 43 = 78 Lines
