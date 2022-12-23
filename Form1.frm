VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Acrylic Demo"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10485
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   435
      Left            =   9060
      TabIndex        =   15
      Top             =   4770
      Width           =   1005
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Border"
      Height          =   435
      Left            =   7890
      TabIndex        =   14
      Top             =   5940
      Width           =   1065
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Frameless"
      Height          =   435
      Left            =   7890
      TabIndex        =   13
      Top             =   5340
      Width           =   1065
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Light BG"
      Height          =   435
      Left            =   9090
      TabIndex        =   12
      Top             =   5340
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   4500
      TabIndex        =   11
      Text            =   "3"
      Top             =   5910
      Width           =   585
   End
   Begin VB.TextBox TextAnimationID 
      Height          =   435
      Left            =   4500
      TabIndex        =   10
      Text            =   "0"
      Top             =   5280
      Width           =   585
   End
   Begin VB.TextBox TextA 
      Height          =   435
      Left            =   3660
      TabIndex        =   9
      Text            =   "48"
      Top             =   5280
      Width           =   585
   End
   Begin VB.TextBox TextB 
      Height          =   435
      Left            =   2910
      TabIndex        =   8
      Text            =   "242"
      Top             =   5280
      Width           =   585
   End
   Begin VB.TextBox TextG 
      Height          =   435
      Left            =   2160
      TabIndex        =   7
      Text            =   "242"
      Top             =   5280
      Width           =   585
   End
   Begin VB.TextBox TextR 
      Height          =   435
      Left            =   1440
      TabIndex        =   6
      Text            =   "242"
      Top             =   5280
      Width           =   585
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Custom State"
      Height          =   435
      Left            =   2910
      TabIndex        =   5
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dark BG"
      Height          =   435
      Left            =   9090
      TabIndex        =   4
      Top             =   5940
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aero"
      Height          =   435
      Left            =   330
      TabIndex        =   2
      Top             =   5910
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Disable Effect"
      Height          =   435
      Left            =   1410
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acrylic"
      Height          =   435
      Left            =   330
      TabIndex        =   0
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   4455
      Left            =   300
      TabIndex        =   3
      Top             =   210
      Width           =   9765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowCompositionAttribute Lib "user32.dll" (ByVal hWnd As Long, ByRef data As WindowsCompostionAttributeData) As Long

Private Type WindowsCompostionAttributeData
    Attribute As Long
    data As Long
    SizeOfData As Long
End Type

Public Enum AccentState
    ACCENT_DISABLED = 0
    ACCENT_ENABLE_GRADIENT = 1
    ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
    ACCENT_ENABLE_BLURBEHIND = 3
    ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    ACCENT_ENABLE_MICABLURBEHIND = 5
    ACCENT_INVALIDSTATE = 6
End Enum

Const WCA_ACCENT_POLICY = 19

Private Type AccentPolicy
    AccentState As Long
    AccentFlags As Long
    GradientColor As Long
    AnimationID As Long
End Type

Sub SetAccent(State As AccentState)
    Dim data As WindowsCompostionAttributeData
    
    Dim accent As AccentPolicy
    
    accent.AccentState = State
    
    data.Attribute = WCA_ACCENT_POLICY
    data.SizeOfData = Len(accent)
    data.data = VarPtr(accent)
    
    Dim lret As Long
    lret = SetWindowCompositionAttribute(hWnd, data)
End Sub
Sub SetArcylic(hWnd As Long, Optional GradientColor As Long = 821228274, Optional EnableShadow As Boolean = True, Optional AnimationID As Integer = 0)
    Dim data As WindowsCompostionAttributeData
    
    Dim accent As AccentPolicy
    
    accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
    accent.GradientColor = GradientColor
    
    If EnableShadow Then
        accent.AccentFlags = &H20 Or &H40 Or &H80 Or &H100
    Else
        accent.AccentFlags = &H0
    End If
    
    accent.AnimationID = AnimationID
    
    data.Attribute = WCA_ACCENT_POLICY
    data.SizeOfData = Len(accent)
    data.data = VarPtr(accent)
    
    Dim lret As Long
    lret = SetWindowCompositionAttribute(hWnd, data)
End Sub

Public Function RGBA(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional ByVal Alpha As Byte) As Long
    If Alpha > 127 Then
        RGBA = RGB(Red, Green, Blue) Or (Alpha - 128) * &H1000000 Or &H80000000
    Else
        RGBA = RGB(Red, Green, Blue) Or Alpha * &H1000000
    End If
End Function

Private Sub Command1_Click()
    SetArcylic Me.hWnd, RGBA(CInt(TextR.Text), CInt(TextG.Text), CInt(TextB.Text), CInt(TextA.Text)), True, 0
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
    SetAccent ACCENT_ENABLE_BLURBEHIND
End Sub


Private Sub Command4_Click()
Me.BackColor = vbBlack
Label1.ForeColor = vbWhite
End Sub
Private Sub Command5_Click()
    SetAccent CInt(Text1.Text)
End Sub

Private Sub Command6_Click()
    SetAccent ACCENT_DISABLED
End Sub

Private Sub Command7_Click()
Me.BackColor = vbWhite
Label1.ForeColor = vbBlack
End Sub

Private Sub Command8_Click()
Me.BorderStyle = 0
Me.Caption = Me.Caption
End Sub

Private Sub Command9_Click()
Me.BorderStyle = 2
Me.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    Label1.Caption = "This is a simple demo of acrylic material in Visual Basic 6." & vbCrLf & "Acrylic is a type of Brush that creates a translucent texture. You can apply acrylic to app surfaces to add depth and help establish a visual hierarchy." & _
    vbCrLf & "The Fluent Design System helps you create modern, bold UI that incorporates light, depth, motion, material, and scale. Acrylic is a Fluent Design System component that adds physical texture (material) and depth to your app. To learn more, see the Fluent Design overview." & _
    vbCrLf & "If you are using in-app acrylic on navigation surfaces, consider extending content beneath the acrylic pane to improve the flow in your app. Using NavigationView will do this for you automatically. However, to avoid creating a striping effect, try not to place multiple pieces of acrylic edge-to-edge - this can create an unwanted seam between the two blurred surfaces. Acrylic is a tool to bring visual harmony to your designs, but when used incorrectly can result in visual noise."

End Sub

