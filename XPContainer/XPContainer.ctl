VERSION 5.00
Begin VB.UserControl XPContainer 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "XPContainer.ctx":0000
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ToolboxBitmap   =   "XPContainer.ctx":0014
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   15
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   15
      Width           =   2415
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XPContainer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   75
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   15
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "XPContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'---------------------------------------------------------------------------------------
' XP Container Control
'---------------------------------------------------------------------------------------
'
' Control        : XPContainer
' Date           : 08/04/2005
' Author         : Cameron Groves
' Copyright      : Copyright © 2005 Cameron Groves. All rights reserved.
' Purpose        : A simple lightweight Container Control.
'
'---------------------------------------------------------------------------------------
'
' TERMS & CONDITIONS
'
' Redistribution and use in source and binary forms, with or
' without modification, are permitted provided that the following
' conditions are met:
'
' 1. Redistributions of source code must retain the above copyright
'    notice, this list of conditions and the following disclaimer.
'
' 2. The end-user documentation included with the redistribution, if any,
'    must include the following acknowledgment:
'
'    "This product includes software developed by Cameron Groves"
'
'    Alternately, this acknowledgment may appear in the software itself, if
'    and wherever such third-party acknowledgments normally appear.
'
' THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
' INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
' AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
' CAMERON GROVES OR ANY CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
' INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
' BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
' USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
' THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'---------------------------------------------------------------------------------------

Option Explicit

' API Declarations

Private Declare Function OleTranslateColor _
                Lib "olepro32.dll" (ByVal OLE_COLOR As Long, _
                                    ByVal HPALETTE As Long, _
                                    pccolorref As Long) As Long
Private Const CLR_INVALID = -1

' Used to prevent crashes on Windows XP
                
Private Declare Function LoadLibrary _
                Lib "KERNEL32" _
                Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary _
                Lib "KERNEL32" (ByVal hLibModule As Long) As Long

Private m_hMod As Long

' Enumerations

Public Enum XPContainerStyles
    [Header Visible] = 0
    [Header Invisible] = 1
End Enum

Public Enum XPContainerThemes
    [XP Blue] = 0
    [XP Dark Blue] = 1
    [XP Dark Green] = 2
    [XP Green] = 3
    [XP Light Blue] = 4
    [XP Light Green] = 5
    [XP Orange] = 6
    [XP Pastel Green] = 7
    [XP Purple] = 8
    [XP Red] = 9
    [XP Silver] = 10
    [XP Yellow] = 11
End Enum

' Property Variables

Private m_Theme As XPContainerThemes
Private m_Style As XPContainerStyles
Private m_HeaderLightColor As OLE_COLOR
Private m_HeaderDarkColor As OLE_COLOR
Private m_BackLightColor As OLE_COLOR
Private m_BackDarkColor As OLE_COLOR
Private m_BorderColor As OLE_COLOR
Private m_TextColor As OLE_COLOR
Private m_Caption As String

' Default Property Values

Private Const m_def_Theme = 0
Private Const m_def_Style = 0
Private Const m_def_HeaderLightColor = &HF7E0D3
Private Const m_def_HeaderDarkColor = &HEDC5A7
Private Const m_def_BackLightColor = &HFCF4EF
Private Const m_def_BackDarkColor = &HFAE8DC
Private Const m_def_BorderColor = &HDCC1AD
Private Const m_def_TextColor = &H7B2D02
Private Const m_def_Caption = "XPContainer"

'---------------------------------------------------------------------------------------
' Procedure : ApplyTheme
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Sets the color properties for each theme.
'---------------------------------------------------------------------------------------

Private Sub ApplyTheme()

    Select Case m_Theme

        Case [XP Blue]
            HeaderLightColor = &HF7E0D3
            HeaderDarkColor = &HEDC5A7
            BackLightColor = &HFCF4EF
            BackDarkColor = &HFAE8DC
            BorderColor = &HDCC1AD
            TextColor = &H7B2D02

        Case [XP Dark Blue]
            HeaderLightColor = &HECDCD3
            HeaderDarkColor = &HDABAA8
            BackLightColor = &HF8F2EF
            BackDarkColor = &HF1E5DD
            BorderColor = &HD6B4A0
            TextColor = &H4B2A17

        Case [XP Dark Green]
            HeaderLightColor = &HD8E5C8
            HeaderDarkColor = &HB1CB92
            BackLightColor = &HF1F5EB
            BackDarkColor = &HE1EBD5
            BorderColor = &HAAC688
            TextColor = &H213B00

        Case [XP Green]
            HeaderLightColor = &HE0EAE8
            HeaderDarkColor = &HC2D6D1
            BackLightColor = &HF4F8F7
            BackDarkColor = &HE7EFED
            BorderColor = &HBCD3CD
            TextColor = &H324741

        Case [XP Light Blue]
            HeaderLightColor = &HF1E3C8
            HeaderDarkColor = &HE4C992
            BackLightColor = &HFAF5EB
            BackDarkColor = &HF5EAD5
            BorderColor = &HE2C488
            TextColor = &H553900

        Case [XP Light Green]
            HeaderLightColor = &HDAF2E3
            HeaderDarkColor = &HB5E5C8
            BackLightColor = &HF1FAF5
            BackDarkColor = &HE3F5EA
            BorderColor = &HAEE3C3
            TextColor = &H245738

        Case [XP Orange]
            HeaderLightColor = &HD2E2FD
            HeaderDarkColor = &HA7C6FA
            BackLightColor = &HEFF5FE
            BackDarkColor = &HDDE9FD
            BorderColor = &H9FC0FA
            TextColor = &H16366D

        Case [XP Pastel Green]
            HeaderLightColor = &HE3E3D6
            HeaderDarkColor = &HC9C9AE
            BackLightColor = &HF5F5F0
            BackDarkColor = &HEAEAE0
            BorderColor = &HC4C4A6
            TextColor = &H39391D

        Case [XP Purple]
            HeaderLightColor = &HEAD7DF
            HeaderDarkColor = &HD5B0BF
            BackLightColor = &HF7F1F3
            BackDarkColor = &HEFE1E6
            BorderColor = &HD1A9B9
            TextColor = &H46202F

        Case [XP Red]
            HeaderLightColor = &HD6D2FB
            HeaderDarkColor = &HAEA6F8
            BackLightColor = &HF0EFFE
            BackDarkColor = &HE0DDFC
            BorderColor = &HA79EF7
            TextColor = &H1D156A

        Case [XP Silver]
            HeaderLightColor = &HECEAE9
            HeaderDarkColor = &HD9D6D3
            BackLightColor = &HF8F7F7
            BackDarkColor = &HF1EFEE
            BorderColor = &HD6D2CF
            TextColor = &H4A4744

        Case [XP Yellow]
            HeaderLightColor = &HE4FAFC
            HeaderDarkColor = &HB9EEF4
            BackLightColor = &HEEFCFD
            BackDarkColor = &HDCF7FA
            BorderColor = &H95E1EA
            TextColor = &H66D5E1
    End Select

End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawBackground
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Draws the background gradient.
'---------------------------------------------------------------------------------------

Private Function DrawBackground(lLightColor As OLE_COLOR, _
                                lDarkColor As OLE_COLOR)

    On Error GoTo ErrHandler
    
    Dim xx, R1, R2, G1, G2, B1, B2, Rs, Gs, Bs, Rx, Gx, Bx
    Dim lColor As Long, lColor2 As Long
    
    lColor = TranslateColor(lLightColor)
    lColor2 = TranslateColor(lDarkColor)
            
    R1 = GetRed(lColor): R2 = GetRed(lColor2)
    G1 = GetGreen(lColor): G2 = GetGreen(lColor2)
    B1 = GetBlue(lColor): B2 = GetBlue(lColor2)
    
    If Style = [Header Visible] Then
        Rx = R1: Gx = G1: Bx = B1
        Rs = (R1 - R2) / (Picture2.ScaleHeight - 1)
        Gs = (G1 - G2) / (Picture2.ScaleHeight - 1)
        Bs = (B1 - B2) / (Picture2.ScaleHeight - 1)
            
        For xx = 24 To UserControl.ScaleHeight - 1
            UserControl.Line (0, xx)-(Picture2.ScaleWidth, xx), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next xx

    Else
        Rx = R1: Gx = G1: Bx = B1
        Rs = (R1 - R2) / (UserControl.ScaleHeight - 1)
        Gs = (G1 - G2) / (UserControl.ScaleHeight - 1)
        Bs = (B1 - B2) / (UserControl.ScaleHeight - 1)
    
        For xx = 0 To UserControl.ScaleHeight - 1
            UserControl.Line (0, xx)-(UserControl.ScaleWidth, xx), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next xx

    End If
    
ErrHandler:
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawBorder
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Draws a border around the control.
'---------------------------------------------------------------------------------------

Private Function DrawBorder(lBorderColor As OLE_COLOR)

    On Error GoTo ErrHandler
    
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), lBorderColor, B
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), lBorderColor
    UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), lBorderColor
    
ErrHandler:
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawHeader
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Draws the header gradient.
'---------------------------------------------------------------------------------------

Private Function DrawHeader(lLightColor As OLE_COLOR, _
                            lDarkColor As OLE_COLOR, _
                            lTextColor As OLE_COLOR)

    On Error GoTo ErrHandler
    
    Dim xx, R1, R2, G1, G2, B1, B2, Rs, Gs, Bs, Rx, Gx, Bx
    Dim lColor As Long, lColor2 As Long
    
    lColor = TranslateColor(lLightColor)
    lColor2 = TranslateColor(lDarkColor)
            
    R1 = GetRed(lColor): R2 = GetRed(lColor2)
    G1 = GetGreen(lColor): G2 = GetGreen(lColor2)
    B1 = GetBlue(lColor): B2 = GetBlue(lColor2)

    Rx = R1: Gx = G1: Bx = B1
    Rs = (R1 - R2) / (Picture1.ScaleHeight - 1)
    Gs = (G1 - G2) / (Picture1.ScaleHeight - 1)
    Bs = (B1 - B2) / (Picture1.ScaleHeight - 1)

    For xx = 0 To Picture1.ScaleHeight - 1
        Picture1.Line (0, xx)-(Picture1.ScaleWidth, xx), RGB(Rx, Gx, Bx)
        Rx = Rx - Rs
        Gx = Gx - Gs
        Bx = Bx - Bs
    Next xx
        
    Label1.ForeColor = lTextColor
        
ErrHandler:
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetBlue
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Gets the Blue value of an RGB Color.
'---------------------------------------------------------------------------------------

Private Function GetBlue(iColor As Long) As Integer
    GetBlue = ((iColor And &HFF0000) / 65536)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetGreen
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Gets the Green value of an RGB Color.
'---------------------------------------------------------------------------------------

Private Function GetGreen(iColor As Long) As Integer
    GetGreen = ((iColor And &HFF00FF00) / 256&)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRed
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Gets the Red value of an RGB Color.
'---------------------------------------------------------------------------------------

Private Function GetRed(iColor As Long) As Integer
    GetRed = iColor Mod 256
End Function

'---------------------------------------------------------------------------------------
' Procedure : RedrawControl
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Draws the control and sets the label caption.
'---------------------------------------------------------------------------------------

Private Sub RedrawControl()

    UserControl.Cls
    Label1.Caption = m_Caption
    
    If Style = [Header Visible] Then
        Picture1.Visible = True
        DrawHeader m_HeaderLightColor, m_HeaderDarkColor, m_TextColor
        DrawBackground m_BackLightColor, m_BackDarkColor
        DrawBorder m_BorderColor
    Else
        Picture1.Visible = False
        DrawBackground m_BackLightColor, m_BackDarkColor
        DrawBorder m_BorderColor
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TranslateColor
' Date      : 08/04/2005
' Author    : Cameron Groves
' Purpose   : Translates an OLE color to a long color value.
'---------------------------------------------------------------------------------------

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Initialize
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    ' Used to prevent crashes on Windows XP
    m_hMod = LoadLibrary("shell32.dll")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_InitProperties
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_InitProperties()
    m_HeaderLightColor = m_def_HeaderLightColor
    m_HeaderDarkColor = m_def_HeaderDarkColor
    m_BackLightColor = m_def_BackLightColor
    m_BackDarkColor = m_def_BackDarkColor
    m_BorderColor = m_def_BorderColor
    m_TextColor = m_def_TextColor
    m_Caption = m_def_Caption
    m_Style = m_def_Style
    m_Theme = m_def_Theme
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Paint
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_Paint()
    RedrawControl
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_ReadProperties
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_HeaderLightColor = PropBag.ReadProperty("HeaderLightColor", m_def_HeaderLightColor)
    m_HeaderDarkColor = PropBag.ReadProperty("HeaderDarkColor", m_def_HeaderDarkColor)
    m_BackLightColor = PropBag.ReadProperty("BackLightColor", m_def_BackLightColor)
    m_BackDarkColor = PropBag.ReadProperty("BackDarkColor", m_def_BackDarkColor)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Resize
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_Resize()

    On Error GoTo ErrHandler

    If UserControl.Width <> 0 Then
        Label1.Top = (Picture1.ScaleHeight - Label1.Height) / 2
        Picture1.Width = UserControl.ScaleWidth - 2
        Picture2.Width = Picture1.Width
        Picture2.Height = UserControl.ScaleHeight - (Picture1.Height + 2)
    End If
    
    RedrawControl

ErrHandler:
    Exit Sub
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Show
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Terminate
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_Terminate()
    ' Used to prevent crashes on Windows XP
    FreeLibrary m_hMod
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_WriteProperties
' Date      : 08/04/2005
' Author    : Cameron Groves
'---------------------------------------------------------------------------------------

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HeaderLightColor", m_HeaderLightColor, m_def_HeaderLightColor)
    Call PropBag.WriteProperty("HeaderDarkColor", m_HeaderDarkColor, m_def_HeaderDarkColor)
    Call PropBag.WriteProperty("BackLightColor", m_BackLightColor, m_def_BackLightColor)
    Call PropBag.WriteProperty("BackDarkColor", m_BackDarkColor, m_def_BackDarkColor)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("TextColor", m_TextColor, m_def_TextColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
End Sub

'---------------------------------------------------------------------------------------
' Properties
'---------------------------------------------------------------------------------------

Public Property Get BackDarkColor() As OLE_COLOR
    BackDarkColor = m_BackDarkColor
End Property

Public Property Let BackDarkColor(ByVal New_BackDarkColor As OLE_COLOR)
    m_BackDarkColor = New_BackDarkColor
    PropertyChanged "BackDarkColor"
    RedrawControl
End Property

Public Property Get BackLightColor() As OLE_COLOR
    BackLightColor = m_BackLightColor
End Property

Public Property Let BackLightColor(ByVal New_BackLightColor As OLE_COLOR)
    m_BackLightColor = New_BackLightColor
    PropertyChanged "BackLightColor"
    RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    RedrawControl
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    RedrawControl
End Property

Public Property Get HeaderDarkColor() As OLE_COLOR
    HeaderDarkColor = m_HeaderDarkColor
End Property

Public Property Let HeaderDarkColor(ByVal New_HeaderDarkColor As OLE_COLOR)
    m_HeaderDarkColor = New_HeaderDarkColor
    PropertyChanged "HeaderDarkColor"
    RedrawControl
End Property

Public Property Get HeaderLightColor() As OLE_COLOR
    HeaderLightColor = m_HeaderLightColor
End Property

Public Property Let HeaderLightColor(ByVal New_HeaderLightColor As OLE_COLOR)
    m_HeaderLightColor = New_HeaderLightColor
    PropertyChanged "HeaderLightColor"
    RedrawControl
End Property

Public Property Get Style() As XPContainerStyles
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As XPContainerStyles)
    m_Style = New_Style
    PropertyChanged "Style"
    RedrawControl
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    m_TextColor = New_TextColor
    PropertyChanged "TextColor"
    RedrawControl
End Property

Public Property Get Theme() As XPContainerThemes
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As XPContainerThemes)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    ApplyTheme
End Property

