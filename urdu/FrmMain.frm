VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENGLISH AND URDU WRITE TOGATHER"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "PRESS BUTTON FOR ENGLISH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "14208;6588"
      FontHeight      =   480
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
' This Code will help you to write Two Scripts
' English and Urdu
' Infact I use Arbic Unicode, some Urdu Characters
' are missing.
' If you are interesting in developing Urdu and
' English Application, this Code will help you
' alot, and if you feel any problem then write me
' I will try my best to help you.
' Note:
' First of all you have to incdule MS Form2 Library
' Bcz only form2 library support unicodes

' Email me for more assistance
' perpetual_elate@msn.com
' URL : www.webweavers.cyberkings.com
'**************************************************

Dim ModeValue As Boolean
Dim UniCode
Private Sub Command1_Click()
If ModeValue = False Then
    ModeValue = True
    Command1.Caption = "PRESS BUTTON FOR URDU"
ElseIf ModeValue = True Then
    ModeValue = False
    Command1.Caption = "PRESS BUTTON FOR ENGLISH"
End If
End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
        If KeyCode = 32 Then 'space
        UniCode = &H20
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        ElseIf KeyCode = 13 Then 'enter
        UniCode = &HA
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        ElseIf KeyCode = 9 Then 'horizontal tab
        UniCode = &H9
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        ElseIf KeyCode = 127 Then 'delete
        UniCode = &H7F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        KeyCode = 0
        End If
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If ModeValue = False Then
        If KeyAscii = 97 Or TextBox1.SelText <> "" Then 'a
        TextBox1.SelText = ""
        UniCode = &H627
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 98 Or TextBox1.SelText <> "" Then 'b
        TextBox1.SelText = ""
        UniCode = &H628
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 99 Or TextBox1.SelText <> "" Then 'c
        TextBox1.SelText = ""
        UniCode = &H686
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 100 Or TextBox1.SelText <> "" Then 'd
        TextBox1.SelText = ""
        UniCode = &H62F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 101 Or TextBox1.SelText <> "" Then 'e
        TextBox1.SelText = ""
        UniCode = &H639
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 102 Or TextBox1.SelText <> "" Then 'f
        TextBox1.SelText = ""
        UniCode = &H641
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 103 Or TextBox1.SelText <> "" Then 'g
        TextBox1.SelText = ""
        UniCode = &H6AF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 104 Or TextBox1.SelText <> "" Then 'h
        TextBox1.SelText = ""
        UniCode = &H6C1
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 105 Or TextBox1.SelText <> "" Then 'i
        TextBox1.SelText = ""
        UniCode = &H64A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 106 Or TextBox1.SelText <> "" Then 'j
        TextBox1.SelText = ""
        UniCode = &H62C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 107 Or TextBox1.SelText <> "" Then 'k
        TextBox1.SelText = ""
        UniCode = &H6A9
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 108 Or TextBox1.SelText <> "" Then 'l
        TextBox1.SelText = ""
        UniCode = &H644
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 109 Or TextBox1.SelText <> "" Then 'm
        TextBox1.SelText = ""
        UniCode = &H645
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 110 Or TextBox1.SelText <> "" Then 'n
        TextBox1.SelText = ""
        UniCode = &H646
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 111 Or TextBox1.SelText <> "" Then 'o
        TextBox1.SelText = ""
        UniCode = &H647
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 112 Or TextBox1.SelText <> "" Then 'p
        TextBox1.SelText = ""
        UniCode = &H67E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 113 Or TextBox1.SelText <> "" Then 'q
        TextBox1.SelText = ""
        UniCode = &H642
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 114 Or TextBox1.SelText <> "" Then 'r
        TextBox1.SelText = ""
        UniCode = &H631
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 115 Or TextBox1.SelText <> "" Then 's
        TextBox1.SelText = ""
        UniCode = &H633
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 116 Or TextBox1.SelText <> "" Then 't
        TextBox1.SelText = ""
        UniCode = &H62A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 117 Or TextBox1.SelText <> "" Then 'u
        TextBox1.SelText = ""
        UniCode = &H621
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 118 Or TextBox1.SelText <> "" Then 'v
        TextBox1.SelText = ""
        UniCode = &H637
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 119 Or TextBox1.SelText <> "" Then 'w
        TextBox1.SelText = ""
        UniCode = &H648
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 120 Or TextBox1.SelText <> "" Then 'x
        TextBox1.SelText = ""
        UniCode = &H634
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 121 Or TextBox1.SelText <> "" Then 'y
        TextBox1.SelText = ""
        UniCode = &H6D12
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 122 Or TextBox1.SelText <> "" Then 'z
        TextBox1.SelText = ""
        UniCode = &H632
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        ' for capital chracters
        
        ElseIf KeyAscii = 65 Or TextBox1.SelText <> "" Then 'A
        TextBox1.SelText = ""
        UniCode = &H622
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 66 Or TextBox1.SelText <> "" Then 'B
        TextBox1.SelText = ""
        UniCode = &HFBB0
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 67 Or TextBox1.SelText <> "" Then 'C
        TextBox1.SelText = ""
        UniCode = &H62B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 68 Or TextBox1.SelText <> "" Then 'D
        TextBox1.SelText = ""
        UniCode = &H688
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 69 Or TextBox1.SelText <> "" Then 'E
        TextBox1.SelText = ""
        UniCode = &H650
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 70 Or TextBox1.SelText <> "" Then 'F
        TextBox1.SelText = ""
        UniCode = &H652
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 71 Or TextBox1.SelText <> "" Then 'G
        TextBox1.SelText = ""
        UniCode = &H63A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 72 Or TextBox1.SelText <> "" Then 'H
        TextBox1.SelText = ""
        UniCode = &H62D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 73 Or TextBox1.SelText <> "" Then 'I
        TextBox1.SelText = ""
        UniCode = &H649
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 74 Or TextBox1.SelText <> "" Then 'J
        TextBox1.SelText = ""
        UniCode = &H636
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 75 Or TextBox1.SelText <> "" Then 'K
        TextBox1.SelText = ""
        UniCode = &H62E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 76 Or TextBox1.SelText <> "" Then 'L
        TextBox1.SelText = ""
        UniCode = &HFEFB
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 77 Or TextBox1.SelText <> "" Then 'M
        TextBox1.SelText = ""
        UniCode = &H66B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 78 Or TextBox1.SelText <> "" Then 'N
        TextBox1.SelText = ""
        UniCode = &H6BA
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 79 Or TextBox1.SelText <> "" Then 'O
        TextBox1.SelText = ""
        UniCode = &H629
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 80 Or TextBox1.SelText <> "" Then 'P
        TextBox1.SelText = ""
        UniCode = &H64F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 81 Or TextBox1.SelText <> "" Then 'Q
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 82 Or TextBox1.SelText <> "" Then 'R
        TextBox1.SelText = ""
        UniCode = &H691
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 83 Or TextBox1.SelText <> "" Then 'S
        TextBox1.SelText = ""
        UniCode = &H635
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 84 Or TextBox1.SelText <> "" Then 'T
        TextBox1.SelText = ""
        UniCode = &H679
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 85 Or TextBox1.SelText <> "" Then 'U
        TextBox1.SelText = ""
        UniCode = &H626
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 86 Or TextBox1.SelText <> "" Then 'V
        TextBox1.SelText = ""
        UniCode = &H638
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 87 Or TextBox1.SelText <> "" Then 'W
        TextBox1.SelText = ""
        UniCode = &H624
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 88 Or TextBox1.SelText <> "" Then 'X
        TextBox1.SelText = ""
        UniCode = &H698
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 89 Or TextBox1.SelText <> "" Then 'Y
        TextBox1.SelText = ""
        UniCode = &HFBAF
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 90 Or TextBox1.SelText <> "" Then 'Z
        TextBox1.SelText = ""
        UniCode = &H630
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'for nummaric values
        ElseIf KeyAscii = 48 Or TextBox1.SelText <> "" Then '0
        TextBox1.SelText = ""
        UniCode = &H660
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 49 Or TextBox1.SelText <> "" Then '1
        TextBox1.SelText = ""
        UniCode = &H661
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 50 Or TextBox1.SelText <> "" Then '2
        TextBox1.SelText = ""
        UniCode = &H662
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 51 Or TextBox1.SelText <> "" Then '3
        TextBox1.SelText = ""
        UniCode = &H663
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 52 Or TextBox1.SelText <> "" Then '4
        TextBox1.SelText = ""
        UniCode = &H664
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 53 Or TextBox1.SelText <> "" Then '5
        TextBox1.SelText = ""
        UniCode = &H665
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 54 Or TextBox1.SelText <> "" Then '6
        TextBox1.SelText = ""
        UniCode = &H666
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 55 Or TextBox1.SelText <> "" Then '7
        TextBox1.SelText = ""
        UniCode = &H667
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 56 Or TextBox1.SelText <> "" Then '8
        UniCode = &H668
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 57 Or TextBox1.SelText <> "" Then '9
        TextBox1.SelText = ""
        UniCode = &H669
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 41 Or TextBox1.SelText <> "" Then ')
        TextBox1.SelText = ""
        UniCode = &HFD3F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 33 Or TextBox1.SelText <> "" Then '!
        TextBox1.SelText = ""
        UniCode = &H21
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 64 Or TextBox1.SelText <> "" Then '@
        TextBox1.SelText = ""
        UniCode = &H40
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 35 Or TextBox1.SelText <> "" Then '#
        TextBox1.SelText = ""
        UniCode = &H23
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 36 Or TextBox1.SelText <> "" Then '$
        TextBox1.SelText = ""
        UniCode = &H24
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 37 Or TextBox1.SelText <> "" Then '%
        TextBox1.SelText = ""
        UniCode = &H66A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 94 Or TextBox1.SelText <> "" Then '^
        TextBox1.SelText = ""
        UniCode = &H5E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 38 Or TextBox1.SelText <> "" Then '&
        TextBox1.SelText = ""
        UniCode = &H26
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 42 Or TextBox1.SelText <> "" Then '*
        TextBox1.SelText = ""
        UniCode = &H66D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 40 Or TextBox1.SelText <> "" Then  '(
        TextBox1.SelText = ""
        UniCode = &HFD3E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        
        'for special characters
        'symbols
        ElseIf KeyAscii = 63 Or TextBox1.SelText <> "" Then '?
        TextBox1.SelText = ""
        UniCode = &H61F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 47 Or TextBox1.SelText <> "" Then '/
        TextBox1.SelText = ""
        UniCode = &H2F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 44 Or TextBox1.SelText <> "" Then ',
        TextBox1.SelText = ""
        UniCode = &H60C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 46 Or TextBox1.SelText <> "" Then '.
        TextBox1.SelText = ""
        UniCode = &H640
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 95 Or TextBox1.SelText <> "" Then '_
        TextBox1.SelText = ""
        UniCode = &H5F
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 45 Or TextBox1.SelText <> "" Then '-
        TextBox1.SelText = ""
        UniCode = &H2D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 43 Or TextBox1.SelText <> "" Then '+
        TextBox1.SelText = ""
        UniCode = &H2B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 61 Or TextBox1.SelText <> "" Then '=
        TextBox1.SelText = ""
        UniCode = &H3D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 58 Or TextBox1.SelText <> "" Then ':
        TextBox1.SelText = ""
        UniCode = &H3A
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 59 Or TextBox1.SelText <> "" Then ';
        TextBox1.SelText = ""
        UniCode = &H61B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 60 Or TextBox1.SelText <> "" Then '<
        TextBox1.SelText = ""
        UniCode = &H3C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 62 Or TextBox1.SelText <> "" Then '>
        TextBox1.SelText = ""
        UniCode = &H3E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 123 Or TextBox1.SelText <> "" Then '{
        TextBox1.SelText = ""
        UniCode = &H7B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 125 Or TextBox1.SelText <> "" Then '}
        TextBox1.SelText = ""
        UniCode = &H7D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 91 Or TextBox1.SelText <> "" Then '[
        TextBox1.SelText = ""
        UniCode = &H5B
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 93 Or TextBox1.SelText <> "" Then ']
        TextBox1.SelText = ""
        UniCode = &H5D
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 124 Or TextBox1.SelText <> "" Then '|
        TextBox1.SelText = ""
        UniCode = &H7C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 92 Or TextBox1.SelText <> "" Then '\
        TextBox1.SelText = ""
        UniCode = &H5C
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 126 Or TextBox1.SelText <> "" Then '~
        TextBox1.SelText = ""
        UniCode = &H7E
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 96 Or TextBox1.SelText <> "" Then '`
        TextBox1.SelText = ""
        UniCode = &H60
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 34 Or TextBox1.SelText <> "" Then '"
        TextBox1.SelText = ""
        UniCode = &H22
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        ElseIf KeyAscii = 39 Or TextBox1.SelText <> "" Then ''
        TextBox1.SelText = ""
        UniCode = &H27
        TextBox1.Text = TextBox1.Text + ChrW(UniCode)
        End If
        KeyAscii = 0
End If
End Sub

