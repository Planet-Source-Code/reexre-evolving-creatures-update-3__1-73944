VERSION 5.00
Begin VB.Form frmEDITOR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Creature Editor"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   2625
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   Begin VB.CommandButton cmdDELETE 
      Caption         =   "Delete From Disk"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   68
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmEDITOR.frx":0000
      Top             =   0
      Width           =   8415
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE AS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   6120
      Pattern         =   "*.cre"
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdCLear 
      Caption         =   "CLR ALL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.CheckBox chBrainLink 
      Caption         =   "with Brain"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton oAddLink 
      Caption         =   "Add Links"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton oAddPoint 
      Caption         =   "Add Points"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Double Click to Load"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "frmEDITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CR             As New clsCreature
Dim P1             As Long
Dim P2             As Long

Private Sub cmdCLear_Click()
    CR.ClearAll
    Me.Cls
    oAddPoint = True

End Sub

Private Sub cmdDELETE_Click()
    Dim R          As Long
    R = MsgBox("Are you Sure to Delete  " & Chr$(34) & CreFN & Chr$(34) & " ?", vbYesNo)

    If R = vbYes Then
        If Dir(App.Path & "\Creatures\" & CreFN) <> vbNullString Then
            Kill App.Path & "\Creatures\" & CreFN
            If Dir(App.Path & "\Creatures\" & CreFN & "Pop.txt") <> vbNullString Then Kill App.Path & "\Creatures\" & CreFN & "Pop.txt"
        Else
            MsgBox Chr$(34) & CreFN & Chr$(34) & "  already deleted."
        End If
        File1.Refresh
    End If


End Sub

Private Sub cmdSAVE_Click()
    Dim FN         As String
ag:
    FN = InputBox("Type FileName", "Save as ", "new")
    If Len(FN) < 3 Then MsgBox "Too short Name": GoTo ag
    CR.SaveMe "Creatures\" & FN
    File1.Refresh
    CreFN = FN

End Sub

Private Sub File1_DblClick()
    BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hdc, 0, 0, vbBlack    'ness
    CR.LoadMe "Creatures\" & File1, 50, Me.ScaleHeight - 50
    CR.DRAW Me.hdc, 0, True
    Me.Refresh
    CreFN = File1
End Sub

Private Sub Form_Activate()
    File1.Path = App.Path & "\Creatures"
    File1.Refresh

    CR.DRAW Me.hdc, 0, True
    Me.Refresh
End Sub

Private Sub Form_Load()
    CR.LoadMe "Creatures\" & "default.cre", 50, Me.ScaleHeight - 50


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I          As Long
    Dim dx         As Single
    Dim dy         As Single
    Dim D          As Single
    Dim Dmin       As Single

    If oAddPoint Then
        CR.ADDpoint X, Y
        For I = 1 To CR.NP
            MyCircle Me.hdc, CR.getpointX(I), CR.getpointY(I), 3, 2, vbBlue
            Me.Refresh

        Next
    Else
        If P1 = 0 Then
            Dmin = 99999999999#
            For I = 1 To CR.NP
                dx = CR.getpointX(I) - X
                dy = CR.getpointY(I) - Y
                D = Sqr(dx * dx + dy * dy)
                If D < Dmin Then Dmin = D: P1 = I
            Next

        Else

        End If
    End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I          As Long

    If Button = 1 And oAddLink And P1 <> 0 Then
        BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hdc, 0, 0, vbBlack    'ness
        For I = 1 To CR.NP
            MyCircle Me.hdc, CR.getpointX(I), CR.getpointY(I), 3, 2, vbBlue
        Next
        FastLine Me.hdc, X \ 1, Y \ 1, CR.getpointX(P1), CR.getpointY(P1), 1, vbWhite
        CR.DRAW Me.hdc, 0, True
        Me.Refresh

    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I          As Long
    Dim dx         As Single
    Dim dy         As Single
    Dim D          As Single
    Dim Dmin       As Single

    If oAddLink Then
        If P1 <> 0 Then
            Dmin = 99999999999#
            For I = 1 To CR.NP
                dx = CR.getpointX(I) - X
                dy = CR.getpointY(I) - Y
                D = Sqr(dx * dx + dy * dy)
                If D < Dmin And I <> P1 Then Dmin = D: P2 = I
            Next
        End If
        Me.Caption = P1 & "  " & P2
        CR.AddLink P1, P2, ST, chBrainLink
        P1 = 0
        P2 = 0
        BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hdc, 0, 0, vbBlack    'ness
        CR.DRAW Me.hdc, 0, True
        Me.Refresh
    End If

End Sub

Private Sub oAddLink_Click()
    chBrainLink.Visible = oAddLink

End Sub

Private Sub oAddPoint_Click()
    chBrainLink.Visible = oAddLink
End Sub
