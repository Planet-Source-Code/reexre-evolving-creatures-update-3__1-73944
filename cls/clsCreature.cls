VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NP          As Long
Public InvNP       As Single

Private P()        As tPoint

Public NL          As Long
Private L()        As tLink

Public LinkLeft    As Long
Public LinkRight   As Long


Public CurrCenX    As Single
Public CurrCenY    As Single
Public CurrVX      As Single
Public CurrVY      As Single
Public CurrDX      As Single
Public CurrDY      As Single
Public CurrHeightfromGround As Single

Public InvSIZE     As Single

Private Const BrainedLinkMAXExcursion As Single = 0.8    '.6  '1=from 50% to 150%

Public TaskWalk    As Single
Public TaskAVGheight As Single
Public TaskMAXHeight As Single
Public TaskDontJUMP As Single
Public TaskVerticality As Single
Public TaskHorizontality As Single

Private Const Walk_Mult As Single = 0.01
Private Const AVGH_Mult As Single = 0.14
Private Const MaxH_Mult As Single = 0.05
Private Const DontJump_Mult As Single = 1
Private Const Orient_Mult As Single = 20

Public DrawOffset As Long
Public CurrANG As Single
Public OLDAng As Single
Public RotVEL As Single

Public Property Get getpointX(wP As Long) As Single
    getpointX = P(wP).X
End Property
Public Property Get getpointY(wP As Long) As Single
    getpointY = P(wP).Y
End Property

Public Property Let AdjAddedBrainLenToReach(wL As Long, Value As Single)
    'Value Must be between -0.5 and 0.5
    L(wL).AddedBrainLenToReach = L(wL).MainLen * Value * BrainedLinkMAXExcursion

End Property
Public Property Get GetLinkP1(wL As Long) As Single
    GetLinkP1 = L(wL).P1
End Property
Public Property Get GetLinkP2(wL As Long) As Single
    GetLinkP2 = L(wL).P2
End Property
Public Property Get GetLinkTension(wL As Long) As Single
    GetLinkTension = L(wL).TENSION
End Property
Public Property Get GetLinkMainL(wL As Long) As Single
    GetLinkMainL = L(wL).MainLen
End Property
Public Property Get GetCurrAddedBrainLen(wL As Long) As Single
GetCurrAddedBrainLen = L(wL).CurrAddedBrainLen
End Property
Public Property Get GetTouchGround(wP As Long) As Single
    GetTouchGround = P(wP).IsOnGround
End Property

Public Property Get IsLinkDynamic(wL As Long) As Boolean
    IsLinkDynamic = L(wL).Dynamic
End Property

Public Property Let IsLinkDynamic(wL As Long, V As Boolean)
    L(wL).Dynamic = V
End Property

'Public Property Get GetDX() As Single
'    'GetDX = (P(NP).X - P(1).X)
'    GetDX = 0.5 * (P(L(NL).P1).X + P(L(NL).P2).X - P(L(1).P1).X - P(L(1).P2).X)
'End Property
'Public Property Get GetDY() As Single
'    'Getdy = (P(NP).y - P(1).y)
'    GetDY = 0.5 * (P(L(NL).P1).Y + P(L(NL).P2).Y - P(L(1).P1).Y - P(L(1).P2).Y)
'End Property
'Public Property Get GetPOS() As Single
'    GetPOS = (P(1).X + P(NP).X) * 0.5
'End Property
'Public Property Get GetHeightAVG() As Single
'    Dim I          As Long
'    For I = 1 To NP Step 1        '3
'        GetHeightAVG = GetHeightAVG + P(I).Y
'    Next
'    GetHeightAVG = GetHeightAVG / (NP)
'End Property
'Public Property Get GetHeightfromGround() As Single
'    Dim I          As Long
'    Dim Max        As Single
'    Max = 0
'    For I = 1 To NP
'        If (MaxY - P(I).Y) > Max Then Max = MaxY - P(I).Y
'        GetHeightfromGround = Max
'    Next
'End Property
'Public Property Get GetVY() As Single
'    Dim I          As Long
'    For I = 1 To NP
'        GetVY = GetVY + P(I).vY
'    Next
'    GetVY = GetVY / NP
'End Property
Public Sub ADDpoint(X As Single, Y As Single)
    NP = NP + 1
    ReDim Preserve P(NP)
    With P(NP)
        .X = X
        .Y = Y
        .vX = 0
        .vY = 0
        .IsOnGround = 0
        .IvX = 0
        .IvY = 0
    End With
    InvNP = 1 / NP
End Sub

Public Sub AddLink(P1 As Long, P2 As Long, ByVal Stren As Single, Optional IsDynamic As Boolean = False)
    Dim Xmax       As Single
    Dim Xmin       As Single
    Dim Ymax       As Single
    Dim Ymin       As Single

    Dim X          As Single
    Dim Y          As Single
    Dim I          As Long
    Dim V          As Single
    Dim lTOP       As Long
    Dim lBOTTOM    As Long

    'If IsDynamic Then Stren = Stren * 0.75


    NL = NL + 1
    ReDim Preserve L(NL)
    With L(NL)
        .P1 = P1
        .P2 = P2
        .Dynamic = IsDynamic
        .MainLen = PointDist(P(P1), P(P2))
        .InvMainLen = 1 / .MainLen
        .Stren = Stren

        .AddedBrainLenToReach = 0
        .CurrAddedBrainLen = 0
        
        .TENSION = 0
       
        
    End With

    Xmax = -999999999
    Xmin = 99999999999#
    Ymax = -999999999
    Ymin = 99999999999#


    For I = 1 To NL
        X = (P(L(I).P1).X + P(L(I).P2).X) * 0.5
        If X > Xmax Then Xmax = X: LinkRight = I
        If X < Xmin Then Xmin = X: LinkLeft = I
        Y = (P(L(I).P1).Y + P(L(I).P2).Y) * 0.5
        If Y > Ymax Then Ymax = Y: lBOTTOM = I
        If Y < Ymin Then Ymin = Y: lTOP = I
    Next


    If Xmax - Xmin > Ymax - Ymin Then
        V = (P(L(LinkRight).P1).X + P(L(LinkRight).P2).X) * 0.5 - _
            (P(L(LinkLeft).P1).X + P(L(LinkLeft).P2).X) * 0.5
    Else
        V = (P(L(lBOTTOM).P1).Y + P(L(lBOTTOM).P2).Y) * 0.5 - _
            (P(L(lTOP).P1).Y + P(L(lTOP).P2).Y) * 0.5
    End If

    If V <> 0 Then InvSIZE = 1 / V

End Sub

Public Sub ClearAll()
    NP = 0
    NL = 0
    CurrCenX = 0
    CurrCenY = 0
    CurrVX = 0
    CurrVY = 0
    CurrDX = 0
    CurrDY = 0
    CurrHeightfromGround = 0
    
    CurrANG = 0
    OLDAng = 0
    RotVEL = 0
    
    ResetTaskValues

End Sub
Public Sub ResetTaskValues()
    TaskWalk = 30
    TaskAVGheight = 0
    TaskMAXHeight = 1E+17
    TaskDontJUMP = 0
    TaskVerticality = 0
    TaskHorizontality = 0
End Sub
Public Sub DRAW(hdc As Long, ByVal drPANX As Long, Optional DrawPts As Boolean = False, Optional Antialias As Boolean = False)
    Dim x1         As Long
    Dim y1         As Long
    Dim X2         As Long
    Dim Y2         As Long

    Dim I          As Long

    Dim C          As Long
    Dim C1         As Long
    Dim C2         As Long

    Dim V          As Single

    For I = 1 To NL
        With L(I)
            x1 = P(.P1).X - drPANX + DrawOffset
            y1 = P(.P1).Y
            X2 = P(.P2).X - drPANX + DrawOffset
            Y2 = P(.P2).Y

            C1 = .TENSION * 255
            If Abs(C1) > 255 Then C1 = Sgn(C1) * 255

            If Antialias Then
                If .Dynamic Then
                    C = GradDynamic2.GetGrad(C1)
                Else
                    C = GradStatic2.GetGrad(C1)
                End If
                AAA.LineDIB x1, y1, X2, Y2, C
            Else
                If .Dynamic Then
                    C = GradDynamic1.GetGrad(C1)
                Else
                    C = GradStatic1.GetGrad(C1)
                End If
                FastLine hdc, x1, y1, X2, Y2, 2, C
            End If
        End With


        'Y1 = CurrCenY \ 1
        'X1 = CurrCenX \ 1
        'Y2 = Y1 + CurrVY * 100
        'X2 = X1 + CurrVX * 100
        'FastLine HDC, X1, Y1, X2, Y2, 2, vbRed
        '
        'X2 = X1 + (CurrDX * 0.5) \ 1
        'Y2 = Y1 + (CurrDY * 0.5) \ 1
        'FastLine HDC, X1, Y1, X2, Y2, 2, vbGreen
    Next

    If DrawPts Then
        For I = 1 To NP
            C = P(I).IsOnGround * 510 - 255
            If Antialias Then
                C = GradGround2.GetGrad(C)
                AAA.CircleDIB P(I).X \ 1 - drPANX + DrawOffset, P(I).Y \ 1, 1, 1, C
                AAA.CircleDIB P(I).X \ 1 - drPANX + DrawOffset, P(I).Y \ 1, 2, 2, C
            Else
                C = GradGround1.GetGrad(C)
                MyCircle hdc, P(I).X \ 1 - drPANX + DrawOffset, P(I).Y \ 1, 3, 2, C
            End If
        Next
    End If



End Sub

Public Sub DoPhysics2(ByVal MyIDX As Long)
    Dim dx         As Single
    Dim dy         As Single
    Dim ndx        As Single
    Dim ndy        As Single
    Dim D          As Single
    Dim I          As Long
    Dim resF       As Single
    Dim vX         As Single
    Dim vY         As Single
    Dim CX         As Single
    Dim CY         As Single
    Dim DDD        As Single

    Dim A          As Single
    Dim vx1        As Single
    Dim vy1        As Single
    Dim vx2        As Single
    Dim vy2        As Single
    Dim MAG        As Single


    For I = 1 To NL
        With L(I)

            dx = P(.P2).X - P(.P1).X
            dy = P(.P2).Y - P(.P1).Y
            dx = dx + (P(.P2).vX - P(.P1).vX) * 10
            dy = dy + (P(.P2).vY - P(.P1).vY) * 10

            D = Sqr(dx * dx + dy * dy)
            If D <> 0 Then
                ndx = dx / D
                ndy = dy / D
            End If

            'Update3:
            .CurrAddedBrainLen = .CurrAddedBrainLen * 0.98 + .AddedBrainLenToReach * 0.02
            'smooth

            If .Dynamic Then
                resF = (.MainLen - D + .CurrAddedBrainLen)
                .TENSION = 3 * resF / (.MainLen + .CurrAddedBrainLen)
                resF = resF * .Stren
            Else
                resF = (.MainLen - D)
                .TENSION = 3 * resF * .InvMainLen
                resF = resF * .Stren
            End If
            resF = resF * 0.01    '0.01   ' 0.012

            dx = ndx * resF
            dy = ndy * resF

            If Abs(.TENSION) > 1 Then .TENSION = Sgn(.TENSION)

            P(.P1).IvX = P(.P1).IvX - dx
            P(.P1).IvY = P(.P1).IvY - dy
            P(.P2).IvX = P(.P2).IvX + dx
            P(.P2).IvY = P(.P2).IvY + dy
        End With
    Next


    CurrHeightfromGround = 9999999
    For I = 1 To NP
        With P(I)
            .IsOnGround = .IsOnGround * 0.985

            .vX = .vX + .IvX: .IvX = 0
            .vY = .vY + .IvY: .IvY = 0

            .vY = .vY + GravityY
            .vX = .vX * AirResistence

            .X = .X + .vX
            .Y = .Y + .vY

            If .Y > MaxY Then
                ' Update2: smoothed touch ground sensor (Sort of Pressure) ---------------
                .IsOnGround = .IsOnGround + .vY * 15    ' 100 * Sqr(.vY * .vY + .vX * .vX)
                If .IsOnGround > 1 Then .IsOnGround = 1
                '----------------------------------------

                If MyIDX = CurBEST Then

                    A = Atan2(.vX, -.vY)
                    MAG = Sqr(.vX * .vX + .vY * .vY) * 1.1
                    If MAG > 0.02 Then
                        vx1 = Cos(A + 0.25) * MAG
                        vy1 = Sin(A + 0.25) * MAG
                        vx2 = Cos(A - 0.25) * MAG
                        vy2 = Sin(A - 0.25) * MAG
                        AddParticle .X + DrawOffset, .Y - 1, .vX * 1.2, -.vY * 1.2 - 0.01
                        AddParticle .X + DrawOffset, .Y - 1, vx1, vy1 - 0.01
                        AddParticle .X + DrawOffset, .Y - 1, vx2, vy2 - 0.01
                    End If

                End If

                .Y = MaxY
                .vY = -.vY * Bounce
                .vX = .vX * Friction

            End If

            CX = CX + .X
            CY = CY + .Y
            vX = vX + .vX
            vY = vY + .vY
            If (MaxY - .Y) < CurrHeightfromGround Then CurrHeightfromGround = MaxY - .Y
        End With
    Next

    CurrCenX = CX * InvNP
    CurrCenY = CY * InvNP
    CurrVY = vY * InvNP
    CurrVX = vX * InvNP
    CurrDX = (P(L(LinkRight).P1).X + P(L(LinkRight).P2).X - P(L(LinkLeft).P1).X - P(L(LinkLeft).P2).X)    '*.5
    CurrDY = (P(L(LinkRight).P1).Y + P(L(LinkRight).P2).Y - P(L(LinkLeft).P1).Y - P(L(LinkLeft).P2).Y)    '*.5

    OLDAng = CurrANG
    CurrANG = Atan2(CurrDX, CurrDY)
    RotVEL = 100 * AngleDIFF(OLDAng, CurrANG)
    If RotVEL > 1 Then
        RotVEL = 1
    ElseIf RotVEL < -1 Then RotVEL = -1
    End If

    DDD = Sqr(CurrDX * CurrDX + CurrDY * CurrDY)
    If DDD <> 0 Then CurrDX = CurrDX / DDD: CurrDY = CurrDY / DDD

    TaskWalk = TaskWalk - (CurrVX) * Walk_Mult
    TaskAVGheight = TaskAVGheight + CurrCenY * AVGH_Mult
    If CurrCenY < TaskMAXHeight Then TaskMAXHeight = CurrCenY * MaxH_Mult
    TaskDontJUMP = TaskDontJUMP + CurrHeightfromGround * DontJump_Mult
    TaskVerticality = TaskVerticality + Abs(CurrDX) * Orient_Mult
    TaskHorizontality = TaskHorizontality + Abs(CurrDY) * Orient_Mult






End Sub
Public Sub SaveMe(ByVal filename As String)
    Dim MinX       As Long
    Dim MinY       As Long
    Dim MaxX       As Long
    Dim MaxY       As Long
    Dim I          As Long

    If Right$(filename, 4) <> ".cre" Then filename = filename & ".cre"
    filename = App.Path & "\" & filename


    MaxX = -1
    MaxY = -1
    MinX = 9999999
    MinY = 9999999

    For I = 1 To NP
        If P(I).X > MaxX Then MaxX = P(I).X
        If P(I).Y > MaxY Then MaxY = P(I).Y
        If P(I).X < MinX Then MinX = P(I).X
        If P(I).Y < MinY Then MinY = P(I).Y
    Next


    Open filename For Output As 1
    Print #1, "Points:"
    Print #1, NP
    For I = 1 To NP
        Print #1, P(I).X - MinX
        Print #1, P(I).Y - MaxY
    Next
    Print #1, "Links:"
    Print #1, NL
    For I = 1 To NL
        Print #1, L(I).P1
        Print #1, L(I).P2
        Print #1, Replace(L(I).Stren, ",", ".")
        Print #1, CLng(L(I).Dynamic)
    Next

    Close 1

End Sub

Public Sub LoadMe(ByVal filename As String, Left As Long, Bottom As Long)

    Dim I          As Long
    Dim S          As String
    Dim V1         As Long
    Dim V2         As Long
    Dim Stre       As Single
    Dim v4         As Long
    Dim N          As Long

    If Right$(filename, 4) <> ".cre" Then filename = filename & ".cre"
    filename = App.Path & "\" & filename


    ClearAll

    Open filename For Input As 1
    Input #1, S                   '"Points:"
    Input #1, N
    For I = 1 To N
        Input #1, V1
        Input #1, V2
        ADDpoint V1 + Left, V2 + Bottom
    Next
    Input #1, S                   ' "Links:"
    Input #1, N
    For I = 1 To N
        Input #1, V1
        Input #1, V2
        Input #1, Stre
        Input #1, v4
        AddLink V1, V2, Stre, CBool(v4)
    Next

    Close 1


End Sub
