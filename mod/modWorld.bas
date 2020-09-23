Attribute VB_Name = "modWorld"
Option Explicit

Public BR          As New simplyBrainsPOP
Public GA          As New SimplyGA2

Public MaxY        As Single
Public GravityY    As Single

Public TrainTime   As Long

Public MVV         As Single


Public GEN         As Long
Public pGEN        As Long

Public CNT         As Long

Public ST          As Single

Public Const Bounce As Single = 0.98    '0.5   '0.87
Public Const Friction As Single = 0.5 '0.33    '25
Public Const AirResistence As Single = 0.99975


Public Const Pi    As Single = 3.14159265358979
Public Const Pi2    As Single = 6.28318530717959
Public Const PIh   As Single = 1.5707963267949

Public Const CyclesForScreenFrame As Long = 20 '16 '14
Public Const CyclesForVideoFrame As Long = CyclesForScreenFrame * 2

Public Const BrainClock As Long = CyclesForScreenFrame * 0.25 '4

Public CreFN       As String

Public DoVideo     As Boolean
Public Frame       As Long
Public TextFramesDone As Boolean
Public INTROframesDone As Boolean

Public DoVideoBEST As Long

Public AAA         As New LineGS

Public TrainString As String

Public CurBEST     As Long

Public ROTMAX As Single

Public CRE()      As New clsCreature

Public Function Atan2(X As Single, Y As Single) As Single
    If X Then
        Atan2 = -Pi + Atn(Y / X) - (X > 0) * Pi
    Else
        Atan2 = -PIh - (Y > 0) * Pi
    End If

    ' While Atan2 < 0: Atan2 = Atan2 + Pi2: Wend
    ' While Atan2 > Pi2: Atan2 = Atan2 - Pi2: Wend

    If Atan2 < 0 Then Atan2 = Atan2 + Pi2
    
End Function
Public Function AngleDIFF(ByRef A1 As Single, ByRef A2 As Single) As Single
    AngleDIFF = A2 - A1
    While AngleDIFF < -Pi
        AngleDIFF = AngleDIFF + Pi2
    Wend
    While AngleDIFF > Pi
        AngleDIFF = AngleDIFF - Pi2
    Wend
End Function



Public Sub SaveFrame()

    If CNT Mod (CyclesForVideoFrame) = 0 Then
        BitBlt frmMAIN.PicFrame.hdc, 0, 0, 640, 360, frmMAIN.hdc, frmMAIN.ScaleWidth * 0.5 - 320, frmMAIN.ScaleHeight - 360, vbSrcCopy
        frmMAIN.PicFrame.Refresh

        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98

        Frame = Frame + 1
    End If

End Sub

Public Sub MakeTextFramesOLD(ByVal S As String)
    Dim I          As Long

    BitBlt frmMAIN.PicFrame.hdc, 0, 0, 640, 360, frmMAIN.PicFrame.hdc, 0, 0, vbBlack

    frmMAIN.PicFrame.CurrentY = 10
    frmMAIN.PicFrame.Print "GENERATION " & GEN & vbCrLf & vbCrLf & S
    frmMAIN.PicFrame.Refresh

    SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98

    For I = 1 To 30 * 2
        'SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        FileCopy App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", App.Path & "\Frames\" & Format(Frame + 1, "00000") & ".jpg"
        DoEvents
        DoEvents
        Frame = Frame + 1
    Next
    TextFramesDone = True


End Sub


Public Sub MakeTextFrames(ByVal S As String)
    Dim I          As Long
    Dim S2         As String
    S2 = frmMAIN.Text1 & vbCrLf
    S2 = S2 & "Population Size: " & GA.NOI & vbCrLf & "Creature Neural Network: Inputs " & BR.GetNofInputs(1) & "  Outputs " & BR.GetNofOutputs(1) & ",  Nof Genes: " & BR.GetNofTotalGenes & vbCrLf
    S2 = S2 & TrainString
    S2 = S2 & vbCrLf & "Software by Roberto Mior."
    S2 = S2 & Space$(20)

    S = "GENERATION " & GEN & vbCrLf & vbCrLf & S & "  "

    frmMAIN.PicFrame.FontSize = 14
    For I = 1 To Len(S) Step 2
        BitBlt frmMAIN.PicFrame.hdc, 0, 0, 640, 360, frmMAIN.PicFrame.hdc, 0, 0, vbWhite
        frmMAIN.PicFrame.CurrentY = 8
        frmMAIN.PicFrame.Print Left$(S, I)
        frmMAIN.PicFrame.Refresh
        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        Frame = Frame + 1
    Next
    For I = 1 To Len(S2) Step 20
        BitBlt frmMAIN.PicFrame.hdc, 0, 0, 640, 360, frmMAIN.PicFrame.hdc, 0, 0, vbWhite
        frmMAIN.PicFrame.CurrentY = 8
        frmMAIN.PicFrame.FontSize = 14
        frmMAIN.PicFrame.Print S
        frmMAIN.PicFrame.CurrentY = 122
        frmMAIN.PicFrame.FontSize = 8
        frmMAIN.PicFrame.Print Left$(S2, I)
        frmMAIN.PicFrame.Refresh
        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        Frame = Frame + 1
    Next
    Frame = Frame - 1

    For I = 1 To 15
        'SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        FileCopy App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", App.Path & "\Frames\" & Format(Frame + 1, "00000") & ".jpg"
        DoEvents
        DoEvents
        Frame = Frame + 1
    Next
    TextFramesDone = True


End Sub

Public Sub MakeINTROFrames()
    Dim I          As Long
    Dim J          As Long
    Dim R          As Single

    Dim X          As Single
    Dim Y          As Single
    Dim dx         As Single
    Dim dy         As Single
    Dim D          As Single

    Dim x1         As Single
    Dim y1         As Single
    Dim X2         As Single
    Dim Y2         As Single
    Dim IncX       As Single
    Dim IncY       As Single
    Dim CX         As Single
    Dim CY         As Single
    Dim TL         As Long
    Dim S          As String
    Dim Inob       As Long
    Dim IB         As Long
    Dim MX         As Long
    Dim MY         As Long
    Dim MXs        As Single
    Dim MYs        As Single

    Dim C          As Long

    Const MAG      As Single = 1.5

    CRE(1).DoPhysics2 1
    CX = -CRE(1).CurrCenX * MAG + 640 \ 2
    CY = -CRE(1).CurrCenY * MAG + 360 \ 2



    BitBlt frmMAIN.PicFrame.hdc, 0, 0, 640, 360, frmMAIN.PicFrame.hdc, 0, 0, vbWhite
    AAA.DIB frmMAIN.PicFrame.hdc, frmMAIN.PicFrame.Image.Handle, 640, 360


    '--NAME-------------------------
    frmMAIN.PicFrame.FontSize = 18
    S = "  " & CreFN
    For I = 1 To Len(S)
        frmMAIN.PicFrame.CurrentX = 1
        frmMAIN.PicFrame.CurrentY = 1
        frmMAIN.PicFrame.Print Left$(S, I)
        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        Frame = Frame + 1
    Next
    AAA.DIB frmMAIN.PicFrame.hdc, frmMAIN.PicFrame.Image.Handle, 640, 360
    frmMAIN.PicFrame.FontSize = 8
    '--------------------------------------------


    '--Points------------------------------------------
    For I = 1 To CRE(1).NP
        frmMAIN.Caption = "drawing pt " & I & " of " & CRE(1).NP: DoEvents

        x1 = MAG * CRE(1).getpointX(I) + CX
        y1 = MAG * CRE(1).getpointY(I) + CY

        Do
            AAA.Array2Pic
            MXs = MXs * 0.75 + x1 * 0.25
            MYs = MYs * 0.75 + y1 * 0.25
            MX = MXs \ 1
            MY = MYs \ 1
            GoSub DrawMouse
            SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
            Frame = Frame + 1
            dx = MX - x1
            dy = MY - y1
            D = dx * dx + dy * dy
        Loop While D > 9



        For R = 1 To 3 Step 1
            AAA.CircleDIB x1 \ 1, y1 \ 1, R, R, vbRed
            AAA.Array2Pic
            MX = x1
            MY = y1
            GoSub DrawMouse
            SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
            Frame = Frame + 1
        Next

        ' Points TEXT
        AAA.Array2Pic             'Becasue of Drawmouse
        BitBlt frmMAIN.PicFrame.hdc, 500, 0, 140, 140, frmMAIN.PicFrame.hdc, 0, 0, vbWhite
        frmMAIN.PicFrame.CurrentX = 500
        frmMAIN.PicFrame.CurrentY = 1
        frmMAIN.PicFrame.Print "Pts: " & I
        AAA.DIB frmMAIN.PicFrame.hdc, frmMAIN.PicFrame.Image.Handle, 640, 360
        '---------------------------

        '-Wait-------------------------------------------
        Frame = Frame - 1
        For J = 1 To 3
            'SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
            FileCopy App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", App.Path & "\Frames\" & Format(Frame + 1, "00000") & ".jpg"
            DoEvents
            DoEvents
            Frame = Frame + 1
        Next
    Next

    '-------------------------------
    'Last Point to first Link
    C = -1
    x1 = MAG * CRE(1).getpointX(CRE(1).GetLinkP1(1)) + CX
    y1 = MAG * CRE(1).getpointY(CRE(1).GetLinkP1(1)) + CY

    Do
        AAA.Array2Pic
        MXs = MXs * 0.75 + x1 * 0.25
        MYs = MYs * 0.75 + y1 * 0.25
        MX = MXs \ 1
        MY = MYs \ 1
        GoSub DrawMouse
        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        Frame = Frame + 1
        dx = MX - x1
        dy = MY - y1
        D = dx * dx + dy * dy
    Loop While D > 9
    '------------------------------------







    '-------------------------------


    Inob = 0
    IB = 0

    '-LINKS-------------------------------------------
    For I = 1 To CRE(1).NL
        frmMAIN.Caption = "drawing link " & I & " of " & CRE(1).NL: DoEvents

        IncX = CRE(1).getpointX(CRE(1).GetLinkP2(I)) - CRE(1).getpointX(CRE(1).GetLinkP1(I))
        IncY = CRE(1).getpointY(CRE(1).GetLinkP2(I)) - CRE(1).getpointY(CRE(1).GetLinkP1(I))
        IncX = IncX * 0.2
        IncY = IncY * 0.2

        x1 = MAG * CRE(1).getpointX(CRE(1).GetLinkP1(I)) + CX
        y1 = MAG * CRE(1).getpointY(CRE(1).GetLinkP1(I)) + CY
        X2 = x1
        Y2 = y1

        C = IIf(CRE(1).IsLinkDynamic(I), RGB(100, 100, 255), vbBlack)

        For J = 1 To 5 * MAG
            AAA.LineDIB x1 \ 1, y1 \ 1, X2 \ 1, Y2 \ 1, C
            AAA.Array2Pic
            MX = X2: MY = Y2
            GoSub DrawMouse
            SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
            Frame = Frame + 1
            X2 = X2 + IncX
            Y2 = Y2 + IncY
        Next
            SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
            Frame = Frame + 1



        If CRE(1).IsLinkDynamic(I) Then IB = IB + 1 Else: Inob = Inob + 1


        'Link TEXT
        AAA.Array2Pic             'Becasue of Drawmouse
        BitBlt frmMAIN.PicFrame.hdc, 510, 12, 140, 140, frmMAIN.PicFrame.hdc, 0, 0, vbWhite
        frmMAIN.PicFrame.CurrentX = 500
        frmMAIN.PicFrame.CurrentY = 12
        frmMAIN.PicFrame.Print "NoBrain Links: " & Inob
        frmMAIN.PicFrame.CurrentX = 500
        frmMAIN.PicFrame.CurrentY = 24
        frmMAIN.PicFrame.Print "  Brain Links: " & IB
        AAA.DIB frmMAIN.PicFrame.hdc, frmMAIN.PicFrame.Image.Handle, 640, 360
        '------------------

        'Variable WAIT
        'If I <> CRE(1).NL Then
        '    x1 = MAG * CRE(1).getpointX(CRE(1).GetLinkP1(I + 1)) + CX
        '    y1 = MAG * CRE(1).getpointY(CRE(1).GetLinkP1(I + 1)) + CY
        '    dx = X2 - x1: dy = Y2 - y1
        '    D = 1 + Sqr(dx * dx + dy * dy) * 0.1
        '    For J = 1 To D
        '        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        '        Frame = Frame + 1
        '    Next
        'End If
        '__LINK to LINK______________________________________
        If I <> CRE(1).NL Then
            C = -1
            x1 = CRE(1).getpointX(CRE(1).GetLinkP2(I)) + CX
            y1 = CRE(1).getpointY(CRE(1).GetLinkP2(I)) + CY
            X2 = CRE(1).getpointX(CRE(1).GetLinkP1(I + 1)) + CX
            Y2 = CRE(1).getpointY(CRE(1).GetLinkP1(I + 1)) + CY
            dx = X2 - x1: dy = Y2 - y1
            D = Sqr(dx * dx + dy * dy)
            If D <> 0 Then
                'IncX = dx         'CRE(1).getpointX(CRE(1).GetLinkP1(I + 1)) - CRE(1).getpointX(CRE(1).GetLinkP2(I))
                'IncY = dy         'CRE(1).getpointY(CRE(1).GetLinkP1(I + 1)) - CRE(1).getpointY(CRE(1).GetLinkP2(I))
                'IncX = IncX / 3
                'IncY = IncY / 3
                'x1 = MAG * CRE(1).getpointX(CRE(1).GetLinkP2(I)) + CX
                'y1 = MAG * CRE(1).getpointY(CRE(1).GetLinkP2(I)) + CY
                'X2 = x1
                'Y2 = y1
                'For J = 1 To 3 * MAG
                '    'AAA.LineDIB x1 \ 1, y1 \ 1, X2 \ 1, Y2 \ 1, IIf(CRE(1).IsLinkDynamic(I), RGB(100, 100, 255), vbBlack)
                '    'AAA.Array2Pic
                '    MX = X2: MY = Y2
                '    GoSub DrawMouse
                '    SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
                '    AAA.Array2Pic    'Becasue of Drawmouse
                '    Frame = Frame + 1
                '    X2 = X2 + IncX
                '    Y2 = Y2 + IncY
                'Next
                MXs = MAG * CRE(1).getpointX(CRE(1).GetLinkP2(I)) + CX
                MYs = MAG * CRE(1).getpointY(CRE(1).GetLinkP2(I)) + CY
                x1 = MAG * CRE(1).getpointX(CRE(1).GetLinkP1(I + 1)) + CX
                y1 = MAG * CRE(1).getpointY(CRE(1).GetLinkP1(I + 1)) + CY

                Do
                    AAA.Array2Pic
                    MXs = MXs * 0.8 + x1 * 0.2
                    MYs = MYs * 0.8 + y1 * 0.2
                    MX = MXs \ 1
                    MY = MYs \ 1
                    GoSub DrawMouse
                    SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
                    Frame = Frame + 1
                    dx = MX - x1
                    dy = MY - y1
                    D = dx * dx + dy * dy
                Loop While D > 9
                '------------------------------------
            End If
        End If
        '-------------------------------

    Next

    S = "Population Size: " & GA.NOI & vbCrLf & "Creature Neural Network: Inputs " & BR.GetNofInputs(1) & "  Outputs " & BR.GetNofOutputs(1) & ",  Nof Genes: " & BR.GetNofTotalGenes & vbCrLf
    S = S & vbCrLf & TrainString
    S = S & vbCrLf & "Software by Roberto Mior."
    S = S & Space$(20)



    I = 1
    J = 0
    Do
        If Mid$(S, I, 1) = " " Then IB = I
        If Mid$(S, I, 1) = vbTab Then IB = I
        If Mid$(S, I, 2) = vbCrLf Then J = 0
        J = J + 1
        If J > 27 Then S = Left$(S, IB) & vbCrLf & Right$(S, Len(S) - IB): J = 0
        I = I + 1
    Loop While I < Len(S)


    'Go to other TEXT-------------------------
    C = -1
    x1 = 1
    y1 = 100
    MXs = MAG * CRE(1).getpointX(CRE(1).GetLinkP2(CRE(1).NL)) + CX
    MYs = MAG * CRE(1).getpointY(CRE(1).GetLinkP2(CRE(1).NL)) + CY
    Do
        AAA.Array2Pic
        MXs = MXs * 0.8 + x1 * 0.2
        MYs = MYs * 0.8 + y1 * 0.2
        MX = MXs \ 1
        MY = MYs \ 1
        GoSub DrawMouse
        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        Frame = Frame + 1
        dx = MX - x1
        dy = MY - y1
        D = dx * dx + dy * dy
    Loop While D > 9
    AAA.Array2Pic             'Becasue of Drawmouse
    '------------------------------------


    '-Other Text-------------------------------------------
    For I = 1 To Len(S) Step 3
        frmMAIN.PicFrame.CurrentX = 1
        frmMAIN.PicFrame.CurrentY = 100
        frmMAIN.PicFrame.Print Left$(S, I)
        SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        Frame = Frame + 1
    Next
    '-Wait-------------------------------------------
    Frame = Frame - 1
    For I = 1 To 60
        'SaveJPG frmMAIN.PicFrame.Image, App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", 98
        FileCopy App.Path & "\Frames\" & Format(Frame, "00000") & ".jpg", App.Path & "\Frames\" & Format(Frame + 1, "00000") & ".jpg"
        DoEvents
        DoEvents
        Frame = Frame + 1
    Next
    Exit Sub

DrawMouse:
    If C = -1 Then C = RGB(0, 200, 0)
    AAA.DIB frmMAIN.PicFrame.hdc, frmMAIN.PicFrame.Image.Handle, 640, 360
    AAA.LineGP frmMAIN.PicFrame.hdc, MX, MY, MX + 20, MY + 20, C
    AAA.LineGP frmMAIN.PicFrame.hdc, MX, MY, MX, MY + 28, C
    AAA.LineGP frmMAIN.PicFrame.hdc, MX + 20, MY + 20, MX, MY + 28, C
    'If C <> -1 Then AAA.CircleGP frmMAIN.PicFrame.hdc, MX, MY, 2, 2, C
    Return


End Sub
