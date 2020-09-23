Attribute VB_Name = "modCreature"
Option Explicit

Public Type tPoint
    X              As Single
    Y              As Single
    vX             As Single
    vY             As Single
    IsOnGround     As Single
    IvX            As Single
    IvY            As Single
End Type


Public Type tLink
    P1             As Long
    P2             As Long
    MainLen        As Single
    AddedBrainLenToReach  As Single
    CurrAddedBrainLen As Single
    
    InvMainLen     As Single
    Stren          As Single
    TENSION        As Single
    Dynamic        As Boolean

End Type



Public GradStatic1 As New clsGradient3
Public GradStatic2 As New clsGradient3
Public GradDynamic1 As New clsGradient3
Public GradDynamic2 As New clsGradient3
Public GradGround1 As New clsGradient3
Public GradGround2 As New clsGradient3


Public InputValues() As Single
Public OutputValues() As Single




Public Function PointDist(P1 As tPoint, P2 As tPoint) As Single
    Dim dx         As Single
    Dim dy         As Single
    dx = P2.X - P1.X
    dy = P2.Y - P1.Y
    PointDist = Sqr(dx * dx + dy * dy)
End Function




Public Sub InitGrad()
    GradStatic1.Init 0, 255, 0, 220, 220, 220, 255, 0, 0
    GradDynamic1.Init 0, 200, 0, 140, 140, 255, 200, 0, 0
    GradGround1.Init 200, 200, 255, 255, 255, 255, 255, 0, 0

    GradStatic2.Init 0, 120, 0, 0, 0, 0, 120, 0, 0
    GradDynamic2.Init 50, 180, 50, 130, 130, 180, 180, 50, 50
    GradGround2.Init 80, 80, 80, 255, 128, 0, 255, 0, 0
End Sub
