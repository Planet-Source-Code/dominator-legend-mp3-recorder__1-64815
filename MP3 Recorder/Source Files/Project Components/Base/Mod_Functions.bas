Attribute VB_Name = "Mod_Functions"
Option Explicit
Public Sub DrawVolume(ByRef PicVolume As PictureBox)
    If Not Frm_Main.DX Is Nothing Then Frm_Main.DX.SoundVolume = DigialVolume
    PicVolume.Line (0, 0)- _
        (DigialVolume * 100, PicVolume.ScaleHeight), PicVolume.ForeColor, BF
    PicVolume.Line (DigialVolume * 100, PicVolume.ScaleHeight)- _
        (PicVolume.ScaleWidth, 0), PicVolume.BackColor, BF
    Frm_Main.Lbl_VolPercent.Caption = Format(DigialVolume * 100, "##0") & "%"
End Sub
Public Sub ShowWAVData16Bit(ByRef LeftPic As PictureBox, ByRef RightPic As PictureBox, ByRef Pic_WavDisplay As PictureBox, ByRef DataBuff() As Byte, Optional ByVal DispType As Integer, Optional ByVal Channels As Integer = 1)
    On Error Resume Next
    Dim HBuffer                 As Long
    Dim K                       As Long
    Dim Col                     As Long
    Dim LAve                    As Single
    Dim RAve                    As Single
    Dim LX                      As Single
    Dim LY                      As Single
    Dim RX                      As Single
    Dim RY                      As Single
    Dim LInt                    As Single
    Dim RInt                    As Single
    Dim vPI                     As Single
    Dim Stp                     As Integer
    Dim Pic1                    As Object
    Set Pic1 = Pic_WavDisplay
    HBuffer = UBound(DataBuff)
    If False Then
        LeftPic.Scale (0, 0)-(HBuffer, 1)
        RightPic.Scale (0, 0)-(HBuffer, 1)
    Else
        LeftPic.Scale (1, HBuffer)-(0, 0)
        RightPic.Scale (1, HBuffer)-(0, 0)
    End If
    Pic1.PSet (0, 0)
    If Channels = 1 Then
        Pic1.Scale (0, -0.5)-(HBuffer, 0.5)
        Stp = 16
        Select Case DispType
        Case 0
            Pic1.Cls
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                Pic1.Line -(K, LInt)
                LAve = LAve + (Abs(LInt) * 2)
            Next K
        Case 1
            Pic1.Cls
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                Pic1.Line (LX, 0)-(K, LInt), RGB(255 * (1 - (LInt + 0.5)), 0, 255 * (LInt + 0.5)), BF
                LX = K
                LAve = LAve + (Abs(LInt) * 2)
            Next K
        Case 2
            Pic1.Cls
            Col = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                Pic1.Line -(K, LInt), RGB(255 * (1 - (LInt + 0.5)), 0, 255 * (LInt + 0.5))
                LAve = LAve + (Abs(LInt) * 2)
            Next K
        Case 3
            Pic1.Cls
            Pic1.Scale (-1, -1)-(1, 1)
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                vPI = (PI / 8) * (K / (HBuffer / 16))
                If K = 0 Then
                    Pic1.PSet (Abs(LInt + 0.5) * Cos(vPI), Abs(LInt + 0.5) * Sin(vPI))
                Else
                    Pic1.Line -(Abs(LInt + 0.5) * Cos(vPI), Abs(LInt + 0.5) * Sin(vPI))
                End If
                LAve = LAve + (Abs(LInt) * 2)
            Next K
        End Select
        RAve = LAve
    Else
        Pic1.Scale (0, -1)-(HBuffer, 1)
        Stp = 20
        Select Case DispType
        Case 0
            Pic1.Cls
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                RInt = BytesToInt(DataBuff(K + 2), DataBuff(K + 3)) / 65536#
                Pic1.Line (LX, LY)-(K, LInt - 0.5)
                Pic1.Line (RX, RY)-(K, RInt + 0.5)
                LX = K
                LY = LInt - 0.5
                RX = K
                RY = RInt + 0.5
                LAve = LAve + (Abs(LInt) * 2)
                RAve = RAve + (Abs(RInt) * 2)
            Next K
            Pic1.Line (LX, LY)-(HBuffer, LY)
            Pic1.Line (RX, RY)-(HBuffer, RY)
        Case 1
            Pic1.Cls
            Stp = 48
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                RInt = BytesToInt(DataBuff(K + 2), DataBuff(K + 3)) / 65536#
                Pic1.Line (LX, -0.5)-(K, LInt - 0.5), RGB(255 * (1 - (LInt + 0.5)), 0, 255 * (LInt + 0.5)), BF
                Pic1.Line (LX, 0.5)-(K, RInt + 0.5), RGB(255 * (1 - (RInt + 0.5)), 0, 255 * (RInt + 0.5)), BF
                LX = K
                LAve = LAve + (Abs(LInt) * 2)
                RAve = RAve + (Abs(RInt) * 2)
            Next K
        Case 2
            Pic1.Cls
            Col = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd)
            Stp = 48
            For K = 0 To HBuffer - 2 Step Stp
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536#
                RInt = BytesToInt(DataBuff(K + 2), DataBuff(K + 3)) / 65536#
                Pic1.Line (LX, LY)-(K, LInt - 0.5), RGB(255 * (1 - (LInt + 0.5)), 0, 255 * (LInt + 0.5))
                Pic1.Line (RX, RY)-(K, RInt + 0.5), RGB(255 * (1 - (RInt + 0.5)), 0, 255 * (RInt + 0.5))
                LX = K
                LY = LInt - 0.5
                RX = K
                RY = RInt + 0.5
                LAve = LAve + (Abs(LInt) * 2)
                RAve = RAve + (Abs(RInt) * 2)
            Next K
        Case 3
            Dim LPX As Single, LPY As Single, RPX As Single, RPY As Single
            Pic1.Cls
            Pic1.Scale (-2, -1)-(2, 1)
            For K = 0 To HBuffer - 2 Step Stp * 2
                LInt = BytesToInt(DataBuff(K), DataBuff(K + 1)) / 65536# + 0.5
                RInt = BytesToInt(DataBuff(K + 2), DataBuff(K + 3)) / 65536# + 0.5
                vPI = (PI / 16) * (K / (HBuffer / 32))
                If K = 0 Then
                    LX = -1 + Abs(LInt) * Cos(vPI)
                    LY = Abs(LInt) * Sin(vPI)
                    RX = 1 + Abs(RInt) * Cos(vPI)
                    RY = Abs(RInt) * Sin(vPI)
                Else
                    LPX = -1 + Abs(LInt) * Cos(vPI)
                    LPY = Abs(LInt) * Sin(vPI)
                    RPX = 1 + Abs(RInt) * Cos(vPI)
                    RPY = Abs(RInt) * Sin(vPI)
                    Pic1.Line (LX, LY)-(LPX, LPY)
                    Pic1.Line (RX, RY)-(RPX, RPY)
                    LX = LPX
                    LY = LPY
                    RX = RPX
                    RY = RPY
                End If
                LAve = LAve + (Abs(LInt - 0.5) * 4)
                RAve = RAve + (Abs(RInt - 0.5) * 4)
            Next K
        End Select
    End If
    If False Then
        LeftPic.Line (0, 0)-(LAve * Stp * 2, 1), RGB(10, 105, 10), BF
        LeftPic.Line (LAve * Stp * 2, 0)-(HBuffer, 1), LeftPic.BackColor, BF
        RightPic.Line (0, 0)-(RAve * Stp * 2, 1), RGB(10, 105, 10), BF
        RightPic.Line (RAve * Stp * 2, 0)-(HBuffer, 1), LeftPic.BackColor, BF
    Else
        LeftPic.Line (0, 0)-(1, LAve * Stp * 2), RGB(10, 105, 10), BF
        LeftPic.Line (0, LAve * Stp * 2)-(1, HBuffer), LeftPic.BackColor, BF
        RightPic.Line (0, 0)-(1, RAve * Stp * 2), RGB(10, 105, 10), BF
        RightPic.Line (0, RAve * Stp * 2)-(1, HBuffer), LeftPic.BackColor, BF
    End If
End Sub
Private Function BytesToInt(ByVal FirstByte As Byte, ByVal SecondByte As Byte) As Long
    Dim B2      As Bytes2
    Dim i       As IntType
    B2.Byte1 = FirstByte
    B2.Byte2 = SecondByte
    LSet i = B2
    BytesToInt = i.IntVal
End Function
