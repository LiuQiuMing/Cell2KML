Attribute VB_Name = "Computation"
Const pi = 3.1415926535898
Private a, b, c, alpha, e, e2, w, V As Double
Private B1, L1, B2, L2 As Double
Private s As Double
Private A1, A2 As Double
Private Sub getellipseparameter()
    a = 6378245
    b = 6356752.3142
    c = a ^ 2 / b
    alpha = (a - b) / a
    e = Sqr(a ^ 2 - b ^ 2) / a
    e2 = Sqr(a ^ 2 - b ^ 2) / b
End Sub
Private Function computerw()
    w = Sqr(1 - e ^ 2 * (Sin(B1) ^ 2))
    V = w * (a / b)
End Function
Function Computation(STARTLAT, STARTLONG, ANGLE1, DISTANCE As Double) As String
    B1 = STARTLAT
    L1 = STARTLONG
    A1 = ANGLE1
    s = DISTANCE
    Call getellipseparameter
   If B1 = 0 Then
        If A1 = 90 Then
            A2 = 270
            B2 = 0
            L2 = L1 + s / a * 180 / pi
        End If
        If A1 = 270 Then
            A2 = 90
            B2 = 0
            L2 = L1 - s / a * 180 / pi
        End If
        Exit Function
    End If
    B1 = rad(B1)
    L1 = rad(L1)
    A1 = rad(A1)
    Call computerw
    Dim W1 As Double
    E1 = e
    W1 = w
    sinu1 = Sin(B1) * Sqr(1 - E1 * E1) / W1
    cosu1 = Cos(B1) / W1
    sinA0 = cosu1 * Sin(A1)
    cotq1 = cosu1 * Cos(A1)
    sin2q1 = 2 * cotq1 / (cotq1 ^ 2 + 1)
    cos2q1 = (cotq1 ^ 2 - 1) / (cotq1 ^ 2 + 1)
    cos2A0 = 1 - sinA0 ^ 2
    e2 = Sqr(a ^ 2 - b ^ 2) / b
    k2 = e2 * e2 * cos2A0
    Dim aa, BB, cc, EE22, AAlpha, BBeta As Double
    aa = b * (1 + k2 / 4 - 3 * k2 * k2 / 64 + 5 * k2 * k2 * k2 / 256)
    BB = b * (k2 / 8 - k2 * k2 / 32 + 15 * k2 * k2 * k2 / 1024)
    cc = b * (k2 * k2 / 128 - 3 * k2 * k2 * k2 / 512)
    e2 = E1 * E1
    AAlpha = (e2 / 2 + e2 * e2 / 8 + e2 * e2 * e2 / 16) - (e2 * e2 / 16 + e2 * e2 * e2 / 16) * cos2A0 + (3 * e2 * e2 * e2 / 128) * cos2A0 * cos2A0
    BBeta = (e2 * e2 / 32 + e2 * e2 * e2 / 32) * cos2A0 - (e2 * e2 * e2 / 64) * cos2A0 * cos2A0
    q0 = (s - (BB + cc * cos2q1) * sin2q1) / aa
    sin2q1q0 = sin2q1 * Cos(2 * q0) + cos2q1 * Sin(2 * q0)
    cos2q1q0 = cos2q1 * Cos(2 * q0) - sin2q1 * Sin(2 * q0)
    q = q0 + (BB + 5 * cc * cos2q1q0) * sin2q1q0 / aa
    theta = (AAlpha * q + BBeta * (sin2q1q0 - sin2q1)) * sinA0
    sinu2 = sinu1 * Cos(q) + cosu1 * Cos(A1) * Sin(q)
    B2 = Atn(sinu2 / (Sqr(1 - E1 * E1) * Sqr(1 - sinu2 * sinu2))) * 180 / pi
    lamuda = Atn(Sin(A1) * Sin(q) / (cosu1 * Cos(q) - sinu1 * Sin(q) * Cos(A1))) * 180 / pi
                 If (Sin(A1) > 0) Then
                    If (Sin(A1) * Sin(q) / (cosu1 * Cos(q) - sinu1 * Sin(q) * Cos(A1)) > 0) Then
                        lamuda = Abs(lamuda)
                    Else
                        lamuda = 180 - Abs(lamuda)
                    End If
                Else
                    If (Sin(A1) * Sin(q) / (cosu1 * Cos(q) - sinu1 * Sin(q) * Cos(A1)) > 0) Then
                        lamuda = Abs(lamuda) - 180
                    Else
                        lamuda = -Abs(lamuda)
                    End If
                End If
                L2 = L1 * 180 / pi + lamuda - theta * 180 / pi
                
'                A2 = Atn(cosu1 * Sin(A1) / (cosu1 * Cos(q) * Cos(A1) - sinu1 * Sin(q))) * 180 / pi
'                If (Sin(A1) > 0) Then
'                    If (cosu1 * Sin(A1) / (cosu1 * Cos(q) * Cos(A1) - sinu1 * Sin(q)) > 0) Then
'                        A2 = 180 + Abs(A2)
'                    Else
'                        A2 = 360 - Abs(A2)
'                    End If
'                Else
'                    If (cosu1 * Sin(A1) / (cosu1 * Cos(q) * Cos(A1) - sinu1 * Sin(q)) > 0) Then
'                        A2 = Abs(A2)
'                    Else
'                        A2 = 180 - Abs(A2)
'                    End If
'                End If
     Computation = format(L2, "0.00000000") & "," & format(B2, "0.00000000")
End Function
Private Function rad(ByVal angle_d As Double) As Double
    rad = angle_d * pi / 180
End Function
