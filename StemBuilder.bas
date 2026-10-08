' ============================================================================================
' STEM BUILDER   (41168 Biomedical Engineering Design, Seminar 2, femoral stem stream)
'
' Builds the femoral stem models in the modelling plan, ready for an ISO 7206-4 static FEA.
' Every saved part has
'   * the course stem outline (taken from "Femoral stem - solid.SLDPRT") at the plate thickness
'     and neck diameter you choose, with the course elliptical neck collar,
'   * the course 30 mm femoral head as a separate body, bored to fit the neck,
'   * the whole model tilted into the ISO 7206-4 test position, so the load axis is the global
'     Y axis (alpha = 10 deg, beta = 0 deg) and displacement along the load is simply UY,
'   * split lines on the stem at the potting level (80 mm below the head centre) and 3 mm above,
'   * a flat face on top of the head for the 2300 N load, normal to the face,
'   * for lattice models, a graded or uniform simple cubic strut lattice in a window above the
'     potting level, with the uniform lattice sized to hold the same metal as the graded one.
'
' HOW TO RUN
'   1. In SolidWorks choose Tools > Macro > New and save it as StemBuilder.swp.
'   2. In the VBA window that opens, select all the code in the module and delete it.
'   3. Paste this whole file in. Save, close the VBA window.
'   4. Tools > Macro > Run > StemBuilder.swp and answer the questions.
'   The parts, a log and a results table go into a folder called "Stem models" next to the
'   macro file. An existing part is never overwritten. A rebuilt one gets " v2" added to its name.
'   A solid model takes well under a minute. A lattice model takes a few minutes, and
'   SolidWorks shows nothing on screen until each part is finished.
'
' All lengths in this file are millimetres. The SolidWorks API works in metres, so the wrapper
' functions near the bottom convert every value on the way in.
' ============================================================================================
Option Explicit

' ---------------------------------- settings you may change ----------------------------------
Private Const DEFAULT_THICKNESS_LIST As String = "3,4,5,6,7"   ' feature 1 sweep (course stem is 5 mm)
Private Const DEFAULT_NECK_LIST As String = "10,11,12,13,14"   ' feature 2 sweep (course neck is 12 mm)
Private Const COURSE_THICKNESS As Double = 5
Private Const COURSE_NECK As Double = 12
Private Const CELL_MM As Double = 2                  ' lattice cell size
Private Const WALL_MM As Double = 3                  ' solid frame kept around the lattice window
Private Const ROD_INTO_WALL_MM As Double = 1         ' rods run this far into the frame so they fuse
Private Const WINDOW_GAP_MM As Double = 2            ' solid band between the reading line and the window
Private Const ZONE_POROSITIES As String = "0.45,0.60,0.50,0.40,0.30"   ' zone 1 (top) to zone 5 (bottom)
Private Const MIN_STRUT_MM As Double = 0.75          ' printing limit from Seminar 1
Private Const HEAD_RADIUS_MM As Double = 15          ' course head is 30 mm in diameter
Private Const HEAD_BOTTOM_MM As Double = 10          ' flat bottom of the head, measured from its centre
Private Const HEAD_FLAT_MM As Double = 1             ' cap removed from the top of the head for the load face
Private Const COLLAR_HALF_MM As Double = 14.5        ' collar half length along the neck base line
Private Const COLLAR_LEN_MM As Double = 2            ' collar length along the neck axis
Private Const COLLAR_EXTRA_MM As Double = 0.5        ' collar is this much wider than the neck, front to back
Private Const NECK_START_MM As Double = 1            ' neck starts 1 mm inside the collar so they fuse cleanly
Private Const NECK_END_MM As Double = 27             ' neck ends at the head centre, 27 mm from its base
Private Const ALPHA_DEG As Double = 10               ' ISO 7206-4 frontal angle
Private Const POTTING_D_MM As Double = 80            ' ISO 7206-4 potting depth for 120 < CT <= 250 mm
Private Const READING_OFFSET_MM As Double = 3        ' read peak stress above this line

' ---------------------------------- fixed values -----------------------------------------------
' Stem frame: stem roughly along +Y, medial (neck) side towards +X, front face towards +Z.
' The course coordinates are shifted by SHIFT_X and SHIFT_Y so every extrusion start is positive.
Private Const SHIFT_X As Double = 15
Private Const SHIFT_Y As Double = 10
Private Const PI_ As Double = 3.14159265358979
Private Const ERR_STEM As Long = 1000
Private Const GRID_NI As Long = 50
Private Const GRID_NJ As Long = 80

' outline primitives: kind 1 = line, 2 = arc
Private nPrim As Long
Private pKind() As Long, pX0() As Double, pY0() As Double, pX1() As Double, pY1() As Double
Private pCX() As Double, pCY() As Double, pR() As Double, pDir() As Long

' key points of the outline
Private tipCX As Double, tipCY As Double, tipR As Double
Private latBotX As Double, latBotY As Double, latTopX As Double, latTopY As Double
Private shDX As Double, shDY As Double
Private filCX As Double, filCY As Double, filR As Double
Private filShX As Double, filShY As Double, filTopX As Double, filTopY As Double
Private topY As Double, topRX As Double, topRY As Double
Private medTopX As Double, medTopY As Double, medCX As Double, medCY As Double, medR As Double
Private baseX As Double, baseY As Double, uX As Double, uY As Double, vX As Double, vY As Double
Private headCX As Double, headCY As Double

' ISO 7206-4 geometry
Private isoCT As Double, isoKX As Double, isoKY As Double, isoLX As Double, isoLY As Double
Private isoLean As Double, isoTheta As Double, sinT As Double, cosT As Double
Private isoCX As Double, isoCY As Double, yPot As Double, yRead As Double
Private splLpX As Double, splLpY As Double, splLrX As Double, splLrY As Double
Private splMpX As Double, splMpY As Double, splMrX As Double, splMrY As Double

' lattice window
Private gridX0 As Double, gridY0 As Double
Private inWin() As Boolean
Private nCells As Long
Private cellI() As Long, cellJ() As Long, cellX() As Double, cellY() As Double, cellZone() As Long
Private nZones As Long, zonePor() As Double, zoneDia() As Double, zoneRows() As Long, zoneCells() As Long
Private rowZone() As Long
Private loopN As Long, loopX() As Double, loopY() As Double, loopEdgesOK As Boolean

' rods of the current lattice model
Private nXG As Long, xgX0() As Double, xgX1() As Double, xgD() As Double, xgCount() As Long, xgRows() As String
Private nYG As Long, ygY0() As Double, ygY1() As Double, ygD() As Double, ygCount() As Long, ygCols() As String
Private nzLev As Long, zLev() As Double
Private currentKind As Long, currentDUni As Double

' run state
Private swApp As Object, swModel As Object, swMath As Object
Private planeFront As Object, planeTop As Object, planeRight As Object
Private curSketch As Object, curXform As Object, curArcFlip As Boolean
Private lastFeat As Object, lastCenterline As Object
Private offCombo(1 To 3) As Long
Private TestMode As Boolean, testFile As Integer
Private OutDir As String, LogPath As String, CsvPath As String
Private summary As String, stepName As String, currentFile As String
Private solidStemVolume As Double, solidStemT As Double, solidStemN As Double

' =============================================================================================
'                                         ENTRY POINT
' =============================================================================================
Sub main()
    Dim choice As String, tList As String, nList As String, kind As String
    Dim tStar As Double, nStar As Double, ok As Boolean, t0 As Double

    TestMode = False
    Set swApp = Application.SldWorks
    Set swMath = swApp.GetMathUtility
    SetupOutput
    t0 = Timer
    InitGeometry
    WriteLog "Geometry ready in " & Format(Timer - t0, "0.0") & " s. ISO tilt " & Format(isoTheta, "0.000") & _
        " deg, CT " & Format(isoCT, "0.00") & " mm, lattice window " & nCells & " cells."

    choice = InputBox("Which models do you want to build?" & vbCrLf & vbCrLf & _
        "1   Thickness sweep, F1-1 to F1-5 (solid)" & vbCrLf & _
        "2   Neck sweep, F2-1 to F2-5 (solid)" & vbCrLf & _
        "3   Material-reduction set, MR-S, MR-FG and MR-U" & vbCrLf & _
        "4   One model with your own values" & vbCrLf & vbCrLf & _
        "Models are saved in " & OutDir, "Stem builder", "1")
    If Len(Trim(choice)) = 0 Then Exit Sub

    summary = ""
    Select Case Trim(choice)
    Case "1"
        tList = InputBox("Plate thicknesses to build, in mm, separated by commas." & vbCrLf & _
            "The course stem is 5 mm thick.", "Thickness sweep", DEFAULT_THICKNESS_LIST)
        If Len(Trim(tList)) = 0 Then Exit Sub
        nStar = AskNumber("Neck diameter to use for the thickness sweep, in mm.", "Thickness sweep", COURSE_NECK)
        If nStar <= 0 Then Exit Sub
        ok = BuildSweep("F1", tList, True, nStar)
    Case "2"
        tStar = AskNumber("Thickness you chose from the thickness sweep, in mm.", "Neck sweep", COURSE_THICKNESS)
        If tStar <= 0 Then Exit Sub
        nList = InputBox("Neck diameters to build, in mm, separated by commas.", "Neck sweep", DEFAULT_NECK_LIST)
        If Len(Trim(nList)) = 0 Then Exit Sub
        ok = BuildSweep("F2", nList, False, tStar)
    Case "3"
        tStar = AskNumber("Thickness you chose from the thickness sweep, in mm.", "Material-reduction set", COURSE_THICKNESS)
        If tStar <= 0 Then Exit Sub
        nStar = AskNumber("Neck diameter you chose from the neck sweep, in mm.", "Material-reduction set", COURSE_NECK)
        If nStar <= 0 Then Exit Sub
        ok = BuildModel("MR-S", tStar, nStar, 0, False)
        If ok Then ok = BuildModel("MR-FG", tStar, nStar, 1, False)
        If ok Then ok = BuildModel("MR-U", tStar, nStar, 2, True)
    Case "4"
        tStar = AskNumber("Plate thickness, in mm.", "Custom model", COURSE_THICKNESS)
        If tStar <= 0 Then Exit Sub
        nStar = AskNumber("Neck diameter, in mm.", "Custom model", COURSE_NECK)
        If nStar <= 0 Then Exit Sub
        kind = UCase(Trim(InputBox("Type S for solid, G for graded lattice or U for uniform lattice.", "Custom model", "S")))
        If kind = "G" Then
            ok = BuildModel("Custom", tStar, nStar, 1, True)
        ElseIf kind = "U" Then
            ok = BuildModel("Custom", tStar, nStar, 2, True)
        Else
            ok = BuildModel("Custom", tStar, nStar, 0, True)
        End If
    Case Else
        MsgBox "Please type 1, 2, 3 or 4.", vbExclamation, "Stem builder"
        Exit Sub
    End Select

    If ok Then
        MsgBox "Finished." & vbCrLf & vbCrLf & summary & vbCrLf & _
            "The parts, the log and the results table are in" & vbCrLf & OutDir, vbInformation, "Stem builder"
    Else
        MsgBox "Stopped because a step failed." & vbCrLf & vbCrLf & summary & vbCrLf & _
            "The log file says which step and why." & vbCrLf & LogPath, vbExclamation, "Stem builder"
    End If
End Sub

' Builds one model per value in the list. sweepThickness says which parameter the list holds.
Private Function BuildSweep(ByVal prefix As String, ByVal valueList As String, ByVal sweepThickness As Boolean, _
                            ByVal fixedVal As Double) As Boolean
    Dim parts() As String, k As Long, v As Double, nDone As Long, last As Long
    parts = Split(valueList, ",")
    last = UBound(parts)
    BuildSweep = True
    nDone = 0
    For k = 0 To last
        v = Val(Trim(parts(k)))
        If v > 0 Then
            nDone = nDone + 1
            If sweepThickness Then
                BuildSweep = BuildModel(prefix & "-" & nDone, v, fixedVal, 0, (k = last))
            Else
                BuildSweep = BuildModel(prefix & "-" & nDone, fixedVal, v, 0, (k = last))
            End If
            If Not BuildSweep Then Exit Function
        End If
    Next k
End Function

Private Function AskNumber(ByVal prompt As String, ByVal title As String, ByVal dflt As Double) As Double
    Dim s As String
    s = InputBox(prompt, title, NumStr(dflt))
    If Len(Trim(s)) = 0 Then AskNumber = 0 Else AskNumber = Val(Trim(s))
End Function

' =============================================================================================
'                                       GEOMETRY SET-UP
' =============================================================================================
Private Sub InitGeometry()
    Dim tc As Double, cornerX As Double, cornerY As Double, phi As Double, tl As Double
    ' course outline, shifted
    tipCX = 0 + SHIFT_X: tipCY = 0 + SHIFT_Y: tipR = 3
    latBotX = -3 + SHIFT_X: latBotY = 0 + SHIFT_Y
    latTopX = -8.2336 + SHIFT_X: latTopY = 99.863 + SHIFT_Y
    shDX = 0.34202014: shDY = 0.93969262
    filCX = 7.2381 + SHIFT_X: filCY = 127.7522 + SHIFT_Y: filR = 5
    topY = 132.7522 + SHIFT_Y
    topRX = 8.7371 + SHIFT_X: topRY = 132.7522 + SHIFT_Y
    medTopX = 29.4908 + SHIFT_X: medTopY = 111.9985 + SHIFT_Y
    medCX = 253 + SHIFT_X: medCY = 0 + SHIFT_Y: medR = 250

    ' tangent points of the R5 fillet between the 20 degree shoulder line and the top line
    tc = (topY - latTopY) / shDY
    cornerX = latTopX + tc * shDX: cornerY = latTopY + tc * shDY
    phi = ArcCos(-shDX)
    tl = filR / Tan(phi / 2)
    filShX = cornerX - tl * shDX: filShY = cornerY - tl * shDY
    filTopX = cornerX + tl: filTopY = cornerY

    ' neck axis (u, from the stem towards the head) and its in-plane normal v
    uX = 1 / Sqr(2): uY = 1 / Sqr(2)
    vX = -uX: vY = uY
    baseX = (topRX + medTopX) / 2: baseY = (topRY + medTopY) / 2
    headCX = baseX + NECK_END_MM * uX: headCY = baseY + NECK_END_MM * uY

    ComputeIso
    BuildOutline
    ParseZones
    ComputeWindow
End Sub

Private Sub ComputeIso()
    Dim tX As Double, tY As Double, yK As Double, yL As Double
    tX = tipCX: tY = tipCY - tipR
    isoCT = Sqr((headCX - tX) ^ 2 + (headCY - tY) ^ 2)
    yK = tY + 0.1 * isoCT: yL = tY + 0.4 * isoCT
    isoKX = (XLeft(yK) + XRight(yK)) / 2: isoKY = yK
    isoLX = (XLeft(yL) + XRight(yL)) / 2: isoLY = yL
    isoLean = Atn((isoLX - isoKX) / (isoLY - isoKY)) * 180 / PI_
    isoTheta = ALPHA_DEG + isoLean
    sinT = Sin(isoTheta * PI_ / 180): cosT = Cos(isoTheta * PI_ / 180)
    isoCX = headCX * cosT - headCY * sinT
    isoCY = headCX * sinT + headCY * cosT
    yPot = isoCY - POTTING_D_MM
    yRead = yPot + READING_OFFSET_MM
    LatCross yRead, splLrX, splLrY
    LatCross yPot, splLpX, splLpY
    MedCross yPot, splMpX, splMpY
    MedCross yRead, splMrX, splMrY
End Sub

' Height of a stem-frame point once the model is tilted into the test frame.
Private Function YTest(ByVal x As Double, ByVal y As Double) As Double
    YTest = x * sinT + y * cosT
End Function

Private Sub LatCross(ByVal yt As Double, ByRef px As Double, ByRef py As Double)
    Dim a As Double, b As Double, t As Double
    a = YTest(latBotX, latBotY): b = YTest(latTopX, latTopY)
    t = (yt - a) / (b - a)
    px = latBotX + t * (latTopX - latBotX)
    py = latBotY + t * (latTopY - latBotY)
End Sub

Private Sub MedCross(ByVal yt As Double, ByRef px As Double, ByRef py As Double)
    Dim lo As Double, hi As Double, m As Double, k As Long, flo As Double, fm As Double
    lo = tipCY: hi = medTopY
    For k = 1 To 200
        m = (lo + hi) / 2
        flo = YTest(XMedArc(lo), lo) - yt
        fm = YTest(XMedArc(m), m) - yt
        If flo * fm <= 0 Then hi = m Else lo = m
    Next k
    py = (lo + hi) / 2: px = XMedArc(py)
End Sub

Private Function XLatLine(ByVal y As Double) As Double
    XLatLine = latBotX + (y - latBotY) * (latTopX - latBotX) / (latTopY - latBotY)
End Function

Private Function XMedArc(ByVal y As Double) As Double
    XMedArc = medCX - Sqr(medR * medR - (y - medCY) ^ 2)
End Function

Private Function XLeft(ByVal y As Double) As Double
    If y <= latTopY Then
        XLeft = XLatLine(y)
    ElseIf y <= filShY Then
        XLeft = latTopX + (y - latTopY) * shDX / shDY
    Else
        XLeft = filCX - Sqr(Max0(filR * filR - (y - filCY) ^ 2))
    End If
End Function

Private Function XRight(ByVal y As Double) As Double
    If y <= medTopY Then
        XRight = XMedArc(y)
    Else
        XRight = medTopX - (y - medTopY)
    End If
End Function

Private Function InsidePlate(ByVal x As Double, ByVal y As Double) As Boolean
    If y < tipCY Then
        InsidePlate = ((x - tipCX) ^ 2 + (y - tipCY) ^ 2 < tipR * tipR)
    ElseIf y > topY Then
        InsidePlate = False
    Else
        InsidePlate = (x > XLeft(y)) And (x < XRight(y))
    End If
End Function

Private Sub AddPrimLine(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double)
    nPrim = nPrim + 1
    pKind(nPrim) = 1: pX0(nPrim) = x0: pY0(nPrim) = y0: pX1(nPrim) = x1: pY1(nPrim) = y1
End Sub

Private Sub AddPrimArc(ByVal cx As Double, ByVal cy As Double, ByVal x0 As Double, ByVal y0 As Double, _
                       ByVal x1 As Double, ByVal y1 As Double, ByVal dirv As Long)
    nPrim = nPrim + 1
    pKind(nPrim) = 2: pCX(nPrim) = cx: pCY(nPrim) = cy
    pX0(nPrim) = x0: pY0(nPrim) = y0: pX1(nPrim) = x1: pY1(nPrim) = y1
    pR(nPrim) = Sqr((x0 - cx) ^ 2 + (y0 - cy) ^ 2): pDir(nPrim) = dirv
End Sub

' Counter-clockwise outline with extra vertices where the potting and reading lines cross it, so
' the extruded side faces are already split at those heights.
Private Sub BuildOutline()
    ReDim pKind(1 To 12): ReDim pX0(1 To 12): ReDim pY0(1 To 12): ReDim pX1(1 To 12): ReDim pY1(1 To 12)
    ReDim pCX(1 To 12): ReDim pCY(1 To 12): ReDim pR(1 To 12): ReDim pDir(1 To 12)
    nPrim = 0
    AddPrimLine latTopX, latTopY, splLrX, splLrY
    AddPrimLine splLrX, splLrY, splLpX, splLpY
    AddPrimLine splLpX, splLpY, latBotX, latBotY
    AddPrimArc tipCX, tipCY, latBotX, latBotY, tipCX + tipR, tipCY, 1
    AddPrimArc medCX, medCY, tipCX + tipR, tipCY, splMpX, splMpY, -1
    AddPrimArc medCX, medCY, splMpX, splMpY, splMrX, splMrY, -1
    AddPrimArc medCX, medCY, splMrX, splMrY, medTopX, medTopY, -1
    AddPrimLine medTopX, medTopY, topRX, topRY
    AddPrimLine topRX, topRY, filTopX, filTopY
    AddPrimArc filCX, filCY, filTopX, filTopY, filShX, filShY, 1
    AddPrimLine filShX, filShY, latTopX, latTopY
End Sub

Private Function DistToOutline(ByVal x As Double, ByVal y As Double) As Double
    Dim k As Long, best As Double, d As Double
    best = 1E+30
    For k = 1 To nPrim
        If pKind(k) = 1 Then
            d = DistPointSeg(x, y, pX0(k), pY0(k), pX1(k), pY1(k))
        Else
            d = DistPointArc(x, y, k)
        End If
        If d < best Then best = d
    Next k
    DistToOutline = best
End Function

Private Function DistPointSeg(ByVal x As Double, ByVal y As Double, ByVal x0 As Double, ByVal y0 As Double, _
                              ByVal x1 As Double, ByVal y1 As Double) As Double
    Dim dx As Double, dy As Double, t As Double
    dx = x1 - x0: dy = y1 - y0
    t = ((x - x0) * dx + (y - y0) * dy) / (dx * dx + dy * dy)
    If t < 0 Then t = 0
    If t > 1 Then t = 1
    DistPointSeg = Sqr((x - x0 - t * dx) ^ 2 + (y - y0 - t * dy) ^ 2)
End Function

Private Function DistPointArc(ByVal x As Double, ByVal y As Double, ByVal k As Long) As Double
    Dim a0 As Double, a1 As Double, a As Double, sweep As Double, rel As Double, d0 As Double, d1 As Double
    a0 = Atan2(pY0(k) - pCY(k), pX0(k) - pCX(k))
    a1 = Atan2(pY1(k) - pCY(k), pX1(k) - pCX(k))
    a = Atan2(y - pCY(k), x - pCX(k))
    If pDir(k) > 0 Then
        sweep = WrapAngle(a1 - a0): rel = WrapAngle(a - a0)
    Else
        sweep = WrapAngle(a0 - a1): rel = WrapAngle(a0 - a)
    End If
    If rel <= sweep Then
        DistPointArc = Abs(Sqr((x - pCX(k)) ^ 2 + (y - pCY(k)) ^ 2) - pR(k))
    Else
        d0 = Sqr((x - pX0(k)) ^ 2 + (y - pY0(k)) ^ 2)
        d1 = Sqr((x - pX1(k)) ^ 2 + (y - pY1(k)) ^ 2)
        If d0 < d1 Then DistPointArc = d0 Else DistPointArc = d1
    End If
End Function

' --------------------------------------- lattice window --------------------------------------
Private Sub ParseZones()
    Dim parts() As String, k As Long
    parts = Split(ZONE_POROSITIES, ",")
    nZones = UBound(parts) + 1
    ReDim zonePor(1 To nZones): ReDim zoneDia(1 To nZones): ReDim zoneRows(1 To nZones): ReDim zoneCells(1 To nZones)
    For k = 1 To nZones
        zonePor(k) = Val(Trim(parts(k - 1)))
        zoneDia(k) = DiameterForPorosity(zonePor(k), CELL_MM)
    Next k
End Sub

' A cell joins the window when its corners, edge midpoints and centre are all inside the plate,
' at least WALL_MM from the outline, and its corners sit WINDOW_GAP_MM or more above the
' reading line once the model is tilted.
Private Function CellOK(ByVal xc As Double, ByVal yc As Double) As Boolean
    Dim ix As Long, iy As Long, qx As Double, qy As Double, ylimit As Double
    ylimit = yRead + WINDOW_GAP_MM
    CellOK = False
    For ix = -1 To 1
        For iy = -1 To 1
            qx = xc + ix * CELL_MM / 2: qy = yc + iy * CELL_MM / 2
            If ix <> 0 And iy <> 0 Then
                If YTest(qx, qy) < ylimit Then Exit Function
            End If
            If Not InsidePlate(qx, qy) Then Exit Function
            If DistToOutline(qx, qy) < WALL_MM - 0.000000001 Then Exit Function
        Next iy
    Next ix
    CellOK = True
End Function

Private Function CountCells(ByVal x0 As Double, ByVal y0 As Double, ByVal fill As Boolean) As Long
    Dim i As Long, j As Long, n As Long, xc As Double, yc As Double
    n = 0
    For j = 0 To GRID_NJ - 1
        yc = y0 + j * CELL_MM
        If yc > tipCY And yc < topY Then
            For i = 0 To GRID_NI - 1
                xc = x0 + i * CELL_MM
                If InsidePlate(xc, yc) Then
                    If CellOK(xc, yc) Then
                        n = n + 1
                        If fill Then inWin(i, j) = True
                    End If
                End If
            Next i
        End If
    Next j
    CountCells = n
End Function

Private Sub ComputeWindow()
    Dim ox As Long, oy As Long, n As Long, best As Long, bx As Double, by As Double
    best = -1
    For ox = 0 To 3
        For oy = 0 To 3
            n = CountCells(-20 + ox * CELL_MM / 4, oy * CELL_MM / 4, False)
            If n > best Then
                best = n: bx = -20 + ox * CELL_MM / 4: by = oy * CELL_MM / 4
            End If
        Next oy
    Next ox
    gridX0 = bx: gridY0 = by
    ReDim inWin(-2 To GRID_NI + 1, -2 To GRID_NJ + 1)
    n = CountCells(gridX0, gridY0, True)
    CleanWindow
    CollectCells
    AssignZones
    TraceWindow
End Sub

Private Function EdgeNeighbours(ByVal i As Long, ByVal j As Long) As Long
    Dim n As Long
    n = 0
    If inWin(i + 1, j) Then n = n + 1
    If inWin(i - 1, j) Then n = n + 1
    If inWin(i, j + 1) Then n = n + 1
    If inWin(i, j - 1) Then n = n + 1
    EdgeNeighbours = n
End Function

' Removes cells that touch the rest only at a corner, and loose single cells, so the window is
' one simple outline.
Private Sub CleanWindow()
    Dim changed As Boolean, i As Long, j As Long, dj As Long
    Do
        changed = False
        For j = 0 To GRID_NJ - 1
            For i = 0 To GRID_NI - 1
                For dj = -1 To 1 Step 2
                    If inWin(i, j) And inWin(i + 1, j + dj) And Not inWin(i + 1, j) And Not inWin(i, j + dj) Then
                        If EdgeNeighbours(i + 1, j + dj) < EdgeNeighbours(i, j) Then
                            inWin(i + 1, j + dj) = False
                        Else
                            inWin(i, j) = False
                        End If
                        changed = True
                    End If
                Next dj
            Next i
        Next j
        For j = 0 To GRID_NJ - 1
            For i = 0 To GRID_NI - 1
                If inWin(i, j) Then
                    If EdgeNeighbours(i, j) = 0 Then
                        inWin(i, j) = False
                        changed = True
                    End If
                End If
            Next i
        Next j
    Loop While changed
End Sub

Private Sub CollectCells()
    Dim i As Long, j As Long
    nCells = 0
    For j = 0 To GRID_NJ - 1
        For i = 0 To GRID_NI - 1
            If inWin(i, j) Then nCells = nCells + 1
        Next i
    Next j
    If nCells = 0 Then Exit Sub
    ReDim cellI(1 To nCells): ReDim cellJ(1 To nCells): ReDim cellX(1 To nCells): ReDim cellY(1 To nCells)
    ReDim cellZone(1 To nCells)
    nCells = 0
    For j = 0 To GRID_NJ - 1
        For i = 0 To GRID_NI - 1
            If inWin(i, j) Then
                nCells = nCells + 1
                cellI(nCells) = i: cellJ(nCells) = j
                cellX(nCells) = gridX0 + i * CELL_MM: cellY(nCells) = gridY0 + j * CELL_MM
            End If
        Next i
    Next j
End Sub

Private Function RowHasCells(ByVal j As Long) As Boolean
    Dim i As Long
    RowHasCells = False
    For i = 0 To GRID_NI - 1
        If inWin(i, j) Then
            RowHasCells = True
            Exit Function
        End If
    Next i
End Function

' Zone 1 is the top band of rows. Rows are shared out as evenly as possible, extra rows go to
' the top zones.
Private Sub AssignZones()
    Dim j As Long, nRows As Long, perZone As Long, extra As Long, z As Long, cnt As Long, k As Long, cap As Long
    ReDim rowZone(0 To GRID_NJ - 1)
    nRows = 0
    For j = 0 To GRID_NJ - 1
        If RowHasCells(j) Then nRows = nRows + 1
    Next j
    If nRows = 0 Then Exit Sub
    perZone = nRows \ nZones: extra = nRows Mod nZones
    z = 1: cnt = 0
    For j = GRID_NJ - 1 To 0 Step -1
        If RowHasCells(j) Then
            cap = perZone + IIf(z <= extra, 1, 0)
            If cnt >= cap And z < nZones Then
                z = z + 1: cnt = 0
            End If
            rowZone(j) = z
            cnt = cnt + 1
        End If
    Next j
    For z = 1 To nZones
        zoneRows(z) = 0: zoneCells(z) = 0
    Next z
    For j = 0 To GRID_NJ - 1
        If RowHasCells(j) Then zoneRows(rowZone(j)) = zoneRows(rowZone(j)) + 1
    Next j
    For k = 1 To nCells
        cellZone(k) = rowZone(cellJ(k))
        zoneCells(cellZone(k)) = zoneCells(cellZone(k)) + 1
    Next k
End Sub

' Traces the boundary of the window cells as one counter-clockwise loop of corner points.
' Corner (I, J) is the lower-left corner of cell (I, J).
Private Sub TraceWindow()
    Dim i As Long, j As Long, nEdges As Long, sI As Long, sJ As Long, found As Boolean
    Dim nxtI() As Long, nxtJ() As Long
    Dim ci As Long, cj As Long, ni As Long, nj As Long, guard As Long
    Dim rawI() As Long, rawJ() As Long, nRaw As Long, k As Long, kp As Long, kn As Long, cross As Long
    ReDim nxtI(0 To GRID_NI, 0 To GRID_NJ): ReDim nxtJ(0 To GRID_NI, 0 To GRID_NJ)
    found = False: nEdges = 0
    For j = 0 To GRID_NJ - 1
        For i = 0 To GRID_NI - 1
            If inWin(i, j) Then
                If Not inWin(i, j - 1) Then SetNext nxtI, nxtJ, i, j, i + 1, j, nEdges
                If Not inWin(i + 1, j) Then SetNext nxtI, nxtJ, i + 1, j, i + 1, j + 1, nEdges
                If Not inWin(i, j + 1) Then SetNext nxtI, nxtJ, i + 1, j + 1, i, j + 1, nEdges
                If Not inWin(i - 1, j) Then SetNext nxtI, nxtJ, i, j + 1, i, j, nEdges
                If Not found Then
                    sI = i: sJ = j: found = True
                End If
            End If
        Next i
    Next j
    loopN = 0: loopEdgesOK = False
    If Not found Then Exit Sub
    ReDim rawI(1 To nEdges + 1): ReDim rawJ(1 To nEdges + 1)
    ci = sI: cj = sJ: nRaw = 0: guard = 0
    Do
        nRaw = nRaw + 1: rawI(nRaw) = ci: rawJ(nRaw) = cj
        ni = nxtI(ci, cj): nj = nxtJ(ci, cj)
        ci = ni: cj = nj
        guard = guard + 1
        If guard > nEdges Then Exit Do
    Loop Until ci = sI And cj = sJ
    loopEdgesOK = (nRaw = nEdges)
    ReDim loopX(1 To nRaw): ReDim loopY(1 To nRaw)
    For k = 1 To nRaw
        kp = k - 1: If kp < 1 Then kp = nRaw
        kn = k + 1: If kn > nRaw Then kn = 1
        cross = (rawI(k) - rawI(kp)) * (rawJ(kn) - rawJ(k)) - (rawJ(k) - rawJ(kp)) * (rawI(kn) - rawI(k))
        If cross <> 0 Then
            loopN = loopN + 1
            loopX(loopN) = gridX0 + (rawI(k) - 0.5) * CELL_MM
            loopY(loopN) = gridY0 + (rawJ(k) - 0.5) * CELL_MM
        End If
    Next k
End Sub

Private Sub SetNext(nxtI() As Long, nxtJ() As Long, ByVal i0 As Long, ByVal j0 As Long, _
                    ByVal i1 As Long, ByVal j1 As Long, ByRef nEdges As Long)
    nxtI(i0, j0) = i1: nxtJ(i0, j0) = j1
    nEdges = nEdges + 1
End Sub

' ------------------------------------------- rods --------------------------------------------
' Rods sit on the cell centres. X rods run along each row, Y rods up each column, Z rods through
' the plate at every cell. Rods stop ROD_INTO_WALL_MM inside the solid frame, and Y rods change
' diameter exactly on the zone boundaries (cell faces). kind 1 graded, 2 uniform.
Private Sub PlanRods(ByVal t As Double, ByVal kind As Long, ByVal dUni As Double)
    Dim k As Long, i As Long, j As Long, a As Double, e As Double
    Dim iA As Long, x0 As Double, x1 As Double, d As Double
    Dim jA As Long, jB As Long, js As Long, je As Long, y0 As Double, y1 As Double, dSeg As Double
    a = CELL_MM: e = ROD_INTO_WALL_MM
    nzLev = NzFor(t)
    ReDim zLev(1 To nzLev)
    For k = 1 To nzLev
        zLev(k) = ((k - 1) - (nzLev - 1) / 2) * a
    Next k

    nXG = 0
    ReDim xgX0(1 To GRID_NJ * 4): ReDim xgX1(1 To GRID_NJ * 4): ReDim xgD(1 To GRID_NJ * 4)
    ReDim xgCount(1 To GRID_NJ * 4): ReDim xgRows(1 To GRID_NJ * 4)
    For j = 0 To GRID_NJ - 1
        i = 0
        Do While i < GRID_NI
            If inWin(i, j) Then
                iA = i
                Do While i < GRID_NI
                    If Not inWin(i, j) Then Exit Do
                    i = i + 1
                Loop
                x0 = gridX0 + iA * a - a / 2 - e
                x1 = gridX0 + (i - 1) * a + a / 2 + e
                d = RodDia(rowZone(j), kind, dUni)
                AddXGroup x0, x1, d, gridY0 + j * a
            Else
                i = i + 1
            End If
        Loop
    Next j

    nYG = 0
    ReDim ygY0(1 To GRID_NI * 20): ReDim ygY1(1 To GRID_NI * 20): ReDim ygD(1 To GRID_NI * 20)
    ReDim ygCount(1 To GRID_NI * 20): ReDim ygCols(1 To GRID_NI * 20)
    For i = 0 To GRID_NI - 1
        j = 0
        Do While j < GRID_NJ
            If inWin(i, j) Then
                jA = j
                Do While j < GRID_NJ
                    If Not inWin(i, j) Then Exit Do
                    j = j + 1
                Loop
                jB = j - 1
                js = jA
                Do While js <= jB
                    dSeg = RodDia(rowZone(js), kind, dUni)
                    je = js
                    Do While je + 1 <= jB
                        If Abs(RodDia(rowZone(je + 1), kind, dUni) - dSeg) > 0.0000001 Then Exit Do
                        je = je + 1
                    Loop
                    y0 = gridY0 + js * a - a / 2
                    If js = jA Then y0 = y0 - e
                    y1 = gridY0 + je * a + a / 2
                    If je = jB Then y1 = y1 + e
                    AddYGroup y0, y1, dSeg, gridX0 + i * a
                    js = je + 1
                Loop
            Else
                j = j + 1
            End If
        Loop
    Next i
End Sub

Private Function RodDia(ByVal zone As Long, ByVal kind As Long, ByVal dUni As Double) As Double
    If kind = 2 Then RodDia = dUni Else RodDia = zoneDia(zone)
End Function

Private Sub AddXGroup(ByVal x0 As Double, ByVal x1 As Double, ByVal d As Double, ByVal y As Double)
    Dim g As Long
    For g = 1 To nXG
        If Abs(xgX0(g) - x0) < 0.000001 And Abs(xgX1(g) - x1) < 0.000001 And Abs(xgD(g) - d) < 0.000001 Then
            xgRows(g) = xgRows(g) & ";" & NumStr(y): xgCount(g) = xgCount(g) + 1
            Exit Sub
        End If
    Next g
    nXG = nXG + 1
    xgX0(nXG) = x0: xgX1(nXG) = x1: xgD(nXG) = d: xgRows(nXG) = NumStr(y): xgCount(nXG) = 1
End Sub

Private Sub AddYGroup(ByVal y0 As Double, ByVal y1 As Double, ByVal d As Double, ByVal x As Double)
    Dim g As Long
    For g = 1 To nYG
        If Abs(ygY0(g) - y0) < 0.000001 And Abs(ygY1(g) - y1) < 0.000001 And Abs(ygD(g) - d) < 0.000001 Then
            ygCols(g) = ygCols(g) & ";" & NumStr(x): ygCount(g) = ygCount(g) + 1
            Exit Sub
        End If
    Next g
    nYG = nYG + 1
    ygY0(nYG) = y0: ygY1(nYG) = y1: ygD(nYG) = d: ygCols(nYG) = NumStr(x): ygCount(nYG) = 1
End Sub

' ---------------------------------------- lattice sizing -------------------------------------
' Strut diameter giving the target porosity in a simple cubic cell of size a.
' Metal fraction = (3 pi / 4)(d/a)^2 - sqrt(2)(d/a)^3.
Private Function DiameterForPorosity(ByVal por As Double, ByVal a As Double) As Double
    Dim target As Double, x As Double, f As Double, df As Double, k As Long
    target = 1 - por
    x = Sqr(target / (0.75 * PI_))
    For k = 1 To 60
        f = 0.75 * PI_ * x * x - Sqr(2) * x ^ 3 - target
        df = 1.5 * PI_ * x - 3 * Sqr(2) * x * x
        x = x - f / df
    Next k
    DiameterForPorosity = x * a
End Function

Private Function NzFor(ByVal t As Double) As Long
    NzFor = Int(t / CELL_MM + 0.000000001)
    If NzFor < 1 Then NzFor = 1
End Function

' Metal in one window column (a by a footprint through the plate thickness t), with nz rod layers.
Private Function CellMetal(ByVal r As Double, ByVal t As Double) As Double
    Dim nz As Long
    nz = NzFor(t)
    CellMetal = nz * (3 * PI_ * r * r * CELL_MM - 8 * Sqr(2) * r ^ 3) + PI_ * r * r * (t - nz * CELL_MM)
End Function

Private Function GradedLatticeVolume(ByVal t As Double) As Double
    Dim k As Long, s As Double
    s = 0
    For k = 1 To nCells
        s = s + CellMetal(zoneDia(cellZone(k)) / 2, t)
    Next k
    GradedLatticeVolume = s
End Function

' Uniform strut diameter that puts exactly the same metal in the window as the graded lattice.
Private Function UniformDiameter(ByVal t As Double) As Double
    Dim target As Double, r As Double, f As Double, df As Double, k As Long, h As Double
    target = GradedLatticeVolume(t) / nCells
    r = zoneDia(1) / 2
    h = 0.000001
    For k = 1 To 80
        f = CellMetal(r, t) - target
        df = (CellMetal(r + h, t) - CellMetal(r - h, t)) / (2 * h)
        r = r - f / df
    Next k
    UniformDiameter = 2 * r
End Function

' =============================================================================================
'                                        MODEL BUILDER
' kind: 0 solid, 1 graded lattice, 2 uniform lattice
' =============================================================================================
Private Function BuildModel(ByVal runId As String, ByVal t As Double, ByVal n As Double, _
                            ByVal kind As Long, ByVal keepOpen As Boolean) As Boolean
    Dim fileName As String, kindName As String, startTime As Double, secs As Double
    Dim dUni As Double, vStem As Double, vHead As Double, warnings As String
    Dim sk As Object, splitOK As Boolean, rotOK As Boolean, collarOK As Boolean, path As String

    BuildModel = False
    startTime = Timer
    warnings = ""
    dUni = 0
    Select Case kind
    Case 1: kindName = "graded lattice"
    Case 2: kindName = "uniform lattice"
    Case Else: kindName = "solid"
    End Select
    fileName = runId & " T" & NumStr(t) & " N" & NumStr(n)
    If kind = 1 Then fileName = fileName & " graded"
    If kind = 2 Then fileName = fileName & " uniform"
    WriteLog "----- " & fileName & " (" & kindName & ", plate " & NumStr(t) & " mm, neck " & NumStr(n) & " mm)"
    currentFile = fileName

    On Error GoTo Failed

    SetStep "check the inputs"
    If t < 1 Or t > 30 Then Err.Raise ERR_STEM + 1, , "Thickness " & NumStr(t) & " mm is outside 1 to 30 mm."
    If n < 4 Or n > 2 * (HEAD_RADIUS_MM - 4) Then Err.Raise ERR_STEM + 1, , "Neck diameter " & NumStr(n) & " mm does not fit the 30 mm head."
    currentKind = kind
    currentDUni = 0
    If kind > 0 Then
        If nCells < 10 Then Err.Raise ERR_STEM + 1, , "The lattice window holds only " & nCells & " cells."
        If Not loopEdgesOK Then Err.Raise ERR_STEM + 1, , "The lattice window outline is not one simple loop."
        dUni = UniformDiameter(t)
        currentDUni = dUni
        PlanRods t, kind, dUni
        If kind = 2 Then WriteLog "Uniform strut diameter " & Format(dUni, "0.0000") & " mm matches the graded lattice metal volume."
        If kind = 2 And dUni < MIN_STRUT_MM Then warnings = warnings & "Uniform strut is below the printing limit. "
    End If

    SetStep "create a new part"
    SW_NewPart
    offCombo(1) = -1: offCombo(2) = -1: offCombo(3) = -1

    ' ---- stem plate
    SetStep "sketch and extrude the stem plate"
    SW_BeginSketchPlane 1
    DrawOutline
    Set sk = SW_EndSketch("Sketch stem outline")
    If Not SW_ExtrudeMid(sk, t, True, "Stem plate") Then Err.Raise ERR_STEM + 2, , "The plate extrusion failed."

    ' ---- split the front and back faces along the potting and reading lines
    SetStep "split the front and back faces at the potting level"
    splitOK = SplitFrontBack(t)
    If Not splitOK Then warnings = warnings & "Front and back faces not split. "

    ' ---- lattice
    If kind > 0 Then
        SetStep "cut the lattice window"
        SW_BeginSketchPlane 1
        DrawWindowLoop
        Set sk = SW_EndSketch("Sketch lattice window")
        If Not SW_CutMid(sk, t + 2, "Lattice window") Then Err.Raise ERR_STEM + 3, , "The window cut failed."
        SetStep "check the extrusion direction on the Right and Top planes"
        offCombo(3) = CalibrateOffset(3)
        offCombo(2) = CalibrateOffset(2)
        SetStep "add the X rods"
        AddXRods
        SetStep "add the Y rods"
        AddYRods
        SetStep "add the Z rods"
        AddZRods t
        SetStep "check the lattice"
        CheckLattice warnings
    End If

    ' ---- collar, neck and head
    SetStep "add the neck collar"
    collarOK = AddCollar(t, n)
    If Not collarOK Then warnings = warnings & "Neck collar missing, so the neck starts 1 mm inside the plate. "
    SetStep "add the neck"
    SW_BeginSketchPlane 1
    If collarOK Then
        DrawNeckProfile n, NECK_START_MM
    Else
        DrawNeckProfile n, -1
    End If
    Set sk = SW_EndSketch("Sketch neck")
    If Not SW_Revolve(sk, True, "Neck") Then Err.Raise ERR_STEM + 4, , "The neck revolve failed."
    SetStep "add the head"
    SW_BeginSketchPlane 1
    DrawHeadProfile n
    Set sk = SW_EndSketch("Sketch head")
    If Not SW_Revolve(sk, False, "Head") Then Err.Raise ERR_STEM + 5, , "The head revolve failed."
    SetStep "cut the load face on the head"
    SW_BeginSketchPlane 1
    DrawHeadFlatCut
    Set sk = SW_EndSketch("Sketch head load face")
    If Not SW_CutMid(sk, 2 * HEAD_RADIUS_MM + 10, "Head load face") Then warnings = warnings & "Head load face not cut. "
    SetStep "name the bodies"
    NameBodies headCX, headCY

    ' ---- tilt into the ISO 7206-4 test position
    SetStep "tilt the model into the ISO 7206-4 position"
    rotOK = SW_RotateAll(isoTheta)
    If Not rotOK Then warnings = warnings & "Model not tilted, orient it as in the Week 8 video. "
    If rotOK Then
        SetStep "add the reference planes"
        SW_PlaneParallelTop yPot, "Potting level (D = 80 mm)"
        SW_PlaneParallelTop yRead, "Reading line (3 mm above potting)"
        If Not splitOK Then
            SetStep "split the faces with the potting planes"
            If SW_SplitIntersect() Then
                warnings = Replace(warnings, "Front and back faces not split. ", "")
            End If
        End If
    End If

    ' ---- measure, label and save
    SetStep "measure the volumes"
    If rotOK Then
        MeasureBodies isoCX, isoCY, vStem, vHead
    Else
        MeasureBodies headCX, headCY, vStem, vHead
    End If
    If kind > 0 Then LatticeVolumeCheck t, n, kind, dUni, vStem, warnings
    SetStep "write the custom properties"
    WriteProperties runId, t, n, kindName, dUni
    SetStep "save the part"
    path = FreePath(fileName)
    SW_SaveAndClose path, keepOpen

    secs = Timer - startTime
    If secs < 0 Then secs = secs + 86400
    WriteResult runId, t, n, kindName, vStem, vHead, dUni, warnings
    summary = summary & fileName & "    stem " & Format(vStem, "0") & " mm3"
    If Len(warnings) > 0 Then summary = summary & "    (see log)"
    summary = summary & vbCrLf
    WriteLog "Saved " & path & " in " & Format(secs, "0") & " s. " & warnings
    BuildModel = True
    Exit Function

Failed:
    WriteLog "FAILED while trying to " & stepName & ". " & Err.Description
    summary = summary & fileName & "    FAILED while trying to " & stepName & vbCrLf
    Resume AfterFailure
AfterFailure:
    ' The failed part stays open, so the step that went wrong can be seen in the feature tree.
    On Error Resume Next
    If Not TestMode Then
        If Not swModel.SketchManager.ActiveSketch Is Nothing Then swModel.SketchManager.InsertSketch True
    End If
    RestoreView
    BuildModel = False
End Function

' ------------------------------------- sketch content ----------------------------------------
Private Sub DrawOutline()
    Dim k As Long
    For k = 1 To nPrim
        If pKind(k) = 1 Then
            SW_Line pX0(k), pY0(k), 0, pX1(k), pY1(k), 0
        Else
            SW_Arc pCX(k), pCY(k), 0, pX0(k), pY0(k), 0, pX1(k), pY1(k), 0, pDir(k)
        End If
    Next k
End Sub

Private Sub DrawWindowLoop()
    Dim k As Long, kn As Long
    For k = 1 To loopN
        kn = k + 1
        If kn > loopN Then kn = 1
        SW_Line loopX(k), loopY(k), 0, loopX(kn), loopY(kn), 0
    Next k
End Sub

' Neck: rectangle on the axis from startMM to NECK_END_MM, revolved about the neck axis.
' startMM is measured from the neck base face, positive away from the plate.
Private Sub DrawNeckProfile(ByVal n As Double, ByVal startMM As Double)
    Dim ax As Double, ay As Double, bx As Double, by As Double, r As Double
    r = n / 2
    ax = baseX + startMM * uX: ay = baseY + startMM * uY
    bx = baseX + NECK_END_MM * uX: by = baseY + NECK_END_MM * uY
    SW_Line ax, ay, 0, bx, by, 0
    SW_Line bx, by, 0, bx + r * vX, by + r * vY, 0
    SW_Line bx + r * vX, by + r * vY, 0, ax + r * vX, ay + r * vY, 0
    SW_Line ax + r * vX, ay + r * vY, 0, ax, ay, 0
    SW_CenterLine baseX - 6 * uX, baseY - 6 * uY, 0, baseX - 2 * uX, baseY - 2 * uY, 0
End Sub

' Head: sphere of HEAD_RADIUS_MM about the neck end, flat bottom HEAD_BOTTOM_MM below the centre,
' bored to the neck diameter from the bottom face up to the centre.
Private Sub DrawHeadProfile(ByVal n As Double)
    Dim Rh As Double, rb As Double, r As Double
    Dim p0x As Double, p0y As Double, pbx As Double, pby As Double, pcx As Double, pcy As Double
    Dim pdx As Double, pdy As Double, botX As Double, botY As Double
    Rh = HEAD_RADIUS_MM: r = n / 2
    rb = Sqr(Rh * Rh - HEAD_BOTTOM_MM * HEAD_BOTTOM_MM)
    p0x = headCX + Rh * uX: p0y = headCY + Rh * uY
    botX = headCX - HEAD_BOTTOM_MM * uX: botY = headCY - HEAD_BOTTOM_MM * uY
    pbx = botX + rb * vX: pby = botY + rb * vY
    pcx = botX + r * vX: pcy = botY + r * vY
    pdx = headCX + r * vX: pdy = headCY + r * vY
    SW_Arc headCX, headCY, 0, p0x, p0y, 0, pbx, pby, 0, 1
    SW_Line pbx, pby, 0, pcx, pcy, 0
    SW_Line pcx, pcy, 0, pdx, pdy, 0
    SW_Line pdx, pdy, 0, headCX, headCY, 0
    SW_Line headCX, headCY, 0, p0x, p0y, 0
    SW_CenterLine headCX + (Rh + 3) * uX, headCY + (Rh + 3) * uY, 0, headCX + (Rh + 8) * uX, headCY + (Rh + 8) * uY, 0
End Sub

' Removes the top HEAD_FLAT_MM of the head along the ISO load axis, leaving a flat face that is
' horizontal once the model is tilted into the test position.
Private Sub DrawHeadFlatCut()
    Dim nx As Double, ny As Double, tx As Double, ty As Double, h0 As Double, h1 As Double, w As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, cx As Double, cy As Double, dx As Double, dy As Double
    nx = sinT: ny = cosT                 ' test-frame "up", written in the stem frame
    tx = cosT: ty = -sinT                ' test-frame "right", written in the stem frame
    h0 = HEAD_RADIUS_MM - HEAD_FLAT_MM: h1 = HEAD_RADIUS_MM + 10: w = HEAD_RADIUS_MM + 5
    ax = headCX + h0 * nx - w * tx: ay = headCY + h0 * ny - w * ty
    bx = headCX + h0 * nx + w * tx: by = headCY + h0 * ny + w * ty
    cx = headCX + h1 * nx + w * tx: cy = headCY + h1 * ny + w * ty
    dx = headCX + h1 * nx - w * tx: dy = headCY + h1 * ny - w * ty
    SW_Line ax, ay, 0, bx, by, 0
    SW_Line bx, by, 0, cx, cy, 0
    SW_Line cx, cy, 0, dx, dy, 0
    SW_Line dx, dy, 0, ax, ay, 0
End Sub

' ------------------------------------------ lattice ------------------------------------------
Private Sub AddXRods()
    Dim g As Long, k As Long, zz As Long, rows() As String, sk As Object, y As Double
    For g = 1 To nXG
        SW_BeginSketchPlane 3
        rows = Split(xgRows(g), ";")
        For k = 0 To UBound(rows)
            y = Val(rows(k))
            For zz = 1 To nzLev
                SW_Circle 0, y, zLev(zz), xgD(g) / 2
            Next zz
        Next k
        Set sk = SW_EndSketch("Sketch X rods " & Format(g, "00"))
        If Not SW_ExtrudeFromOffset(sk, xgX0(g), xgX1(g) - xgX0(g), True, "X rods " & Format(g, "00"), offCombo(3)) Then
            Err.Raise ERR_STEM + 6, , "X rod group " & g & " did not extrude."
        End If
    Next g
End Sub

Private Sub AddYRods()
    Dim g As Long, k As Long, zz As Long, cols() As String, sk As Object, x As Double
    For g = 1 To nYG
        SW_BeginSketchPlane 2
        cols = Split(ygCols(g), ";")
        For k = 0 To UBound(cols)
            x = Val(cols(k))
            For zz = 1 To nzLev
                SW_Circle x, 0, zLev(zz), ygD(g) / 2
            Next zz
        Next k
        Set sk = SW_EndSketch("Sketch Y rods " & Format(g, "00"))
        If Not SW_ExtrudeFromOffset(sk, ygY0(g), ygY1(g) - ygY0(g), True, "Y rods " & Format(g, "00"), offCombo(2)) Then
            Err.Raise ERR_STEM + 7, , "Y rod group " & g & " did not extrude."
        End If
    Next g
End Sub

Private Sub AddZRods(ByVal t As Double)
    Dim k As Long, sk As Object, d As Double
    SW_BeginSketchPlane 1
    For k = 1 To nCells
        If currentKind = 2 Then d = currentDUni Else d = zoneDia(cellZone(k))
        SW_Circle cellX(k), cellY(k), 0, d / 2
    Next k
    Set sk = SW_EndSketch("Sketch Z rods")
    If Not SW_ExtrudeMid(sk, t, True, "Z rods") Then Err.Raise ERR_STEM + 8, , "The Z rods did not extrude."
End Sub

Private Sub CheckLattice(ByRef warnings As String)
    Dim nb As Long
    If TestMode Then Exit Sub
    nb = BodyCount()
    If nb <> 1 Then
        warnings = warnings & "Stem has " & nb & " bodies after the lattice. "
        WriteLog "WARNING the stem has " & nb & " bodies after adding the lattice (expected 1)."
    Else
        WriteLog "Lattice fused into one stem body."
    End If
End Sub

' Logs the lattice metal and, when the solid version was built in this run, compares the
' measured stem volume with the prediction (solid - window + lattice).
Private Sub LatticeVolumeCheck(ByVal t As Double, ByVal n As Double, ByVal kind As Long, ByVal dUni As Double, _
                               ByVal vMeasured As Double, ByRef warnings As String)
    Dim vLat As Double, vPred As Double
    If kind = 1 Then vLat = GradedLatticeVolume(t) Else vLat = nCells * CellMetal(dUni / 2, t)
    WriteLog "Window " & nCells & " cells. Lattice metal " & Format(vLat, "0.0") & " mm3 (" & _
        Format(100 * vLat / (nCells * CELL_MM * CELL_MM * t), "0.0") & "% of the window). Measured stem " & _
        Format(vMeasured, "0.0") & " mm3."
    If TestMode Then Exit Sub
    If solidStemVolume > 0 And Abs(solidStemT - t) < 0.000001 And Abs(solidStemN - n) < 0.000001 Then
        vPred = solidStemVolume - nCells * CELL_MM * CELL_MM * t + vLat
        WriteLog "Predicted stem volume " & Format(vPred, "0.0") & " mm3. Difference " & _
            Format(100 * (vMeasured - vPred) / vPred, "0.00") & "%."
        If Abs(vMeasured - vPred) > 0.01 * vPred Then warnings = warnings & "Stem volume is more than 1% off the prediction. "
    End If
End Sub

' --------------------------------------- collar --------------------------------------------
' Elliptical pad on the neck base face, as in the course model. Sketched on that face and
' extruded COLLAR_LEN_MM outwards. The volume it adds shows whether it went the right way.
' Returns False, with nothing left behind but an unused sketch, if the collar cannot be made.
Private Function AddCollar(ByVal t As Double, ByVal n As Double) As Boolean
    Dim semiZ As Double, vBefore As Double, vAdded As Double, expected As Double, sk As Object
    Dim k As Long, nStart As Long
    AddCollar = False
    semiZ = n / 2 + COLLAR_EXTRA_MM
    If Abs(semiZ - t / 2) < 0.3 Then semiZ = t / 2 + 0.3          ' keeps the collar edge off the plate faces
    expected = PI_ * COLLAR_HALF_MM * semiZ * COLLAR_LEN_MM
    On Error GoTo CollarFailed
    ' the ray travels along -u, into the plate, so it meets the neck base face square on
    If Not SW_BeginSketchOnFace(baseX, baseY, 0, -uX, -uY, 0) Then
        WriteLog "Could not select the neck base face for the collar."
        Exit Function
    End If
    SW_Ellipse baseX, baseY, 0, baseX + COLLAR_HALF_MM * vX, baseY + COLLAR_HALF_MM * vY, 0, baseX, baseY, semiZ
    Set sk = SW_EndSketch("Sketch neck collar")
    If TestMode Then
        AddCollar = SW_ExtrudeBlind(sk, COLLAR_LEN_MM, False, True, "Neck collar")
        Exit Function
    End If
    vBefore = BodyVolume(0, 0, 0)
    For k = 0 To 1
        nStart = swModel.GetFeatureCount
        If SW_ExtrudeBlind(sk, COLLAR_LEN_MM, (k = 1), True, "Neck collar") Then
            vAdded = BodyVolume(0, 0, 0) - vBefore
            If Abs(vAdded - expected) <= 0.05 * expected Then
                AddCollar = True
                Exit Function
            End If
            WriteLog "Collar try " & (k + 1) & " added " & Format(vAdded, "0.0") & " mm3, expected " & _
                Format(expected, "0.0") & " mm3. Removing it."
            SW_DeleteBackTo nStart
        End If
    Next k
    WriteLog "The collar could not be added."
    Exit Function

CollarFailed:
    WriteLog "The collar step stopped: " & Err.Description
    Resume CollarCleanUp
CollarCleanUp:
    On Error Resume Next
    If Not TestMode Then
        If Not swModel.SketchManager.ActiveSketch Is Nothing Then swModel.SketchManager.InsertSketch True
    End If
    AddCollar = False
End Function

' ------------------------------------- split lines -------------------------------------------
' The potting and reading lines are sketched on the front plane and projected onto the front and
' back faces. The side faces are already split because the outline has vertices on those lines.
Private Function SplitFrontBack(ByVal t As Double) As Boolean
    Dim dxl As Double, dyl As Double, sk As Object, py As Double, px As Double
    dxl = cosT: dyl = -sinT
    SW_BeginSketchPlane 1
    SW_Line splLpX - 3 * dxl, splLpY - 3 * dyl, 0, splMpX + 3 * dxl, splMpY + 3 * dyl, 0
    SW_Line splLrX - 3 * dxl, splLrY - 3 * dyl, 0, splMrX + 3 * dxl, splMrY + 3 * dyl, 0
    Set sk = SW_EndSketch("Sketch potting and reading lines")
    py = tipCY + 30: px = (XLeft(py) + XRight(py)) / 2
    SplitFrontBack = SW_SplitProject(sk, px, py, t / 2, px, py, -t / 2, 4)
End Function

' =============================================================================================
'                         SOLIDWORKS WRAPPERS (millimetres in, metres out)
' In test mode every call is written to a text file instead, so the geometry can be checked
' without SolidWorks.
' =============================================================================================
Private Sub Rec(ByVal s As String)
    Print #testFile, s
End Sub

Private Sub SW_NewPart()
    Dim tpl As String
    If TestMode Then
        Rec "NEWPART"
        Exit Sub
    End If
    tpl = ""
    On Error Resume Next
    tpl = swApp.GetUserPreferenceStringValue(8)                 ' default part template
    If LCase(Right(tpl, 7)) <> ".prtdot" Then tpl = swApp.GetDocumentTemplate(1, "", 0, 0, 0)
    On Error GoTo 0
    If Len(tpl) = 0 Then Err.Raise ERR_STEM + 20, , "No default part template is set (Tools > Options > Default Templates)."
    Set swModel = swApp.NewDocument(tpl, 0, 0, 0)
    If swModel Is Nothing Then Err.Raise ERR_STEM + 21, , "SolidWorks could not open a new part from " & tpl
    FindDefaultPlanes
    On Error Resume Next
    swModel.SketchManager.AddToDB = True
    swModel.SketchManager.DisplayWhenAdded = False
    swModel.FeatureManager.EnableFeatureTree = False
    swModel.ActiveView.EnableGraphicsUpdate = False
    On Error GoTo 0
End Sub

Private Sub RestoreView()
    If TestMode Then Exit Sub
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.SketchManager.DisplayWhenAdded = True
    swModel.FeatureManager.EnableFeatureTree = True
    swModel.ActiveView.EnableGraphicsUpdate = True
    swModel.ForceRebuild3 False
    swModel.ViewZoomtofit2
End Sub

' The first three reference planes of a part template are Front, Top and Right. Their normals are
' checked as well, so a template with another order still works.
Private Sub FindDefaultPlanes()
    Dim f As Object, k As Long, a As Variant, cand(1 To 3) As Object
    Dim fF As Object, fT As Object, fR As Object
    k = 0
    Set f = swModel.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "RefPlane" Then
            k = k + 1
            Set cand(k) = f
            If k = 3 Then Exit Do
        End If
        Set f = f.GetNextFeature
    Loop
    If k < 3 Then Err.Raise ERR_STEM + 22, , "The part template has fewer than three reference planes."
    Set planeFront = cand(1): Set planeTop = cand(2): Set planeRight = cand(3)
    On Error Resume Next
    For k = 1 To 3
        Err.Clear
        a = cand(k).GetSpecificFeature2.Transform.ArrayData
        If Err.Number = 0 Then
            If Abs(a(8)) > 0.9 Then
                Set fF = cand(k)
            ElseIf Abs(a(5)) > 0.9 Or Abs(a(7)) > 0.9 Then
                Set fT = cand(k)
            ElseIf Abs(a(2)) > 0.9 Or Abs(a(6)) > 0.9 Then
                Set fR = cand(k)
            End If
        End If
    Next k
    Err.Clear
    On Error GoTo 0
    If Not (fF Is Nothing Or fT Is Nothing Or fR Is Nothing) Then
        Set planeFront = fF: Set planeTop = fT: Set planeRight = fR
    End If
End Sub

Private Sub SW_BeginSketchPlane(ByVal which As Long)
    If TestMode Then
        Rec "SKETCH," & Choose(which, "FRONT", "TOP", "RIGHT")
        Exit Sub
    End If
    swModel.ClearSelection2 True
    Select Case which
    Case 1: planeFront.Select2 False, 0
    Case 2: planeTop.Select2 False, 0
    Case Else: planeRight.Select2 False, 0
    End Select
    swModel.SketchManager.InsertSketch True
    PrepareSketchTransform
End Sub

' Opens a sketch on the face through (x, y, z). (dx, dy, dz) is the direction of a ray that meets
' the face square on from outside the part.
Private Function SW_BeginSketchOnFace(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                                      ByVal dx As Double, ByVal dy As Double, ByVal dz As Double) As Boolean
    If TestMode Then
        Rec "SKETCHFACE," & P3(x, y, z)
        SW_BeginSketchOnFace = True
        Exit Function
    End If
    swModel.ClearSelection2 True
    If Not SelectFace(x, y, z, dx, dy, dz, False, 0) Then
        SW_BeginSketchOnFace = False
        Exit Function
    End If
    swModel.SketchManager.InsertSketch True
    PrepareSketchTransform
    SW_BeginSketchOnFace = True
End Function

' Selects the face through the point (x, y, z). A ray that starts 20 mm back along (dx, dy, dz)
' and travels towards the point is tried first, because a ray does not depend on the view.
' Selecting at the point itself is the fallback.
Private Function SelectFace(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                            ByVal dx As Double, ByVal dy As Double, ByVal dz As Double, _
                            ByVal appendIt As Boolean, ByVal mk As Long) As Boolean
    Dim ok As Boolean
    ok = False
    On Error Resume Next
    ok = swModel.Extension.SelectByRay((x - 20 * dx) / 1000#, (y - 20 * dy) / 1000#, (z - 20 * dz) / 1000#, _
        dx, dy, dz, 0.0001, 2, appendIt, mk, 0)
    If Not ok Then
        Err.Clear
        ok = swModel.Extension.SelectByID2("", "FACE", x / 1000#, y / 1000#, z / 1000#, appendIt, mk, Nothing, 0)
    End If
    Err.Clear
    On Error GoTo 0
    SelectFace = ok
End Function

Private Sub PrepareSketchTransform()
    Dim ox As Double, oy As Double, oz As Double, ax As Double, ay As Double, az As Double
    Dim bx As Double, by As Double, bz As Double
    Set curSketch = swModel.SketchManager.ActiveSketch
    If curSketch Is Nothing Then Err.Raise ERR_STEM + 23, , "No active sketch after InsertSketch."
    Set curXform = curSketch.ModelToSketchTransform
    ToSketch 0, 0, 0, ox, oy, oz
    ToSketch 1, 0, 0, ax, ay, az
    ToSketch 0, 1, 0, bx, by, bz
    curArcFlip = ((ax - ox) * (by - oy) - (ay - oy) * (bx - ox)) < 0
End Sub

Private Sub ToSketch(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                     ByRef sx As Double, ByRef sy As Double, ByRef sz As Double)
    Dim d(2) As Double, vIn As Variant, mp As Object, vOut As Variant
    d(0) = x / 1000#: d(1) = y / 1000#: d(2) = z / 1000#
    vIn = d
    Set mp = swMath.CreatePoint(vIn)
    Set mp = mp.MultiplyTransform(curXform)
    vOut = mp.ArrayData
    sx = vOut(0): sy = vOut(1): sz = vOut(2)
End Sub

Private Function SW_EndSketch(ByVal nm As String) As Object
    Dim f As Object
    If TestMode Then
        Rec "ENDSKETCH," & nm
        Set SW_EndSketch = Nothing
        Exit Function
    End If
    swModel.SketchManager.InsertSketch True
    Set f = swModel.FeatureByPositionReverse(0)
    On Error Resume Next
    f.Name = nm
    On Error GoTo 0
    Set SW_EndSketch = f
End Function

Private Sub SW_Line(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
                    ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double
    If TestMode Then
        Rec "LINE," & P3(x1, y1, z1) & "," & P3(x2, y2, z2)
        Exit Sub
    End If
    ToSketch x1, y1, z1, a, b, c
    ToSketch x2, y2, z2, d, e, f
    If swModel.SketchManager.CreateLine(a, b, c, d, e, f) Is Nothing Then Err.Raise ERR_STEM + 24, , "A sketch line could not be created."
End Sub

Private Sub SW_CenterLine(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
                          ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double
    If TestMode Then
        Rec "CENTERLINE," & P3(x1, y1, z1) & "," & P3(x2, y2, z2)
        Exit Sub
    End If
    ToSketch x1, y1, z1, a, b, c
    ToSketch x2, y2, z2, d, e, f
    Set lastCenterline = swModel.SketchManager.CreateCenterLine(a, b, c, d, e, f)
    If lastCenterline Is Nothing Then Err.Raise ERR_STEM + 25, , "A centreline could not be created."
End Sub

Private Sub SW_Arc(ByVal cx As Double, ByVal cy As Double, ByVal cz As Double, _
                   ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, _
                   ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal dirv As Long)
    Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double
    Dim g As Double, h As Double, i As Double
    If TestMode Then
        Rec "ARC," & P3(cx, cy, cz) & "," & P3(x1, y1, z1) & "," & P3(x2, y2, z2) & "," & dirv
        Exit Sub
    End If
    ToSketch cx, cy, cz, a, b, c
    ToSketch x1, y1, z1, d, e, f
    ToSketch x2, y2, z2, g, h, i
    If curArcFlip Then dirv = -dirv
    If swModel.SketchManager.CreateArc(a, b, c, d, e, f, g, h, i, dirv) Is Nothing Then Err.Raise ERR_STEM + 26, , "A sketch arc could not be created."
End Sub

Private Sub SW_Circle(ByVal cx As Double, ByVal cy As Double, ByVal cz As Double, ByVal r As Double)
    Dim a As Double, b As Double, c As Double
    If TestMode Then
        Rec "CIRCLE," & P3(cx, cy, cz) & "," & NumStr(r)
        Exit Sub
    End If
    ToSketch cx, cy, cz, a, b, c
    If swModel.SketchManager.CreateCircleByRadius(a, b, c, r / 1000#) Is Nothing Then Err.Raise ERR_STEM + 27, , "A sketch circle could not be created."
End Sub

Private Sub SW_Ellipse(ByVal cx As Double, ByVal cy As Double, ByVal cz As Double, _
                       ByVal mx As Double, ByVal my As Double, ByVal mz As Double, _
                       ByVal nx As Double, ByVal ny As Double, ByVal nz As Double)
    Dim a As Double, b As Double, c As Double, d As Double, e As Double, f As Double
    Dim g As Double, h As Double, i As Double
    If TestMode Then
        Rec "ELLIPSE," & P3(cx, cy, cz) & "," & P3(mx, my, mz) & "," & P3(nx, ny, nz)
        Exit Sub
    End If
    ToSketch cx, cy, cz, a, b, c
    ToSketch mx, my, mz, d, e, f
    ToSketch nx, ny, nz, g, h, i
    If swModel.SketchManager.CreateEllipse(a, b, c, d, e, f, g, h, i) Is Nothing Then Err.Raise ERR_STEM + 28, , "The collar ellipse could not be created."
End Sub

Private Sub SelectSketch(sk As Object, ByVal mark As Long)
    swModel.ClearSelection2 True
    sk.Select2 False, mark
End Sub

Private Function SW_ExtrudeMid(sk As Object, ByVal depthMM As Double, ByVal mergeIt As Boolean, ByVal nm As String) As Boolean
    If TestMode Then
        Rec "EXTRUDE_MID," & NumStr(depthMM) & "," & mergeIt & "," & nm
        SW_ExtrudeMid = True
        Exit Function
    End If
    SelectSketch sk, 0
    Set lastFeat = swModel.FeatureManager.FeatureExtrusion3(True, False, False, 6, 0, depthMM / 1000#, 0#, _
        False, False, False, False, 0#, 0#, False, False, False, False, mergeIt, True, True, 0, 0#, False)
    SW_ExtrudeMid = FinishFeature(nm)
End Function

' Extrudes from startMM to startMM + lenMM along the sketch plane normal. combo picks the
' Dir / FlipStartOffset pair that CalibrateOffset found to give that result.
Private Function SW_ExtrudeFromOffset(sk As Object, ByVal startMM As Double, ByVal lenMM As Double, _
                                      ByVal mergeIt As Boolean, ByVal nm As String, ByVal combo As Long) As Boolean
    Dim dirFlip As Boolean, startFlip As Boolean
    If TestMode Then
        Rec "EXTRUDE_OFFSET," & NumStr(startMM) & "," & NumStr(lenMM) & "," & mergeIt & "," & nm
        SW_ExtrudeFromOffset = True
        Exit Function
    End If
    dirFlip = (combo = 1 Or combo = 2)
    startFlip = (combo = 1 Or combo = 3)
    SelectSketch sk, 0
    Set lastFeat = swModel.FeatureManager.FeatureExtrusion3(True, False, dirFlip, 0, 0, lenMM / 1000#, 0#, _
        False, False, False, False, 0#, 0#, False, False, False, False, mergeIt, True, True, 3, startMM / 1000#, startFlip)
    SW_ExtrudeFromOffset = FinishFeature(nm)
End Function

Private Function SW_ExtrudeBlind(sk As Object, ByVal depthMM As Double, ByVal flipIt As Boolean, _
                                 ByVal mergeIt As Boolean, ByVal nm As String) As Boolean
    If TestMode Then
        Rec "EXTRUDE_BLIND," & NumStr(depthMM) & "," & flipIt & "," & mergeIt & "," & nm
        SW_ExtrudeBlind = True
        Exit Function
    End If
    SelectSketch sk, 0
    Set lastFeat = swModel.FeatureManager.FeatureExtrusion3(True, False, flipIt, 0, 0, depthMM / 1000#, 0#, _
        False, False, False, False, 0#, 0#, False, False, False, False, mergeIt, True, True, 0, 0#, False)
    SW_ExtrudeBlind = FinishFeature(nm)
End Function

Private Function SW_CutMid(sk As Object, ByVal depthMM As Double, ByVal nm As String) As Boolean
    If TestMode Then
        Rec "CUT_MID," & NumStr(depthMM) & "," & nm
        SW_CutMid = True
        Exit Function
    End If
    SelectSketch sk, 0
    Set lastFeat = swModel.FeatureManager.FeatureCut4(True, False, False, 6, 0, depthMM / 1000#, 0#, _
        False, False, False, False, 0#, 0#, False, False, False, False, False, True, True, True, True, False, _
        0, 0#, False, False)
    SW_CutMid = FinishFeature(nm)
End Function

' Revolves the sketch 360 degrees about its single centreline. If SolidWorks does not pick the
' centreline up by itself, it is selected explicitly and the revolve is tried again.
Private Function SW_Revolve(sk As Object, ByVal mergeIt As Boolean, ByVal nm As String) As Boolean
    Dim marks As Variant, k As Long, sd As Object
    If TestMode Then
        Rec "REVOLVE," & mergeIt & "," & nm
        SW_Revolve = True
        Exit Function
    End If
    SelectSketch sk, 0
    Set lastFeat = swModel.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, _
        2 * PI_, 0#, False, False, 0#, 0#, 0, 0#, 0#, mergeIt, True, True)
    If lastFeat Is Nothing And Not lastCenterline Is Nothing Then
        marks = Array(4, 16, 1)
        On Error Resume Next
        For k = 0 To UBound(marks)
            SelectSketch sk, 0
            Set sd = swModel.SelectionManager.CreateSelectData
            sd.Mark = marks(k)
            lastCenterline.Select4 True, sd
            Set lastFeat = swModel.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, _
                2 * PI_, 0#, False, False, 0#, 0#, 0, 0#, 0#, mergeIt, True, True)
            If Not lastFeat Is Nothing Then Exit For
        Next k
        On Error GoTo 0
    End If
    SW_Revolve = FinishFeature(nm)
End Function

Private Function FinishFeature(ByVal nm As String) As Boolean
    If lastFeat Is Nothing Then
        FinishFeature = False
        WriteLog "  Feature '" & nm & "' was not created."
        Exit Function
    End If
    On Error Resume Next
    lastFeat.Name = nm
    On Error GoTo 0
    swModel.ClearSelection2 True
    FinishFeature = True
End Function

' Deletes the most recent feature only. A sketch used by that feature stays in the tree.
Private Sub SW_DeleteLastFeature()
    Dim f As Object
    If TestMode Then
        Rec "DELETE_LAST"
        Exit Sub
    End If
    Set f = swModel.FeatureByPositionReverse(0)
    swModel.ClearSelection2 True
    f.Select2 False, 0
    swModel.EditDelete
    swModel.ClearSelection2 True
End Sub

' Deletes features from the end of the tree until the part has keepCount features again, so a
' trial feature and its sketch go, and nothing older than them can.
Private Sub SW_DeleteBackTo(ByVal keepCount As Long)
    Dim guard As Long
    If TestMode Then
        Rec "DELETE_BACK_TO"
        Exit Sub
    End If
    guard = 0
    Do While swModel.GetFeatureCount > keepCount And guard < 4
        SW_DeleteLastFeature
        guard = guard + 1
    Loop
End Sub

' Finds which Dir / FlipStartOffset pair puts an offset extrusion between +300 and +310 mm along
' the plane normal axis, using a throwaway rod far from the stem. Usually the first pair works.
Private Function CalibrateOffset(ByVal which As Long) As Long
    Dim combo As Long, sk As Object, b As Variant, k As Long, bx As Variant
    Dim lo As Double, hi As Double, found As Boolean, nStart As Long
    If TestMode Then
        CalibrateOffset = 0
        Exit Function
    End If
    For combo = 0 To 3
        nStart = swModel.GetFeatureCount
        SW_BeginSketchPlane which
        If which = 3 Then
            SW_Circle 0, 400, 0, 1
        Else
            SW_Circle 400, 0, 0, 1
        End If
        Set sk = SW_EndSketch("Direction check sketch")
        found = False
        If SW_ExtrudeFromOffset(sk, 300, 10, False, "Direction check", combo) Then
            b = swModel.GetBodies2(0, True)
            For k = 0 To UBound(b)
                bx = b(k).GetBodyBox
                If which = 3 Then
                    If Abs((bx(1) + bx(4)) * 500# - 400) < 5 Then
                        lo = bx(0) * 1000#: hi = bx(3) * 1000#: found = True
                    End If
                Else
                    If Abs((bx(0) + bx(3)) * 500# - 400) < 5 Then
                        lo = bx(1) * 1000#: hi = bx(4) * 1000#: found = True
                    End If
                End If
            Next k
        End If
        SW_DeleteBackTo nStart
        If found Then
            If Abs(lo - 300) < 0.01 And Abs(hi - 310) < 0.01 Then
                CalibrateOffset = combo
                WriteLog "Offset extrusions on the " & Choose(which, "Front", "Top", "Right") & " plane use option set " & combo & "."
                Exit Function
            End If
        End If
    Next combo
    Err.Raise ERR_STEM + 9, , "Could not work out the extrusion direction on the " & Choose(which, "Front", "Top", "Right") & " plane."
End Function

' Projection split line. The selection marks this call expects differ between SolidWorks
' versions, so several pairs are tried. Success means the stem gained at least minNew faces.
' (fx, fy, fz) lies on the front face (+Z side) and (bx, by, bz) on the back face (-Z side).
Private Function SW_SplitProject(sk As Object, ByVal fx As Double, ByVal fy As Double, ByVal fz As Double, _
                                 ByVal bx As Double, ByVal by As Double, ByVal bz As Double, ByVal minNew As Long) As Boolean
    Dim marks As Variant, k As Long, before As Long, after As Long, sm As Long, fm As Long
    If TestMode Then
        Rec "SPLIT_PROJECT," & P3(fx, fy, fz) & "," & P3(bx, by, bz)
        SW_SplitProject = True
        Exit Function
    End If
    marks = Array(4, 1, 0, 0, 0, 1, 1, 0, 1, 2, 2, 1)
    before = FaceCountAll()
    For k = 0 To UBound(marks) Step 2
        sm = marks(k): fm = marks(k + 1)
        On Error Resume Next
        swModel.ClearSelection2 True
        sk.Select2 False, sm
        SelectFace fx, fy, fz, 0, 0, -1, True, fm
        SelectFace bx, by, bz, 0, 0, 1, True, fm
        swModel.InsertSplitLineProject False, False
        swModel.ClearSelection2 True
        Err.Clear
        On Error GoTo 0
        after = FaceCountAll()
        If after >= before + minNew Then
            WriteLog "Split line made (sketch mark " & sm & ", face mark " & fm & "). Faces " & before & " to " & after & "."
            Set lastFeat = swModel.FeatureByPositionReverse(0)
            On Error Resume Next
            lastFeat.Name = "Split at potting and reading lines"
            On Error GoTo 0
            SW_SplitProject = True
            Exit Function
        ElseIf after <> before Then
            WriteLog "Split try with marks " & sm & "/" & fm & " only took the faces from " & before & " to " & after & ". Undoing it."
            SW_DeleteLastFeature
        End If
    Next k
    WriteLog "Projection split line could not be made automatically."
    SW_SplitProject = False
End Function

' Fallback after the tilt: split the stem where the potting and reading planes cut it.
Private Function SW_SplitIntersect() As Boolean
    Dim marks As Variant, k As Long, before As Long, after As Long, pm As Long, bm As Long, opt As Long
    Dim fPot As Object, fRead As Object, b As Variant, j As Long, sd As Object, hk As Long
    If TestMode Then
        Rec "SPLIT_INTERSECT"
        SW_SplitIntersect = True
        Exit Function
    End If
    Set fRead = swModel.FeatureByPositionReverse(0)
    Set fPot = swModel.FeatureByPositionReverse(1)
    marks = Array(1, 2, 4, 1, 0, 0, 1, 1, 2, 1)
    before = FaceCountAll()
    For k = 0 To UBound(marks) Step 2
        For opt = 0 To 1
            pm = marks(k): bm = marks(k + 1)
            On Error Resume Next
            swModel.ClearSelection2 True
            fPot.Select2 False, pm
            fRead.Select2 True, pm
            Set sd = swModel.SelectionManager.CreateSelectData
            sd.Mark = bm
            b = swModel.GetBodies2(0, True)
            hk = HeadIndex(b, isoCX, isoCY)
            For j = 0 To UBound(b)
                If j <> hk Then b(j).Select2 True, sd
            Next j
            swModel.FeatureManager.InsertSplitLineIntersect opt
            swModel.ClearSelection2 True
            Err.Clear
            On Error GoTo 0
            after = FaceCountAll()
            If after > before Then
                WriteLog "Split by plane intersection worked (marks " & pm & "/" & bm & ", option " & opt & "). Faces " & before & " to " & after & "."
                SW_SplitIntersect = True
                Exit Function
            End If
        Next opt
    Next k
    WriteLog "Intersection split also failed. Add it by hand with Insert > Curve > Split Line > Intersection, picking both planes and the Stem body."
    SW_SplitIntersect = False
End Function

' Tilts every body about the global Z axis through the origin. The head's centre of mass is
' checked afterwards, and the sign or selection mark is changed if SolidWorks did something else.
Private Function SW_RotateAll(ByVal angDeg As Double) As Boolean
    Dim hx As Double, hy As Double, hz As Double, ex As Double, ey As Double, qx As Double, qy As Double, qz As Double
    Dim ang As Double, attempt As Long, mk As Long
    If TestMode Then
        Rec "ROTATE," & NumStr(angDeg)
        SW_RotateAll = True
        Exit Function
    End If
    SW_RotateAll = False
    If Not NearestBodyCentre(headCX, headCY, hx, hy, hz) Then
        WriteLog "No bodies found before the tilt."
        Exit Function
    End If
    ex = hx * cosT - hy * sinT: ey = hx * sinT + hy * cosT
    For attempt = 1 To 4
        If attempt <= 2 Then ang = angDeg Else ang = -angDeg
        If attempt Mod 2 = 1 Then mk = 1 Else mk = 0
        SelectAllBodies mk
        Set lastFeat = Nothing
        On Error Resume Next
        Set lastFeat = swModel.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, 0, 0, ang * PI_ / 180, False, 1)
        On Error GoTo 0
        swModel.ClearSelection2 True
        If Not lastFeat Is Nothing Then
            NearestBodyCentre ex, ey, qx, qy, qz
            If Abs(qx - ex) < 0.01 And Abs(qy - ey) < 0.01 Then
                On Error Resume Next
                lastFeat.Name = "ISO 7206-4 tilt " & Format(angDeg, "0.00") & " deg"
                On Error GoTo 0
                WriteLog "Tilted by " & Format(angDeg, "0.000") & " deg. Head centre now at x " & Format(isoCX, "0.00") & _
                    " mm, y " & Format(isoCY, "0.00") & " mm."
                SW_RotateAll = True
                Exit Function
            End If
            WriteLog "Tilt try " & attempt & " moved the head to the wrong place. Undoing it."
            SW_DeleteLastFeature
        End If
    Next attempt
    WriteLog "Move/Copy Body could not tilt the model."
End Function

Private Sub SelectAllBodies(ByVal mk As Long)
    Dim b As Variant, k As Long, sd As Object
    swModel.ClearSelection2 True
    b = swModel.GetBodies2(0, True)
    If IsEmpty(b) Then Exit Sub
    Set sd = swModel.SelectionManager.CreateSelectData
    sd.Mark = mk
    For k = 0 To UBound(b)
        b(k).Select2 True, sd
    Next k
End Sub

Private Sub SW_PlaneParallelTop(ByVal yMM As Double, ByVal nm As String)
    Dim f As Object
    If TestMode Then
        Rec "PLANE_Y," & NumStr(yMM) & "," & nm
        Exit Sub
    End If
    On Error Resume Next
    swModel.ClearSelection2 True
    planeTop.Select2 False, 0
    swModel.FeatureManager.InsertRefPlane 8, yMM / 1000#, 0, 0, 0, 0
    Set f = swModel.FeatureByPositionReverse(0)
    If f.GetTypeName2 = "RefPlane" Then
        f.Name = nm
    Else
        WriteLog "Plane '" & nm & "' was not created."
    End If
    swModel.ClearSelection2 True
    On Error GoTo 0
End Sub

' The head is the body whose centre of mass is nearest the head centre. Every other body is stem.
Private Sub NameBodies(ByVal hx As Double, ByVal hy As Double)
    Dim b As Variant, k As Long, best As Long, bestD As Double, d As Double, v As Variant
    If TestMode Then
        Rec "NAMEBODIES"
        Exit Sub
    End If
    b = swModel.GetBodies2(0, True)
    If IsEmpty(b) Then Exit Sub
    best = -1: bestD = 1E+30
    For k = 0 To UBound(b)
        v = b(k).GetMassProperties(1#)
        d = Sqr((v(0) * 1000# - hx) ^ 2 + (v(1) * 1000# - hy) ^ 2)
        If d < bestD Then
            bestD = d: best = k
        End If
    Next k
    On Error Resume Next
    For k = 0 To UBound(b)
        If k = best Then b(k).Name = "Head" Else b(k).Name = "Stem"
    Next k
    On Error GoTo 0
    If UBound(b) > 1 Then WriteLog "WARNING the part has " & (UBound(b) + 1) & " bodies, expected 2."
End Sub

' Splits the total volume into head (body nearest (hx, hy)) and stem (everything else), in mm3.
Private Sub MeasureBodies(ByVal hx As Double, ByVal hy As Double, ByRef vStem As Double, ByRef vHead As Double)
    Dim b As Variant, k As Long, best As Long, bestD As Double, d As Double, v As Variant
    vStem = 0: vHead = 0
    If TestMode Then Exit Sub
    b = swModel.GetBodies2(0, True)
    If IsEmpty(b) Then Exit Sub
    best = -1: bestD = 1E+30
    For k = 0 To UBound(b)
        v = b(k).GetMassProperties(1#)
        d = Sqr((v(0) * 1000# - hx) ^ 2 + (v(1) * 1000# - hy) ^ 2)
        If d < bestD Then
            bestD = d: best = k
        End If
    Next k
    For k = 0 To UBound(b)
        v = b(k).GetMassProperties(1#)
        If k = best Then vHead = v(3) * 1000000000# Else vStem = vStem + v(3) * 1000000000#
    Next k
End Sub

' Total volume of all solid bodies in mm3 (the arguments are unused, kept for clarity at the call).
Private Function BodyVolume(ByVal a As Double, ByVal b2 As Double, ByVal c As Double) As Double
    Dim b As Variant, k As Long, v As Variant, s As Double
    BodyVolume = 0
    If TestMode Then Exit Function
    b = swModel.GetBodies2(0, True)
    s = 0
    If Not IsEmpty(b) Then
        For k = 0 To UBound(b)
            v = b(k).GetMassProperties(1#)
            s = s + v(3) * 1000000000#
        Next k
    End If
    BodyVolume = s
End Function

' Index in b of the body whose centre of mass is nearest (hx, hy), which is the head.
Private Function HeadIndex(b As Variant, ByVal hx As Double, ByVal hy As Double) As Long
    Dim k As Long, d As Double, bestD As Double, v As Variant
    HeadIndex = -1: bestD = 1E+30
    For k = 0 To UBound(b)
        v = b(k).GetMassProperties(1#)
        d = Sqr((v(0) * 1000# - hx) ^ 2 + (v(1) * 1000# - hy) ^ 2)
        If d < bestD Then
            bestD = d: HeadIndex = k
        End If
    Next k
End Function

Private Function NearestBodyCentre(ByVal px As Double, ByVal py As Double, _
                                   ByRef x As Double, ByRef y As Double, ByRef z As Double) As Boolean
    Dim b As Variant, k As Long, v As Variant, d As Double, bestD As Double
    NearestBodyCentre = False
    b = swModel.GetBodies2(0, True)
    If IsEmpty(b) Then Exit Function
    bestD = 1E+30
    For k = 0 To UBound(b)
        v = b(k).GetMassProperties(1#)
        d = Sqr((v(0) * 1000# - px) ^ 2 + (v(1) * 1000# - py) ^ 2)
        If d < bestD Then
            bestD = d
            x = v(0) * 1000#: y = v(1) * 1000#: z = v(2) * 1000#
            NearestBodyCentre = True
        End If
    Next k
End Function

Private Function BodyCount() As Long
    Dim b As Variant
    b = swModel.GetBodies2(0, True)
    If IsEmpty(b) Then BodyCount = 0 Else BodyCount = UBound(b) + 1
End Function

Private Function FaceCountAll() As Long
    Dim b As Variant, k As Long, s As Long
    b = swModel.GetBodies2(0, True)
    s = 0
    If Not IsEmpty(b) Then
        For k = 0 To UBound(b)
            s = s + b(k).GetFaceCount
        Next k
    End If
    FaceCountAll = s
End Function

Private Sub WriteProperties(ByVal runId As String, ByVal t As Double, ByVal n As Double, ByVal kindName As String, ByVal dUni As Double)
    Dim cpm As Object
    If TestMode Then Exit Sub
    On Error Resume Next
    Set cpm = swModel.Extension.CustomPropertyManager("")
    cpm.Add3 "Run", 30, runId, 2
    cpm.Add3 "Plate thickness (mm)", 30, NumStr(t), 2
    cpm.Add3 "Neck diameter (mm)", 30, NumStr(n), 2
    cpm.Add3 "Stem type", 30, kindName, 2
    cpm.Add3 "ISO tilt (deg)", 30, Format(isoTheta, "0.000"), 2
    cpm.Add3 "CT (mm)", 30, Format(isoCT, "0.00"), 2
    cpm.Add3 "Head centre x (mm)", 30, Format(isoCX, "0.00"), 2
    cpm.Add3 "Head centre y (mm)", 30, Format(isoCY, "0.00"), 2
    cpm.Add3 "Potting level y (mm)", 30, Format(yPot, "0.00"), 2
    cpm.Add3 "Reading line y (mm)", 30, Format(yRead, "0.00"), 2
    If dUni > 0 Then cpm.Add3 "Uniform strut diameter (mm)", 30, Format(dUni, "0.000"), 2
    On Error GoTo 0
End Sub

' A file name in the output folder that is not taken yet. An existing part is never overwritten,
' because it may already hold simulation studies. Later copies get " v2", " v3" and so on.
Private Function FreePath(ByVal fileName As String) As String
    Dim k As Long, p As String
    p = OutDir & fileName & ".SLDPRT"
    If Not TestMode Then
        k = 1
        Do While Len(Dir(p)) > 0
            k = k + 1
            p = OutDir & fileName & " v" & k & ".SLDPRT"
        Loop
    End If
    FreePath = p
End Function

Private Sub SW_SaveAndClose(ByVal path As String, ByVal keepOpen As Boolean)
    Dim errs As Long, warns As Long, ok As Boolean
    If TestMode Then
        Rec "SAVE," & path
        Exit Sub
    End If
    swModel.ForceRebuild3 False
    ok = swModel.Extension.SaveAs(path, 0, 1, Nothing, errs, warns)
    If Not ok Then Err.Raise ERR_STEM + 30, , "Saving to " & path & " failed (error code " & errs & ")."
    If keepOpen Then
        RestoreView
    Else
        swApp.CloseDoc swModel.GetTitle
    End If
End Sub

' =============================================================================================
'                                        OUTPUT AND UTILITIES
' =============================================================================================
Private Sub SetupOutput()
    Dim macroPath As String
    macroPath = swApp.GetCurrentMacroPathName
    If Len(macroPath) = 0 Then macroPath = CurDir() & "\StemBuilder.swp"
    OutDir = Left(macroPath, InStrRev(macroPath, "\")) & "Stem models\"
    On Error Resume Next
    MkDir Left(OutDir, Len(OutDir) - 1)
    On Error GoTo 0
    LogPath = OutDir & "Stem builder log.txt"
    CsvPath = OutDir & "Stem builder results.csv"
    If Len(Dir(CsvPath)) = 0 Then
        AppendLine CsvPath, "Date,Run,Plate thickness (mm),Neck (mm),Type,Stem volume (mm3),Head volume (mm3)," & _
            "Uniform strut (mm),Lattice cells,ISO tilt (deg),CT (mm),Head centre x (mm),Head centre y (mm)," & _
            "Potting level y (mm),Reading line y (mm),Warnings"
    End If
    WriteLog "===== Stem builder started."
End Sub

' Records the step now running, for the log, and shows it in the SolidWorks status bar.
Private Sub SetStep(ByVal s As String)
    stepName = s
    If TestMode Then Exit Sub
    On Error Resume Next
    swApp.Frame.SetStatusBarText "Stem builder, " & currentFile & ": " & s
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal s As String)
    If TestMode Then
        Rec "LOG," & Replace(s, ",", ";")
        Exit Sub
    End If
    AppendLine LogPath, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & s
End Sub

Private Sub WriteResult(ByVal runId As String, ByVal t As Double, ByVal n As Double, ByVal kindName As String, _
                        ByVal vStem As Double, ByVal vHead As Double, ByVal dUni As Double, ByVal warnings As String)
    Dim cells As String, uni As String
    If kindName = "solid" Then
        solidStemVolume = vStem: solidStemT = t: solidStemN = n
    End If
    If TestMode Then Exit Sub
    If kindName = "solid" Then cells = "" Else cells = CStr(nCells)
    If dUni > 0 And kindName = "uniform lattice" Then uni = Format(dUni, "0.000") Else uni = ""
    AppendLine CsvPath, Format(Now, "yyyy-mm-dd hh:nn") & "," & runId & "," & NumStr(t) & "," & NumStr(n) & "," & _
        kindName & "," & Format(vStem, "0.0") & "," & Format(vHead, "0.0") & "," & uni & "," & cells & "," & _
        Format(isoTheta, "0.000") & "," & Format(isoCT, "0.00") & "," & Format(isoCX, "0.00") & "," & _
        Format(isoCY, "0.00") & "," & Format(yPot, "0.00") & "," & Format(yRead, "0.00") & "," & Replace(warnings, ",", ";")
End Sub

Private Sub AppendLine(ByVal path As String, ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open path For Append As #f
    Print #f, s
    Close #f
End Sub

Private Function NumStr(ByVal x As Double) As String
    NumStr = Replace(CStr(Round(x, 6)), ",", ".")
End Function

Private Function P3(ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
    P3 = NumStr(x) & "," & NumStr(y) & "," & NumStr(z)
End Function

Private Function Max0(ByVal x As Double) As Double
    If x < 0 Then Max0 = 0 Else Max0 = x
End Function

Private Function ArcCos(ByVal x As Double) As Double
    If x >= 1 Then
        ArcCos = 0
    ElseIf x <= -1 Then
        ArcCos = PI_
    Else
        ArcCos = Atn(-x / Sqr(1 - x * x)) + PI_ / 2
    End If
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2 = Atn(y / x) + PI_
        Else
            Atan2 = Atn(y / x) - PI_
        End If
    ElseIf y > 0 Then
        Atan2 = PI_ / 2
    ElseIf y < 0 Then
        Atan2 = -PI_ / 2
    Else
        Atan2 = 0
    End If
End Function

Private Function WrapAngle(ByVal a As Double) As Double
    Do While a < 0
        a = a + 2 * PI_
    Loop
    Do While a >= 2 * PI_
        a = a - 2 * PI_
    Loop
    WrapAngle = a
End Function
