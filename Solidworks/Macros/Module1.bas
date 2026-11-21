Attribute VB_Name = "Module1"
Option Explicit
'=====================================================================
' SCARA_PrepareForSimscape.bas
' Restructures SR_00000_SCARA_Robot_Assy for Simscape Multibody Link
' export and assigns real printed masses (docx plate weights).
' SAFETY: full backup, works on a _SIM copy, checkpoint saves,
' per-step error isolation, throttling, low RAM, detailed log.
' Run from SolidWorks VBA with the top assembly open.
'=====================================================================

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'--------------------------- CONFIG (edit freely) --------------------
Private Const TOTAL_PRINTED_GRAMS As Double = 1337.96 ' sum of 23 plates
Private Const CLAMP_GRAMS_EACH    As Double = 7.8     ' plate 7, per piece
Private Const SUPPRESS_BELTS      As Boolean = True
Private Const SUPPRESS_FASTENERS  As Boolean = True
Private Const SUPPRESS_HIDDEN_VAR As Boolean = True
Private Const MAKE_GRIPPER_FLEX   As Boolean = True
Private Const APPLY_COTS_MASSES   As Boolean = True
Private Const DELAY_MS            As Long = 400       ' throttle per op
Private Const MASS_TOL_PCT        As Double = 2#

' numeric SolidWorks constants (late binding = compile-safe)
Private Const swDocPART As Long = 2
Private Const swDocASSEMBLY As Long = 3
Private Const swOpenSilent As Long = 1
Private Const swSaveSilent As Long = 1
Private Const swCompSuppressed As Long = 0
Private Const swThisConfig As Long = 1
Private Const swSolidBody As Long = 0
Private Const swCustomText As Long = 30

'--------------------------- GLOBALS ---------------------------------
Private gApp As Object
Private gLogPath As String
Private gWorkDir As String
Private gBackupDir As String
Private gSimPath As String
Private dName As Object, dPath As Object, dCount As Object
Private dVol As Object, dMass0 As Object, dTarget As Object
Private dActual As Object, dPrinted As Object, dPin As Object
Private gSuppBelts As Long, gSuppFast As Long, gSuppVar As Long

'=====================================================================
Public Sub SCARA_PrepareForSimscape()
    On Error GoTo MainEH
    Dim swModel As Object, swAssy As Object, fso As Object
    Dim e As Long, w As Long, ok As Boolean, tStart As Single
    tStart = Timer

    Set gApp = Application.SldWorks
    If gApp Is Nothing Then MsgBox "Start SolidWorks first.", vbExclamation: Exit Sub
    Set swModel = gApp.ActiveDoc
    If swModel Is Nothing Then MsgBox "Open the SCARA top assembly first.", vbExclamation: Exit Sub
    If swModel.GetType <> swDocASSEMBLY Then MsgBox "Active document is not an assembly.", vbExclamation: Exit Sub
    If swModel.GetPathName = "" Then MsgBox "Save the assembly once first.", vbExclamation: Exit Sub
    Set swAssy = swModel
    Set fso = CreateObject("Scripting.FileSystemObject")

    '--- work dirs + log
    Dim baseDir As String: baseDir = fso.GetParentFolderName(swModel.GetPathName)
    gWorkDir = baseDir & "\SIM_PREP_" & Format(Now, "yyyymmdd_hhnnss")
    fso.CreateFolder gWorkDir
    gBackupDir = gWorkDir & "\Backup"
    fso.CreateFolder gBackupDir
    gLogPath = gWorkDir & "\SCARA_SimPrep_Log.txt"
    LogMsg "=== SCARA Simscape preparation started ==="
    LogMsg "Source assembly: " & swModel.GetPathName

    InitDicts

    '--- performance / UI throttling (all reversible, all guarded)
    On Error Resume Next
    gApp.CommandInProgress = True
    swModel.FeatureManager.EnableFeatureTree = False
    swModel.ActiveView.EnableGraphicsUpdate = False
    On Error GoTo MainEH

    '--- PHASE 1: backup every referenced file (CRITICAL)
    LogMsg "PHASE 1: backup"
    If Not BackupAll(swAssy, fso) Then
        LogMsg "FATAL: backup incomplete - nothing was modified."
        GoTo SafeExit
    End If
    SafeWait 1000

    '--- PHASE 2: create _SIM copy (original file stays untouched)
    LogMsg "PHASE 2: create SIM copy"
    If InStr(1, swModel.GetPathName, "_SIM", vbTextCompare) > 0 Then
        gSimPath = swModel.GetPathName
        LogMsg "Already a _SIM file - working in place."
    Else
        gSimPath = fso.GetParentFolderName(swModel.GetPathName) & "\" & _
                   fso.GetBaseName(swModel.GetPathName) & "_SIM.sldasm"
        If fso.FileExists(gSimPath) Then gSimPath = Left(gSimPath, Len(gSimPath) - 7) & _
                   "_" & Format(Now, "hhnnss") & ".sldasm"
        ok = swModel.Extension.SaveAs(gSimPath, 0, swSaveSilent, Nothing, e, w)
        SafeWait 1500
        If Not ok Or Not fso.FileExists(gSimPath) Then
            LogMsg "FATAL: SaveAs SIM copy failed (err=" & e & ") - nothing modified."
            GoTo SafeExit
        End If
        LogMsg "SIM assembly: " & gSimPath
    End If
    Set swModel = gApp.ActiveDoc
    Set swAssy = swModel

    '--- PHASE 3: resolve + restructure (suppress / flexible)
    LogMsg "PHASE 3: restructure"
    On Error Resume Next
    swAssy.ResolveAllLightWeightComponents True
    swAssy.EditAssembly
    On Error GoTo MainEH
    SafeWait 800
    RestructureAssembly swAssy
    SafeWait 800
    swModel.ForceRebuild3 False
    SafeWait 1200
    ok = swModel.Save3(swSaveSilent, e, w)
    LogMsg "Checkpoint save after restructure: " & ok
    SafeWait 1000

    '--- collect inventory BEFORE closing (counts of visible parts)
    CollectInventory swAssy
    LogMsg "Unique active parts: " & dName.Count

    '--- close assembly to keep RAM low during part processing
    Dim assyTitle As String: assyTitle = swModel.GetTitle
    On Error Resume Next
    swModel.ActiveView.EnableGraphicsUpdate = True
    swModel.FeatureManager.EnableFeatureTree = True
    On Error GoTo MainEH
    gApp.CloseDoc assyTitle
    Set swModel = Nothing: Set swAssy = Nothing
    SafeWait 1500

    '--- PHASE 4: volumes (one part at a time)
    LogMsg "PHASE 4: read volumes"
    ReadAllVolumes
    SafeWait 1000

    '--- PHASE 5: targets + material library
    LogMsg "PHASE 5: compute mass targets"
    LoadPins baseDir
    ComputeTargets
    Dim libPath As String
    libPath = gWorkDir & "\SCARA_SIM_Materials_" & Format(Now, "hhnnss") & ".sldmat"
    If Not BuildMaterialLib(libPath) Then
        LogMsg "FATAL: material library write failed."
        GoTo ReopenAndExit
    End If
    WriteTargetsCSV baseDir

    '--- PHASE 6: apply materials + verify + save each part
    LogMsg "PHASE 6: apply masses"
    ApplyAllMaterials libPath

    '--- PHASE 7: reopen SIM assembly, final save, report
ReopenAndExit:
    LogMsg "PHASE 7: finalize"
    Set swModel = gApp.OpenDoc6(gSimPath, swDocASSEMBLY, swOpenSilent, "", e, w)
    SafeWait 1500
    If Not swModel Is Nothing Then
        swModel.ForceRebuild3 False
        SafeWait 1200
        ok = swModel.Save3(swSaveSilent, e, w)
        LogMsg "Final SIM save: " & ok
    End If
    WriteReport baseDir
    WriteReadme baseDir

SafeExit:
    On Error Resume Next
    gApp.CommandInProgress = False
    If Not swModel Is Nothing Then
        swModel.ActiveView.EnableGraphicsUpdate = True
        swModel.FeatureManager.EnableFeatureTree = True
    End If
    On Error GoTo 0
    Dim mins As Double: mins = (Timer - tStart) / 60#
    LogMsg "=== Finished in " & Format(mins, "0.0") & " min ==="
    MsgBox "SCARA Simscape preparation finished." & vbCrLf & _
           "SIM file : " & gSimPath & vbCrLf & _
           "Log/CSV  : " & gWorkDir & vbCrLf & vbCrLf & _
           "Next: Tools > Simscape Multibody Link > Export." & vbCrLf & _
           "Edit SCARA_Mass_Targets.csv (Pinned=1) and re-run to pin exact part weights.", _
           vbInformation, "SCARA SimPrep"
    Exit Sub

MainEH:
    LogMsg "ERROR " & Err.Number & ": " & Err.Description & " (macro stopped safely; backups in " & gBackupDir & ")"
    On Error Resume Next
    gApp.CommandInProgress = False
    If Not swModel Is Nothing Then
        swModel.ActiveView.EnableGraphicsUpdate = True
        swModel.FeatureManager.EnableFeatureTree = True
        swModel.Save3 swSaveSilent, e, w
    End If
    MsgBox "Macro stopped safely at a logged step." & vbCrLf & _
           "Nothing is corrupted. See log: " & gLogPath, vbExclamation
End Sub

'=====================================================================
Private Sub InitDicts()
    Set dName = CreateObject("Scripting.Dictionary")
    Set dPath = CreateObject("Scripting.Dictionary")
    Set dCount = CreateObject("Scripting.Dictionary")
    Set dVol = CreateObject("Scripting.Dictionary")
    Set dMass0 = CreateObject("Scripting.Dictionary")
    Set dTarget = CreateObject("Scripting.Dictionary")
    Set dActual = CreateObject("Scripting.Dictionary")
    Set dPrinted = CreateObject("Scripting.Dictionary")
    Set dPin = CreateObject("Scripting.Dictionary")
End Sub

Private Sub LogMsg(ByVal s As String)
    On Error Resume Next
    Debug.Print s
    If gLogPath = "" Then Exit Sub
    Dim f As Integer: f = FreeFile
    Open gLogPath For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & " | " & s
    Close #f
End Sub

Private Sub SafeWait(ByVal ms As Long)
    If ms > 60000 Then ms = 60000
    Dim t0 As Single, t As Single
    t0 = Timer
    Do
        Sleep 25
        DoEvents
        t = Timer
        If t < t0 Then t = t + 86400
    Loop While (t - t0) * 1000# < ms
End Sub

Private Function CompState(comp As Object) As Long
    On Error Resume Next
    Err.Clear
    CompState = CallByName(comp, "GetSuppression2", VbMethod)
    If Err.Number <> 0 Then
        Err.Clear
        CompState = CallByName(comp, "GetSuppression", VbMethod)
        If Err.Number <> 0 Then CompState = 2
    End If
End Function

'--------------------------- PHASE 1 ---------------------------------
Private Function BackupAll(swAssy As Object, fso As Object) As Boolean
    On Error GoTo fail
    Dim vComps As Variant, i As Long, comp As Object, p As String
    BackupAll = False
    ' backup both assembly files
    Dim rootPath As String: rootPath = swAssy.GetPathName
    If Not BackupOne(fso, rootPath) Then Exit Function
    vComps = swAssy.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function
    For i = LBound(vComps) To UBound(vComps)
        Set comp = vComps(i)
        If Not comp Is Nothing Then
            p = comp.GetPathName
            If p <> "" Then
                If Not BackupOne(fso, p) Then
                    LogMsg "Backup failed: " & p
                    Exit Function
                End If
            End If
        End If
        If i Mod 10 = 0 Then SafeWait 100
    Next i
    BackupAll = True
    LogMsg "Backup OK -> " & gBackupDir
    Exit Function
fail:
    LogMsg "Backup error: " & Err.Description
    BackupAll = False
End Function

Private Function BackupOne(fso As Object, ByVal src As String) As Boolean
    On Error GoTo fail
    If Not fso.FileExists(src) Then BackupOne = True: Exit Function
    Dim dst As String: dst = gBackupDir & "\" & fso.GetFileName(src)
    If fso.FileExists(dst) Then BackupOne = True: Exit Function
    fso.CopyFile src, dst, True
    BackupOne = fso.FileExists(dst)
    Exit Function
fail:
    BackupOne = False
End Function

'--------------------------- PHASE 3 ---------------------------------
Private Sub RestructureAssembly(swAssy As Object)
    On Error Resume Next
    Dim vComps As Variant, i As Long, comp As Object, md As Object
    Dim nm As String, gripComp As Object, gripDocDirty As Boolean
    vComps = swAssy.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub
    For i = LBound(vComps) To UBound(vComps)
        Set comp = vComps(i)
        If Not comp Is Nothing Then
            If CompState(comp) <> swCompSuppressed Then
                Set md = comp.GetModelDoc2
                If Not md Is Nothing Then
                    nm = md.GetTitle
                    Dim doSupp As Boolean: doSupp = False
                    If SUPPRESS_BELTS And Left(nm, 9) = "COTS_BLT_" Then doSupp = True: gSuppBelts = gSuppBelts + 1
                    If SUPPRESS_FASTENERS And Left(nm, 9) = "COTS_FST_" Then doSupp = True: gSuppFast = gSuppFast + 1
                    If SUPPRESS_HIDDEN_VAR Then
                        If nm = "SR_50_010_Wide_Gripper_End" Or nm = "SR_50_011_Wide_Gripper_End_Left" Then
                            doSupp = True: gSuppVar = gSuppVar + 1: gripDocDirty = True
                        End If
                    End If
                    If nm = "SR_50000_Gripper_Mechanism_Assy" Then Set gripComp = comp
                    If doSupp Then
                        Err.Clear
                        comp.SetSuppression2 swCompSuppressed, swThisConfig, Empty
                        If Err.Number <> 0 Then LogMsg "  suppress failed: " & nm & " (" & Err.Description & ")"
                        SafeWait 150
                    End If
                End If
            End If
        End If
    Next i
    LogMsg "Suppressed: belts=" & gSuppBelts & " fasteners=" & gSuppFast & " hidden variants=" & gSuppVar
    ' save gripper sub-assembly if its children changed
    If gripDocDirty And Not gripComp Is Nothing Then
        Dim gDoc As Object: Set gDoc = gripComp.GetModelDoc2
        If Not gDoc Is Nothing Then
            Dim ge As Long, gw As Long
            gDoc.Save3 swSaveSilent, ge, gw
            SafeWait 800
        End If
    End If
    ' flexible gripper (try several API member names, late bound)
    If MAKE_GRIPPER_FLEX And Not gripComp Is Nothing Then
        Err.Clear
        If TrySetFlexible(gripComp, True) Then
            SafeWait 500
            Dim rb As Boolean: rb = swAssy.ForceRebuild3(False)
            SafeWait 1000
            If rb Then
                LogMsg "Gripper set FLEXIBLE (internal joints will export)."
            Else
                TrySetFlexible gripComp, False
                swAssy.ForceRebuild3 False
                LogMsg "WARN: flexible solve failed - reverted to rigid (set manually if needed)."
            End If
        Else
            LogMsg "WARN: flexible API not available - set gripper flexible manually."
        End If
    End If
    On Error GoTo 0
End Sub

Private Function TrySetFlexible(comp As Object, ByVal onOff As Boolean) As Boolean
    Dim v As Long: v = IIf(onOff, 1, 0)
    On Error Resume Next
    Err.Clear: CallByName comp, "SolveAs", VbLet, v
    If Err.Number = 0 Then TrySetFlexible = True: Exit Function
    Err.Clear: CallByName comp, "IsFlexible", VbLet, onOff
    If Err.Number = 0 Then TrySetFlexible = True: Exit Function
    Err.Clear: CallByName comp, "SetFlexible", VbMethod, onOff
    If Err.Number = 0 Then TrySetFlexible = True: Exit Function
    Err.Clear: CallByName comp, "Flexible", VbLet, onOff
    If Err.Number = 0 Then TrySetFlexible = True: Exit Function
    TrySetFlexible = False
End Function

'--------------------------- INVENTORY -------------------------------
Private Sub CollectInventory(swAssy As Object)
    On Error Resume Next
    Dim vComps As Variant, i As Long, comp As Object, md As Object
    Dim nm As String, k As String
    vComps = swAssy.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub
    For i = LBound(vComps) To UBound(vComps)
        Set comp = vComps(i)
        If Not comp Is Nothing Then
            If CompState(comp) <> swCompSuppressed Then
                Set md = comp.GetModelDoc2
                If Not md Is Nothing Then
                    If md.GetType = swDocPART Then
                        nm = md.GetTitle
                        k = LCase(nm)
                        If Not dName.Exists(k) Then
                            dName.Add k, nm
                            dPath.Add k, comp.GetPathName
                            dCount.Add k, 0
                            dVol.Add k, 0#
                            dMass0.Add k, 0#
                            dTarget.Add k, 0#
                            dActual.Add k, 0#
                            Dim isPr As Boolean: isPr = False
                            If Left(nm, 3) = "SR_" Then isPr = True
                            If nm = "COTS_LIN_Rod_Clamp_D10" Then isPr = True
                            dPrinted.Add k, isPr
                        End If
                        dCount(k) = dCount(k) + 1
                    End If
                End If
            End If
        End If
    Next i
End Sub

'--------------------------- PHASE 4 ---------------------------------
Private Sub ReadAllVolumes()
    Dim k As Variant, v As Double, m As Double
    For Each k In dName.Keys
        v = 0#: m = 0#
        If ReadPartVolMass(CStr(dPath(k)), v, m) Then
            dVol(k) = v: dMass0(k) = m: dActual(k) = m
        Else
            LogMsg "  WARN: no volume for " & dName(k)
        End If
        SafeWait DELAY_MS
    Next k
End Sub

Private Function ReadPartVolMass(ByVal path As String, ByRef volM3 As Double, ByRef massKg As Double) As Boolean
    Dim pDoc As Object, mp As Object, vB As Variant, e As Long, w As Long
    On Error GoTo fail
    Set pDoc = gApp.OpenDoc6(path, swDocPART, swOpenSilent, "", e, w)
    If pDoc Is Nothing Then GoTo fail
    SafeWait DELAY_MS
    Set mp = pDoc.Extension.CreateMassProperty
    vB = pDoc.GetBodies2(swSolidBody, True)
    If IsEmpty(vB) Then GoTo cleanup
    mp.AddBodies vB
    volM3 = mp.Volume: massKg = mp.Mass
    ReadPartVolMass = (volM3 > 0#)
cleanup:
    On Error Resume Next
    gApp.CloseDoc pDoc.GetTitle
    Set mp = Nothing: Set pDoc = Nothing
    SafeWait DELAY_MS
    Exit Function
fail:
    LogMsg "  WARN volume read: " & path & " (" & Err.Description & ")"
    ReadPartVolMass = False
    Resume cleanup
End Function

'--------------------------- PHASE 5 ---------------------------------
Private Sub LoadPins(ByVal baseDir As String)
    On Error Resume Next
    ' built-in realistic COTS masses (grams per instance) - editable
    If APPLY_COTS_MASSES Then
        dPin("cots_mot_nema17_20mm_shaft") = 280#
        dPin("cots_mot_nema17_l24_20mm_shaft") = 130#
        dPin("cots_srv_mg996r") = 55#
        dPin("cots_srv_mg996r_horn") = 2#
        dPin("cots_sw_micro_limit_switch") = 5#
        dPin("cots_brg_608_8x22x7") = 11.5
        dPin("cots_brg_6806_30x42x7") = 25#
        dPin("cots_brg_6807_35x47x7") = 28#
        dPin("cots_brg_51107_35x52x12") = 80#
        dPin("cots_brg_51108_40x60x13") = 110#
        dPin("cots_lin_lm10uu_10x19x29") = 24#
        dPin("cots_lin_lead_screw_nut_d8") = 8#
    End If
    dPin("cots_lin_rod_clamp_d10") = CLAMP_GRAMS_EACH
    ' user CSV overrides (ModelName,Grams,Pinned)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim csv As String: csv = baseDir & "\SCARA_Mass_Targets.csv"
    If Not fso.FileExists(csv) Then Exit Sub
    Dim f As Integer: f = FreeFile
    Dim ln As String, parts() As String, k As String
    Open csv For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If InStr(ln, ",") > 0 Then
            parts = Split(ln, ",")
            If UBound(parts) >= 2 Then
                k = LCase(Trim(parts(0)))
                If dName.Exists(k) And Trim(parts(2)) = "1" Then
                    dPin(k) = Val(Trim(parts(1)))
                End If
            End If
        End If
    Loop
    Close #f
    LogMsg "User CSV pins loaded."
End Sub

Private Sub ComputeTargets()
    Dim k As Variant, pinnedSum As Double, volSum As Double
    pinnedSum = 0#: volSum = 0#
    For Each k In dName.Keys
        If dPrinted(k) Then
            If dPin.Exists(k) Then
                pinnedSum = pinnedSum + dPin(k) * dCount(k)
            Else
                volSum = volSum + dVol(k) * dCount(k)
            End If
        End If
    Next k
    Dim remainG As Double: remainG = TOTAL_PRINTED_GRAMS - pinnedSum
    Dim rho As Double
    If volSum > 0# And remainG > 0# Then
        rho = (remainG / 1000#) / volSum          ' kg/m3
    Else
        rho = 650#
        LogMsg "WARN: pool exhausted or zero volume - fallback density 650 kg/m3."
    End If
    LogMsg "Printed pool: total=" & TOTAL_PRINTED_GRAMS & "g pinned=" & Format(pinnedSum, "0.0") & _
           "g distributed=" & Format(remainG, "0.0") & "g eff.density=" & Format(rho, "0.0") & " kg/m3"
    For Each k In dName.Keys
        If dPin.Exists(k) Then
            dTarget(k) = dPin(k)
        ElseIf dPrinted(k) And dVol(k) > 0# Then
            dTarget(k) = rho * dVol(k) * 1000#    ' grams per instance
        End If
    Next k
End Sub

Private Function BuildMaterialLib(ByVal libPath As String) As Boolean
    On Error GoTo fail
    Dim s As String, k As Variant, mid As Long, rho As Double
    s = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    s = s & "<mstns:materials xmlns:mstns=""http://www.solidworks.com/sldmaterials"">" & vbCrLf
    s = s & "<classification name=""SCARA_SIM"">" & vbCrLf
    mid = 1000
    For Each k In dName.Keys
        If dTarget(k) > 0# And dVol(k) > 0# Then
            rho = (dTarget(k) / 1000#) / dVol(k)
            mid = mid + 1
            s = s & "<material name=""SIM_" & dName(k) & """ matid=""" & mid & """>" & vbCrLf
            s = s & "<physicalproperties>" & vbCrLf
            s = s & "<density value=""" & Trim(Str(Round(rho, 4))) & """ useproperty=""1""/>" & vbCrLf
            s = s & "<elastic_modulus value=""3500000000"" useproperty=""1""/>" & vbCrLf
            s = s & "<poissons_ratio value=""0.36"" useproperty=""1""/>" & vbCrLf
            s = s & "</physicalproperties>" & vbCrLf
            s = s & "</material>" & vbCrLf
        End If
    Next k
    s = s & "</classification>" & vbCrLf & "</mstns:materials>" & vbCrLf
    Dim f As Integer: f = FreeFile
    Open libPath For Output As #f
    Print #f, s
    Close #f
    BuildMaterialLib = True
    LogMsg "Material library: " & libPath
    Exit Function
fail:
    BuildMaterialLib = False
End Function

Private Sub WriteTargetsCSV(ByVal baseDir As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile, k As Variant, pinFlag As String
    Open baseDir & "\SCARA_Mass_Targets.csv" For Output As #f
    Print #f, "ModelName,TargetGramsPerInstance,Pinned"
    For Each k In dName.Keys
        pinFlag = "0"
        If dPin.Exists(k) Then pinFlag = "1"
        If dTarget(k) > 0# Then
            Print #f, dName(k) & "," & Trim(Str(Round(dTarget(k), 3))) & "," & pinFlag
        Else
            Print #f, dName(k) & "," & Trim(Str(Round(dMass0(k) * 1000#, 3))) & ",0"
        End If
    Next k
    Close #f
End Sub

'--------------------------- PHASE 6 ---------------------------------
Private Sub ApplyAllMaterials(ByVal libPath As String)
    Dim k As Variant
    For Each k In dName.Keys
        If dTarget(k) > 0# And dVol(k) > 0# Then
            ApplyOne CStr(dPath(k)), "SIM_" & dName(k), libPath, CDbl(dTarget(k)), k
            SafeWait DELAY_MS
        End If
    Next k
End Sub

Private Sub ApplyOne(ByVal path As String, ByVal matName As String, ByVal libPath As String, _
                     ByVal targetG As Double, ByVal k As Variant)
    Dim pDoc As Object, mp As Object, vB As Variant, e As Long, w As Long
    Dim m As Double, dev As Double, lbl As String
    On Error GoTo fail
    Set pDoc = gApp.OpenDoc6(path, swDocPART, swOpenSilent, "", e, w)
    If pDoc Is Nothing Then LogMsg "  WARN open failed: " & path: Exit Sub
    SafeWait DELAY_MS
    Err.Clear
    pDoc.SetMaterialPropertyName2 "", libPath, matName
    If Err.Number <> 0 Then LogMsg "  WARN material set: " & matName & " (" & Err.Description & ")"
    SafeWait 200
    pDoc.ForceRebuild3 False
    SafeWait DELAY_MS
    Set mp = pDoc.Extension.CreateMassProperty
    vB = pDoc.GetBodies2(swSolidBody, True)
    If Not IsEmpty(vB) Then
        mp.AddBodies vB
        m = mp.Mass
        dActual(k) = m
        If targetG > 0# Then
            dev = Abs(m * 1000# - targetG) / targetG * 100#
            If dev > MASS_TOL_PCT Then
                LogMsg "  WARN mass mismatch " & dName(k) & ": got " & Format(m * 1000#, "0.00") & _
                       "g want " & Format(targetG, "0.00") & "g (" & Format(dev, "0.0") & "%)"
            End If
        End If
    End If
    lbl = "PLA+ printed (mass-matched)"
    If Not dPrinted(k) Then lbl = "COTS realistic (mass-matched)"
    On Error Resume Next
    Dim cpm As Object: Set cpm = pDoc.Extension.CustomPropertyManager("")
    cpm.Set2 "Material", lbl
    If Err.Number <> 0 Then Err.Clear: cpm.Add3 "Material", swCustomText, lbl, 0
    On Error GoTo fail
    pDoc.Save3 swSaveSilent, e, w
    SafeWait 300
cleanup:
    On Error Resume Next
    gApp.CloseDoc pDoc.GetTitle
    Set mp = Nothing: Set pDoc = Nothing
    SafeWait DELAY_MS
    Exit Sub
fail:
    LogMsg "  WARN apply failed: " & path & " (" & Err.Description & ")"
    Resume cleanup
End Sub

'--------------------------- PHASE 7 ---------------------------------
Private Sub WriteReport(ByVal baseDir As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile, k As Variant
    Dim totAll As Double, totPrint As Double, actG As Double
    Open gWorkDir & "\SCARA_MassReport.csv" For Output As #f
    Print #f, "ModelName,Instances,Printed,TargetGrams,ActualGrams,TotalGrams"
    For Each k In dName.Keys
        actG = dActual(k) * 1000#
        totAll = totAll + actG * dCount(k)
        If dPrinted(k) Then totPrint = totPrint + actG * dCount(k)
        Print #f, dName(k) & "," & dCount(k) & "," & dPrinted(k) & "," & _
                  Trim(Str(Round(dTarget(k), 3))) & "," & Trim(Str(Round(actG, 3))) & "," & _
                  Trim(Str(Round(actG * dCount(k), 3)))
    Next k
    Close #f
    LogMsg "Printed total = " & Format(totPrint, "0.0") & " g (target " & TOTAL_PRINTED_GRAMS & " g)"
    LogMsg "Exported robot total = " & Format(totAll / 1000#, "0.000") & " kg (suppressed items excluded)"
End Sub

Private Sub WriteReadme(ByVal baseDir As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open gWorkDir & "\MATLAB_NOTES.txt" For Output As #f
    Print #f, "After Simscape Multibody Link export, add in Simulink:"
    Print #f, "1) Belt constraints (belts were suppressed on purpose):"
    Print #f, "   J1: motor 20T -> idler 80T/25T -> driven 110T  => ratio 17.6 : 1"
    Print #f, "   J2: motor 20T -> idler 80T/26T -> driven 92T   => ratio 14.15 : 1"
    Print #f, "   J3: motor 20T -> driven 90T                    => ratio 4.5 : 1"
    Print #f, "2) Lead screw constraint between Z-screw revolute and Z-platform prismatic"
    Print #f, "   (use your screw lead, e.g. T8 lead = 8 mm/rev)."
    Print #f, "3) Joint position limits at the micro-switch positions."
    Print #f, "4) Fasteners were suppressed; their mass can be lumped into links if needed."
    Print #f, "5) Rods/lead screw are modeled 500 mm although named L400/L380 (verify)."
    Close #f
End Sub
