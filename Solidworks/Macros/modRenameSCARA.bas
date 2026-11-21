Attribute VB_Name = "modRenameSCARA"
Option Explicit
' ============================================================================
'  SCARA_Robot  ->  MATLAB / Simscape preparation macro  (SCARA edition v1)
'  Engine: proven MODUMILL v3 low-memory "one file at a time" flow.
'  Source audit: SCARA_Robot_DeepAudit.md / .json (SOLIDWORKS 34.1.1, 2026-11-20)
'  Project folder: C:\Users\User\Desktop\SR   (61 documents, 112 instances)
'
'  SINGLE LIVE RUN - no dry run, no backup copies (per request).
'    Phase 1: open each doc ALONE (children first, Base first) ->
'             rename configuration to "Default" ->
'             VERIFY material (KEEP if already correct, SET only if wrong/missing) ->
'             write 6 custom properties -> SAVE -> CLOSE -> pause.
'    Phase 2: per-assembly windows (deepest first), lightweight open ->
'             repair child config references -> rename child files
'             (official RenameDocument + SaveAs fallback) -> SaveReferenced ->
'             close all -> delete leftover old files -> rename top file.
'    Phase 3: reopen assemblies, verify all references, write CSV report.
'
'  NAMING: PartNumber + PartName, characters [A-Z a-z 0-9 _] ONLY.
'          No spaces, hyphens, plus, slash or math symbols.
'  MATERIAL: read back first; correct material is NEVER reassigned.
'          Custom "PLA+" library auto-detected from first PLA+ part (Base).
'  RESUME: if interrupted, run again - finished work is detected and skipped.
'  Open documents are auto-saved and closed before the run starts.
' ============================================================================

#If VBA7 Then
Private Declare PtrSafe Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' ------------------------------ USER SETTINGS ------------------------------
Public Const PROJ_DIR As String = "C:\Users\User\Desktop\SR"
Public Const TOP_ASM As String = "C:\Users\User\Desktop\SR\SCARA_Robot.SLDASM"
Public Const TOP_NEW As String = "C:\Users\User\Desktop\SR\SR_00000_SCARA_Robot_Assy.SLDASM"
Public Const MAT_LIB As String = "solidworks materials.sldmat"
Public Const OPEN_LIGHTWEIGHT As Boolean = True
Public Const PAUSE_MS As Long = 250
Public Const CLEAN_LEFTOVER_OLD As Boolean = True
Public Const PROJECT_NAME As String = "SCARA_Robot"
Public Const REV As String = "A"
Public Const NEW_CFG As String = "Default"
Public Const LOG_PATH As String = "C:\Users\User\Desktop\SR\_SW_Rename_SCARA_Log.txt"
Public Const REPORT_CSV As String = "C:\Users\User\Desktop\SR\_SCARA_Component_Report.csv"

' ---------------------------- EXPECTED COUNTS ------------------------------
Private Const EXP_FILES As Long = 61
Private Const EXP_RENAMES As Long = 61
Private Const EXP_CFG As Long = 61
Private Const EXP_MAT_SET As Long = 3        ' Micro_switch + 2 Arduino cases
Private Const EXP_MAT_KEEP As Long = 56      ' parts already correct (verify only)
Private Const EXP_INSTANCES As Long = 112

' ------------------------------ TYPES / STATE ------------------------------
Private Type FileRec
    Id As Long
    Action As String
    DocType As String
    Depth As Long
    OldPath As String
    NewPath As String
    MatAction As String          ' CHECK / NONE
    MatTarget As String
    NeedsReview As Boolean
    PropNames(1 To 6) As String
    PropValues(1 To 6) As String
    Renamed As Boolean
    UsedFallback As Boolean
End Type

Private Type CfgRec
    FileId As Long
    OldCfg As String
    NewCfg As String
    Done As Boolean
    Skipped As Boolean
End Type

Private Files() As FileRec
Private Cfgs() As CfgRec
Private OrderIds() As Long
Private Wins() As Long
Private PairPar() As Long
Private PairChild() As Long
Private gNFiles As Long, gNCfgs As Long
Private gNOrder As Long, gNWins As Long, gNPairs As Long

Private swApp As SldWorks.SldWorks
Private mIdByOld As Object
Private mIdByNew As Object
Private mCfgMap As Object
Private gMatLibs As Object       ' lcase(material) -> library file (auto-detected)
Private gWinOk As Object
Private gRenamedInWindow As Collection
Private gTmpKills As Collection

Private gLogNum As Integer
Private gTopFid As Long
Private gVerifyCount As Long

Private nAutoClosed As Long
Private nCfgRenamed As Long, nCfgSkipped As Long
Private nMatVerified As Long, nMatSet As Long, nMatFail As Long
Private nPropSet As Long
Private nRefRepaired As Long
Private nRenamed As Long, nRenameFail As Long
Private nMismatch As Long, nOldLeft As Long, nStubsCleaned As Long
Private nDocsEdited As Long, nWindowsDone As Long
Private nWarnings As Long

' --------------------------------- MAIN ------------------------------------
Public Sub MainSCARA()
    On Error GoTo MainErr

    ResetCounters
    OpenLog
    LogLine "=============================================================="
    LogLine "SCARA_Robot rename + material verify/set + properties (live run, one file at a time)"
    LogLine "Started: " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    LoadData
    LogLine "Plan: " & EXP_RENAMES & " file renames, " & EXP_CFG & " config renames -> '" & NEW_CFG & _
            "', " & EXP_MAT_SET & " materials to assign, " & EXP_MAT_KEEP & " materials verify-only."

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        LogLine "FATAL: cannot get the SOLIDWORKS application object."
        GoTo MainEnd
    End If

    AutoCloseOpenDocs          ' save + close anything open, so the run starts clean

    InitLookups
    LogLine "Execution: " & gNOrder & " documents one-by-one (phase 1), " & gNWins & _
            " rename windows (phase 2), pause " & PAUSE_MS & " ms, lightweight=" & OPEN_LIGHTWEIGHT

    Dim i As Long
    For i = 1 To gNFiles
        If Files(i).NeedsReview Then
            LogLine "  REVIEW (inferred - verify): " & StemOf(Files(i).OldPath) & "  ->  " & StemOf(Files(i).NewPath)
        End If
    Next i

    Phase1_EditEachDocument
    Phase2_RenameWindows
    RenameTopFile
    Phase3_VerifyAndReport
    WriteSummary
    GoTo MainEnd

MainErr:
    LogLine "FATAL ERROR " & Err.Number & ": " & Err.Description
    MsgBox "Macro stopped with error " & Err.Number & ":" & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Everything finished so far is SAVED on disk." & vbCrLf & _
           "Run the macro again - it resumes where it stopped." & vbCrLf & _
           "Log: " & LOG_PATH, vbCritical, "SCARA prep"
MainEnd:
    CloseAllDocs
    LogLine "Finished: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    CloseLog
End Sub

Private Sub ResetCounters()
    gVerifyCount = 0
    nAutoClosed = 0
    nCfgRenamed = 0: nCfgSkipped = 0
    nMatVerified = 0: nMatSet = 0: nMatFail = 0
    nPropSet = 0: nRefRepaired = 0
    nRenamed = 0: nRenameFail = 0
    nMismatch = 0: nOldLeft = 0: nStubsCleaned = 0
    nDocsEdited = 0: nWindowsDone = 0: nWarnings = 0
End Sub

Private Sub InitLookups()
    Dim i As Long, j As Long, d As Object
    Set mIdByOld = CreateObject("Scripting.Dictionary"): mIdByOld.CompareMode = 1
    Set mIdByNew = CreateObject("Scripting.Dictionary"): mIdByNew.CompareMode = 1
    Set mCfgMap = CreateObject("Scripting.Dictionary"): mCfgMap.CompareMode = 1
    Set gMatLibs = CreateObject("Scripting.Dictionary"): gMatLibs.CompareMode = 1
    Set gWinOk = CreateObject("Scripting.Dictionary"): gWinOk.CompareMode = 1
    For i = 1 To gNFiles
        mIdByOld.Item(LCase(Files(i).OldPath)) = i
        If LCase(Files(i).NewPath) <> LCase(Files(i).OldPath) Then
            mIdByNew.Item(LCase(Files(i).NewPath)) = i
        End If
    Next i
    For j = 1 To gNCfgs
        i = Cfgs(j).FileId
        If Not mCfgMap.Exists(i) Then
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = 1
            mCfgMap.Add i, d
        End If
        mCfgMap.Item(i).Item(LCase(Cfgs(j).OldCfg)) = Cfgs(j).NewCfg
    Next j
    gTopFid = 0
    If mIdByOld.Exists(LCase(TOP_ASM)) Then gTopFid = CLng(mIdByOld.Item(LCase(TOP_ASM)))
End Sub

Private Sub AutoCloseOpenDocs()
    Dim vDocs As Variant, m As Variant
    Dim md As SldWorks.ModelDoc2
    Dim e As Long, w As Long, st As Boolean
    On Error Resume Next
    vDocs = swApp.GetDocuments
    If IsArray(vDocs) Then
        For Each m In vDocs
            Set md = m
            If Not md Is Nothing Then
                If Len(md.GetPathName) > 0 Then
                    st = md.Save3(swSaveAsOptions_Silent, e, w)
                    If st Then
                        LogLine "AUTO: saved open document before run: " & md.GetPathName
                    Else
                        LogWarn "AUTO: could not save open document (closed anyway): " & md.GetPathName
                    End If
                End If
                swApp.CloseDoc md.GetTitle
                nAutoClosed = nAutoClosed + 1
            End If
            Set md = Nothing
        Next m
    End If
    Err.Clear
    On Error GoTo 0
    If nAutoClosed > 0 Then
        LogLine "AUTO: " & nAutoClosed & " open document(s) saved and closed."
        Pace
    End If
End Sub

' ------------------------------- PHASE 1 ------------------------------------
Private Sub Phase1_EditEachDocument()
    LogLine "--- Phase 1: open -> config rename -> material verify/set -> props -> save -> close ---"
    Dim k As Long
    For k = 1 To gNOrder
        EditOneDocument OrderIds(k), k
        CloseAllDocs
        Pace
    Next k
    LogLine "Phase 1 done: docs=" & nDocsEdited & "/" & gNOrder & _
            ", cfg renames=" & nCfgRenamed & " (skipped " & nCfgSkipped & ")" & _
            ", materials verified(kept)=" & nMatVerified & ", set=" & nMatSet & " (fail " & nMatFail & ")" & _
            ", props=" & nPropSet
End Sub

Private Sub EditOneDocument(ByVal fid As Long, ByVal k As Long)
    Dim eff As String, doc As SldWorks.ModelDoc2, tag As String
    On Error GoTo EH
    tag = "[" & k & "/" & gNOrder & "] "
    eff = EffectivePath(Files(fid).OldPath, Files(fid).NewPath)
    If Len(eff) = 0 Then
        LogWarn tag & "file not found on disk (old and new missing) - skipped: " & Files(fid).OldPath
        Exit Sub
    End If
    If LCase(eff) <> LCase(Files(fid).OldPath) Then
        LogLine tag & "resume: already renamed on disk - editing " & eff
    End If
    Set doc = OpenSmart(eff)
    If doc Is Nothing Then
        LogWarn tag & "cannot open - skipped: " & eff
        Exit Sub
    End If
    ActivateDoc doc
    LogLine tag & "edit: " & eff & "  (" & Files(fid).DocType & ")  " & Format(Now, "hh:nn:ss")

    ApplyCfgRenamesFor fid, doc
    ApplyMaterial Files(fid), doc
    ApplyProps Files(fid), doc
    SaveDoc doc, tag
    nDocsEdited = nDocsEdited + 1
    Exit Sub
EH:
    LogWarn tag & "error " & Err.Number & ": " & Err.Description & " - continuing with next file"
    Err.Clear
End Sub

Private Sub ApplyCfgRenamesFor(ByVal fid As Long, ByVal doc As SldWorks.ModelDoc2)
    Dim i As Long, dtChecked As Boolean, hasDT As Boolean
    For i = 1 To gNCfgs
        If Cfgs(i).FileId = fid Then
            If Not dtChecked Then
                dtChecked = True
                hasDT = DocHasDesignTable(doc)
                If hasDT Then LogWarn "CFG: design table detected - skipping config rename for " & StemOf(doc.GetPathName)
            End If
            If hasDT Then
                Cfgs(i).Skipped = True
                nCfgSkipped = nCfgSkipped + 1
            Else
                DoCfgRename i, doc
            End If
        End If
    Next i
End Sub

Private Sub DoCfgRename(ByVal i As Long, ByVal doc As SldWorks.ModelDoc2)
    Dim swCfg As SldWorks.Configuration
    Dim oldN As String, newN As String, docStem As String
    oldN = Cfgs(i).OldCfg
    newN = Cfgs(i).NewCfg
    docStem = StemOf(doc.GetPathName)
    On Error GoTo EH
    Set swCfg = FindCfgSmart(doc, oldN)
    If swCfg Is Nothing Then
        If ConfigExists(doc, newN) Then
            Cfgs(i).Done = True
            LogLine "CFG: already renamed: '" & oldN & "' -> '" & newN & "' (" & docStem & ")"
        Else
            Cfgs(i).Skipped = True
            LogWarn "CFG: configuration not found: '" & oldN & "' in " & docStem
        End If
        Exit Sub
    End If
    If StrComp(swCfg.Name, newN, vbTextCompare) = 0 Then
        Cfgs(i).Done = True
        LogLine "CFG: already correct: '" & newN & "' (" & docStem & ")"
        Exit Sub
    End If
    If ConfigExists(doc, newN) Then
        Cfgs(i).Skipped = True
        LogWarn "CFG: target name already exists - skipped: '" & newN & "' in " & docStem
        Exit Sub
    End If
    swCfg.Name = newN
    If StrComp(swCfg.Name, newN, vbTextCompare) = 0 Then
        Cfgs(i).Done = True
        nCfgRenamed = nCfgRenamed + 1
        LogLine "CFG: renamed '" & oldN & "' -> '" & newN & "' in " & docStem
    Else
        Cfgs(i).Skipped = True
        LogWarn "CFG: rename did not take effect: '" & oldN & "' in " & docStem
    End If
    Exit Sub
EH:
    Cfgs(i).Skipped = True
    LogWarn "CFG: rename failed '" & oldN & "' in " & docStem & ": " & Err.Description
    Err.Clear
End Sub

' exact match -> trimmed match -> single-config fallback (robust first run)
Private Function FindCfgSmart(ByVal doc As SldWorks.ModelDoc2, ByVal oldN As String) As SldWorks.Configuration
    Dim vNames As Variant, v As Variant
    On Error Resume Next
    Set FindCfgSmart = doc.GetConfigurationByName(oldN)
    On Error GoTo 0
    If Not FindCfgSmart Is Nothing Then Exit Function
    On Error Resume Next
    vNames = doc.GetConfigurationNames
    If IsArray(vNames) Then
        For Each v In vNames
            If StrComp(Trim(CStr(v)), Trim(oldN), vbTextCompare) = 0 Then
                Set FindCfgSmart = doc.GetConfigurationByName(CStr(v))
                Exit Function
            End If
        Next v
        If UBound(vNames) = LBound(vNames) Then
            Set FindCfgSmart = doc.GetConfigurationByName(CStr(vNames(LBound(vNames))))
        End If
    End If
    On Error GoTo 0
End Function

' ---- material: VERIFY first, SET only if missing/wrong ---------------------
Private Sub ApplyMaterial(ByRef f As FileRec, ByVal doc As SldWorks.ModelDoc2)
    Dim swPart As SldWorks.PartDoc
    Dim vNames As Variant, vN As Variant
    Dim cur As String, got As String, libCur As String, libUse As String
    If f.MatAction <> "CHECK" Then Exit Sub
    If Len(f.MatTarget) = 0 Then Exit Sub
    If doc.GetType <> swDocPART Then Exit Sub
    Set swPart = doc
    On Error GoTo EH
    vNames = doc.GetConfigurationNames
    For Each vN In vNames
        cur = ReadbackMaterial(doc, CStr(vN), libCur)
        CacheMatLib f.MatTarget, libCur
        If StrComp(Trim(cur), Trim(f.MatTarget), vbTextCompare) = 0 Then
            nMatVerified = nMatVerified + 1
            LogLine "MAT KEEP (verified): " & StemOf(doc.GetPathName) & " [" & vN & "] = " & cur
        Else
            libUse = ResolveMatLib(f.MatTarget)
            swPart.SetMaterialPropertyName2 CStr(vN), libUse, f.MatTarget
            got = ReadbackMaterial(doc, CStr(vN), libCur)
            If StrComp(Trim(got), Trim(f.MatTarget), vbTextCompare) = 0 Then
                nMatSet = nMatSet + 1
                CacheMatLib f.MatTarget, libUse
                LogLine "MAT SET: " & StemOf(doc.GetPathName) & " [" & vN & "] '" & cur & "' -> " & got & "  (lib: " & libUse & ")"
            Else
                nMatFail = nMatFail + 1
                LogWarn "MAT: " & StemOf(doc.GetPathName) & " expected '" & f.MatTarget & "', read back '" & got & "'"
            End If
        End If
    Next vN
    Exit Sub
EH:
    LogWarn "MAT error on " & StemOf(doc.GetPathName) & ": " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

Private Function ReadbackMaterial(ByVal doc As SldWorks.ModelDoc2, ByVal cfgName As String, ByRef libOut As String) As String
    Dim oP As Object
    Set oP = doc
    libOut = ""
    On Error Resume Next
    ReadbackMaterial = oP.GetMaterialPropertyName2(cfgName, libOut)
    If Err.Number <> 0 Then
        ReadbackMaterial = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub CacheMatLib(ByVal matName As String, ByVal libName As String)
    If Len(matName) = 0 Or Len(libName) = 0 Then Exit Sub
    gMatLibs.Item(LCase(Trim(matName))) = libName   ' remembers custom libs (e.g. PLA+)
End Sub

Private Function ResolveMatLib(ByVal matName As String) As String
    If gMatLibs.Exists(LCase(Trim(matName))) Then
        ResolveMatLib = CStr(gMatLibs.Item(LCase(Trim(matName))))
    Else
        ResolveMatLib = MAT_LIB
    End If
End Function

' ---- custom properties ------------------------------------------------------
Private Sub ApplyProps(ByRef f As FileRec, ByVal doc As SldWorks.ModelDoc2)
    Dim cpm As Object
    Dim j As Long, rc As Long
    Dim vVal As Variant, bRes As Boolean, sLink As String
    On Error GoTo EH
    Set cpm = doc.Extension.CustomPropertyManager("")
    For j = 1 To 6
        If Len(f.PropNames(j)) > 0 And Len(f.PropValues(j)) > 0 Then
            rc = cpm.Add3(f.PropNames(j), swCustomInfoText, f.PropValues(j), swCustomPropertyDeleteAndAdd)
            vVal = Empty: bRes = False: sLink = ""
            On Error Resume Next
            rc = cpm.Get5(f.PropNames(j), False, vVal, bRes, sLink)
            If Err.Number <> 0 Then
                Err.Clear
                vVal = f.PropValues(j)   ' trust Add3
            End If
            On Error GoTo EH
            If CStr(vVal) = f.PropValues(j) Then
                nPropSet = nPropSet + 1
            Else
                LogWarn "PROP: readback mismatch " & f.PropNames(j) & " on " & StemOf(doc.GetPathName) & ": got '" & CStr(vVal) & "'"
            End If
        End If
    Next j
    Exit Sub
EH:
    LogWarn "PROP error on " & StemOf(doc.GetPathName) & ": " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

Private Sub SaveDoc(ByVal doc As SldWorks.ModelDoc2, ByVal tag As String)
    Dim st As Boolean, e As Long, w As Long
    On Error Resume Next
    st = doc.Save3(swSaveAsOptions_Silent, e, w)
    If Err.Number <> 0 Then
        LogWarn tag & "Save3 raised " & Err.Number & ": " & Err.Description
        Err.Clear
        st = False
    End If
    On Error GoTo 0
    If st Then
        LogLine tag & "saved: " & doc.GetPathName
    Else
        LogWarn tag & "SAVE FAILED for " & doc.GetPathName & " (errors=" & e & ", warnings=" & w & ") - run again to redo"
    End If
End Sub

' ------------------------------- PHASE 2 ------------------------------------
Private Sub Phase2_RenameWindows()
    LogLine "--- Phase 2: per-assembly rename windows (deepest first) ---"
    Dim wi As Long
    For wi = 1 To gNWins
        WindowProc wi
        CloseAllDocs
        Pace
    Next wi
    CleanupOldFiles
    LogLine "Phase 2 done: windows=" & nWindowsDone & "/" & gNWins & ", renamed=" & nRenamed & _
            " (failed " & nRenameFail & "), refs repaired=" & nRefRepaired
End Sub

Private Sub WindowProc(ByVal wi As Long)
    Dim fid As Long, eff As String, tag As String
    Dim W As SldWorks.ModelDoc2
    Dim changed As Boolean, saved As Boolean
    On Error GoTo EH
    fid = Wins(wi)
    tag = "[win " & wi & "/" & gNWins & " " & StemOf(Files(fid).OldPath) & "] "
    gWinOk.Item(fid) = False
    Set gRenamedInWindow = New Collection
    Set gTmpKills = New Collection
    eff = EffectivePath(Files(fid).OldPath, Files(fid).NewPath)
    If Len(eff) = 0 Then
        LogWarn tag & "assembly not found on disk - window skipped"
        Exit Sub
    End If
    Set W = OpenSmart(eff)
    If W Is Nothing Then
        LogWarn tag & "cannot open assembly - window skipped: " & eff
        Exit Sub
    End If
    ActivateDoc W
    LogLine tag & "window open (" & Format(Now, "hh:nn:ss") & "): " & eff
    changed = RepairRefsInWindow(W, tag)
    If RenameChildrenInWindow(W, fid, tag) Then changed = True
    If changed Then
        saved = SaveWindow(W, tag)
        gWinOk.Item(fid) = saved
    Else
        LogLine tag & "nothing to change in this window"
        gWinOk.Item(fid) = True
    End If
    nWindowsDone = nWindowsDone + 1
    CloseAllDocs
    Exit Sub
EH:
    LogWarn tag & "window error " & Err.Number & ": " & Err.Description & " - continuing"
    Err.Clear
    RestorePendingTmp
    CloseAllDocs
End Sub

Private Function RepairRefsInWindow(ByVal W As SldWorks.ModelDoc2, ByVal tag As String) As Boolean
    Dim vNames As Variant, vN As Variant, firstCfg As String
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    Dim cfid As Long, cur As String, nw As String
    On Error GoTo EH
    vNames = W.GetConfigurationNames
    If Not IsArray(vNames) Then Exit Function
    If UBound(vNames) >= LBound(vNames) Then firstCfg = CStr(vNames(LBound(vNames)))
    For Each vN In vNames
        If W.ShowConfiguration2(CStr(vN)) Then
            Set swAssy = W
            vComps = swAssy.GetComponents(False)
            If IsArray(vComps) Then
                For Each v In vComps
                    Set swComp = v
                    cfid = FileIdByPath(swComp.GetPathName)
                    If cfid > 0 Then
                        cur = ""
                        On Error Resume Next
                        cur = swComp.ReferencedConfiguration
                        On Error GoTo EH
                        nw = NewCfgName(cfid, cur)
                        If Len(nw) > 0 And StrComp(nw, cur, vbBinaryCompare) <> 0 Then
                            If SetCompRef(swComp, nw, tag & "cfg-repair " & swComp.Name2 & " '" & cur & "' -> '" & nw & "'") Then
                                RepairRefsInWindow = True
                            End If
                        End If
                    End If
                Next v
            End If
        End If
    Next vN
    If Len(firstCfg) > 0 Then
        On Error Resume Next
        W.ShowConfiguration2 firstCfg
        On Error GoTo 0
    End If
    Exit Function
EH:
    LogWarn tag & "ref-repair error " & Err.Number & ": " & Err.Description
    Err.Clear
End Function

Private Function SetCompRef(ByVal swComp As SldWorks.Component2, ByVal newName As String, ByVal logTxt As String) As Boolean
    On Error Resume Next
    swComp.ReferencedConfiguration = newName
    If Err.Number <> 0 Then
        LogWarn "REF set failed (" & logTxt & "): " & Err.Description
        Err.Clear
    ElseIf StrComp(swComp.ReferencedConfiguration, newName, vbTextCompare) = 0 Then
        SetCompRef = True
        nRefRepaired = nRefRepaired + 1
        LogLine "REF: " & logTxt
    Else
        LogWarn "REF did not stick: " & logTxt
    End If
    On Error GoTo 0
End Function

Private Function RenameChildrenInWindow(ByVal W As SldWorks.ModelDoc2, ByVal fid As Long, ByVal tag As String) As Boolean
    Dim pi As Long, cid As Long
    Dim done As Object
    Set done = CreateObject("Scripting.Dictionary")
    For pi = 1 To gNPairs
        If PairPar(pi) = fid Then
            cid = PairChild(pi)
            If Files(cid).Action = "RENAME" Then
                If Not done.Exists(cid) Then
                    done.Add cid, True
                    If RenameOneChild(W, cid, tag) Then RenameChildrenInWindow = True
                End If
            End If
        End If
    Next pi
End Function

Private Function RenameOneChild(ByVal W As SldWorks.ModelDoc2, ByVal cid As Long, ByVal tag As String) As Boolean
    Dim oldP As String, newP As String, compName As String, compNameNew As String
    Dim oldE As Boolean, newE As Boolean, tmpP As String, okRen As Boolean
    On Error GoTo EH
    oldP = Files(cid).OldPath
    newP = Files(cid).NewPath
    oldE = FileExists(oldP)
    newE = FileExists(newP)
    compName = FindCompNameByPath(W, LCase(oldP))
    compNameNew = FindCompNameByPath(W, LCase(newP))

    ' ---- resume awareness -------------------------------------------------
    If newE And Not oldE Then
        If Len(compNameNew) > 0 Then
            Files(cid).Renamed = True
            LogLine tag & "child already renamed & referenced: " & StemOf(newP)
        Else
            nRenameFail = nRenameFail + 1
            LogWarn tag & "PROBLEM: '" & StemOf(newP) & "' exists but this assembly has no component for it. Run again."
        End If
        Exit Function
    End If
    If Not oldE And Not newE Then
        nRenameFail = nRenameFail + 1
        LogWarn tag & "child file missing on disk (old and new): " & oldP
        Exit Function
    End If
    If Len(compName) = 0 Then
        If newE Then
            Files(cid).Renamed = True
            LogLine tag & "child already renamed: " & StemOf(newP)
        Else
            nRenameFail = nRenameFail + 1
            LogWarn tag & "no component found for '" & StemOf(oldP) & "' in this window - skipped"
        End If
        Exit Function
    End If

    ' ---- target exists -> move aside temporarily ---------------------------
    tmpP = ""
    If newE Then
        tmpP = newP & ".rntmp"
        On Error Resume Next
        If FileExists(tmpP) Then Kill tmpP
        Name newP As tmpP
        If Err.Number <> 0 Then
            LogWarn tag & "cannot move existing target aside (" & Err.Description & ") - skipped: " & StemOf(newP)
            Err.Clear
            On Error GoTo 0
            nRenameFail = nRenameFail + 1
            Exit Function
        End If
        On Error GoTo 0
        LogLine tag & "existing target moved aside temporarily: " & StemOf(newP)
    End If

    ' ---- official rename flow (with fallbacks) -----------------------------
    okRen = TryRenameDocument(W, compName, StemOf(newP), tag)
    If Not okRen Then okRen = TryRenameDocument(W, compName, newP, tag)
    If Not okRen Then okRen = TrySaveAsFallback(cid, tag)
    If Not okRen Then
        LogLine tag & "last resort: resolving component: " & compName
        ResolveCompByName W, compName
        okRen = TryRenameDocument(W, compName, StemOf(newP), tag)
    End If

    If okRen Then
        Files(cid).Renamed = True
        nRenamed = nRenamed + 1
        RenameOneChild = True
        gRenamedInWindow.Add cid
        If Len(tmpP) > 0 Then gTmpKills.Add tmpP
        LogLine tag & "RENAMED: " & StemOf(oldP) & "  ->  " & StemOf(newP) & "  (" & Format(Now, "hh:nn:ss") & ")"
    Else
        nRenameFail = nRenameFail + 1
        LogWarn tag & "rename FAILED for " & oldP
        If Len(tmpP) > 0 Then RestoreOneTmp tmpP
    End If
    Exit Function
EH:
    nRenameFail = nRenameFail + 1
    LogWarn tag & "rename error " & Err.Number & ": " & Err.Description & " (" & Files(cid).OldPath & ")"
    Err.Clear
    If Len(tmpP) > 0 Then RestoreOneTmp tmpP
End Function

Private Function TryRenameDocument(ByVal W As SldWorks.ModelDoc2, ByVal compName As String, _
                                   ByVal newName As String, ByVal tag As String) As Boolean
    Dim renErr As Long, selOk As Boolean
    On Error GoTo EH
    W.ClearSelection2 True
    selOk = W.Extension.SelectByID2(compName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
    If Not selOk Then
        LogLine tag & "SelectByID2 failed for '" & compName & "'"
        Exit Function
    End If
    renErr = -1
    On Error Resume Next
    renErr = W.Extension.RenameDocument(newName)
    If Err.Number <> 0 Then
        LogLine tag & "RenameDocument('" & newName & "') raised " & Err.Number & ": " & Err.Description
        Err.Clear
        renErr = -2
    End If
    On Error GoTo EH
    If renErr = 0 Then
        TryRenameDocument = True
        Exit Function
    End If
    LogLine tag & "RenameDocument('" & newName & "') returned " & renErr
    Exit Function
EH:
    LogLine tag & "RenameDocument error " & Err.Number & ": " & Err.Description
    Err.Clear
End Function

Private Sub ResolveCompByName(ByVal W As SldWorks.ModelDoc2, ByVal compName As String)
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    On Error Resume Next
    Set swAssy = W
    vComps = swAssy.GetComponents(False)
    If IsArray(vComps) Then
        For Each v In vComps
            Set swComp = v
            If swComp.Name2 = compName Then
                If swComp.SuppressionState = swComponentLightweight Then
                    LogLine "REN: resolving lightweight component before rename: " & compName
                    swComp.SetSuppressionState2 swComponentFullyResolved
                    DoEvents
                    If PAUSE_MS > 0 Then ApiSleep PAUSE_MS
                End If
                Exit For
            End If
        Next v
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function TrySaveAsFallback(ByVal cid As Long, ByVal tag As String) As Boolean
    Dim childDoc As SldWorks.ModelDoc2, bRet As Boolean, e As Long, w2 As Long
    On Error Resume Next
    Set childDoc = OpenSmart(Files(cid).OldPath)
    If childDoc Is Nothing Then
        LogWarn tag & "SaveAs fallback: cannot open child: " & Files(cid).OldPath
        Exit Function
    End If
    bRet = childDoc.Extension.SaveAs(Files(cid).NewPath, swSaveAsCurrentVersion, _
                                     swSaveAsOptions_Silent, Nothing, e, w2)
    If Err.Number <> 0 Then
        LogWarn tag & "SaveAs fallback raised " & Err.Number & ": " & Err.Description
        Err.Clear
        bRet = False
    End If
    On Error GoTo 0
    If bRet Then
        Files(cid).UsedFallback = True
        LogLine tag & "SaveAs fallback OK: " & StemOf(Files(cid).OldPath) & " -> " & StemOf(Files(cid).NewPath)
        TrySaveAsFallback = True
    Else
        LogWarn tag & "SaveAs fallback FAILED (errors=" & e & ", warnings=" & w2 & ")"
    End If
End Function

Private Function SaveWindow(ByVal W As SldWorks.ModelDoc2, ByVal tag As String) As Boolean
    Dim st As Boolean, e As Long, w2 As Long
    On Error Resume Next
    st = W.Save3(swSaveAsOptions_Silent + swSaveAsOptions_SaveReferenced, e, w2)
    If Err.Number <> 0 Then
        LogWarn tag & "Save3(SaveReferenced) raised " & Err.Number & ": " & Err.Description
        Err.Clear
        st = False
    End If
    On Error GoTo 0
    If st Then
        LogLine tag & "window saved (renames committed): " & W.GetPathName
        KillPendingTmp
        SaveWindow = True
        Exit Function
    End If
    If gRenamedInWindow.Count > 0 Then
        LogWarn tag & "SAVE FAILED with pending renames - window discarded. Run the macro again."
        RestorePendingTmp
        SaveWindow = False
        Exit Function
    End If
    LogWarn tag & "SaveReferenced failed (errors=" & e & ") - retrying plain save (ref repairs only)"
    On Error Resume Next
    st = W.Save3(swSaveAsOptions_Silent, e, w2)
    If Err.Number <> 0 Then Err.Clear: st = False
    On Error GoTo 0
    If st Then
        LogLine tag & "plain save OK: " & W.GetPathName
        SaveWindow = True
    Else
        LogWarn tag & "SAVE FAILED - window discarded. Run the macro again."
        SaveWindow = False
    End If
End Function

Private Sub KillPendingTmp()
    Dim v As Variant
    On Error Resume Next
    For Each v In gTmpKills
        If FileExists(CStr(v)) Then
            Kill CStr(v)
            If Err.Number <> 0 Then Err.Clear
        End If
    Next v
    On Error GoTo 0
    Set gTmpKills = New Collection
End Sub

Private Sub RestorePendingTmp()
    Dim i As Long
    If gTmpKills Is Nothing Then Exit Sub
    For i = gTmpKills.Count To 1 Step -1
        RestoreOneTmp CStr(gTmpKills(i))
    Next i
End Sub

Private Sub RestoreOneTmp(ByVal t As String)
    Dim orig As String, i As Long
    If Len(t) = 0 Then Exit Sub
    On Error Resume Next
    orig = Left$(t, Len(t) - Len(".rntmp"))
    If FileExists(t) And Not FileExists(orig) Then
        Name t As orig
        If Err.Number <> 0 Then
            LogWarn "TMP: could not restore " & orig & ": " & Err.Description
            Err.Clear
        Else
            LogLine "TMP: restored " & orig
        End If
    End If
    If Not gTmpKills Is Nothing Then
        For i = gTmpKills.Count To 1 Step -1
            If StrComp(CStr(gTmpKills(i)), t, vbTextCompare) = 0 Then gTmpKills.Remove i
        Next i
    End If
    On Error GoTo 0
End Sub

Private Sub CleanupOldFiles()
    LogLine "--- Phase 2b: cleanup of leftover old-name files ---"
    Dim i As Long, f As String
    For i = 1 To gNFiles
        If Files(i).Action = "RENAME" And i <> gTopFid Then
            If FileExists(Files(i).NewPath) Then
                If FileExists(Files(i).OldPath) Then
                    If CLEAN_LEFTOVER_OLD And AllParentsOk(i) Then
                        On Error Resume Next
                        Kill Files(i).OldPath
                        If Err.Number <> 0 Then
                            LogWarn "CLEAN: could not delete " & Files(i).OldPath & ": " & Err.Description
                            Err.Clear
                            nOldLeft = nOldLeft + 1
                        Else
                            LogLine "CLEAN: removed leftover old file: " & StemOf(Files(i).OldPath)
                            nStubsCleaned = nStubsCleaned + 1
                        End If
                        On Error GoTo 0
                    Else
                        nOldLeft = nOldLeft + 1
                        LogLine "NOTE: old file KEPT for review: " & Files(i).OldPath
                    End If
                End If
            Else
                LogWarn "CLEAN: renamed file missing on disk: " & Files(i).NewPath
            End If
        End If
    Next i
    On Error Resume Next
    f = Dir(PROJ_DIR & "\*.rntmp")
    Do While Len(f) > 0
        Kill PROJ_DIR & "\" & f
        LogLine "CLEAN: removed temp file: " & f
        f = Dir()
    Loop
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Function AllParentsOk(ByVal cid As Long) As Boolean
    Dim pi As Long
    AllParentsOk = True
    For pi = 1 To gNPairs
        If PairChild(pi) = cid Then
            If Not gWinOk.Exists(PairPar(pi)) Then
                AllParentsOk = False
                Exit Function
            End If
            If Not CBool(gWinOk.Item(PairPar(pi))) Then
                AllParentsOk = False
                Exit Function
            End If
        End If
    Next pi
End Function

Private Sub RenameTopFile()
    LogLine "--- Phase 2c: rename the top assembly file itself ---"
    On Error GoTo EH
    If FileExists(TOP_NEW) Then
        LogLine "TOP: already renamed: " & TOP_NEW
        Exit Sub
    End If
    If Not FileExists(TOP_ASM) Then
        LogWarn "TOP: top assembly not found (old and new missing): " & TOP_ASM
        Exit Sub
    End If
    Name TOP_ASM As TOP_NEW
    LogLine "TOP RENAMED: " & TOP_ASM & "  ->  " & TOP_NEW
    Exit Sub
EH:
    LogWarn "TOP: rename failed " & Err.Number & ": " & Err.Description & " - rename manually or run again"
    Err.Clear
End Sub

' ------------------------------- PHASE 3 ------------------------------------
Private Sub Phase3_VerifyAndReport()
    LogLine "--- Phase 3: verification + CSV report ---"
    Dim i As Long, okRen As Long
    For i = 1 To gNFiles
        If Files(i).Action = "RENAME" Then
            If FileExists(Files(i).NewPath) Then
                okRen = okRen + 1
            Else
                LogWarn "VERIFY: renamed file missing on disk: " & Files(i).NewPath
            End If
        End If
    Next i
    LogLine "VERIFY: renamed files present on disk: " & okRen & " / " & EXP_RENAMES

    Dim csvLines As Collection
    Set csvLines = New Collection
    csvLines.Add "ParentAssembly;ScanConfig;Instance;ModelName;ModelPath;RefConfig;State;Status"
    gVerifyCount = 0: nMismatch = 0
    For i = 1 To gNFiles
        If Files(i).DocType = "Assembly" Then
            VerifyOneAssembly i, csvLines
            CloseAllDocs
            Pace
        End If
    Next i
    LogLine "VERIFY: instances scanned=" & gVerifyCount & " (expected " & EXP_INSTANCES & "), mismatches=" & nMismatch
    If gVerifyCount <> EXP_INSTANCES Then LogWarn "instance count differs from audit (" & EXP_INSTANCES & ")"
    WriteCsv csvLines
End Sub

Private Sub VerifyOneAssembly(ByVal fid As Long, ByVal csvLines As Collection)
    Dim eff As String, doc As SldWorks.ModelDoc2, scanCfg As String, firstCfg As String
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant, vNames As Variant
    Dim swComp As SldWorks.Component2
    Dim p As String, cfid As Long, ref As String, stat As String
    Dim supp As Boolean
    On Error GoTo EH
    eff = EffectivePath(Files(fid).OldPath, Files(fid).NewPath)
    If Len(eff) = 0 Then
        LogWarn "VERIFY: assembly missing on disk: " & Files(fid).OldPath
        Exit Sub
    End If
    Set doc = OpenSmart(eff)
    If doc Is Nothing Then
        LogWarn "VERIFY: cannot open " & eff
        Exit Sub
    End If
    ActivateDoc doc
    vNames = doc.GetConfigurationNames
    If IsArray(vNames) Then
        If UBound(vNames) >= LBound(vNames) Then firstCfg = CStr(vNames(LBound(vNames)))
    End If
    scanCfg = firstCfg
    If ConfigExists(doc, NEW_CFG) Then scanCfg = NEW_CFG
    On Error Resume Next
    doc.ShowConfiguration2 scanCfg
    On Error GoTo EH
    Set swAssy = doc
    vComps = swAssy.GetComponents(False)     ' direct children only (no double counting)
    If Not IsArray(vComps) Then Exit Sub
    For Each v In vComps
        Set swComp = v
        p = swComp.GetPathName
        cfid = FileIdByPath(p)
        supp = False: ref = ""
        On Error Resume Next
        supp = (swComp.SuppressionState = swComponentSuppressed)
        ref = swComp.ReferencedConfiguration
        On Error GoTo EH
        stat = "OK"
        gVerifyCount = gVerifyCount + 1
        If cfid > 0 Then
            If Files(cfid).Action = "RENAME" And LCase(p) <> LCase(Files(cfid).NewPath) Then
                stat = "OLD_PATH"
                nMismatch = nMismatch + 1
                LogWarn "VERIFY: still references OLD file: " & p
            ElseIf Len(NewCfgName(cfid, ref)) > 0 Then
                stat = "OLD_CFG_REF"
                nMismatch = nMismatch + 1
                LogWarn "VERIFY: component references OLD config '" & ref & "': " & p
            End If
        Else
            If Len(p) > 0 Then
                stat = "NOT_IN_PLAN"
                nMismatch = nMismatch + 1
                LogWarn "VERIFY: unexpected reference: " & p
            End If
        End If
        csvLines.Add CsvQ(StemOf(eff)) & ";" & CsvQ(scanCfg) & ";" & CsvQ(swComp.Name2) & ";" & _
                     CsvQ(StemOf(p)) & ";" & CsvQ(p) & ";" & CsvQ(ref) & ";" & _
                     CsvQ(IIf(supp, "SUPPRESSED", "SHOWN")) & ";" & CsvQ(stat)
    Next v
    Exit Sub
EH:
    LogWarn "VERIFY error in " & StemOf(eff) & ": " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

Private Sub WriteCsv(ByVal lines As Collection)
    Dim fn As Integer, i As Long
    On Error GoTo EH
    fn = FreeFile
    Open REPORT_CSV For Output As #fn
    For i = 1 To lines.Count
        Print #fn, lines(i)
    Next i
    Close #fn
    LogLine "CSV report written: " & REPORT_CSV & " (" & (lines.Count - 1) & " instances)"
    Exit Sub
EH:
    LogWarn "CSV write failed: " & Err.Description
End Sub

' -------------------------------- SUMMARY -----------------------------------
Private Sub WriteSummary()
    Dim s As String
    LogLine "=================== SUMMARY (SCARA edition) ==================="
    LogLine "Documents processed (phase 1) : " & nDocsEdited & " / " & gNOrder
    LogLine "Configuration renames applied : " & nCfgRenamed & " (skipped: " & nCfgSkipped & ")"
    LogLine "Materials verified (kept)     : " & nMatVerified & "  (expected " & EXP_MAT_KEEP & ")"
    LogLine "Materials assigned            : " & nMatSet & " (failed: " & nMatFail & ", expected " & EXP_MAT_SET & ")"
    LogLine "Custom properties set         : " & nPropSet
    LogLine "Refs repaired (phase 2)       : " & nRefRepaired
    LogLine "Rename windows processed      : " & nWindowsDone & " / " & gNWins
    LogLine "Files renamed                 : " & nRenamed & " (failed: " & nRenameFail & ")"
    LogLine "Leftover old files deleted    : " & nStubsCleaned & " (kept for review: " & nOldLeft & ")"
    LogLine "Verify instances / mismatches : " & gVerifyCount & " / " & nMismatch
    LogLine "Warnings                      : " & nWarnings
    LogLine "NEW top assembly              : " & TOP_NEW
    LogLine "Log                           : " & LOG_PATH
    If nRenameFail > 0 Or nMismatch > 0 Or nMatFail > 0 Then
        LogLine ">>> Not perfect: RUN THE MACRO AGAIN - it resumes and retries only what is missing."
    End If

    s = "SCARA_Robot preparation - COMPLETE" & vbCrLf & vbCrLf & _
        "Documents processed: " & nDocsEdited & " / " & gNOrder & vbCrLf & _
        "Config renames: " & nCfgRenamed & " (skipped " & nCfgSkipped & ")" & vbCrLf & _
        "Materials kept (verified): " & nMatVerified & vbCrLf & _
        "Materials assigned: " & nMatSet & " (failed " & nMatFail & ")" & vbCrLf & _
        "Custom properties: " & nPropSet & vbCrLf & _
        "Files renamed: " & nRenamed & " (failed " & nRenameFail & ")" & vbCrLf & _
        "Verify mismatches: " & nMismatch & "   Warnings: " & nWarnings & vbCrLf & vbCrLf & _
        "NEW top assembly:" & vbCrLf & TOP_NEW & vbCrLf & vbCrLf & _
        "Log: " & LOG_PATH & vbCrLf & "CSV: " & REPORT_CSV
    If nRenameFail > 0 Or nMismatch > 0 Then
        s = s & vbCrLf & vbCrLf & "Not perfect? Run the macro AGAIN - it resumes where it stopped."
    End If
    MsgBox s, vbInformation, "SCARA prep"
End Sub

' -------------------------------- HELPERS -----------------------------------
Private Sub Pace()
    On Error Resume Next
    DoEvents
    If PAUSE_MS > 0 Then ApiSleep PAUSE_MS
    DoEvents
    On Error GoTo 0
End Sub

Private Function OpenSmart(ByVal p As String) As SldWorks.ModelDoc2
    Dim t As Long, opts As Long, e As Long, w As Long
    t = DocTypeFromExt(p)
    If t = 0 Then Exit Function
    opts = swOpenDocOptions_Silent
    If OPEN_LIGHTWEIGHT And t = swDocASSEMBLY Then
        opts = swOpenDocOptions_Silent Or swOpenDocOptions_OverrideDefaultLoadLightweight Or swOpenDocOptions_LoadLightweight
    End If
    On Error Resume Next
    Set OpenSmart = swApp.OpenDoc6(p, t, opts, "", e, w)
    If Err.Number <> 0 Then
        LogWarn "OpenDoc6 raised for " & p & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If OpenSmart Is Nothing And opts <> swOpenDocOptions_Silent Then
        LogLine "OPEN: lightweight failed for " & p & " - retrying normal open"
        On Error Resume Next
        Set OpenSmart = swApp.OpenDoc6(p, t, swOpenDocOptions_Silent, "", e, w)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
    If OpenSmart Is Nothing Then LogWarn "OPEN: failed: " & p & " (errors=" & e & ", warnings=" & w & ")"
End Function

Private Sub ActivateDoc(ByVal doc As SldWorks.ModelDoc2)
    Dim e As Long
    On Error Resume Next
    swApp.ActivateDoc2 doc.GetTitle, False, e
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub CloseAllDocs()
    Dim vDocs As Variant, m As Variant
    Dim md As SldWorks.ModelDoc2
    If swApp Is Nothing Then Exit Sub
    On Error Resume Next
    vDocs = swApp.GetDocuments
    If IsArray(vDocs) Then
        For Each m In vDocs
            Set md = m
            swApp.CloseDoc md.GetTitle
            Set md = Nothing
        Next m
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function EffectivePath(ByVal oldP As String, ByVal newP As String) As String
    If FileExists(oldP) Then
        EffectivePath = oldP
        Exit Function
    End If
    If Len(newP) > 0 Then
        If FileExists(newP) Then
            EffectivePath = newP
            Exit Function
        End If
    End If
    EffectivePath = ""
End Function

Private Function FileIdByPath(ByVal p As String) As Long
    Dim lk As String
    FileIdByPath = 0
    If Len(p) = 0 Then Exit Function
    lk = LCase(p)
    If mIdByOld.Exists(lk) Then
        FileIdByPath = CLng(mIdByOld.Item(lk))
        Exit Function
    End If
    If mIdByNew.Exists(lk) Then FileIdByPath = CLng(mIdByNew.Item(lk))
End Function

Private Function NewCfgName(ByVal cfid As Long, ByVal cur As String) As String
    Dim d As Object
    NewCfgName = ""
    If Len(cur) = 0 Then Exit Function
    If mCfgMap Is Nothing Then Exit Function
    If mCfgMap.Exists(cfid) Then
        Set d = mCfgMap.Item(cfid)
        If d.Exists(LCase(cur)) Then NewCfgName = CStr(d.Item(LCase(cur)))
    End If
End Function

Private Function FindCompNameByPath(ByVal W As SldWorks.ModelDoc2, ByVal pathLower As String) As String
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vComps As Variant, v As Variant
    Dim swComp As SldWorks.Component2
    FindCompNameByPath = ""
    On Error GoTo EH
    Set swAssy = W
    vComps = swAssy.GetComponents(False)
    If Not IsArray(vComps) Then Exit Function
    For Each v In vComps
        Set swComp = v
        If LCase(swComp.GetPathName) = pathLower Then
            FindCompNameByPath = swComp.Name2
            Exit Function
        End If
    Next v
    Exit Function
EH:
    Err.Clear
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(p) <> "")
    If Err.Number <> 0 Then
        FileExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function StemOf(ByVal p As String) As String
    Dim s As String, i As Long
    s = p
    i = InStrRev(s, "\")
    If i > 0 Then s = Mid(s, i + 1)
    i = InStrRev(s, ".")
    If i > 0 Then s = Left(s, i - 1)
    StemOf = s
End Function

Private Function DocTypeFromExt(ByVal p As String) As Long
    Dim e As String
    If Len(p) >= 7 Then e = LCase(Right(p, 7))
    If e = ".sldprt" Then
        DocTypeFromExt = swDocPART
    ElseIf e = ".sldasm" Then
        DocTypeFromExt = swDocASSEMBLY
    Else
        DocTypeFromExt = 0
    End If
End Function

Private Function DocHasDesignTable(ByVal doc As SldWorks.ModelDoc2) As Boolean
    On Error Resume Next
    DocHasDesignTable = doc.Extension.HasDesignTable
    If Err.Number <> 0 Then
        DocHasDesignTable = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ConfigExists(ByVal doc As SldWorks.ModelDoc2, ByVal name As String) As Boolean
    Dim vNames As Variant, v As Variant
    ConfigExists = False
    On Error Resume Next
    vNames = doc.GetConfigurationNames
    If IsArray(vNames) Then
        For Each v In vNames
            If StrComp(CStr(v), name, vbTextCompare) = 0 Then
                ConfigExists = True
                Exit For
            End If
        Next v
    End If
    On Error GoTo 0
End Function

Private Function CsvQ(ByVal s As String) As String
    CsvQ = """" & Replace(s, """", """""") & """"
End Function

Private Sub OpenLog()
    On Error Resume Next
    gLogNum = FreeFile
    Open LOG_PATH For Append As #gLogNum
    If Err.Number <> 0 Then
        gLogNum = 0
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub LogLine(ByVal s As String)
    Debug.Print s
    If gLogNum <> 0 Then
        On Error Resume Next
        Print #gLogNum, s
        On Error GoTo 0
    End If
End Sub

Private Sub LogWarn(ByVal s As String)
    nWarnings = nWarnings + 1
    LogLine "WARNING: " & s
End Sub

Private Sub CloseLog()
    On Error Resume Next
    If gLogNum <> 0 Then Close #gLogNum
    gLogNum = 0
    On Error GoTo 0
End Sub

' ============================ PLAN DATA (from SCARA_Robot_DeepAudit) =========
' 61 documents = 59 unique parts + Gripper_Mechanism + SCARA_Robot (root).
' ids 1..45 = root parts (Base FIRST for PLA+ library auto-detection),
' ids 46..59 = gripper parts, 60 = gripper assembly, 61 = root assembly.
' All new names: [A-Z a-z 0-9 _] only. All configurations -> "Default".
' =============================================================================
Private Sub LoadData()
    Dim k As Long, n As Long, c As Long
    gNFiles = 61: gNCfgs = 61: gNOrder = 61: gNWins = 2: gNPairs = 60
    ReDim Files(1 To gNFiles)
    ReDim Cfgs(1 To gNCfgs)
    ReDim OrderIds(1 To gNOrder)
    ReDim Wins(1 To gNWins)
    ReDim PairPar(1 To gNPairs)
    ReDim PairChild(1 To gNPairs)

    ' --- files: id, action, type, depth, oldStem, newStem, matAction, matTarget, review, description, category
    AddFile 1, "RENAME", "Part", 1, "Base", "SR_10_001_Base", "CHECK", "PLA+", False, "Robot base frame, 3D printed", "FAB"
    AddFile 2, "RENAME", "Part", 1, "J1_coupler", "SR_10_002_J1_Coupler", "CHECK", "PLA+", False, "J1 axis coupler plate, connects J1 output to Z axis base", "FAB"
    AddFile 3, "RENAME", "Part", 1, "Base_cover", "SR_10_003_Base_Cover", "CHECK", "PLA+", False, "Base top cover shell", "FAB"
    AddFile 4, "RENAME", "Part", 1, "GT2_Pulley_22_80teeth", "SR_10_004_GT2_J1_Idler_Pulley_80T", "CHECK", "PLA+", False, "GT2 idler pulley 80T on M8 bolt, J1 first stage", "FAB"
    AddFile 5, "RENAME", "Part", 1, "GT2_Pulley_110teeth_J1", "SR_10_005_GT2_J1_Driven_Pulley_110T", "CHECK", "PLA+", False, "GT2 driven pulley 110T, J1 output", "FAB"
    AddFile 6, "RENAME", "Part", 1, "Z_axis_Bottom_Plate", "SR_20_001_Z_Bottom_Plate", "CHECK", "PLA+", False, "Z axis bottom plate", "FAB"
    AddFile 7, "RENAME", "Part", 1, "Z_axis_Top_Plate", "SR_20_002_Z_Top_Plate", "CHECK", "PLA+", False, "Z axis top plate with motor mount", "FAB"
    AddFile 8, "RENAME", "Part", 1, "Z_axis_Mount_Platform", "SR_20_003_Z_Mount_Platform", "CHECK", "PLA+", False, "Z axis moving mount platform", "FAB"
    AddFile 9, "RENAME", "Part", 1, "Top_cover", "SR_20_004_Z_Top_Cover", "CHECK", "PLA+", False, "Z axis top cover", "FAB"
    AddFile 10, "RENAME", "Part", 1, "Arm_1", "SR_30_001_Arm_1", "CHECK", "PLA+", False, "Arm 1 link, J1 to J2", "FAB"
    AddFile 11, "RENAME", "Part", 1, "Arm_1_Cover", "SR_30_002_Arm_1_Cover", "CHECK", "PLA+", False, "Arm 1 top cover", "FAB"
    AddFile 12, "RENAME", "Part", 1, "J2_Coupler", "SR_30_003_J2_Coupler", "CHECK", "PLA+", False, "J2 axis coupler", "FAB"
    AddFile 13, "RENAME", "Part", 1, "GT2_Pulley_23_80teeth", "SR_30_004_GT2_J2_Idler_Pulley_80T", "CHECK", "PLA+", False, "GT2 idler pulley 80T, J2 stage", "FAB"
    AddFile 14, "RENAME", "Part", 1, "GT2_Pulley_92teeth_J2", "SR_30_005_GT2_J2_Driven_Pulley_92T", "CHECK", "PLA+", False, "GT2 driven pulley 92T, J2 output", "FAB"
    AddFile 15, "RENAME", "Part", 1, "Arm_2", "SR_40_001_Arm_2", "CHECK", "PLA+", False, "Arm 2 link, J2 to J3", "FAB"
    AddFile 16, "RENAME", "Part", 1, "Arm_2_Cover", "SR_40_002_Arm_2_Cover", "CHECK", "PLA+", False, "Arm 2 cover with J3 motor mount", "FAB"
    AddFile 17, "RENAME", "Part", 1, "J3_Coupler", "SR_40_003_J3_Coupler", "CHECK", "PLA+", False, "J3 axis coupler", "FAB"
    AddFile 18, "RENAME", "Part", 1, "GT2_Pulley_90teeth_J3", "SR_40_004_GT2_J3_Driven_Pulley_90T", "CHECK", "PLA+", False, "GT2 driven pulley 90T, J3 output", "FAB"
    AddFile 19, "RENAME", "Part", 1, "Arduino UNO case p1", "SR_60_001_Arduino_UNO_Case_P1", "CHECK", "PLA+", True, "Arduino UNO enclosure part 1 (hidden); material inferred as printed PLA+", "FAB"
    AddFile 20, "RENAME", "Part", 1, "Arduino UNO case p2", "SR_60_002_Arduino_UNO_Case_P2", "CHECK", "PLA+", True, "Arduino UNO enclosure part 2 (hidden); material inferred as printed PLA+", "FAB"
    AddFile 21, "RENAME", "Part", 1, "GT2_Pulley_Parametric", "COTS_PUL_GT2_Metal_Drive_Pulley", "CHECK", "Alloy Steel", False, "GT2 metal drive pulley, 5 mm bore, on motor shafts", "COTS"
    AddFile 22, "RENAME", "Part", 1, "Stepper_NEMA_17_20mm_shaft", "COTS_MOT_NEMA17_20mm_Shaft", "CHECK", "Alloy Steel", False, "NEMA 17 stepper motor, 20 mm shaft (imported STEP)", "COTS"
    AddFile 23, "RENAME", "Part", 1, "NEMA_17_Stepper_L24mm_20mm_shaft", "COTS_MOT_NEMA17_L24_20mm_Shaft", "CHECK", "Alloy Steel", False, "NEMA 17 stepper L24 mm body, 20 mm shaft (imported STEP)", "COTS"
    AddFile 24, "RENAME", "Part", 1, "Micro_switch", "COTS_SW_Micro_Limit_Switch", "CHECK", "Alloy Steel", True, "Micro limit switch (imported STEP); material inferred from project metadata Steel", "COTS"
    AddFile 25, "RENAME", "Part", 1, "Ball_Bearing_608_8_22_7mm", "COTS_BRG_608_8x22x7", "CHECK", "Alloy Steel", False, "Deep groove ball bearing 608, 8x22x7 mm", "COTS"
    AddFile 26, "RENAME", "Part", 1, "Ball_Bearing_30x42x7mm", "COTS_BRG_6806_30x42x7", "CHECK", "Alloy Steel", False, "Ball bearing 6806, 30x42x7 mm", "COTS"
    AddFile 27, "RENAME", "Part", 1, "Ball_Bearing_35x47x7mm", "COTS_BRG_6807_35x47x7", "CHECK", "Alloy Steel", True, "Ball bearing 6807, 35x47x7 mm (legacy config label says 30x42x7)", "COTS"
    AddFile 28, "RENAME", "Part", 1, "Thrust_Bearing_40_60_13", "COTS_BRG_51108_40x60x13", "CHECK", "Alloy Steel", False, "Thrust ball bearing 51108, 40x60x13 mm", "COTS"
    AddFile 29, "RENAME", "Part", 1, "Thrust_Bearing_35x52x12mm", "COTS_BRG_51107_35x52x12", "CHECK", "Alloy Steel", False, "Thrust ball bearing 51107, 35x52x12 mm", "COTS"
    AddFile 30, "RENAME", "Part", 1, "Linear_Bearing_10x19x29mm", "COTS_LIN_LM10UU_10x19x29", "CHECK", "Alloy Steel", False, "Linear ball bearing LM10UU, 10x19x29 mm", "COTS"
    AddFile 31, "RENAME", "Part", 1, "Smooth_Rod_D10mm_L400mm", "COTS_LIN_Rod_D10_L400", "CHECK", "Alloy Steel", True, "Smooth rod D10 mm, designated L400 (model measures 500 mm - verify)", "COTS"
    AddFile 32, "RENAME", "Part", 1, "Smooth_Rod_Clamp", "COTS_LIN_Rod_Clamp_D10", "CHECK", "Alloy Steel", False, "Shaft support clamp for D10 rod", "COTS"
    AddFile 33, "RENAME", "Part", 1, "Lead_Screw_D8mm_L380mm", "COTS_LIN_Lead_Screw_D8_L380", "CHECK", "Alloy Steel", True, "Lead screw D8 mm, designated L380 (model measures 500 mm - verify)", "COTS"
    AddFile 34, "RENAME", "Part", 1, "Lead_Screw_Nut_8mm", "COTS_LIN_Lead_Screw_Nut_D8", "CHECK", "Alloy Steel", False, "Lead screw nut D8 mm", "COTS"
    AddFile 35, "RENAME", "Part", 1, "GT2_Belt_200mm", "COTS_BLT_GT2_L200", "CHECK", "BUTYL", False, "GT2 belt 200 mm, J1 motor stage", "COTS"
    AddFile 36, "RENAME", "Part", 1, "GT2_Belt_300mm", "COTS_BLT_GT2_L300_A", "CHECK", "BUTYL", False, "GT2 belt 300 mm, J2 stage", "COTS"
    AddFile 37, "RENAME", "Part", 1, "GT2_Belt_300mm_2", "COTS_BLT_GT2_L300_B", "CHECK", "BUTYL", False, "GT2 belt 300 mm, J1 second stage", "COTS"
    AddFile 38, "RENAME", "Part", 1, "GT2_Belt_400mm", "COTS_BLT_GT2_L400_A", "CHECK", "BUTYL", False, "GT2 belt 400 mm, J2 motor stage", "COTS"
    AddFile 39, "RENAME", "Part", 1, "GT2_Belt_400mm_2", "COTS_BLT_GT2_L400_B", "CHECK", "BUTYL", False, "GT2 belt 400 mm, J3 stage", "COTS"
    AddFile 40, "RENAME", "Part", 1, "M4_50mm_Bolts", "COTS_FST_Bolt_M4x50", "CHECK", "Ductile Iron", False, "Hex bolt M4x50", "COTS"
    AddFile 41, "RENAME", "Part", 1, "M4_60mm_Bolts", "COTS_FST_Bolt_M4x60", "CHECK", "Ductile Iron", False, "Hex bolt M4x60", "COTS"
    AddFile 42, "RENAME", "Part", 1, "M8_45mm_Bolts", "COTS_FST_Bolt_M8x45", "CHECK", "Ductile Iron", False, "Hex bolt M8x45", "COTS"
    AddFile 43, "RENAME", "Part", 1, "M4_Lock_nut", "COTS_FST_Lock_Nut_M4", "CHECK", "Ductile Iron", False, "Lock nut M4", "COTS"
    AddFile 44, "RENAME", "Part", 1, "M8_Lock_nut", "COTS_FST_Lock_Nut_M8", "CHECK", "Ductile Iron", False, "Lock nut M8", "COTS"
    AddFile 45, "RENAME", "Part", 1, "M8_Washer", "COTS_FST_Washer_M8", "CHECK", "Alloy Steel", False, "Washer M8", "COTS"
    AddFile 46, "RENAME", "Part", 2, "Gripper_to_J3_connector", "SR_50_001_Gripper_To_J3_Connector", "CHECK", "PLA+", False, "Gripper to J3 connector plate", "FAB"
    AddFile 47, "RENAME", "Part", 2, "Servo_Holder", "SR_50_002_Servo_Holder", "CHECK", "PLA+", False, "MG996R servo holder", "FAB"
    AddFile 48, "RENAME", "Part", 2, "Gripper_Cover", "SR_50_003_Gripper_Cover", "CHECK", "PLA+", False, "Gripper cover", "FAB"
    AddFile 49, "RENAME", "Part", 2, "Gripper_Mechanism_Slider", "SR_50_004_Gripper_Slider", "CHECK", "PLA+", False, "Gripper slider", "FAB"
    AddFile 50, "RENAME", "Part", 2, "Gripper_Hand_Left", "SR_50_005_Gripper_Hand_Left", "CHECK", "PLA+", False, "Gripper hand left", "FAB"
    AddFile 51, "RENAME", "Part", 2, "Gripper_Hand_Right", "SR_50_006_Gripper_Hand_Right", "CHECK", "PLA+", False, "Gripper hand right", "FAB"
    AddFile 52, "RENAME", "Part", 2, "Gripper_link1", "SR_50_007_Gripper_Link_1", "CHECK", "PLA+", False, "Gripper link 1", "FAB"
    AddFile 53, "RENAME", "Part", 2, "Gripper_link2", "SR_50_008_Gripper_Link_2", "CHECK", "PLA+", False, "Gripper link 2", "FAB"
    AddFile 54, "RENAME", "Part", 2, "Gripper_End", "SR_50_009_Gripper_End", "CHECK", "PLA+", False, "Gripper end finger", "FAB"
    AddFile 55, "RENAME", "Part", 2, "Wide_gripper_end", "SR_50_010_Wide_Gripper_End", "CHECK", "PLA+", False, "Wide gripper end (hidden variant)", "FAB"
    AddFile 56, "RENAME", "Part", 2, "Wide_gripper_end_Left_side", "SR_50_011_Wide_Gripper_End_Left", "CHECK", "PLA+", False, "Wide gripper end left side (hidden variant)", "FAB"
    AddFile 57, "RENAME", "Part", 2, "Gripper_rail_6mm", "COTS_LIN_Rod_D6_L125", "CHECK", "Alloy Steel", False, "Gripper rail rod D6 mm L125", "COTS"
    AddFile 58, "RENAME", "Part", 2, "Servo_Motor_MG996R_3D_Model", "COTS_SRV_MG996R", "CHECK", "ABS", False, "MG996R servo motor (imported STEP)", "COTS"
    AddFile 59, "RENAME", "Part", 2, "Servo_MG996R_Horn", "COTS_SRV_MG996R_Horn", "CHECK", "ABS", False, "MG996R servo horn (imported STEP)", "COTS"
    AddFile 60, "RENAME", "Assembly", 1, "Gripper_Mechanism", "SR_50000_Gripper_Mechanism_Assy", "NONE", "", False, "Gripper mechanism sub assembly", "FAB"
    AddFile 61, "RENAME", "Assembly", 0, "SCARA_Robot", "SR_00000_SCARA_Robot_Assy", "NONE", "", False, "SCARA robot top assembly", "FAB"

    ' --- configuration renames (fileId, oldName) -> all become "Default" ---
    AddCfg 1, 1, "Base"
    AddCfg 2, 2, "J1 coupler"
    AddCfg 3, 3, "Base cover"
    AddCfg 4, 4, "GT2 Pulley - Parametric"
    AddCfg 5, 5, "GT2 Pulley - Parametric"
    AddCfg 6, 6, "Z-axis Bottom Plate"
    AddCfg 7, 7, "Z-axis Top Plate"
    AddCfg 8, 8, "Z-axis Mount Platform"
    AddCfg 9, 9, "Top cover"
    AddCfg 10, 10, "Arm 1"
    AddCfg 11, 11, "Arm 1 Cover"
    AddCfg 12, 12, "J2 Coupler"
    AddCfg 13, 13, "GT2 Pulley - Parametric"
    AddCfg 14, 14, "GT2 Pulley - Parametric"
    AddCfg 15, 15, "Arm 2"
    AddCfg 16, 16, "Arm 2 Cover"
    AddCfg 17, 17, "J3 Coupler"
    AddCfg 18, 18, "GT2 Pulley - Parametric"
    AddCfg 19, 19, "Arduino UNO case p1"
    AddCfg 20, 20, "Arduino UNO case p2"
    AddCfg 21, 21, "GT2 Pulley - Parametric"
    AddCfg 22, 22, "Stepper NEMA 17 -  20mm shaft"
    AddCfg 23, 23, "NEMA 17 Stepper L24mm - 20mm shaft"
    AddCfg 24, 24, "Micro switch"
    AddCfg 25, 25, "Ball Bearing 608 - 8x22x7mm"
    AddCfg 26, 26, "Ball Bearing 30x42x7mm"
    AddCfg 27, 27, "Ball Bearing 30x42x7mm"
    AddCfg 28, 28, "Thrust Bearing - 40x60x13"
    AddCfg 29, 29, "Thrust Bearing 35x52x12mm"
    AddCfg 30, 30, "Linear Bearing 10x19x29mm"
    AddCfg 31, 31, "Smooth Rod D10mm L400mm"
    AddCfg 32, 32, "Smooth Rod Clamp"
    AddCfg 33, 33, "Lead Screw D8mm L380mm"
    AddCfg 34, 34, "Lead Screw Nut 8mm"
    AddCfg 35, 35, "GT2 Belt 400mm"
    AddCfg 36, 36, "GT2 Belt 400mm"
    AddCfg 37, 37, "GT2 Belt 400mm"
    AddCfg 38, 38, "GT2 Belt 400mm"
    AddCfg 39, 39, "GT2 Belt 400mm"
    AddCfg 40, 40, "M4 50mm - Bolts"
    AddCfg 41, 41, "M4 60mm - Bolts"
    AddCfg 42, 42, "M8 45mm - Bolts"
    AddCfg 43, 43, "M4 Lock nut"
    AddCfg 44, 44, "M8 Lock nut"
    AddCfg 45, 45, "M8 Washer"
    AddCfg 46, 46, "Gripper to J3 connector"
    AddCfg 47, 47, "Servo Holder"
    AddCfg 48, 48, "Gripper Cover"
    AddCfg 49, 49, "Gripper Mechanism Slider"
    AddCfg 50, 50, "Gripper Hand Left"
    AddCfg 51, 51, "Gripper Hand Left"
    AddCfg 52, 52, "Gripper link 1"
    AddCfg 53, 53, "Gripper link 2"
    AddCfg 54, 54, "Gripper End"
    AddCfg 55, 55, "Wide gripper end"
    AddCfg 56, 56, "Wide gripper end - Left side"
    AddCfg 57, 57, "Gripper rail 6mm"
    AddCfg 58, 58, "Servo Motor MG996R 3D Model"
    AddCfg 59, 59, "Servo MG996R Horn"
    AddCfg 60, 60, "Gripper Mechanism"
    AddCfg 61, 61, "SCARA Robot"

    ' --- phase 1 order: ids are already children-first (Base id=1 first) ---
    For k = 1 To gNOrder
        AddOrder k, k
    Next k

    ' --- phase 2 windows: gripper (deepest) then root ---
    AddWin 1, 60
    AddWin 2, 61

    ' --- parent/child pairs (60): gripper children, then root children ---
    n = 0
    For c = 46 To 59
        n = n + 1: AddPair n, 60, c
    Next c
    For c = 1 To 45
        n = n + 1: AddPair n, 61, c
    Next c
    n = n + 1: AddPair n, 61, 60
End Sub

Private Sub AddFile(ByVal id As Long, ByVal act As String, ByVal dt As String, ByVal depth As Long, _
                    ByVal oldStem As String, ByVal newStem As String, ByVal ma As String, ByVal mt As String, _
                    ByVal nr As Boolean, ByVal desc As String, ByVal cat As String)
    Dim ext As String
    ext = IIf(dt = "Part", ".SLDPRT", ".SLDASM")
    Files(id).Id = id
    Files(id).Action = act
    Files(id).DocType = dt
    Files(id).Depth = depth
    Files(id).OldPath = PROJ_DIR & "\" & oldStem & ext
    Files(id).NewPath = PROJ_DIR & "\" & newStem & ext
    Files(id).MatAction = ma
    Files(id).MatTarget = mt
    Files(id).NeedsReview = nr
    Files(id).PropNames(1) = "Project":     Files(id).PropValues(1) = PROJECT_NAME
    Files(id).PropNames(2) = "PartNo":      Files(id).PropValues(2) = newStem
    Files(id).PropNames(3) = "Description": Files(id).PropValues(3) = desc
    Files(id).PropNames(4) = "Category":    Files(id).PropValues(4) = cat
    Files(id).PropNames(5) = "Revision":    Files(id).PropValues(5) = REV
    Files(id).PropNames(6) = "Material":    Files(id).PropValues(6) = mt
    Files(id).Renamed = False
    Files(id).UsedFallback = False
End Sub

Private Sub AddCfg(ByVal id As Long, ByVal fileId As Long, ByVal oldC As String)
    Cfgs(id).FileId = fileId
    Cfgs(id).OldCfg = oldC
    Cfgs(id).NewCfg = NEW_CFG
    Cfgs(id).Done = False
    Cfgs(id).Skipped = False
End Sub

Private Sub AddOrder(ByVal n As Long, ByVal fid As Long)
    OrderIds(n) = fid
End Sub

Private Sub AddWin(ByVal n As Long, ByVal fid As Long)
    Wins(n) = fid
End Sub

Private Sub AddPair(ByVal n As Long, ByVal parId As Long, ByVal childId As Long)
    PairPar(n) = parId
    PairChild(n) = childId
End Sub