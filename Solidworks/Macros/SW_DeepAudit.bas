Attribute VB_Name = "SW_DeepAudit"
'==============================================================================
' SW_DeepAudit.bas  -  SolidWorks 2026 recursive assembly deep-inspection macro
'
' Output (next to the saved assembly):
'   <Assembly>_DeepAudit.json
'   <Assembly>_DeepAudit.md
'   <Assembly>_Audit_Assets\Thumbnails\<Part>.bmp   (optional)
'
' Install: SolidWorks > Tools > Macro > New (or edit a .swp) > in the VBA editor
'          use File > Import File... and pick this .bas, then run  main.
' Binding: 100% late binding (Object). No SolidWorks type-library references
'          are needed; every enum value is defined locally below.
' Entry point: Public Sub main()
'==============================================================================
Option Explicit

'------------------------------ USER CONFIGURATION ----------------------------
Private Const AUTO_RESOLVE_LIGHTWEIGHT As Boolean = True
Private Const EXPORT_PART_THUMBNAILS As Boolean = True
Private Const MAX_FACES_ANALYZED As Long = 25000   ' per part; larger parts skip the per-face surface scan
Private Const THUMB_WIDTH As Long = 800
Private Const THUMB_HEIGHT As Long = 600

'------------------------------ SolidWorks enums ------------------------------
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swSolidBody As Long = 0
Private Const SW_COMP_SUPPRESSED As Long = 0
Private Const SW_COMP_LIGHTWEIGHT As Long = 1
Private Const SW_COMP_FULLYRESOLVED As Long = 2
Private Const SW_COMP_RESOLVED As Long = 3
Private Const SW_COMP_FULLYLIGHTWEIGHT As Long = 4
Private Const SW_COMP_IDMISMATCH As Long = 5
Private Const PI As Double = 3.14159265358979

'------------------------------ JSON/Markdown buffers -------------------------
Private Type TBuf
    a() As String
    n As Long
End Type
Private jb As TBuf
Private mb As TBuf

'------------------------------ Module state ----------------------------------
Private swApp As Object
Private swModel As Object
Private mFso As Object
Private mComps As Collection        ' ordered component records
Private mCompDict As Object         ' id -> component record
Private mUnique As Object           ' doc key -> unique document record
Private mUniqueDocs As Object       ' doc key -> ModelDoc2
Private mCompPaths As Object        ' set of lower-case component file paths
Private mMates As Collection
Private mRefs As Collection
Private mWarn As Collection
Private mPropCache As Object
Private mTree As Object
Private mAsmInfo As Object
Private mRootPath As String
Private mRootTitle As String
Private mRootName As String
Private mRootKey As String
Private mOutDir As String
Private mAssetsDir As String
Private mTreeOff As Boolean
Private mGfxOff As Boolean
Private mFailed As Boolean
Private mReportsWritten As Boolean
Private mCntParts As Long
Private mCntSubAsm As Long
Private mCntVirtual As Long
Private mCntSuppressed As Long
Private mCntUnresolved As Long
Private mCntBadMates As Long
Private mMaxDepth As Long
Private mStart As Double

'==============================================================================
'  ENTRY POINT
'==============================================================================
Public Sub main()
    Dim cfg As Object, rootComp As Object, msg As String
    mStart = Timer
    On Error GoTo ErrorHandler

    Set mFso = CreateObject("Scripting.FileSystemObject")
    Set swApp = GetObject(, "SldWorks.Application")
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then Err.Raise vbObjectError + 1001, "SW_DeepAudit", "No active document. Open an assembly first."
    If CLng(swModel.GetType) <> swDocASSEMBLY Then Err.Raise vbObjectError + 1002, "SW_DeepAudit", "The active document is not an assembly (.SLDASM)."

    InitState
    mRootPath = CStr(swModel.GetPathName)
    mRootTitle = CStr(swModel.GetTitle)
    If Len(mRootPath) > 0 Then
        mRootName = FileBase(mRootPath)
        mOutDir = mFso.GetParentFolderName(mRootPath)
        mRootKey = LCase$(mRootPath)
    Else
        mRootName = FileBase(mRootTitle)
        mOutDir = Environ$("USERPROFILE") & "\Documents"
        mRootKey = "unsaved:" & LCase$(mRootTitle)
        Warn "Assembly has never been saved; reports are written to " & mOutDir
    End If
    mRootName = SafeFileName(mRootName)
    mAssetsDir = mOutDir & "\" & mRootName & "_Audit_Assets"

    If AUTO_RESOLVE_LIGHTWEIGHT Then
        On Error Resume Next
        swModel.ResolveAllLightWeightComponents False
        If Err.Number <> 0 Then Warn "ResolveAllLightWeightComponents failed: " & Err.Description: Err.Clear
        On Error GoTo ErrorHandler
    End If

    DisableUI
    Set cfg = swModel.ConfigurationManager.ActiveConfiguration
    Set rootComp = cfg.GetRootComponent3(True)
    If rootComp Is Nothing Then Err.Raise vbObjectError + 1003, "SW_DeepAudit", "Could not obtain the root component."

    CollectRootInfo
    Set mTree = WalkComponent(rootComp, 0, "", "")
    PhaseMates
    PhaseExternalRefs
    PhaseGeometry
    RestoreUI
    If EXPORT_PART_THUMBNAILS Then PhaseThumbnails
    WriteReports

    msg = "Deep audit complete in " & Format$(Timer - mStart, "0.0") & " s." & vbCrLf & vbCrLf & _
          "Components: " & mComps.Count & "   Mates: " & mMates.Count & "   Warnings: " & mWarn.Count & vbCrLf & vbCrLf & _
          mOutDir & "\" & mRootName & "_DeepAudit.json" & vbCrLf & _
          mOutDir & "\" & mRootName & "_DeepAudit.md"

Cleanup:
    On Error Resume Next
    RestoreUI
    If mFailed And Not mReportsWritten And Not mComps Is Nothing Then
        Warn "Writing PARTIAL reports after a fatal error."
        WriteReports
    End If
    If Len(msg) > 0 Then MsgBox msg, IIf(mFailed, vbExclamation, vbInformation), "SW Deep Audit"
    Exit Sub

ErrorHandler:
    mFailed = True
    msg = "SW Deep Audit stopped: " & Err.Description & " (Err " & Err.Number & ")"
    If Not mWarn Is Nothing Then Warn msg
    Resume Cleanup
End Sub

Private Sub InitState()
    Set mComps = New Collection
    Set mCompDict = NewDict()
    Set mUnique = NewDict()
    Set mUniqueDocs = NewDict()
    Set mCompPaths = NewDict()
    Set mMates = New Collection
    Set mRefs = New Collection
    Set mWarn = New Collection
    Set mPropCache = NewDict()
    Set mAsmInfo = NewDict()
    mCntParts = 0: mCntSubAsm = 0: mCntVirtual = 0: mCntSuppressed = 0
    mCntUnresolved = 0: mCntBadMates = 0: mMaxDepth = 0
    mFailed = False: mReportsWritten = False
    jb.n = 0: mb.n = 0
End Sub

Private Sub DisableUI()
    On Error Resume Next
    mTreeOff = False: mGfxOff = False
    swModel.FeatureManager.EnableFeatureTree = False
    If Err.Number = 0 Then mTreeOff = True
    Err.Clear
    swModel.ActiveView.EnableGraphicsUpdate = False
    If Err.Number = 0 Then mGfxOff = True
    Err.Clear
End Sub

Private Sub RestoreUI()
    On Error Resume Next
    If mTreeOff Then swModel.FeatureManager.EnableFeatureTree = True: mTreeOff = False
    If mGfxOff Then swModel.ActiveView.EnableGraphicsUpdate = True: mGfxOff = False
    Err.Clear
End Sub

Private Sub Warn(ByVal s As String)
    On Error Resume Next
    mWarn.Add s
End Sub

'==============================================================================
'  SAFE LATE-BOUND COM HELPERS
'==============================================================================
Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

' Property / zero-arg method returning a scalar or array
Private Function PV(ByVal o As Object, ByVal nm As String) As Variant
    Dim r As Variant
    On Error Resume Next
    If o Is Nothing Then Exit Function
    Err.Clear
    r = CallByName(o, nm, VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        r = CallByName(o, nm, VbMethod)
        If Err.Number <> 0 Then Err.Clear: Exit Function
    End If
    PV = r
End Function

' Property / zero-arg method returning an object
Private Function PO(ByVal o As Object, ByVal nm As String) As Object
    Dim r As Object
    On Error Resume Next
    If o Is Nothing Then Exit Function
    Err.Clear
    Set r = CallByName(o, nm, VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        Set r = Nothing
        Set r = CallByName(o, nm, VbMethod)
        If Err.Number <> 0 Then Err.Clear: Set r = Nothing
    End If
    Set PO = r
End Function

Private Function SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Or IsObject(v) Or IsArray(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function PStr(ByVal o As Object, ByVal nm As String) As String
    PStr = SafeStr(PV(o, nm))
End Function

Private Function PLng(ByVal o As Object, ByVal nm As String, ByVal def As Long) As Long
    Dim v As Variant
    PLng = def
    On Error Resume Next
    v = PV(o, nm)
    If IsEmpty(v) Or IsNull(v) Or IsArray(v) Or IsObject(v) Then Exit Function
    If IsNumeric(v) Then PLng = CLng(v)
End Function

Private Function PDbl(ByVal o As Object, ByVal nm As String) As Double
    Dim v As Variant
    On Error Resume Next
    v = PV(o, nm)
    If IsEmpty(v) Or IsNull(v) Or IsArray(v) Or IsObject(v) Then Exit Function
    If IsNumeric(v) Then PDbl = CDbl(v)
End Function

Private Function PBool(ByVal o As Object, ByVal nm As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = PV(o, nm)
    If IsEmpty(v) Or IsNull(v) Or IsArray(v) Or IsObject(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        PBool = v
    ElseIf IsNumeric(v) Then
        PBool = (CLng(v) <> 0)
    End If
End Function

Private Function R6(ByVal d As Double) As Double
    On Error Resume Next
    R6 = Round(d, 6)
End Function

Private Function DS(ByVal d As Object, ByVal k As String) As String
    On Error Resume Next
    If d.Exists(k) Then DS = SafeStr(d(k))
End Function

Private Function FileBase(ByVal p As String) As String
    Dim s As String, i As Long
    s = p
    i = InStrRev(s, "\"): If i > 0 Then s = Mid$(s, i + 1)
    i = InStrRev(s, "/"): If i > 0 Then s = Mid$(s, i + 1)
    i = InStrRev(s, "."): If i > 1 Then s = Left$(s, i - 1)
    FileBase = s
End Function

Private Function FileExt(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, ".")
    If i > 0 Then FileExt = LCase$(Mid$(p, i + 1))
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = 0 To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    s = Trim$(s)
    If Len(s) = 0 Then s = "Unnamed"
    SafeFileName = s
End Function

Private Function CleanModelName(ByVal inst As String) As String
    Dim s As String, p As Long
    s = inst
    p = InStrRev(s, "/"): If p > 0 Then s = Mid$(s, p + 1)
    p = InStr(s, "<"): If p > 1 Then s = Left$(s, p - 1)
    p = InStr(s, "^"): If p > 1 Then s = Left$(s, p - 1)
    p = InStrRev(s, "-")
    If p > 1 Then
        If IsNumeric(Mid$(s, p + 1)) Then s = Left$(s, p - 1)
    End If
    CleanModelName = s
End Function

Private Function ActiveConfigName(ByVal doc As Object) As String
    Dim cm As Object, c As Object
    Set cm = PO(doc, "ConfigurationManager")
    Set c = PO(cm, "ActiveConfiguration")
    ActiveConfigName = PStr(c, "Name")
End Function

Private Sub EnsureFolder(ByVal p As String)
    Dim par As String
    On Error Resume Next
    If mFso.FolderExists(p) Then Exit Sub
    par = mFso.GetParentFolderName(p)
    If Len(par) > 0 Then
        If Not mFso.FolderExists(par) Then EnsureFolder par
    End If
    mFso.CreateFolder p
End Sub

'==============================================================================
'  ENUM -> STRING MAPPERS
'==============================================================================
Private Function SuppressionName(ByVal s As Long) As String
    Select Case s
        Case SW_COMP_SUPPRESSED: SuppressionName = "Suppressed"
        Case SW_COMP_LIGHTWEIGHT: SuppressionName = "Lightweight"
        Case SW_COMP_FULLYRESOLVED: SuppressionName = "FullyResolved"
        Case SW_COMP_RESOLVED: SuppressionName = "Resolved"
        Case SW_COMP_FULLYLIGHTWEIGHT: SuppressionName = "FullyLightweight"
        Case SW_COMP_IDMISMATCH: SuppressionName = "InternalIdMismatch"
        Case Else: SuppressionName = "Unknown(" & s & ")"
    End Select
End Function

Private Function MateTypeName(ByVal t As Long, ByVal featType As String, ByVal isLimit As Boolean) As String
    Dim s As String
    Select Case t
        Case 0: s = "Coincident"
        Case 1: s = "Concentric"
        Case 2: s = "Perpendicular"
        Case 3: s = "Parallel"
        Case 4: s = "Tangent"
        Case 5: s = "Distance"
        Case 6: s = "Angle"
        Case 7: s = "Unknown"
        Case 8: s = "Symmetric"
        Case 9: s = "CamFollower"
        Case 10: s = "Gear"
        Case 11: s = "Width"
        Case 12: s = "LockToSketch"
        Case 13: s = "RackPinion"
        Case 15: s = "Path"
        Case 16: s = "Lock"
        Case 17: s = "Screw"
        Case 18: s = "LinearCoupler"
        Case 19: s = "UniversalJoint"
        Case 20: s = "Coordinate"
        Case 21: s = "Slot"
        Case 22: s = "Hinge"
        Case 23: s = "Slider"
        Case Else: s = "Other(" & t & ")"
    End Select
    If InStr(1, featType, "ProfileCenter", vbTextCompare) > 0 Then s = "ProfileCenter"
    If InStr(1, featType, "InPlace", vbTextCompare) > 0 Then s = "InPlace"
    If isLimit Then
        If t = 5 Then
            s = "Limit Distance"
        ElseIf t = 6 Then
            s = "Limit Angle"
        Else
            s = "Limit (" & s & ")"
        End If
    End If
    MateTypeName = s
End Function

Private Function AlignName(ByVal a As Long) As String
    Select Case a
        Case 0: AlignName = "Aligned"
        Case 1: AlignName = "Anti-Aligned"
        Case 2: AlignName = "Closest"
        Case Else: AlignName = "N/A"
    End Select
End Function

Private Function EntityTypeName(ByVal t As Long) As String
    Select Case t
        Case 0: EntityTypeName = "Unknown"
        Case 1: EntityTypeName = "Point"
        Case 2: EntityTypeName = "Line"
        Case 3: EntityTypeName = "Circle"
        Case 4: EntityTypeName = "Plane"
        Case 5: EntityTypeName = "Cylinder"
        Case 6: EntityTypeName = "Sphere"
        Case 7: EntityTypeName = "Set"
        Case 8: EntityTypeName = "Cone"
        Case 9: EntityTypeName = "SweptSurface"
        Case Else: EntityTypeName = "Other(" & t & ")"
    End Select
End Function

Private Function HoleTypeName(ByVal t As Long) As String
    Select Case t
        Case 0: HoleTypeName = "Counterbore"
        Case 1: HoleTypeName = "Countersink"
        Case 2: HoleTypeName = "Hole"
        Case 3: HoleTypeName = "Straight Tap (Pipe)"
        Case 4: HoleTypeName = "Tap"
        Case 5: HoleTypeName = "Legacy"
        Case 6: HoleTypeName = "Counterbore Slot"
        Case 7: HoleTypeName = "Countersink Slot"
        Case 8: HoleTypeName = "Hole Slot"
        Case 9: HoleTypeName = "Pipe Tap Slot"
        Case 10: HoleTypeName = "Tap Slot"
        Case 11: HoleTypeName = "Legacy Slot"
        Case -1: HoleTypeName = "Unknown"
        Case Else: HoleTypeName = "Type " & t
    End Select
End Function

Private Function CustomTypeName(ByVal t As Long) As String
    Select Case t
        Case 3: CustomTypeName = "Number"
        Case 5: CustomTypeName = "Double"
        Case 11: CustomTypeName = "YesNo"
        Case 30: CustomTypeName = "Text"
        Case 64: CustomTypeName = "Date"
        Case Else: CustomTypeName = "Unknown"
    End Select
End Function

Private Function ExtToDocType(ByVal ext As String) As String
    Select Case ext
        Case "sldasm": ExtToDocType = "Assembly"
        Case "sldprt": ExtToDocType = "Part"
        Case "slddrw": ExtToDocType = "Drawing"
        Case Else: ExtToDocType = "Other(" & ext & ")"
    End Select
End Function

'==============================================================================
'  PROPERTIES / ROOT INFO
'==============================================================================
Private Function CollectProps(ByVal doc As Object, ByVal cfgName As String) As Collection
    Dim res As Collection, cpm As Object, nms As Variant, i As Long
    Dim n As String, v As String, r As String, w As Boolean, l As Boolean, tp As Long, d As Object
    Set res = New Collection
    On Error Resume Next
    Set cpm = Nothing
    Set cpm = doc.Extension.CustomPropertyManager(cfgName)
    If cpm Is Nothing Then Set CollectProps = res: Exit Function
    nms = PV(cpm, "GetNames")
    If IsArray(nms) Then
        For i = LBound(nms) To UBound(nms)
            n = CStr(nms(i))
            v = "": r = "": w = False: l = False: tp = 0
            Err.Clear
            cpm.Get6 n, False, v, r, w, l
            If Err.Number <> 0 Then
                Err.Clear
                v = SafeStr(cpm.Get(n)): r = v
            End If
            Err.Clear
            tp = CLng(cpm.GetType2(n))
            If Err.Number <> 0 Then tp = 0: Err.Clear
            Set d = NewDict()
            d("name") = n
            d("type") = CustomTypeName(tp)
            d("value") = v
            d("resolved_value") = r
            res.Add d
        Next i
    End If
    Set CollectProps = res
End Function

Private Function CachedProps(ByVal doc As Object, ByVal path As String, ByVal cfgName As String) As Collection
    Dim k As String
    k = LCase$(path) & "|" & cfgName
    If Len(path) = 0 Then k = "doc:" & PStr(doc, "GetTitle") & "|" & cfgName
    If mPropCache.Exists(k) Then
        Set CachedProps = mPropCache(k)
    Else
        Set CachedProps = CollectProps(doc, cfgName)
        mPropCache.Add k, CachedProps
    End If
End Function

Private Function GetMass(ByVal doc As Object) As Object
    Dim ext As Object, mp As Object, d As Object, com As Variant
    Set d = NewDict()
    On Error Resume Next
    Set ext = PO(doc, "Extension")
    Set mp = PO(ext, "CreateMassProperty2")
    If mp Is Nothing Then Set mp = PO(ext, "CreateMassProperty")
    If mp Is Nothing Then
        d("available") = False
        Set GetMass = d
        Exit Function
    End If
    mp.UseSystemUnits = True
    Err.Clear
    d("available") = True
    d("mass_kg") = R6(PDbl(mp, "Mass"))
    d("volume_mm3") = R6(PDbl(mp, "Volume") * 1000000000#)
    d("surface_area_mm2") = R6(PDbl(mp, "SurfaceArea") * 1000000#)
    com = PV(mp, "CenterOfMass")
    If IsArray(com) Then
        d("center_of_mass_mm") = Array(R6(CDbl(com(LBound(com))) * 1000#), R6(CDbl(com(LBound(com) + 1)) * 1000#), R6(CDbl(com(LBound(com) + 2)) * 1000#))
    End If
    Set GetMass = d
End Function

Private Sub CollectRootInfo()
    Dim cfgs As Variant, i As Long, c As Object, byCfg As Object, ver As Variant
    Dim ds As Variant, active As String
    On Error Resume Next
    mAsmInfo("file_path") = mRootPath
    mAsmInfo("file_name") = IIf(Len(mRootPath) > 0, Mid$(mRootPath, InStrRev(mRootPath, "\") + 1), mRootTitle)
    ver = Empty
    ver = swApp.GetDocumentVersion(mRootPath)
    If Err.Number <> 0 Then Err.Clear
    mAsmInfo("file_version_raw") = SafeStr(ver)
    mAsmInfo("solidworks_revision") = SafeStr(PV(swApp, "RevisionNumber"))
    If Len(mRootPath) > 0 Then
        If mFso.FileExists(mRootPath) Then mAsmInfo("last_saved") = Format$(FileDateTime(mRootPath), "yyyy-mm-dd hh:nn:ss")
    End If
    If Err.Number <> 0 Then Err.Clear
    active = ActiveConfigName(swModel)
    mAsmInfo("active_configuration") = active
    cfgs = PV(swModel, "GetConfigurationNames")
    mAsmInfo("configurations") = IIf(IsArray(cfgs), cfgs, Array())
    Set c = PO(PO(swModel, "ConfigurationManager"), "ActiveConfiguration")
    ds = PV(c, "GetDisplayStates")
    mAsmInfo("display_states_active_configuration") = IIf(IsArray(ds), ds, Array())
    Set mAsmInfo("custom_properties_file") = CachedProps(swModel, mRootPath, "")
    Set byCfg = NewDict()
    If IsArray(cfgs) Then
        For i = LBound(cfgs) To UBound(cfgs)
            Set byCfg(CStr(cfgs(i))) = CachedProps(swModel, mRootPath, CStr(cfgs(i)))
        Next i
    End If
    Set mAsmInfo("custom_properties_by_configuration") = byCfg
    Set mAsmInfo("mass_properties") = GetMass(swModel)
End Sub

'==============================================================================
'  RECURSIVE COMPONENT TRAVERSAL
'==============================================================================
Private Sub RegisterUnique(ByVal key As String, ByVal nm As String, ByVal path As String, ByVal docType As String, _
        ByVal isVirt As Boolean, ByVal md As Object, ByVal isRoot As Boolean)
    Dim rec As Object
    On Error Resume Next
    If Not mUnique.Exists(key) Then
        Set rec = NewDict()
        rec("key") = key
        rec("name") = nm
        rec("file_path") = path
        rec("doc_type") = docType
        rec("is_virtual") = isVirt
        rec("is_root") = isRoot
        rec("instances") = 0
        rec("resolved") = False
        mUnique.Add key, rec
    End If
    Set rec = mUnique(key)
    rec("instances") = CLng(rec("instances")) + 1
    If Not md Is Nothing Then
        If Not mUniqueDocs.Exists(key) Then mUniqueDocs.Add key, md
        rec("resolved") = True
    End If
End Sub

Private Function WalkComponent(ByVal comp As Object, ByVal depth As Long, _
        ByVal parentName As String, ByVal parentDocKey As String) As Object
    Dim node As Object, rec As Object, flags As Object, kids As Collection
    Dim nm As String, inst As String, clean As String, path As String, id As String, key As String
    Dim supp As Long, sv As Variant, isVirt As Boolean, md As Object, docType As String, dtv As Variant
    Dim refCfg As String, mat As String, db As String, xf As Object, ad As Variant
    Dim posV As Variant, rotV As Variant, cfgs As Variant, ch As Variant, c As Object, childNode As Object
    Dim i As Long, vis As Long, isFixed As Boolean, isEnv As Boolean, exBom As Boolean, dispState As String
    Dim suppressed As Boolean

    On Error Resume Next
    Set node = NewDict(): Set rec = NewDict(): Set flags = NewDict(): Set kids = New Collection

    nm = PStr(comp, "Name2")
    inst = nm
    If InStrRev(inst, "/") > 0 Then inst = Mid$(inst, InStrRev(inst, "/") + 1)
    clean = CleanModelName(nm)
    isVirt = PBool(comp, "IsVirtual")
    If depth = 0 Then
        path = mRootPath
        Set md = swModel
        If Len(inst) = 0 Then inst = mRootName
        clean = mRootName
    Else
        path = PStr(comp, "GetPathName")
    End If
    id = PStr(comp, "GetSelectByIDString")
    If depth = 0 Then id = "ROOT"
    If Len(id) = 0 Then id = parentName & "/" & inst
    If mCompDict.Exists(id) Then id = id & "#" & mCompDict.Count

    sv = PV(comp, "GetSuppression2")
    If IsEmpty(sv) Or Not IsNumeric(sv) Then supp = -1 Else supp = CLng(sv)
    suppressed = (supp = SW_COMP_SUPPRESSED)

    If depth > 0 And Not suppressed Then Set md = PO(comp, "GetModelDoc2")

    docType = ""
    If Not md Is Nothing Then
        dtv = PV(md, "GetType")
        If IsNumeric(dtv) Then
            If CLng(dtv) = swDocPART Then docType = "Part"
            If CLng(dtv) = swDocASSEMBLY Then docType = "Assembly"
        End If
    End If
    If Len(docType) = 0 Then docType = ExtToDocType(FileExt(path))
    If Left$(docType, 5) = "Other" And InStr(inst, "^") = 0 Then docType = "Unknown"

    key = IIf(depth = 0, mRootKey, IIf(isVirt Or Len(path) = 0, "virtual:" & parentDocKey & ">" & clean, LCase$(path)))
    If depth > 0 And Not isVirt And Len(path) > 0 Then mCompPaths(LCase$(path)) = True
    If depth = 0 And Len(mRootPath) > 0 Then mCompPaths(LCase$(mRootPath)) = True

    isFixed = PBool(comp, "IsFixed")
    isEnv = PBool(comp, "IsEnvelope")
    exBom = PBool(comp, "ExcludeFromBOM")
    vis = PLng(comp, "Visible", -1)
    refCfg = PStr(comp, "ReferencedConfiguration")
    dispState = PStr(comp, "ReferencedDisplayState2")
    If Len(dispState) = 0 Then dispState = PStr(comp, "ReferencedDisplayState")

    flags("Suppressed") = (supp = SW_COMP_SUPPRESSED)
    flags("Lightweight") = (supp = SW_COMP_LIGHTWEIGHT Or supp = SW_COMP_FULLYLIGHTWEIGHT)
    flags("FullyResolved") = (supp = SW_COMP_FULLYRESOLVED)
    flags("Resolved") = (supp = SW_COMP_RESOLVED Or supp = SW_COMP_FULLYRESOLVED)
    flags("InternalIdMismatch") = (supp = SW_COMP_IDMISMATCH)

    Set xf = PO(comp, "Transform2")
    ad = PV(xf, "ArrayData")
    If IsArray(ad) Then
        If UBound(ad) - LBound(ad) >= 11 Then
            posV = Array(R6(CDbl(ad(LBound(ad) + 9)) * 1000#), R6(CDbl(ad(LBound(ad) + 10)) * 1000#), R6(CDbl(ad(LBound(ad) + 11)) * 1000#))
            rotV = Array(R6(CDbl(ad(LBound(ad)))), R6(CDbl(ad(LBound(ad) + 1))), R6(CDbl(ad(LBound(ad) + 2))), _
                         R6(CDbl(ad(LBound(ad) + 3))), R6(CDbl(ad(LBound(ad) + 4))), R6(CDbl(ad(LBound(ad) + 5))), _
                         R6(CDbl(ad(LBound(ad) + 6))), R6(CDbl(ad(LBound(ad) + 7))), R6(CDbl(ad(LBound(ad) + 8))))
        End If
    End If

    cfgs = Empty
    mat = ""
    If Not md Is Nothing Then
        cfgs = PV(md, "GetConfigurationNames")
        If docType = "Part" Then
            db = ""
            Err.Clear
            mat = SafeStr(md.GetMaterialPropertyName2(refCfg, db))
            If Err.Number <> 0 Then Err.Clear: mat = ""
        End If
    End If

    ' ---- record ----
    rec("id") = id
    rec("depth") = depth
    rec("parent") = parentName
    rec("instance_name") = inst
    rec("model_name") = clean
    rec("file_path") = path
    rec("file_name") = IIf(Len(path) > 0, Mid$(path, InStrRev(path, "\") + 1), "")
    rec("doc_type") = docType
    rec("is_virtual") = isVirt
    rec("storage") = IIf(isVirt, "Virtual (stored inside parent assembly)", "External file")
    rec("suppression_state") = SuppressionName(supp)
    Set rec("suppression_flags") = flags
    rec("fixation") = IIf(isFixed, "Fixed [f]", "Floating [-]")
    rec("is_envelope") = isEnv
    rec("exclude_from_bom") = exBom
    rec("visibility") = IIf(vis = 1, "Shown", IIf(vis = 0, "Hidden", "Unknown"))
    rec("referenced_configuration") = refCfg
    rec("available_configurations") = IIf(IsArray(cfgs), cfgs, Array())
    rec("display_state") = dispState
    If IsArray(posV) Then rec("position_mm") = posV
    If IsArray(rotV) Then rec("rotation_matrix") = rotV
    rec("material") = mat
    rec("resolved") = Not (md Is Nothing)
    rec("unique_key") = key
    If Not md Is Nothing Then
        Set rec("custom_properties_file") = CachedProps(md, path, "")
        Set rec("custom_properties_configuration") = CachedProps(md, path, refCfg)
    Else
        Set rec("custom_properties_file") = New Collection
        Set rec("custom_properties_configuration") = New Collection
        mCntUnresolved = mCntUnresolved + 1
    End If
    mComps.Add rec
    mCompDict.Add id, rec

    ' ---- counters / registry ----
    If depth > mMaxDepth Then mMaxDepth = depth
    If depth > 0 Then
        If docType = "Part" Then mCntParts = mCntParts + 1
        If docType = "Assembly" Then mCntSubAsm = mCntSubAsm + 1
        If isVirt Then mCntVirtual = mCntVirtual + 1
        If suppressed Then mCntSuppressed = mCntSuppressed + 1
    End If
    If Not suppressed Or depth = 0 Then
        RegisterUnique key, clean, path, docType, isVirt, md, (depth = 0)
    End If

    ' ---- children ----
    If Not suppressed Then
        ch = PV(comp, "GetChildren")
        If IsArray(ch) Then
            For i = LBound(ch) To UBound(ch)
                Set c = Nothing
                Set c = ch(i)
                If Not c Is Nothing Then
                    Set childNode = Nothing
                    Set childNode = WalkComponent(c, depth + 1, inst, key)
                    If Not childNode Is Nothing Then kids.Add childNode
                End If
            Next i
        End If
    End If

    node("id") = id
    node("depth") = depth
    node("instance_name") = inst
    node("model_name") = clean
    node("doc_type") = docType
    node("suppression_state") = SuppressionName(supp)
    node("is_virtual") = isVirt
    Set node("children") = kids
    Set WalkComponent = node
End Function

'==============================================================================
'  MATES (root + every resolved sub-assembly, each unique assembly once)
'==============================================================================
Private Sub PhaseMates()
    Dim k As Variant, rec As Object
    On Error Resume Next
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Assembly" Then
            If mUniqueDocs.Exists(CStr(k)) Then
                ScanAssemblyMates mUniqueDocs(CStr(k)), CStr(k), rec
            Else
                Warn "Mates not read for unresolved assembly: " & DS(rec, "name")
            End If
        End If
    Next k
End Sub

Private Sub ScanAssemblyMates(ByVal doc As Object, ByVal asmKey As String, ByVal asmRec As Object)
    Dim feat As Object, sub1 As Object, tn As String, cfgName As String
    Dim guard As Long
    On Error Resume Next
    cfgName = ActiveConfigName(doc)
    Set feat = PO(doc, "FirstFeature")
    Do While Not feat Is Nothing
        tn = PStr(feat, "GetTypeName2")
        If tn = "MateGroup" Then
            Set sub1 = PO(feat, "GetFirstSubFeature")
            Do While Not sub1 Is Nothing
                ProcessMate sub1, asmKey, asmRec, cfgName
                Set sub1 = PO(sub1, "GetNextSubFeature")
            Loop
        End If
        Set feat = PO(feat, "GetNextFeature")
    Loop
End Sub

Private Function FeatSuppressed(ByVal feat As Object, ByVal cfgName As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = feat.IsSuppressed2(1, Array(cfgName))
    If Err.Number = 0 Then
        If IsArray(v) Then FeatSuppressed = CBool(v(LBound(v))): Exit Function
        If Not IsEmpty(v) Then FeatSuppressed = CBool(v): Exit Function
    End If
    Err.Clear
    FeatSuppressed = PBool(feat, "IsSuppressed")
End Function

Private Sub ProcessMate(ByVal feat As Object, ByVal asmKey As String, ByVal asmRec As Object, ByVal cfgName As String)
    Dim m As Object, rec As Object, ents As Collection, ent As Object, ed As Object, rc As Object
    Dim tn As String, mt As Long, al As Long, n As Long, i As Long, code As Long, isW As Boolean
    Dim state As String, supp As Boolean, dims As Collection, dd As Object, nxt As Object, dm As Object, dd2 As Object
    Dim dv As Double, unit As String, angular As Boolean, maxV As Double, minV As Double, isLimit As Boolean
    Dim prm As Variant, prmTxt As String, j As Long, guard As Long, et As Long, haveMate As Boolean

    On Error Resume Next
    tn = PStr(feat, "GetTypeName2")
    Set m = PO(feat, "GetSpecificFeature2")
    haveMate = Not (m Is Nothing)
    mt = -1: al = -1
    If haveMate Then
        mt = PLng(m, "Type", -1)
        al = PLng(m, "Alignment", -1)
        maxV = PDbl(m, "MaximumVariation")
        minV = PDbl(m, "MinimumVariation")
    End If
    isLimit = haveMate And (Abs(maxV - minV) > 0.000000001) And (mt = 5 Or mt = 6)
    angular = (mt = 6) Or (InStr(1, tn, "Angle", vbTextCompare) > 0)

    supp = FeatSuppressed(feat, cfgName)
    code = 0: isW = False
    Err.Clear
    code = CLng(feat.GetErrorCode2(isW))
    If Err.Number <> 0 Then code = 0: Err.Clear
    If supp Then
        state = "Suppressed"
    ElseIf code <> 0 And isW Then
        state = "Warning"
    ElseIf code <> 0 Then
        state = "Error (Broken / Over-defining)"
    Else
        state = "Active (OK)"
    End If
    If state <> "Active (OK)" And state <> "Suppressed" Then mCntBadMates = mCntBadMates + 1

    ' numeric parameters
    Set dims = New Collection
    Set dd = PO(feat, "GetFirstDisplayDimension")
    guard = 0
    Do While Not dd Is Nothing And guard < 25
        guard = guard + 1
        Set dm = Nothing
        Set dm = dd.GetDimension2(0)
        If Not dm Is Nothing Then
            dv = PDbl(dm, "SystemValue")
            Set ed = NewDict()
            ed("name") = PStr(dm, "FullName")
            If angular Then
                ed("value") = R6(dv * 180# / PI): ed("unit") = "deg"
            Else
                ed("value") = R6(dv * 1000#): ed("unit") = "mm"
            End If
            dims.Add ed
        End If
        Set nxt = Nothing
        Set nxt = feat.GetNextDisplayDimension(dd)
        Set dd = nxt
    Loop
    Err.Clear

    ' linked entities
    Set ents = New Collection
    If haveMate Then
        n = PLng(m, "GetMateEntityCount", -1)
        If n < 0 Then n = PLng(m, "GetMateEntitiesCount", -1)
        If n < 0 Then n = 0
        For i = 0 To n - 1
            Set ent = Nothing
            Err.Clear
            Set ent = m.MateEntity(i)
            If Err.Number <> 0 Then Err.Clear
            If Not ent Is Nothing Then
                Set ed = NewDict()
                Set rc = PO(ent, "ReferenceComponent")
                ed("index") = i
                If rc Is Nothing Then
                    ed("component") = "(top-level assembly / origin)"
                    ed("component_path") = ""
                Else
                    ed("component") = PStr(rc, "Name2")
                    ed("component_path") = PStr(rc, "GetPathName")
                End If
                et = PLng(ent, "ReferenceType2", -1)
                If et < 0 Then et = PLng(ent, "ReferenceType", -1)
                ed("entity_type") = EntityTypeName(IIf(et < 0, 0, et))
                prm = PV(ent, "EntityParams")
                If IsArray(prm) Then
                    prmTxt = ""
                    For j = LBound(prm) To UBound(prm)
                        prmTxt = prmTxt & IIf(Len(prmTxt) > 0, ", ", "") & CStr(R6(CDbl(prm(j))))
                    Next j
                    ed("entity_params") = prm
                    ed("entity_params_text") = prmTxt
                End If
                ents.Add ed
            End If
        Next i
    End If

    Set rec = NewDict()
    rec("parent_assembly") = DS(asmRec, "name")
    rec("parent_assembly_path") = DS(asmRec, "file_path")
    rec("parent_assembly_key") = asmKey
    rec("name") = PStr(feat, "Name")
    rec("feature_type") = tn
    rec("mate_type") = MateTypeName(mt, tn, isLimit)
    rec("state") = state
    rec("error_code") = code
    rec("alignment") = AlignName(al)
    If haveMate Then
        If PBool(m, "CanBeFlipped") Then rec("flipped") = PBool(m, "Flipped")
    End If
    Set rec("dimensions") = dims
    If isLimit Then
        If angular Then
            rec("maximum_variation") = R6(maxV * 180# / PI)
            rec("minimum_variation") = R6(minV * 180# / PI)
            rec("limit_unit") = "deg"
        Else
            rec("maximum_variation") = R6(maxV * 1000#)
            rec("minimum_variation") = R6(minV * 1000#)
            rec("limit_unit") = "mm"
        End If
    End If
    Set rec("entities") = ents
    mMates.Add rec
End Sub

'==============================================================================
'  EXTERNAL / IN-CONTEXT / TOP-DOWN REFERENCES
'==============================================================================
Private Sub PhaseExternalRefs()
    Dim k As Variant, rec As Object, p As String
    On Error Resume Next
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        p = DS(rec, "file_path")
        If Len(p) > 0 And Not CBool(rec("is_virtual")) Then
            If mFso.FileExists(p) Then
                ScanExternalRefs p, DS(rec, "name"), DS(rec, "doc_type")
            End If
        End If
    Next k
End Sub

Private Sub ScanExternalRefs(ByVal ownerPath As String, ByVal ownerName As String, ByVal ownerType As String)
    Dim deps As Variant, i As Long, s As String, seen As Object, rec As Object
    Dim inTree As Boolean, ext As String, cls As String, st As String
    On Error Resume Next
    Err.Clear
    deps = swApp.GetDocumentDependencies2(ownerPath, False, True, False)
    If Err.Number <> 0 Then
        Warn "GetDocumentDependencies2 failed for " & ownerName & ": " & Err.Description
        Err.Clear
        Exit Sub
    End If
    If Not IsArray(deps) Then Exit Sub
    Set seen = NewDict()
    For i = LBound(deps) To UBound(deps)
        s = SafeStr(deps(i))
        If InStr(s, "\") > 0 Or InStr(s, "/") > 0 Then
            If LCase$(s) <> LCase$(ownerPath) And Not seen.Exists(LCase$(s)) Then
                seen.Add LCase$(s), True
                inTree = mCompPaths.Exists(LCase$(s))
                ext = FileExt(s)
                cls = ""
                If ownerType = "Assembly" Then
                    If Not inTree Then cls = "External reference outside the component tree"
                Else
                    If ext = "sldasm" Then
                        cls = "In-Context candidate (part depends on an assembly)"
                    ElseIf ext = "sldprt" Then
                        cls = "Derived / linked part reference (top-down)"
                    Else
                        cls = "External file link"
                    End If
                End If
                If Len(cls) > 0 Then
                    st = IIf(mFso.FileExists(s), "Resolved (file found)", "Broken (file not found)")
                    Set rec = NewDict()
                    rec("owner") = ownerName
                    rec("owner_path") = ownerPath
                    rec("owner_type") = ownerType
                    rec("references") = s
                    rec("reference_doc_type") = ExtToDocType(ext)
                    rec("classification") = cls
                    rec("status") = st
                    rec("referencing_feature") = "n/a (document-level dependency; per-feature detail not exposed)"
                    mRefs.Add rec
                End If
            End If
        End If
    Next i
End Sub

'==============================================================================
'  PART GEOMETRY / FEATURE DNA
'==============================================================================
Private Sub PhaseGeometry()
    Dim k As Variant, rec As Object, doc As Object, g As Object, cfgName As String
    Dim arche As String, desc As String, mat As String, cfgs As Variant
    On Error Resume Next
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If mUniqueDocs.Exists(CStr(k)) Then
            Set doc = mUniqueDocs(CStr(k))
            cfgName = ActiveConfigName(doc)
            rec("active_configuration") = cfgName
            cfgs = PV(doc, "GetConfigurationNames")
            rec("configurations") = IIf(IsArray(cfgs), cfgs, Array())
            If DS(rec, "doc_type") = "Part" Then
                arche = "": desc = "": mat = ""
                Set g = AnalyzePart(doc, cfgName, DS(rec, "name"), arche, desc, mat)
                rec("material") = mat
                rec("archetype") = arche
                rec("auto_shape_description") = desc
                Set rec("geometry") = g
                rec("geometry_note") = "Geometry reflects the part document's active configuration in memory (" & cfgName & ")."
            End If
        ElseIf DS(rec, "doc_type") = "Part" Then
            rec("analysis_status") = "Skipped: document not resolved (suppressed / missing / lightweight)."
        End If
    Next k
End Sub

Private Function IsNoiseFeature(ByVal tn As String) As Boolean
    Select Case tn
        Case "ProfileFeature", "3DProfileFeature", "RefPlane", "RefAxis", "RefPoint", "CoordSys", _
             "OriginProfileFeature", "Annotations", "DetailCabinet", "MaterialFolder", "CommentsFolder", _
             "HistoryFolder", "SelectionSetFolder", "SensorFolder", "DocsFolder", "FtrFolder", "EnvFolder", _
             "SolidBodyFolder", "SurfaceBodyFolder", "LightFolder", "CameraFolder"
            IsNoiseFeature = True
        Case Else
            IsNoiseFeature = (Len(tn) >= 6 And Right$(tn, 6) = "Folder")
    End Select
End Function

Private Sub Bump(ByVal d As Object, ByVal k As String)
    If d.Exists(k) Then d(k) = CLng(d(k)) + 1 Else d.Add k, 1
End Sub

Private Function EstimateHoleCount(ByVal feat As Object) As Long
    Dim s As Object, sk As Object, pts As Variant, best As Long, c As Long
    On Error Resume Next
    Set s = PO(feat, "GetFirstSubFeature")
    Do While Not s Is Nothing
        Set sk = PO(s, "GetSpecificFeature2")
        If Not sk Is Nothing Then
            pts = PV(sk, "GetSketchPoints2")
            If IsArray(pts) Then
                c = UBound(pts) - LBound(pts) + 1
                If c > best Then best = c
            End If
        End If
        Set s = PO(s, "GetNextSubFeature")
    Loop
    EstimateHoleCount = best
End Function

Private Function ParseFastener(ByVal featName As String) As String
    Dim re As Object, mt As Object
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(M\d+(\.\d+)?(x\d+(\.\d+)?)?|#\d+-\d+|\d+/\d+-\d+|\d+(\.\d+)?\s*mm)"
    re.IgnoreCase = True
    Set mt = re.Execute(featName)
    If mt.Count > 0 Then ParseFastener = mt(0).Value
End Function

Private Function AnalyzePart(ByVal doc As Object, ByVal cfgName As String, ByVal partName As String, _
        ByRef archetype As String, ByRef desc As String, ByRef material As String) As Object
    Dim g As Object, bb As Object, topo As Object, sf As Object, ratios As Object, feats As Object, hist As Object
    Dim mass As Object, holes As Collection, hr As Object, defn As Object
    Dim bodies As Variant, b As Object, bx As Variant, i As Long, nB As Long, nF As Long, nE As Long
    Dim x0 As Double, y0 As Double, z0 As Double, x1 As Double, y1 As Double, z1 As Double, haveBox As Boolean
    Dim cP As Long, cC As Long, cK As Long, cS As Long, cT As Long, cB As Long, cO As Long, cX As Long
    Dim f As Object, su As Object, idn As Long, scanned As Boolean
    Dim dx As Double, dy As Double, dz As Double, dm(0 To 2) As Double, t As Double, a As Long, c As Long
    Dim el As Double, fl As Double, so As Double, vol As Double, boxVol As Double
    Dim feat As Object, tn As String, db As String
    Dim isSheet As Boolean, isWeld As Boolean, isImp As Boolean, modelCnt As Long
    Dim nExt As Long, nCut As Long, nRev As Long, nSweep As Long, nLoft As Long, nHole As Long
    Dim nFil As Long, nCham As Long, nPat As Long, nMir As Long, nShell As Long, nHoleCnt As Long

    On Error Resume Next
    Set g = NewDict(): Set bb = NewDict(): Set topo = NewDict(): Set sf = NewDict()
    Set ratios = NewDict(): Set feats = NewDict(): Set hist = NewDict(): Set holes = New Collection

    ' ---- material ----
    db = ""
    Err.Clear
    material = SafeStr(doc.GetMaterialPropertyName2(cfgName, db))
    If Err.Number <> 0 Then Err.Clear: material = ""
    If Len(material) = 0 Then material = "(none assigned)"

    ' ---- bodies / bbox / topology ----
    Err.Clear
    bodies = doc.GetBodies2(swSolidBody, True)
    If Err.Number <> 0 Then Err.Clear
    If IsArray(bodies) Then
        nB = UBound(bodies) - LBound(bodies) + 1
        For i = LBound(bodies) To UBound(bodies)
            Set b = Nothing
            Set b = bodies(i)
            If Not b Is Nothing Then
                bx = PV(b, "GetBodyBox")
                If IsArray(bx) Then
                    If Not haveBox Then
                        x0 = bx(0): y0 = bx(1): z0 = bx(2): x1 = bx(3): y1 = bx(4): z1 = bx(5): haveBox = True
                    Else
                        If bx(0) < x0 Then x0 = bx(0)
                        If bx(1) < y0 Then y0 = bx(1)
                        If bx(2) < z0 Then z0 = bx(2)
                        If bx(3) > x1 Then x1 = bx(3)
                        If bx(4) > y1 Then y1 = bx(4)
                        If bx(5) > z1 Then z1 = bx(5)
                    End If
                End If
                nF = nF + PLng(b, "GetFaceCount", 0)
                nE = nE + PLng(b, "GetEdgeCount", 0)
            End If
        Next i
        If nF <= MAX_FACES_ANALYZED Then
            scanned = True
            For i = LBound(bodies) To UBound(bodies)
                Set b = Nothing
                Set b = bodies(i)
                Set f = PO(b, "GetFirstFace")
                Do While Not f Is Nothing
                    Set su = PO(f, "GetSurface")
                    idn = PLng(su, "Identity", -1)
                    Select Case idn
                        Case 4001: cP = cP + 1
                        Case 4002: cC = cC + 1
                        Case 4003: cK = cK + 1
                        Case 4004: cS = cS + 1
                        Case 4005: cT = cT + 1
                        Case 4006: cB = cB + 1
                        Case 4008: cO = cO + 1
                        Case Else: cX = cX + 1
                    End Select
                    Set f = PO(f, "GetNextFace")
                Loop
            Next i
        End If
    End If
    If Not haveBox Then
        bx = doc.GetPartBox(True)
        If IsArray(bx) Then
            x0 = bx(0): y0 = bx(1): z0 = bx(2): x1 = bx(3): y1 = bx(4): z1 = bx(5): haveBox = True
        End If
        Err.Clear
    End If
    dx = (x1 - x0) * 1000#: dy = (y1 - y0) * 1000#: dz = (z1 - z0) * 1000#
    dm(0) = dx: dm(1) = dy: dm(2) = dz
    For a = 0 To 1
        For c = 0 To 1 - a
            If dm(c) < dm(c + 1) Then t = dm(c): dm(c) = dm(c + 1): dm(c + 1) = t
        Next c
    Next a

    Set mass = GetMass(doc)
    vol = 0
    If mass.Exists("volume_mm3") Then vol = CDbl(mass("volume_mm3"))
    boxVol = dx * dy * dz
    el = 0: fl = 0: so = 0
    If dm(1) > 0 Then el = dm(0) / dm(1): fl = dm(2) / dm(1)
    If boxVol > 0 Then so = vol / boxVol

    bb("dx_mm") = R6(dx): bb("dy_mm") = R6(dy): bb("dz_mm") = R6(dz)
    bb("Dim_Max") = R6(dm(0)): bb("Dim_Mid") = R6(dm(1)): bb("Dim_Min") = R6(dm(2))
    bb("note") = "Axis-aligned box in the part's own coordinate system."
    ratios("elongation_ratio") = R6(el)
    ratios("flatness_ratio") = R6(fl)
    ratios("solidity_ratio") = R6(so)
    sf("scan_performed") = scanned
    sf("planar") = cP: sf("cylindrical") = cC: sf("conical") = cK: sf("spherical") = cS
    sf("toroidal") = cT: sf("bspline_nurbs") = cB: sf("offset") = cO: sf("other") = cX
    topo("solid_bodies") = nB: topo("faces") = nF: topo("edges") = nE
    Set topo("surface_types") = sf

    ' ---- feature tree DNA ----
    Set feat = PO(doc, "FirstFeature")
    Do While Not feat Is Nothing
        tn = PStr(feat, "GetTypeName2")
        If Len(tn) > 0 Then
            Bump hist, tn
            If Not IsNoiseFeature(tn) Then modelCnt = modelCnt + 1
            If tn = "SheetMetal" Or tn = "FlatPattern" Or tn = "SMBaseFlange" Or tn = "BaseFlange" Then isSheet = True
            If InStr(1, tn, "Weld", vbTextCompare) > 0 Or tn = "StructuralMember" Then isWeld = True
            If InStr(1, tn, "Import", vbTextCompare) > 0 Or tn = "BaseBody" Then isImp = True
            Select Case tn
                Case "Extrusion": nExt = nExt + 1
                Case "Cut": nCut = nCut + 1
                Case "RevBoss", "Revolution": nRev = nRev + 1
                Case "RevCut": nRev = nRev + 1: nCut = nCut + 1
                Case "Sweep": nSweep = nSweep + 1
                Case "SweepCut": nSweep = nSweep + 1: nCut = nCut + 1
                Case "Loft": nLoft = nLoft + 1
                Case "LoftCut": nLoft = nLoft + 1: nCut = nCut + 1
                Case "Fillet": nFil = nFil + 1
                Case "Chamfer": nCham = nCham + 1
                Case "LPattern", "CirPattern", "MirrorPattern", "TablePattern", "SketchPattern", "CurvePattern", "FillPattern", "DerivedLPattern"
                    nPat = nPat + 1
                Case "MirrorSolid": nMir = nMir + 1
                Case "Shell": nShell = nShell + 1
                Case "HoleWzd"
                    nHole = nHole + 1
                    Set hr = NewDict()
                    hr("feature_name") = PStr(feat, "Name")
                    Set defn = PO(feat, "GetDefinition")
                    hr("hole_type") = HoleTypeName(PLng(defn, "Type", -1))
                    hr("standard_code") = PLng(defn, "Standard2", -1)
                    hr("fastener_size_code") = PLng(defn, "FastenerSize", -1)
                    hr("fastener_size_from_name") = ParseFastener(PStr(feat, "Name"))
                    nHoleCnt = EstimateHoleCount(feat)
                    hr("hole_count_estimated") = nHoleCnt
                    holes.Add hr
            End Select
        End If
        Set feat = PO(feat, "GetNextFeature")
    Loop
    feats("is_sheet_metal") = isSheet
    feats("is_weldment") = isWeld
    feats("is_imported_body") = isImp
    feats("modeling_feature_count") = modelCnt
    feats("extrusions") = nExt: feats("cuts") = nCut: feats("revolves") = nRev
    feats("sweeps") = nSweep: feats("lofts") = nLoft: feats("hole_wizard") = nHole
    feats("fillets") = nFil: feats("chamfers") = nCham: feats("patterns") = nPat
    feats("mirrors") = nMir: feats("shells") = nShell
    Set feats("hole_wizard_details") = holes
    Set feats("feature_type_histogram") = hist

    Set g("bounding_box_mm") = bb
    g("volume_mm3") = vol
    If mass.Exists("surface_area_mm2") Then g("surface_area_mm2") = mass("surface_area_mm2")
    If mass.Exists("mass_kg") Then g("mass_kg") = mass("mass_kg")
    Set g("shape_ratios") = ratios
    Set g("topology") = topo
    Set g("features") = feats

    desc = GenerateSemanticPartDescription(g, material, partName, archetype)
    Set AnalyzePart = g
End Function

Private Function ClassifyArchetype(ByVal g As Object) As String
    Dim feats As Object, topo As Object, sf As Object, ratios As Object
    Dim isSheet As Boolean, isWeld As Boolean, isImp As Boolean, fc As Long
    Dim cyl As Long, pl As Long, el As Double, fl As Double, so As Double
    On Error Resume Next
    Set feats = g("features"): Set topo = g("topology"): Set sf = topo("surface_types"): Set ratios = g("shape_ratios")
    isSheet = CBool(feats("is_sheet_metal")): isWeld = CBool(feats("is_weldment")): isImp = CBool(feats("is_imported_body"))
    fc = CLng(feats("modeling_feature_count"))
    cyl = CLng(sf("cylindrical")): pl = CLng(sf("planar"))
    el = CDbl(ratios("elongation_ratio")): fl = CDbl(ratios("flatness_ratio")): so = CDbl(ratios("solidity_ratio"))
    If isSheet Then
        ClassifyArchetype = "Sheet Metal Component"
    ElseIf isWeld Then
        ClassifyArchetype = "Structural Weldment Profile"
    ElseIf isImp And fc <= 2 Then
        ClassifyArchetype = "Imported COTS / Standard Component"
    ElseIf cyl > pl And el > 3# Then
        ClassifyArchetype = "Rotational / Cylindrical Shaft, Pin, or Fastener"
    ElseIf fl < 0.15 And pl >= 6 Then
        ClassifyArchetype = "Flat Plate / Machined Panel"
    ElseIf el > 4# And pl > cyl Then
        ClassifyArchetype = "Elongated Bar / Linear Rail / Structural Beam"
    ElseIf so < 0.25 Then
        ClassifyArchetype = "Hollow Frame / Thin-Walled Bracket or Housing"
    Else
        ClassifyArchetype = "Prismatic Machined / 3D Structural Block"
    End If
End Function

Private Function GenerateSemanticPartDescription(ByVal g As Object, ByVal material As String, _
        ByVal partName As String, ByRef archetypeTag As String) As String
    Dim bb As Object, feats As Object, topo As Object, sf As Object, ratios As Object
    Dim s1 As String, s2 As String, dom As String, domN As Long, tot As Long
    Dim nm As Variant, v As Long, keys As Variant, i As Long, labels As Variant
    On Error Resume Next
    archetypeTag = ClassifyArchetype(g)
    Set bb = g("bounding_box_mm"): Set feats = g("features"): Set topo = g("topology")
    Set sf = topo("surface_types"): Set ratios = g("shape_ratios")
    keys = Array("planar", "cylindrical", "conical", "spherical", "toroidal", "bspline_nurbs", "offset", "other")
    labels = Array("planar", "cylindrical", "conical", "spherical", "toroidal", "freeform/NURBS", "offset", "other")
    For i = 0 To UBound(keys)
        v = CLng(sf(keys(i)))
        tot = tot + v
        If v > domN Then domN = v: dom = labels(i)
    Next i
    s1 = "'" & partName & "' is classified as: " & archetypeTag & " - overall " & _
         Format$(CDbl(bb("Dim_Max")), "0.0") & " x " & Format$(CDbl(bb("Dim_Mid")), "0.0") & " x " & _
         Format$(CDbl(bb("Dim_Min")), "0.0") & " mm (L x W x H, sorted), " & _
         Format$(CDbl(ratios("solidity_ratio")) * 100#, "0") & "% solid fill of its bounding box, " & _
         CLng(topo("solid_bodies")) & " solid " & IIf(CLng(topo("solid_bodies")) = 1, "body", "bodies") & ", material " & material & "."
    If tot > 0 Then
        s2 = "Its " & CLng(topo("faces")) & " faces are mostly " & dom & " (" & domN & " of " & tot & ")"
    Else
        s2 = "Face-type distribution was not scanned"
    End If
    s2 = s2 & "; the feature tree has " & CLng(feats("modeling_feature_count")) & " modeling features including " & _
         CLng(feats("cuts")) & " cut(s), " & CLng(feats("hole_wizard")) & " Hole Wizard feature(s), " & _
         CLng(feats("patterns")) & " pattern(s), " & CLng(feats("mirrors")) & " mirror(s), " & _
         CLng(feats("fillets")) & " fillet(s) and " & CLng(feats("chamfers")) & " chamfer(s)."
    GenerateSemanticPartDescription = s1 & " " & s2
End Function

'==============================================================================
'  THUMBNAILS
'==============================================================================
Private Sub PhaseThumbnails()
    Dim k As Variant, rec As Object, pd As Object, thumbDir As String, fn As String, base As String
    Dim errs As Long, e1 As Long, e2 As Long, opened As Boolean, title As String, ok As Boolean
    Dim used As Object, n As Long, ctr As Long

    On Error Resume Next
    thumbDir = mAssetsDir & "\Thumbnails"
    EnsureFolder thumbDir
    Set used = NewDict()
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Part" And Not CBool(rec("is_virtual")) And CBool(rec("resolved")) And Len(DS(rec, "file_path")) > 0 Then
            title = ""
            title = CStr(mUniqueDocs(CStr(k)).GetTitle)
            Set pd = Nothing
            opened = False
            Err.Clear
            Set pd = swApp.ActivateDoc3(title, False, 1, errs)
            If pd Is Nothing Then
                Err.Clear
                Set pd = swApp.OpenDoc6(DS(rec, "file_path"), swDocPART, 3, "", e1, e2)
                opened = Not (pd Is Nothing)
            End If
            If pd Is Nothing Then
                Warn "Thumbnail skipped (could not activate): " & DS(rec, "name")
            Else
                pd.ShowNamedView2 "*Isometric", 7
                pd.ViewZoomtofit2
                base = SafeFileName(DS(rec, "name"))
                ctr = 1
                Do While used.Exists(LCase$(base))
                    ctr = ctr + 1
                    base = SafeFileName(DS(rec, "name")) & "_" & ctr
                Loop
                used.Add LCase$(base), True
                fn = thumbDir & "\" & base & ".bmp"
                Err.Clear
                ok = pd.SaveBMP(fn, THUMB_WIDTH, THUMB_HEIGHT)
                If mFso.FileExists(fn) Then
                    rec("thumbnail") = mRootName & "_Audit_Assets/Thumbnails/" & base & ".bmp"
                    n = n + 1
                Else
                    Warn "SaveBMP failed for " & DS(rec, "name")
                End If
                If opened Then swApp.CloseDoc title
            End If
        End If
    Next k
    Err.Clear
    Set pd = swApp.ActivateDoc3(mRootTitle, False, 1, errs)
    Err.Clear
End Sub

'==============================================================================
'  JSON / MARKDOWN OUTPUT
'==============================================================================
Private Sub BAdd(ByRef b As TBuf, ByVal s As String)
    If b.n = 0 Then ReDim b.a(0 To 2047)
    If b.n > UBound(b.a) Then ReDim Preserve b.a(0 To (UBound(b.a) + 1) * 2 - 1)
    b.a(b.n) = s
    b.n = b.n + 1
End Sub

Private Function BText(ByRef b As TBuf) As String
    If b.n = 0 Then Exit Function
    ReDim Preserve b.a(0 To b.n - 1)
    BText = Join(b.a, "")
    b.n = 0
End Function

Private Function JStr(ByVal s As String) As String
    Dim i As Long
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    For i = 0 To 31
        If InStr(s, Chr$(i)) > 0 Then s = Replace(s, Chr$(i), "\u" & Right$("0000" & Hex$(i), 4))
    Next i
    JStr = """" & s & """"
End Function

Private Function JNum(ByVal d As Double) As String
    Dim s As String, dec As String
    dec = Mid$(CStr(1.5), 2, 1)
    s = CStr(Round(d, 6))
    If dec <> "." Then s = Replace(s, dec, ".")
    JNum = s
End Function

Private Sub WriteJson(ByVal v As Variant, ByVal ind As Long)
    Dim o As Object, k As Variant, i As Long, pad As String, pad2 As String, first As Boolean
    pad = Space$(ind * 2): pad2 = Space$((ind + 1) * 2)
    If IsObject(v) Then
        If v Is Nothing Then BAdd jb, "null": Exit Sub
        Set o = v
        Select Case TypeName(o)
            Case "Dictionary"
                If o.Count = 0 Then
                    BAdd jb, "{}"
                Else
                    BAdd jb, "{" & vbLf
                    first = True
                    For Each k In o.Keys
                        If Not first Then BAdd jb, "," & vbLf
                        first = False
                        BAdd jb, pad2 & JStr(CStr(k)) & ": "
                        WriteJson o.Item(k), ind + 1
                    Next k
                    BAdd jb, vbLf & pad & "}"
                End If
            Case "Collection"
                If o.Count = 0 Then
                    BAdd jb, "[]"
                Else
                    BAdd jb, "[" & vbLf
                    For i = 1 To o.Count
                        If i > 1 Then BAdd jb, "," & vbLf
                        BAdd jb, pad2
                        WriteJson o.Item(i), ind + 1
                    Next i
                    BAdd jb, vbLf & pad & "]"
                End If
            Case Else
                BAdd jb, "null"
        End Select
    ElseIf IsArray(v) Then
        On Error Resume Next
        Dim lb As Long, ub As Long
        lb = LBound(v): ub = UBound(v)
        If Err.Number <> 0 Or ub < lb Then
            Err.Clear
            BAdd jb, "[]"
        Else
            BAdd jb, "["
            For i = lb To ub
                If i > lb Then BAdd jb, ", "
                WriteJson v(i), ind + 1
            Next i
            BAdd jb, "]"
        End If
        On Error GoTo 0
    Else
        Select Case VarType(v)
            Case vbNull, vbEmpty: BAdd jb, "null"
            Case vbBoolean: BAdd jb, IIf(v, "true", "false")
            Case vbString: BAdd jb, JStr(CStr(v))
            Case vbByte, vbInteger, vbLong: BAdd jb, CStr(v)
            Case vbSingle, vbDouble, vbCurrency, vbDecimal: BAdd jb, JNum(CDbl(v))
            Case vbDate: BAdd jb, JStr(Format$(v, "yyyy-mm-dd hh:nn:ss"))
            Case Else: BAdd jb, JStr(CStr(v))
        End Select
    End If
End Sub

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object, bin As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.Position = 0
    st.Type = 1
    st.Position = 3                      ' skip the BOM
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    st.CopyTo bin
    bin.SaveToFile path, 2
    bin.Close
    st.Close
End Sub

Private Function BuildSummary() As Object
    Dim s As Object, k As Variant, rec As Object, uParts As Long, uAsm As Long, ctx As Long, i As Long
    Set s = NewDict()
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Part" Then uParts = uParts + 1
        If DS(rec, "doc_type") = "Assembly" Then uAsm = uAsm + 1
    Next k
    For i = 1 To mRefs.Count
        If InStr(DS(mRefs(i), "classification"), "In-Context") > 0 Or InStr(DS(mRefs(i), "classification"), "Derived") > 0 Then ctx = ctx + 1
    Next i
    s("total_component_instances_excluding_root") = mComps.Count - 1
    s("total_part_instances") = mCntParts
    s("unique_parts") = uParts
    s("sub_assembly_instances") = mCntSubAsm
    s("unique_assemblies_including_root") = uAsm
    s("virtual_components") = mCntVirtual
    s("suppressed_components") = mCntSuppressed
    s("unresolved_components") = mCntUnresolved
    s("max_depth") = mMaxDepth
    s("total_mates") = mMates.Count
    s("broken_or_warning_mates") = mCntBadMates
    s("in_context_or_top_down_references") = ctx
    s("external_reference_records") = mRefs.Count
    Set BuildSummary = s
End Function

Private Sub WriteReports()
    Dim root As Object, uniq As Object, k As Variant, rec As Object, jsonPath As String, mdPath As String
    Dim summ As Object, asms As Object, parts As Object

    On Error Resume Next
    Set summ = BuildSummary()
    Set parts = NewDict(): Set asms = NewDict()
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Part" Then parts.Add CStr(k), rec Else asms.Add CStr(k), rec
    Next k

    Set root = NewDict()
    root("tool") = "SW_DeepAudit.bas (SolidWorks VBA)"
    root("generated") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set root("assembly") = mAsmInfo
    Set root("summary") = summ
    Set root("tree") = mTree
    Set root("components") = mCompDict
    Set root("unique_parts") = parts
    Set root("assemblies") = asms
    Set root("mates") = mMates
    Set root("external_references") = mRefs
    Set root("warnings") = mWarn

    jsonPath = mOutDir & "\" & mRootName & "_DeepAudit.json"
    mdPath = mOutDir & "\" & mRootName & "_DeepAudit.md"

    jb.n = 0
    WriteJson root, 0
    BuildMarkdown summ
    Err.Clear
    EnsureFolder mOutDir
    WriteUtf8 jsonPath, BText(jb)
    If Err.Number <> 0 Then
        Warn "Primary output folder failed (" & Err.Description & "); using TEMP."
        Err.Clear
        mOutDir = Environ$("TEMP")
        jsonPath = mOutDir & "\" & mRootName & "_DeepAudit.json"
        mdPath = mOutDir & "\" & mRootName & "_DeepAudit.md"
        jb.n = 0
        WriteJson root, 0
        WriteUtf8 jsonPath, BText(jb)
    End If
    WriteUtf8 mdPath, BText(mb)
    mReportsWritten = True
End Sub

'------------------------------ Markdown --------------------------------------
Private Sub ML(ByVal s As String)
    BAdd mb, s & vbLf
End Sub

Private Function MdEsc(ByVal s As String) As String
    s = Replace(s, "|", "\|")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    MdEsc = s
End Function

Private Function YN(ByVal v As Variant) As String
    On Error Resume Next
    YN = IIf(CBool(v), "Yes", "No")
End Function

Private Function Num2(ByVal v As Variant) As String
    On Error Resume Next
    Num2 = Format$(CDbl(v), "0.00")
End Function

Private Sub MdTree(ByVal node As Object, ByVal prefix As String, ByVal isLast As Boolean, ByVal isRoot As Boolean)
    Dim ln As String, kids As Object, i As Long, childPrefix As String
    On Error Resume Next
    If isRoot Then
        ln = ""
        childPrefix = ""
    Else
        ln = prefix & IIf(isLast, "`-- ", "+-- ")
        childPrefix = prefix & IIf(isLast, "    ", "|   ")
    End If
    ln = ln & "[L" & node("depth") & "] " & node("instance_name") & "  ->  " & node("model_name") & "  (" & node("doc_type") & ")"
    If node("suppression_state") <> "Resolved" And node("suppression_state") <> "FullyResolved" Then ln = ln & "  {" & node("suppression_state") & "}"
    If CBool(node("is_virtual")) Then ln = ln & "  {Virtual}"
    ML ln
    Set kids = node("children")
    For i = 1 To kids.Count
        MdTree kids(i), childPrefix, (i = kids.Count), False
    Next i
End Sub

Private Function EntSummary(ByVal rec As Object) As String
    Dim ents As Object, i As Long, s As String, e As Object
    On Error Resume Next
    Set ents = rec("entities")
    For i = 1 To ents.Count
        Set e = ents(i)
        s = s & IIf(i > 1, " ; ", "") & DS(e, "component") & " [" & DS(e, "entity_type") & "]"
    Next i
    EntSummary = s
End Function

Private Function DimSummary(ByVal rec As Object) As String
    Dim d As Object, i As Long, s As String, e As Object
    On Error Resume Next
    Set d = rec("dimensions")
    For i = 1 To d.Count
        Set e = d(i)
        s = s & IIf(i > 1, "; ", "") & DS(e, "name") & "=" & DS(e, "value") & " " & DS(e, "unit")
    Next i
    If rec.Exists("maximum_variation") Then s = s & IIf(Len(s) > 0, "; ", "") & "limits " & DS(rec, "minimum_variation") & " .. " & DS(rec, "maximum_variation") & " " & DS(rec, "limit_unit")
    DimSummary = s
End Function

Private Sub BuildMarkdown(ByVal summ As Object)
    Dim i As Long, k As Variant, rec As Object, g As Object, bb As Object, topo As Object, ratios As Object
    Dim n As Long, any1 As Boolean, pos As Variant, ps As String, cfgs As Variant, ent As Object
    On Error Resume Next
    mb.n = 0
    ML "# Deep Audit Report - " & mRootName
    ML ""
    ML "- **Generated:** " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ML "- **Assembly file:** `" & DS(mAsmInfo, "file_path") & "`"
    ML "- **Last saved:** " & DS(mAsmInfo, "last_saved") & "   |   **SolidWorks revision:** " & DS(mAsmInfo, "solidworks_revision") & "   |   **File version (raw):** " & DS(mAsmInfo, "file_version_raw")
    ML "- **Active configuration:** " & DS(mAsmInfo, "active_configuration")
    cfgs = mAsmInfo("configurations")
    ps = ""
    If IsArray(cfgs) Then
        For i = LBound(cfgs) To UBound(cfgs)
            ps = ps & IIf(Len(ps) > 0, ", ", "") & CStr(cfgs(i))
        Next i
    End If
    ML "- **Configurations:** " & ps
    cfgs = mAsmInfo("display_states_active_configuration")
    ps = ""
    If IsArray(cfgs) Then
        For i = LBound(cfgs) To UBound(cfgs)
            ps = ps & IIf(Len(ps) > 0, ", ", "") & CStr(cfgs(i))
        Next i
    End If
    ML "- **Display states (active configuration):** " & ps
    Set g = mAsmInfo("mass_properties")
    If Not g Is Nothing Then
        If g.Exists("mass_kg") Then
            ps = ""
            If g.Exists("center_of_mass_mm") Then ps = "  CoM (mm): " & g("center_of_mass_mm")(0) & ", " & g("center_of_mass_mm")(1) & ", " & g("center_of_mass_mm")(2)
            ML "- **Mass properties:** mass " & g("mass_kg") & " kg, volume " & g("volume_mm3") & " mm3, surface area " & g("surface_area_mm2") & " mm2" & ps
        End If
    End If
    ML ""
    ML "## 1. Executive Summary"
    ML ""
    ML "| Metric | Value |"
    ML "|---|---|"
    ML "| Total part instances | " & summ("total_part_instances") & " |"
    ML "| Unique parts | " & summ("unique_parts") & " |"
    ML "| Sub-assembly instances | " & summ("sub_assembly_instances") & " |"
    ML "| Virtual components | " & summ("virtual_components") & " |"
    ML "| Suppressed components | " & summ("suppressed_components") & " |"
    ML "| Unresolved components (no model access) | " & summ("unresolved_components") & " |"
    ML "| Maximum tree depth | " & summ("max_depth") & " |"
    ML "| Total mates | " & summ("total_mates") & " |"
    ML "| Broken / warning mates | " & summ("broken_or_warning_mates") & " |"
    ML "| In-context / top-down references | " & summ("in_context_or_top_down_references") & " |"
    ML ""
    If mWarn.Count > 0 Then
        ML "**Warnings (" & mWarn.Count & "):**"
        ML ""
        For i = 1 To mWarn.Count
            ML "- " & MdEsc(CStr(mWarn(i)))
        Next i
        ML ""
    End If

    ML "## 2. Assembly Hierarchy"
    ML ""
    ML "```text"
    MdTree mTree, "", True, True
    ML "```"
    ML ""

    ML "## 3. Component Inventory"
    ML ""
    ML "| # | Depth | Instance | Model | Type | Virtual | State | Fixation | Envelope | BOM-excl. | Visible | Ref. config | Display state | Position mm (X, Y, Z) | Material |"
    ML "|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|"
    For i = 1 To mComps.Count
        Set rec = mComps(i)
        ps = ""
        pos = Empty
        If rec.Exists("position_mm") Then
            pos = rec("position_mm")
            ps = pos(0) & ", " & pos(1) & ", " & pos(2)
        End If
        ML "| " & i & " | " & DS(rec, "depth") & " | " & MdEsc(DS(rec, "instance_name")) & " | " & MdEsc(DS(rec, "model_name")) & " | " & DS(rec, "doc_type") & _
           " | " & YN(rec("is_virtual")) & " | " & DS(rec, "suppression_state") & " | " & DS(rec, "fixation") & " | " & YN(rec("is_envelope")) & _
           " | " & YN(rec("exclude_from_bom")) & " | " & DS(rec, "visibility") & " | " & MdEsc(DS(rec, "referenced_configuration")) & _
           " | " & MdEsc(DS(rec, "display_state")) & " | " & ps & " | " & MdEsc(DS(rec, "material")) & " |"
    Next i
    ML ""

    ML "## 4. Mates & Relationships (grouped by assembly)"
    ML ""
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Assembly" Then
            n = 0
            For i = 1 To mMates.Count
                If DS(mMates(i), "parent_assembly_key") = CStr(k) Then n = n + 1
            Next i
            ML "### " & MdEsc(DS(rec, "name")) & "  (instances: " & DS(rec, "instances") & ", mates: " & n & ")"
            ML ""
            If n > 0 Then
                ML "| Mate | Type | State | Alignment | Values / limits | Linked entities |"
                ML "|---|---|---|---|---|---|"
                For i = 1 To mMates.Count
                    Set ent = mMates(i)
                    If DS(ent, "parent_assembly_key") = CStr(k) Then
                        ML "| " & MdEsc(DS(ent, "name")) & " | " & DS(ent, "mate_type") & " | " & DS(ent, "state") & " | " & DS(ent, "alignment") & _
                           " | " & MdEsc(DimSummary(ent)) & " | " & MdEsc(EntSummary(ent)) & " |"
                    End If
                Next i
                ML ""
            Else
                ML "_No mates found (or assembly not resolved)._"
                ML ""
            End If
        End If
    Next k

    ML "## 5. External References & Skeleton / Top-Down Links"
    ML ""
    If mRefs.Count > 0 Then
        ML "| Owner | Owner type | References | Referenced type | Classification | Status |"
        ML "|---|---|---|---|---|---|"
        For i = 1 To mRefs.Count
            Set rec = mRefs(i)
            ML "| " & MdEsc(DS(rec, "owner")) & " | " & DS(rec, "owner_type") & " | `" & MdEsc(DS(rec, "references")) & "` | " & DS(rec, "reference_doc_type") & " | " & DS(rec, "classification") & " | " & DS(rec, "status") & " |"
        Next i
    Else
        ML "_No references outside the component tree were detected._"
    End If
    ML ""
    ML "> Detection is based on document dependencies. Feature-level in-context status (locked / out-of-context) is not exposed by this listing."
    ML ""

    ML "## 6. Part Geometry, Feature DNA and Previews"
    ML ""
    ML "| Part | Instances | Archetype | Dim_Max x Dim_Mid x Dim_Min (mm) | Solidity | Bodies | Faces | Material |"
    ML "|---|---|---|---|---|---|---|---|"
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Part" And rec.Exists("geometry") Then
            Set g = rec("geometry"): Set bb = g("bounding_box_mm"): Set topo = g("topology"): Set ratios = g("shape_ratios")
            ML "| " & MdEsc(DS(rec, "name")) & " | " & DS(rec, "instances") & " | " & DS(rec, "archetype") & " | " & Num2(bb("Dim_Max")) & " x " & Num2(bb("Dim_Mid")) & " x " & Num2(bb("Dim_Min")) & _
               " | " & Format$(CDbl(ratios("solidity_ratio")) * 100#, "0") & "% | " & topo("solid_bodies") & " | " & topo("faces") & " | " & MdEsc(DS(rec, "material")) & " |"
        End If
    Next k
    ML ""
    For Each k In mUnique.Keys
        Set rec = mUnique(k)
        If DS(rec, "doc_type") = "Part" Then
            ML "### " & MdEsc(DS(rec, "name"))
            ML ""
            If rec.Exists("geometry") Then
                Set g = rec("geometry"): Set bb = g("bounding_box_mm"): Set topo = g("topology"): Set ratios = g("shape_ratios")
                ML "**Archetype:** " & DS(rec, "archetype")
                ML ""
                ML DS(rec, "auto_shape_description")
                ML ""
                ML "- Bounding box (mm): dX " & Num2(bb("dx_mm")) & ", dY " & Num2(bb("dy_mm")) & ", dZ " & Num2(bb("dz_mm"))
                ML "- Volume " & Num2(g("volume_mm3")) & " mm3, surface area " & Num2(g("surface_area_mm2")) & " mm2"
                ML "- Elongation " & Num2(ratios("elongation_ratio")) & ", flatness " & Num2(ratios("flatness_ratio")) & ", solidity " & Num2(ratios("solidity_ratio"))
                ML "- Surfaces: planar " & topo("surface_types")("planar") & ", cylindrical " & topo("surface_types")("cylindrical") & ", conical " & topo("surface_types")("conical") & _
                   ", spherical " & topo("surface_types")("spherical") & ", toroidal " & topo("surface_types")("toroidal") & ", freeform " & topo("surface_types")("bspline_nurbs") & _
                   ", offset " & topo("surface_types")("offset") & ", other " & topo("surface_types")("other")
                ML "- Edges: " & topo("edges")
                If g("features")("hole_wizard_details").Count > 0 Then
                    ML "- Hole Wizard features:"
                    For i = 1 To g("features")("hole_wizard_details").Count
                        Set ent = g("features")("hole_wizard_details")(i)
                        ML "  - " & MdEsc(DS(ent, "feature_name")) & " | " & DS(ent, "hole_type") & " | size: " & DS(ent, "fastener_size_from_name") & " | holes (estimated): " & DS(ent, "hole_count_estimated")
                    Next i
                End If
            Else
                ML "_" & DS(rec, "analysis_status") & "_"
            End If
            If rec.Exists("thumbnail") Then
                ML ""
                ML "![Part Preview](" & Replace(DS(rec, "thumbnail"), " ", "%20") & ")"
            End If
            ML ""
        End If
    Next k
End Sub
