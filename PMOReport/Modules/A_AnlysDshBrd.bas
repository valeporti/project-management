Public exit_all As Boolean
Public Sub startDashboard()
    barraCarga.show
End Sub

Public Sub mainDshBrd()

Dim mainwb As Workbook
Set mainwb = ActiveWorkbook

'ultima columna y última fila en un inicio, antes de hacer nada en el archivo inicial
Dim last_row_st As Integer, last_col_st As Integer, deshacer As Boolean

Dim row As Integer, col As Integer
Dim rowprint As Integer
Dim posible_error As Integer

'Contadores
Dim num_registros As Integer

exit_all = False
'**
'OBTENCIÓN DE LAS RUTAS Y VERIFICAR ARCHIVOS
'**
    Call progressBar(True, 0, 0, 1)
    Dim lista_rutas() As String, dir_arch As Variant, num_arch As Integer, verificar As Boolean, funcionalidades() As String
    ReDim funcionalidades(6)
    'la funcionalidad "neutra" es la del archivo origen
    funcionalidades(0) = "Archivo Principal"
    funcionalidades(1) = "Portafolio Proyectos BOW"
    funcionalidades(2) = "Extracción PlanView de Proyectos"
    funcionalidades(3) = "Extracción Milestones"
    funcionalidades(4) = "Extracción Financiera"
    funcionalidades(5) = "Fechas de Releases"
    funcionalidades(6) = "LayOut para DashBoard"
    'Ruta de donde sacará archivos
    verificar = verificar_entorno(7, num_arch, mainwb, lista_rutas(), dir_arch, funcionalidades(), False)
        If (verificar = False) Then
            barraCarga.Hide
            Exit Sub
        End If
'**
'ANÁLISIS DE LOS ARCHIVOS
'**
    Call progressBar(True, 0, 0, 10)
    Call creaDshBrd(lista_rutas(), 3)
    If exit_all = True Then
        barraCarga.Hide
        Exit Sub
    End If
    
    
End Sub

'****/
'* La manera de hacer el dashboard, será:
'* Tener ya un workbook para el Dashboard, se puede crear pero sólo si se desea en un futuro
'* En el archivo financiero, sólo tomar valores de columnas del proyecto deseado y cerrar
'* En el archivo de milestones, se tiene que buscar el proyecto, y en la columna de milestones, buscar el milestone que nos interesa
'* En el archivo de extracción general, sólo tomar valores de columnas del proyecto deseado y cerrar
'****\
Public Sub creaDshBrd(archivos() As String, pbttl_row As Integer)
    Dim mlttl_row As Integer, fnttl_row As Integer, pbwb As Workbook, pvttl_row As Integer, lyttl_row As Integer, dtttl_row As Integer
    fnttl_row = 2 'más bien como 2 y 3
    mlttl_row = 2
    pvttl_row = 1
    lyttl_row = 3
    dtttl_row = 3 'la fila donde se encuentran los títulos que nos interesan del archivo de fechas
    
    'Se tomará como archivo base para lectura de #Proyecto, el protafolio de Proyectos en el BOW
    Workbooks.Open (archivos(0)) 'archivo portafolio BOW
    Set pbwb = ActiveWorkbook
    pbwb.Activate
    
    Dim LRbow As Integer, LCbow As Integer, Pwip As Worksheet, verq As Variant, ttlBowRng As Range
    verq = verUnaShExt(pbwb, Pwip, LRbow, LCbow, ttlBowRng, pbttl_row, 1, "Favor de escribir la hoja donde está la extracción de Proyectos de PlanView", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        exit_all = True
        Exit Sub
    End If
    'Mostrar todas las columnas
    Range(Columns(1), Columns(LCbow)).Select
    Selection.EntireColumn.Hidden = False
    'Buscar los tpitulos de las columnas deseadas en el rango
    Dim stPBc As Integer, ragPBc As Integer, idPBc As Integer, wtPBc As Integer, sdlcPBc As Integer, cfPBc As Integer _
        , swcapPBc As Integer, faPBc As Integer, pmPBc As Integer, relPBc As Integer, tmPBc As Integer, prPBc As Integer, _
        chdtPBc As Integer
    relPBc = busca_EnRango(ttlBowRng, "Release", "columna", True) 'Release
    tmPBc = busca_EnRango(ttlBowRng, "Equipo", "columna", True) 'team
    idPBc = busca_EnRango(ttlBowRng, "Work Id", "columna", True) '# Proyecto
    ragPBc = busca_EnRango(ttlBowRng, "RAG", "columna", True) 'Rag, para verificar para lo de si está en Completed
    stPBc = busca_EnRango(ttlBowRng, "Status", "columna", True) 'Status, para verificar para lo de si está en Completed
    pmPBc = busca_EnRango(ttlBowRng, "Project Mgr", "columna", True) 'Project Manager
    prPBc = busca_EnRango(ttlBowRng, "Program Mgr", "columna", True) 'Program Manager
    chdtPBc = busca_EnRango(ttlBowRng, "Fecha de cambio de Fase en Planview", "columna", True)
    sdlcPBc = busca_EnRango(ttlBowRng, "SDLC Phase", "columna", True) 'SDLC Phase
    'Verificar que todos los títulos tengan exitencia, sino salir
    If ((stPBc = 0) Or (ragPBc = 0) Or (idPBc = 0) Or (relPBc = 0) Or (tmPBc = 0) Or (pmPBc = 0) Or (prPBc = 0) _
        Or (chdtPBc = 0) Or (sdlcPBc = 0)) Then ' Or (wtPBc = 0) Or (cfPBc = 0) Or (swcapPBc = 0) Or (faPBc = 0) Or (pmPBc = 0) ) Then
        MsgBox "Se saldrá de la aplicación, hay uno o más titulos en el Portafolio de Proyectos BOW que no se encuentran"
        exit_all = True
        Exit Sub
    End If
    
    Call progressBar(True, 0, 0, 12)
    
    'Buscar en Extacción de PlanView
    Dim pvwb As Workbook, ttlPVRng As Range, LRpv As Integer, LCpv As Integer, pvws As Worksheet
    Workbooks.Open (archivos(1)) 'Extracción PV
    Set pvwb = ActiveWorkbook
    verq = verUnaShExt(pvwb, pvws, LRpv, LCpv, ttlPVRng, pvttl_row, 1, "Favor de escribir la hoja donde está la extracción de Proyectos de PlanView", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        exit_all = True
        Exit Sub
    End If
    
    Dim stPVc As Integer, idPVc As Integer, wtPVc As Integer, sdlcPVc As Integer, cfPVc As Integer, swcapPVc As Integer _
        , faPVc As Integer, pmPVc As Integer, evPVc As Integer, pvPVc As Integer, eaPVc As Integer, bePVc As Integer _
        , etPVc As Integer, pnPVc As Integer, stdtPVc As Integer, swculPVc As Integer, swcthPVc  As Integer, _
        ragPVc As Integer, issPVc As Integer, rskPVc As Integer, eddtPVc As Integer
    stPVc = busca_EnRango(ttlPVRng, "Work Status", "columna", True) 'Work Status (commited, targeted, cancelled)
    idPVc = busca_EnRango(ttlPVRng, "Work ID #", "columna", True) 'P00...
    wtPVc = busca_EnRango(ttlPVRng, "Work Type", "columna", True) 'Work Type ( major, minor, maintenance..)
    sdlcPVc = busca_EnRango(ttlPVRng, "SDLC Phase", "columna", True) 'SDLC Phase (Ini, Def, Des, ...)
    cfPVc = busca_EnRango(ttlPVRng, "Capitalization Flag", "columna", True) 'Capital Flag
    swcapPVc = busca_EnRango(ttlPVRng, "SWCAP Qualification", "columna", True) 'SWCAP Qualification (qualified, not started, disqualified)
    swculPVc = busca_EnRango(ttlPVRng, "SWCAP Useful Life", "columna", True) 'SWCAP Useful Life (complete , not started)
    swcthPVc = busca_EnRango(ttlPVRng, "SWCAP Cost Threshold", "columna", True) 'SWCAP Cost Threshold (qualified, started , not started)
    faPVc = busca_EnRango(ttlPVRng, "Finance Approval", "columna", True) 'Finance approval (undetermined, No, Yes)
    pmPVc = busca_EnRango(ttlPVRng, "Project Manager", "columna", True) 'Project Manager
    evPVc = busca_EnRango(ttlPVRng, "EV-Earned Value (h)", "columna", True) 'Earned Value
    pvPVc = busca_EnRango(ttlPVRng, "EV-Planned Value (h)", "columna", True) 'Planned Value
    eaPVc = busca_EnRango(ttlPVRng, "Effort Actual (h)", "columna", True) 'Effort Actual
    etPVc = busca_EnRango(ttlPVRng, "Effort Total (h)", "columna", True) 'Effort Total
    bePVc = busca_EnRango(ttlPVRng, "Baseline Effort (h)", "columna", True) 'Baseline Effort
    pnPVc = busca_EnRango(ttlPVRng, "Name", "columna", True) 'Project Name
    stdtPVc = busca_EnRango(ttlPVRng, "Actual Start", "columna", True) 'Actual Start Date
    ragPVc = busca_EnRango(ttlPVRng, "Overall RAG Status", "columna", True) 'RAG (green, Red, Ambeer)
    issPVc = busca_EnRango(ttlPVRng, "Issues", "columna", True) 'Issues (1,2,0)
    rskPVc = busca_EnRango(ttlPVRng, "Risks", "columna", True) 'Risks (1,2,0)
    eddtPVc = busca_EnRango(ttlPVRng, "Schedule Finish", "columna", True) 'Scheduled End Date

    If ((stPVc = 0) Or (idPVc = 0) Or (wtPVc = 0) Or (sdlcPVc = 0) Or (cfPVc = 0) Or (swcapPVc = 0) Or (faPVc = 0) Or _
        (pmPVc = 0) Or (evPVc = 0) Or (pvPVc = 0) Or (eaPVc = 0) Or (etPVc = 0) Or (bePVc = 0) Or (pnPVc = 0) Or (stdtPVc = 0) Or _
        (swculPVc = 0) Or (swcthPVc = 0) Or (ragPVc = 0) Or (issPVc = 0) Or (rskPVc = 0) Or (eddtPVc = 0)) Then
        MsgBox "Se saldrá de la aplicación, hay uno o más titulos en la extracción de Planview de proyectos que no se encuentran"
        exit_all = True
        Exit Sub
    End If
    
    Call progressBar(True, 0, 0, 15)
    
    'Buscar en Extacción de PlanView (Financial Navigator)
    Dim fnwb As Workbook, ttlFNRng As Range, LRfn As Integer, LCfn As Integer, fnws As Worksheet
    Workbooks.Open (archivos(3)) 'Extracción Financial Navigator
    Set fnwb = ActiveWorkbook
    verq = verUnaShExt(fnwb, fnws, LRfn, LCfn, ttlFNRng, fnttl_row, 1, "Favor de escribir la hoja donde está el Navegador Financiero", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        exit_all = True
        Exit Sub
    End If
    
    Dim ptFNc As Integer, e0FNc As Integer, e1FNc As Integer, e2FNc As Integer, mvFNc As Integer, meFNc  As Integer, eacFNc As Integer, aeFNc As Integer
    ptFNc = busca_EnRango(ttlFNRng, "Project Title", "columna", True) 'Título del proyecto completo, no hay P00
    e0FNc = busca_EnRango(ttlFNRng, "E0*", "columna", True)
    e1FNc = busca_EnRango(ttlFNRng, "E1*", "columna", True)
    e2FNc = busca_EnRango(ttlFNRng, "E2*", "columna", True)
    mvFNc = busca_EnRango(ttlFNRng, "Marked", "columna", True) 'Marked Version
    meFNc = busca_EnRango(ttlFNRng, "Marked Version", "columna", True) 'Marked Version Estimate
    aeFNc = busca_EnRango(ttlFNRng, "Actual*", "columna", True) 'ACTUAL
    eacFNc = busca_EnRango(ttlFNRng, "EAC*", "columna", True) 'EAC

    If ((ptFNc = 0) Or (e0FNc = 0) Or (e1FNc = 0) Or (e2FNc = 0) Or (mvFNc = 0) Or (meFNc = 0) Or (eacFNc = 0) Or (aeFNc = 0)) Then
        MsgBox "Se saldrá de la aplicación, hay uno o más titulos en la extracción de Financial Navigator que no se encuentran"
        exit_all = True
        Exit Sub
    End If
    
    Call progressBar(True, 0, 0, 17)
    
    'Buscar en Extacción de PlanView (Milestones)
    Dim mlwb As Workbook, ttlMLRng As Range, LRml As Integer, LCml As Integer, mlws As Worksheet
    Workbooks.Open (archivos(2)) 'Extracción MILESTONES
    Set mlwb = ActiveWorkbook
    verq = verUnaShExt(mlwb, mlws, LRml, LCml, ttlMLRng, mlttl_row, 1, "Favor de escribir la hoja donde están los Milestones", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        exit_all = True
        Exit Sub
    End If
    
    Dim idMLc As Integer, mlMLc As Integer, stMLc As Integer, bdMLc As Integer
    idMLc = busca_EnRango(ttlMLRng, "Work ID", "columna", True) 'donde están los Work IDs
    mlMLc = busca_EnRango(ttlMLRng, "Milestone", "columna", True)
    stMLc = busca_EnRango(ttlMLRng, "Schedule Start", "columna", True)
    bdMLc = busca_EnRango(ttlMLRng, "Baseline Date", "columna", True)

    If ((idMLc = 0) Or (mlMLc = 0) Or (stMLc = 0) Or (bdMLc = 0)) Then
        MsgBox "Se saldrá de la aplicación, hay uno o más titulos en la extracción de Milestones que no se encuentran"
        exit_all = True
        Exit Sub
    End If
    
    Call progressBar(True, 0, 0, 19)
    
    'Buscar en FECHAS
    Dim dtwb As Workbook, ttlDTmjRng As Range, ttlDTmnRng As Range, lrMn As Integer, lrMj As Integer, lcMn As Integer, lcMj As Integer
    Dim mjWs As Worksheet, mnWs As Worksheet, rMnRng As Range, rMjRng As Range
    Workbooks.Open (archivos(4)) 'FECHAS
    Set dtwb = ActiveWorkbook
    'En este caso no se hará busqueda de hoja pues se hará posteriormente para el match de la fase del proyecto tomado
    'Sin embargo, el entorno se puede ir tomando para agilidad
    verq = entornoDT(dtwb, mjWs, mnWs, ttlDTmjRng, ttlDTmnRng, dtttl_row, lrMn, lrMj, lcMn, lcMj, 2, rMnRng, rMjRng, 2)
    If verq = False Then
        exit_all = True
        Exit Sub
    End If
    
    Call progressBar(True, 0, 0, 21)
    
    'Buscar en LAYOUT
    Dim lywb As Workbook, ttlLYRng As Range, LRly As Integer, LCly As Integer, lyws As Worksheet
    Workbooks.Open (archivos(5)) 'LAYOUT
    Set lywb = ActiveWorkbook
    verq = verUnaShExt(lywb, lyws, LRly, LCly, ttlLYRng, lyttl_row, 1, "Favor de escribir la hoja donde está el Layout", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        exit_all = True
        Exit Sub
    End If
    
    Dim rlLYc As Integer, idLYc As Integer, pnLYc As Integer, tmLYc As Integer, pmLYc As Integer, wtLYc As Integer, wsLYc As Integer, _
        sdlcLYc As Integer, cfLYc As Integer, swcapLYc As Integer, faLYc As Integer, mspiLYc As Integer, _
        e0LYc As Integer, e1LYc As Integer, e2LYc As Integer, mvLYc As Integer, meLYc As Integer, evLYc As Integer, pvLYc As Integer, _
        eaLYc As Integer, blLYc As Integer, etLYc As Integer, cpiLYc As Integer, spiLYc As Integer, cnsLYc As Integer, remLYc As Integer, _
        eacLYc As Integer, veacLYc As Integer, mveacLYc As Integer, msibLYc As Integer, swculLYc As Integer, _
        swcthLYc As Integer, cpexLYc As Integer, onestLYc As Integer, ragLYc As Integer, issLYc As Integer, rskLYc As Integer, _
        issRskLYc  As Integer, relByLYc As Integer, eddtLYc As Integer, pmissLYc As Integer, cntLYc As Integer, wrMMLYc As Integer, _
        actfnLYc As Integer, eacfnLYc As Integer, cpissLYc As Integer, spissLYc As Integer, evissLYc As Integer, _
        mkestLYc As Integer, ontckLYc As Integer, stphLYc As Integer, phagdLYc As Integer, phagLYc As Integer, phagimpLYc As Integer, _
        phchLYc As Integer, noimpdtLYc As Integer
    rlLYc = busca_EnRango(ttlLYRng, "Release", "columna", True) 'Release
    idLYc = busca_EnRango(ttlLYRng, "Work ID", "columna", True) 'Work ID
    pnLYc = busca_EnRango(ttlLYRng, "Project Name", "columna", True) 'Project Name
    tmLYc = busca_EnRango(ttlLYRng, "Team", "columna", True) 'Team
    pmLYc = busca_EnRango(ttlLYRng, "Project Manager", "columna", True) 'Project Manager
    'prLYc = busca_EnRango(ttlLYRng, "Program Manager", "columna", True) 'Program Manager
    wtLYc = busca_EnRango(ttlLYRng, "Work Type", "columna", True) 'Work Type
    wsLYc = busca_EnRango(ttlLYRng, "Work Status", "columna", True) 'Work Status
    sdlcLYc = busca_EnRango(ttlLYRng, "SDLC Phase", "columna", True) 'SDLC Phase
    cfLYc = busca_EnRango(ttlLYRng, "Capitalization Flag", "columna", True) 'Capitalization Flag
    swcapLYc = busca_EnRango(ttlLYRng, "SWCAP Qualification", "columna", True) 'SWCAP Qualification
    swculLYc = busca_EnRango(ttlLYRng, "SWCAP Useful Life", "columna", True) 'SWCAP Useful Life
    swcthLYc = busca_EnRango(ttlLYRng, "SWCAP Cost Threshold", "columna", True) 'SWCAP Cost Threshold
    faLYc = busca_EnRango(ttlLYRng, "Finance Approval", "columna", True) 'Finance Approval
    mspiLYc = busca_EnRango(ttlLYRng, "MS Project Implementation", "columna", True) 'MS Project Implementation
    msibLYc = busca_EnRango(ttlLYRng, "MS Project Implementation (Baseline)", "columna", True) 'MS Project Implementation (Baseline)
    e0LYc = busca_EnRango(ttlLYRng, "E0 (h)", "columna", True) 'E0 (h)
    e1LYc = busca_EnRango(ttlLYRng, "E1 (h)", "columna", True) 'E1 (h)
    e2LYc = busca_EnRango(ttlLYRng, "E2 (h)", "columna", True) 'E2 (h)
    mvLYc = busca_EnRango(ttlLYRng, "Marked Version", "columna", True) 'Marked Version
    meLYc = busca_EnRango(ttlLYRng, "Marked Version Effort", "columna", True) 'Marked Version Effort
    evLYc = busca_EnRango(ttlLYRng, "EV-Earned Value (h)", "columna", True) 'EV-Earned Value (h)
    pvLYc = busca_EnRango(ttlLYRng, "EV-Planned Value (h)", "columna", True) 'EV-Planned Value (h)
    eaLYc = busca_EnRango(ttlLYRng, "Effort Actual (h)", "columna", True) 'Effort Actual (h)
    blLYc = busca_EnRango(ttlLYRng, "Baseline Effort (h)", "columna", True) 'Baseline Effort (h)
    etLYc = busca_EnRango(ttlLYRng, "Effort Total (h)", "columna", True) 'Effort Total (h)
    cpiLYc = busca_EnRango(ttlLYRng, "CPI", "columna", True) 'CPI
    spiLYc = busca_EnRango(ttlLYRng, "SPI", "columna", True) 'SPI
    cnsLYc = busca_EnRango(ttlLYRng, "% Consumed", "columna", True) '% Consumed
    remLYc = busca_EnRango(ttlLYRng, "Remaining Effort", "columna", True) 'Remaining Effort
    eacLYc = busca_EnRango(ttlLYRng, "EAC", "columna", True) 'EAC
    veacLYc = busca_EnRango(ttlLYRng, "EAC Variance", "columna", True) 'EAC Variance
    mveacLYc = busca_EnRango(ttlLYRng, "EAC Variance (Finantial Nav)", "columna", True) 'Marked Version EAC Variance
    stdtLYc = busca_EnRango(ttlLYRng, "Actual Start", "columna", True) 'Actual Start Date
    ontmLYc = busca_EnRango(ttlLYRng, "On-Time Variance", "columna", True) 'On-Time Variance
    esusLYc = busca_EnRango(ttlLYRng, "% Estimated Used", "columna", True) '% Estimated Used
    scvrLYc = busca_EnRango(ttlLYRng, "Schedule Variance", "columna", True) 'Schedule Variance
    pjdrLYc = busca_EnRango(ttlLYRng, "Project Duration", "columna", True) 'Project Duration
    ontmcmLYc = busca_EnRango(ttlLYRng, "On-Time Compliance", "columna", True) 'On-Time Compliance
    'min15LYc = busca_EnRango(ttlLYRng, "Minor Estimate >= 1500h", "columna", True) 'Minor Estiamtes >= 1500h
    'minAc15LYc = busca_EnRango(ttlLYRng, "Minor Actuals >= 1500h", "columna", True) 'Minor Actuals >= 1500h
    'maj15LYc = busca_EnRango(ttlLYRng, "Major Estimate < 1500h", "columna", True) 'Major Estimate < 1500h
    'main6LYc = busca_EnRango(ttlLYRng, "Maintenance > 660", "columna", True) 'Maintenance > 660
    riMMLYc = busca_EnRango(ttlLYRng, "Release & Impl Dt Mismatch", "columna", True) 'Release & Impl Dt Mismatch
    rpdLYc = busca_EnRango(ttlLYRng, "Release Past Due", "columna", True) 'Release Past due
    msTmEsLYc = busca_EnRango(ttlLYRng, "Missing Timely Estimates", "columna", True) 'Missing Timely Estimates
    onestLYc = busca_EnRango(ttlLYRng, "On-Estimate Compliance", "columna", True) 'On-Estimate Compliance
    ragLYc = busca_EnRango(ttlLYRng, "Overall RAG Status", "columna", True) 'RAG (green, amber, red)
    issLYc = busca_EnRango(ttlLYRng, "Issues", "columna", True) 'Issues (1,2,0)
    rskLYc = busca_EnRango(ttlLYRng, "Risks", "columna", True) 'Risks (1,2,0)
    issRskLYc = busca_EnRango(ttlLYRng, "Risk/Issue Comment", "columna", True) 'Colocar comentario derivado de "risk, Rag o Issue"
    relByLYc = busca_EnRango(ttlLYRng, "End Dates Beyond", "columna", True) 'Release Beyond Comment
    eddtLYc = busca_EnRango(ttlLYRng, "Schedule Finish", "columna", True) 'Schedule End Date
    pmissLYc = busca_EnRango(ttlLYRng, "PM Issue", "columna", True) 'PM Issue
    cntLYc = busca_EnRango(ttlLYRng, "# Alertas", "columna", True) 'Counter
    wrMMLYc = busca_EnRango(ttlLYRng, "Work Type & Resources Mismatch", "columna", True) 'Major Estimate < 1500h, Minor Actuals >= 1500h, Minor Estiamtes >= 1500h, Maintenance > 660
    actfnLYc = busca_EnRango(ttlLYRng, "Actual (Finantial Nav)", "columna", True) 'Actuals de la extaccioon de Finantial NAvigator
    eacfnLYc = busca_EnRango(ttlLYRng, "EAC (Finantial Nav)", "columna", True) 'EAC de la extaccioon de Finantial NAvigator
    cpissLYc = busca_EnRango(ttlLYRng, "CPI Issue", "columna", True) 'CPI Issue
    spissLYc = busca_EnRango(ttlLYRng, "SPI Issue", "columna", True) 'SPI isssue
    evissLYc = busca_EnRango(ttlLYRng, "EAC Variance Issue", "columna", True) 'EAC Variance Issue
    mkestLYc = busca_EnRango(ttlLYRng, "Not Marked Estimates", "columna", True) 'Major & Minor Projects estimates not 'marked'
    ontckLYc = busca_EnRango(ttlLYRng, "On Track", "columna", True) 'Donde se pondrán si va bien o mal
    stphLYc = busca_EnRango(ttlLYRng, "Standard Phase", "columna", True) 'Donde se verifica lo de ls fases estándar
    phagdLYc = busca_EnRango(ttlLYRng, "Phase Aging (days)", "columna", True) 'Phase aging dias
    phagLYc = busca_EnRango(ttlLYRng, "Phase Aging", "columna", True) 'Phase aging
    phagimpLYc = busca_EnRango(ttlLYRng, "M&M Projects Imp. 2 months ago", "columna", True) 'Major & Minor Projects implemented but not yet closed from 2 months prior and back
    phchLYc = busca_EnRango(ttlLYRng, "Phase Change Date", "columna", True) 'Fehca de cambio de fase
    noimpdtLYc = busca_EnRango(ttlLYRng, "Missing Implementation MS", "columna", True) 'Se encuentra el proyecto más no el implementeation date
    
    'FALTA MODIFICAR PARA QUE RECONOZCA SI NO SE HA PUESTO ALGUN TITULO DEL LAYOUT
    If ((rlLYc = 0) Or (idLYc = 0) Or (pnLYc = 0) Or (tmLYc = 0) Or (pmLYc = 0) Or (wtLYc = 0) Or (wsLYc = 0) Or (sdlcLYc = 0) _
    Or (cfLYc = 0) Or (swcapLYc = 0) Or (faLYc = 0) Or (mspiLYc = 0) Or (e0LYc = 0) Or (e1LYc = 0) Or (e2LYc = 0) _
    Or (mvLYc = 0) Or (meLYc = 0) Or (evLYc = 0) Or (pvLYc = 0) Or (eaLYc = 0) Or (blLYc = 0) Or (etLYc = 0) Or (cpiLYc = 0) _
    Or (spiLYc = 0) Or (cnsLYc = 0) Or (remLYc = 0) Or (eacLYc = 0) Or (veacLYc = 0) Or (mveacLYc = 0) _
    Or (msibLYc = 0) Or (stdtLYc = 0) Or (ontmLYc = 0) Or (esusLYc = 0) Or (scvrLYc = 0) Or (pjdrLYc = 0) Or (ontmcmLYc = 0) _
    Or (rpdLYc = 0) Or (riMMLYc = 0) Or (swculLYc = 0) Or (swcthLYc = 0) Or (msTmEsLYc = 0) _
    Or (ragLYc = 0) Or (issLYc = 0) Or (rskLYc = 0) Or (onestLYc = 0) Or (mkestLYc = 0) _
    Or (issRskLYc = 0) Or (relByLYc = 0) Or (eddtLYc = 0) Or (cntLYc = 0) Or (pmissLYc = 0) Or (wrMMLYc = 0) Or _
    (actfnLYc = 0) Or (eacfnLYc = 0) Or (evissLYc = 0) Or (spissLYc = 0) Or (cpissLYc = 0) Or (ontckLYc = 0) Or _
    (stphLYc = 0) Or (phagdLYc = 0) Or (phagLYc = 0) Or (phagimpLYc = 0) Or (phchLYc = 0) Or (noimpdtLYc = 0)) Then
    'Or (min15LYc = 0) Or (minAc15LYc = 0) Or (maj15LYc = 0) Or (main6LYc = 0)Or (prLYc = 0)
        MsgBox "Se saldrá de la aplicación, hay uno o más titulos en los títulos del Layout que no se encuentran"
        exit_all = True
        Exit Sub
    End If
    
    Call progressBar(True, 0, 0, 25)
    
    Application.ScreenUpdating = False 'para que no se muevan las paginas/workbooks y se quede en la misma mientras trabaja
    
    'Crear entorno de reporte, por si se tiene alguna falla
    Dim repWs As Worksheet, repRng As Range
    lywb.Activate 'Para que la nueva hojas se cree en el de Proyectos WIP
    Call crearDshBrdRep(lywb, repWs, repRng)
    lyws.Activate
    Call borrar_entorno(lyttl_row, idLYc)
    
    'Varibal ede row, col interna
    Dim idRow As Integer
    'Variables generales que tomarán los valores de las columnas
    Dim stPB As String, relPB As String, tmPB As String, P00 As String, pmPB As String, prPB As String, chdtPB As Date
    Dim stPV As String, idPV As String, wtPV As String, sdlcPV As String, cfPV As String, swcapPV As String _
        , faPV As String, pmPV As String, evPV As Double, pvPV As Double, eaPV As Double, bePV As Double _
        , etPV As Double, pnPV As String, stdtPV As Date, swculPV As String, swcthPV As String, ragPV As String, _
        issPV As Integer, rskPV As Integer, eddtPV As String, sdlcPB As String
    Dim ptFN As String, e0FN As Double, e1FN As Double, e2FN As Double, mvFN As String, meFN  As Double _
        , eacFN As Double, aeFN As Double
    Dim idPVrng As Range, idFNrng As Range, idMLrng As Range
    'nuevas variables (para cálculos)
    Dim cpi As Double, spi As Double, consumido As Double, remEff As Double, EAC As Double, varEAC As Double, _
        varEACmv As Double, estUsed As Double, schVar As Integer, onTmVar As Double, minEac As String, _
        onTmStr As String, minAe As String
    Dim msimp As Variant
    Dim pbRow As Integer, contP00 As Integer, rowDest As Integer, alertCnt As Integer
    'Variables para detección
    Dim fnEx As Boolean, mlEx As Boolean, colocarPV As Boolean, colocarFN As Boolean
    
    Call progressBar(False, 25, LRbow, 0)
    
    contP00 = 0
    'For para revisar todas y cada una de los Proyectos(renglones) del portafolio de proyectos BOW y colocar los deseados en el LayOut PMO Style
    For pbRow = pbttl_row + 1 To LRbow
    
        alertCnt = 0
        'empieza el conteo para la fila destino:
        rowDest = lyttl_row + 1 + contP00
        
        'Abrir el portafolio de proyectos y tomar valores de primer renglón deseado
        Pwip.Activate
        'Primero, NO nos interesan los que esán completed o Canceled
        stPB = Trim(Cells(pbRow, stPBc).value)
        If ((UCase(stPB) = "COMPLETED") Or (UCase(stPB) = "CANCELED") Or (UCase(Trim(Cells(pbRow, ragPBc).value) = "C"))) Then
            GoTo NEXTRow
        End If
        'De aquella que puede interesarnos tomar valores para analizarlos/colocarlos en LayOut
        relPB = Trim(Cells(pbRow, relPBc).value) 'Release
        tmPB = Trim(Cells(pbRow, tmPBc).value) 'team
        P00 = Trim(Cells(pbRow, idPBc).value) '# Proyecto
        'pmPB = Trim(Cells(pbRow, pmPBc).value) 'Project Manager
        prPB = Trim(Cells(pbRow, prPBc).value) 'Program Manager
        chdtPB = Cells(pbRow, chdtPBc).value ' Fehca de cambio de fase
        sdlcPB = Trim(Cells(pbRow, sdlcPBc).value) 'SDLC Phase
        
        'BUSCAR TAMVIE´N LOD E COMPLETED CANCELED EN LA EXTRACCIÓN.. Y SI SI ES.. SALTAR.. Y LLEVAR EL REGSITRO
        '!!!!!!!
        
        'Luego, ir a extracto de PlanView, el tradicional, y buscar el proyecto
        pvws.Activate
        idRow = busca_EnRangoV2(idPVc, pvttl_row + 1, LRpv, idPVrng, P00, "fila", True) 'Valor de renglón deseado en hoja de "PlanView Tradicional"

        If (idRow = 0 Or idRow = -1) Then
            Call toRepDsh(repWs, repRng, 1, 2, "Extracción Proyectos Vigentes", "No se encontró, seguramente falta agregarlo a portafolio de Proyectos Vigentes en PlanView", P00)
            'GoTo NEXTRow 'ya que no existe en PV pues no se puede hacer más comparaciones
            'MsgBox "no encontró algo en Extracción de Portafolios Vigentes: " & P00
            stdtPV = 0
            stPV = ""
            idPV = ""
            wtPV = ""
            sdlcPV = ""
            cfPV = ""
            swcapPV = ""
            swculPV = ""
            faPV = ""
            pmPV = ""
            evPV = 0
            pvPV = 0
            bePV = 0
            etPV = 0
            eaPV = 0
            pnPV = ""
            swculPV = 0
            swcthPV = 0
            ragPV = ""
            issPV = 0
            rskPV = 0
            eddtPV = 0
            colocarPV = False
        Else
            colocarPV = True
            stPV = Trim(Cells(idRow, stPVc).value) 'work Status
            If ((UCase(stPV) = "COMPLETED") Or (UCase(stPV) = "CANCELED") Or (stPV = "Assumed Completed")) Then
                'avisar que no se tiene detectado en log de proyectos
                Call toRepDsh(repWs, repRng, 1, 2, "Work Status", "Cambió, ya es -> " & stPV, P00)
                GoTo NEXTRow
                sdlcPV = ""
                wtPV = ""
            End If
            
            'sabiendo el Row "key" podemos tomar todos los valores deseados para ese proyecto
            stPV = Trim(Cells(idRow, stPVc).value) 'work Status
            idPV = Trim(Cells(idRow, idPVc).value) 'P00...
            wtPV = Trim(Cells(idRow, wtPVc).value) 'Work Type ( major, minor, maintenance..)
            sdlcPV = Trim(Cells(idRow, sdlcPVc).value) 'SDLC Phase (Ini, Def, Des, ...)
            cfPV = Trim(Cells(idRow, cfPVc).value) 'Capital Flag
            swcapPV = Trim(Cells(idRow, swcapPVc).value) 'SWCAP Qualification (qualified, not started, disqualified)
            swculPV = Trim(Cells(idRow, swculPVc).value)
            faPV = Trim(Cells(idRow, faPVc).value) 'Finance approval (undetermined, No, Yes)
            pmPV = Trim(Cells(idRow, pmPVc).value) 'Project Manager
            evPV = Cells(idRow, evPVc).value 'Earned Value
            pvPV = Cells(idRow, pvPVc).value 'Planned Value
            bePV = Cells(idRow, bePVc).value 'Baseline Effort
            etPV = Cells(idRow, etPVc).value 'Effort Total
            eaPV = Cells(idRow, eaPVc).value 'Effort Actual
            pnPV = Trim(Cells(idRow, pnPVc).value) 'Project Name
            stdtPV = Cells(idRow, stdtPVc).value 'Start Date
            swculPV = Cells(idRow, swculPVc).value 'Start Date
            swcthPV = Cells(idRow, swcthPVc).value 'Start Date
            ragPV = Trim(Cells(idRow, ragPVc).value) 'Rag
            issPV = Cells(idRow, issPVc).value ' Issues
            rskPV = Cells(idRow, rskPVc).value ' risks
            eddtPV = Cells(idRow, eddtPVc).value 'End Date
        End If
        
        'ver que sdlc se igual, si no avisar
        If ((UCase(sdlcPB) <> UCase(sdlcPV)) And colocarPV = True) Then
            'Colorear el cuadro en el Portafolio View, o mejor poner en un array antes y después
            Call toRepDsh(repWs, repRng, 1, 2, "SDLC Phase Change", "Cambió de " & sdlcPB & " -> " & sdlcPV & " (Corregir fase y fecha de cambio)", P00)
            pvws.Activate
        End If
        
        'Después, ir al extracto de financial
        fnws.Activate
        idRow = busca_EnRangoV2(ptFNc, fnttl_row + 2, LRfn, idFNrng, pnPV, "fila", True) 'Valor de renglón deseado en hoja de "Financial Navigator"
        If (idRow = 0 Or idRow = -1) Then
            Call toRepDsh(repWs, repRng, 1, 2, "Extracción Financial Navigator", "X", P00)
            'GoTo NEXTRow 'ya que no existe en PV pues no se puede hacer más comparaciones
            'MsgBox "No se encontró el proyecto " & P00 & " en Financial: "
            meFN = 0
            aeFN = 0
            eacFN = 0
            ptFN = ""
            e0FN = 0
            e1FN = 0
            e2FN = 0
            mvFN = ""
            colocarFN = False
        Else 'Si SI se encontró el proyecto entonces se pueden tomar los valores
            colocarFN = True
            ptFN = Trim(Cells(idRow, ptFNc).value) 'Título del proyecto completo, no hay P00
            e0FN = Cells(idRow, e0FNc).value 'E0 Estimación
            e1FN = Cells(idRow, e1FNc).value 'E1 Estimación
            e2FN = Cells(idRow, e2FNc).value 'E2 Estimación
            mvFN = Trim(Cells(idRow, mvFNc).value) 'Marked Version (E0, E2, E2 Revisión)
            meFN = Cells(idRow, meFNc).value 'Marked Version Estimate
            eacFN = Cells(idRow, eacFNc).value 'EAC from finantial
            aeFN = Cells(idRow, aeFNc).value 'actuals
        End If
        
        
        'En seguida, ir a extracto de Milestones
        mlws.Activate
        Dim scndSearch As String, strFechaNormal As String, mYMatch As String, colocarMS As Boolean, blstrFecha As String, _
            fechaNormal As Date, blFecha As Date, schDtMiss As Boolean, blDtMiss As Boolean, noMsImp As Boolean
        msimp = ""
        schDtMiss = True
        blDtMiss = True
        colocarMS = False
        blFecha = 0
        fechaNormal = 0
        noMsImp = False
 
            
        If (wtPV = "Major Project" Or wtPV = "Minor Project" Or wtPV = "Program") Then
            If (wtPV = "Program") Then
                scndSearch = "MS: Overall Implementation" 'Sting a buscar, en segunda busqueda (primero nupero de proyecto, luego este valor
            Else
                scndSearch = "MS: Project Implemented" 'Sting a buscar, en segunda busqueda (primero nupero de proyecto, luego este valor
            End If
            idRow = busca_EnRangoV2(idMLc, mlttl_row + 1, LRml, idMLrng, P00, "fila", True, scndSearch, mlMLc) 'Valor de renglón deseado en hoja de "Milestones"
            If P00 = "P0000053" Then
                ver = "loL"
            End If
            If (idRow = 0) Then
                'Significa que encontró proyecto más no el Milestone de Implementation
                noMsImp = True
                'GoTo NEXTRow 'ya que no existe en PV pues no se puede hacer más comparaciones
                'MsgBox "No se encontró el Milestone de Implementación del proyecto: " & P00
            ElseIf (idRow = -1) Then
                'Que avise que NO encontró el proyecto
                Call toRepDsh(repWs, repRng, 1, 2, "Extracción Milestones", "X", P00)
                noMsImp = True
                'MsgBox "No se encontró el proyecto " & P00 & " en archivo de Milestones."
            Else
                'analizar fecha
                strFechaNormal = dd_mm_aa(Trim(Cells(idRow, stMLc)), False) 'obtener en formato dd/mm/aaaa un string
                blstrFecha = dd_mm_aa(Trim(Cells(idRow, bdMLc)), False)
                If (strFechaNormal <> "") Then
                    mYMatch = analizaRyFecha(strFechaNormal, relPB, alertCnt)
                    fechaNormal = strToDate(strFechaNormal)
                    schDtMiss = False
                End If
                If (blstrFecha <> "") Then
                    blFecha = strToDate(blstrFecha)
                    blDtMiss = False
                End If
                colocarMS = True
            End If
        Else
            If colocarPV = False Then
                Call toRepDsh(repWs, repRng, 1, 2, "Extracción Milestones", "Sin posiblidad de busqueda", P00)
            ElseIf (colocarFN = False) Then
                Call toRepDsh(repWs, repRng, 1, 2, "Extracción Milestones", "N/A", P00)
            End If
        End If
        
        'ANALSIS DE VARIABLES
        lyws.Activate
        '--- Colocar en LAYOUT ---
'St--- Inicia bloque para cuando son Majors o Minors
        If (wtPV = "Minor Project" Or wtPV = "Major Project") Then
            If (eaPV <> 0 And pvPV <> 0 And bePV <> 0) Then 'hacerlo si los valores son diferentes a 0
                If (colocarPV = True) Then
                    'Hacer análisis (y cálculos) con la información obtenida
                    cpi = Round(evPV / eaPV, 6) 'verde +/- 5%, amarillo +/- 10% , rojo > +/-10% (Earned Value / Effort Actual)
                    spi = Round(evPV / pvPV, 6) 'verde +/- 5%, amarillo +/- 10% , rojo > +/-10% (Earned Value / Planned Value)
                    consumido = Round(eaPV / bePV, 6) '(Effort Actual / Baseline Effort)
                    remEff = etPV - eaPV 'Effort total - Effort Actual
                    EAC = eaPV + remEff '(EAC = Estimated At Completion)
                    varEAC = Round((EAC - bePV) / bePV, 6) 'Alertar si es > 10%
                End If
            End If
            If (meFN <> 0) Then
                If (colocarFN = True) Then
                    varEACmv = Round((eacFN - meFN) / meFN, 6) 'Alertar si es > 10%
                    estUsed = Round(aeFN / meFN, 6) '%Estimated Used
                End If
            End If
            
            'On-Time Compliance
            onTmVar = 0
            If (schDtMiss = False And blDtMiss = False And colocarPV = True) Then 'Significa que sí existen los valores necesarios
                schVar = fechaNormal - blFecha 'Shcedule Variance (MS Implemented - MS implemented baseline)
                projDur = blFecha - stdtPV 'Project Duration (Ms implemented baseline - actual start date)
                onTmVar = Round(schVar / projDur, 6) ' On Time Variance (schedule Varianve / Project Duration)
                Cells(rowDest, scvrLYc).value = schVar 'Shcedule Variance
                Cells(rowDest, pjdrLYc).value = projDur 'Project Duration
                Cells(rowDest, ontmLYc).value = onTmVar ' On Time Variance
            ElseIf (blDtMiss = False And colocarPV = True) Then 'Significa que existe la de baseline pero la otra no (poco probable)
                projDur = blFecha - stdtPV 'Project Duration (Ms implemented baseline - actual start date)
            End If

            If (colocarPV = True) Then
                Cells(rowDest, cpiLYc).value = cpi 'CPI 'verde +/- 5%, amarillo +/- 10% , rojo > +/-10% (Earned Value / Effort Actual)
                Cells(rowDest, spiLYc).value = spi 'SPI 'verde +/- 5%, amarillo +/- 10% , rojo > +/-10% (Earned Value / Planned Value)
                Cells(rowDest, cnsLYc).value = consumido '% Consumed '(Effort Actual / Baseline Effort)
                Cells(rowDest, remLYc).value = remEff 'Remaining Effort 'Effort total - Effort Actual
                Cells(rowDest, eacLYc).value = EAC 'EAC '(EAC = Estimated At Completion)
                Cells(rowDest, veacLYc).value = varEAC  'EAC Variance 'Alertar si es > 10%
                Cells(rowDest, cpissLYc).value = cpispi(cpi, True, sdlcPV, alertCnt) 'Análisis del CPI
                Cells(rowDest, spissLYc).value = cpispi(spi, False, sdlcPV, alertCnt) 'Analisis del SPI
            End If
                'depende tanto de FN como de PV, se tratan ambos casos dentro de la función
                Cells(rowDest, evissLYc).value = eacVarAn(varEAC, varEACmv, colocarPV, colocarFN, alertCnt) 'Analisis del EAC Variance (Finantial Navigator y Proyectos Vigentes)
            If (colocarFN = True) Then
                Cells(rowDest, mveacLYc).value = varEACmv 'MV EAC Variance 'Alertar si es > 10%
                Cells(rowDest, esusLYc).value = estUsed '%Estimated Used
                Cells(rowDest, mkestLYc).value = mkVerAn(mvFN, alertCnt)
            End If
            
            'Texto dentro de Ontime y de OnEstiamte Compliance
            If (colocarFN = False) Then
                Cells(rowDest, onestLYc).value = on_estimate(aeFN, meFN, varEACmv, "no fn", sdlcPV, mvFN, stdtPV, alertCnt) 'On-Estimate Compliance
            ElseIf (colocarPV = False) Then
                Cells(rowDest, onestLYc).value = on_estimate(aeFN, meFN, varEACmv, "no pv", sdlcPV, mvFN, stdtPV, alertCnt) 'On-Estimate Compliance
            Else
                Cells(rowDest, onestLYc).value = on_estimate(aeFN, meFN, varEACmv, wtPV, sdlcPV, mvFN, stdtPV, alertCnt) 'On-Estimate Compliance
            End If
            'Solo están los dos casos de cuando no hay MS (donde no nos interesa siquiera analizar el On_time)
            'Y el otro de PV donde podemos prescindir de un dato y continuar con el análisis de lo demás
            If (colocarMS = False) Then
                Cells(rowDest, ontmcmLYc).value = on_time(blFecha, fechaNormal, onTmVar, "no ms", stdtPV, alertCnt)
            ElseIf (colocarPV = False) Then
                Cells(rowDest, ontmcmLYc).value = on_time(blFecha, fechaNormal, onTmVar, "no pv", stdtPV, alertCnt)
            Else
                Cells(rowDest, ontmcmLYc).value = on_time(blFecha, fechaNormal, onTmVar, wtPV, stdtPV, alertCnt)
            End If
            
'----- Bloque para cuando no son ni major ni minor
        ElseIf (wtPV <> "Minor Project" And wtPV <> "Major Project" And colocarPV = True) Then 'Para revisar algo cuando no sea Minor ni Major
            Cells(rowDest, ontmLYc).value = "N/A" ' On Time Variance
            Cells(rowDest, ontmcmLYc).value = "N/A" ' On Time compliance TEXT
            Cells(rowDest, onestLYc).value = "N/A" 'On estimate Text
        End If
'End-- Termina bloque para cuando son Majors o Minors
        
        
'St--- Inicia bloque para colocar valores, los uqe no se colocaron en el bloque de arriba, igualmente, dentro del bloque,
'----- se hacen algunas verificaciones para saber si los valores porcesados se deben de poner o no
        'Colocar en el Layout los valores ya tomados
        'De info general proyectos
        Cells(rowDest, rlLYc).value = relPB 'Release
        Cells(rowDest, idLYc).value = P00 'P00...
        Cells(rowDest, pnLYc).value = pnPV 'Project Name
        Cells(rowDest, tmLYc).value = tmPB 'Team
        Cells(rowDest, pmLYc).value = pmPV 'Project Manager
        'Cells(rowDest, prLYc).value = prPB 'Program Manager
        Cells(rowDest, rpdLYc).value = rPstDue(relPB, alertCnt) 'Release Past Due, coloca algo si ya es un mes o año en el pasado
        Cells(rowDest, phagdLYc).value = phaseAging(chdtPB, True, "", "", alertCnt)
        Cells(rowDest, phagLYc).value = phaseAging(chdtPB, False, "90", wtPV, alertCnt)
        Cells(rowDest, phagimpLYc).value = phaseAging(chdtPB, False, sdlcPV, wtPV, alertCnt)
        If chdtPB <> 0 Then
            Cells(rowDest, phchLYc).value = chdtPB
        End If
        If (colocarPV = True) Then
            Cells(rowDest, wtLYc).value = wtPV 'Work Type
            Cells(rowDest, wsLYc).value = stPV 'Work Status
            Cells(rowDest, sdlcLYc).value = sdlcPV 'SDLC Phase
            Cells(rowDest, cfLYc).value = cfPV 'Capitalization Flag
            Cells(rowDest, swcapLYc).value = swcapPV 'SWCAP qualification
            Cells(rowDest, swculLYc).value = swculPV 'SWCAP Useful Life
            Cells(rowDest, swcthLYc).value = swcthPV 'SWCAP Cost Threshold
            Cells(rowDest, faLYc).value = faPV 'Finance Approval
            Cells(rowDest, evLYc).value = evPV 'EV-Earned Value (h)
            Cells(rowDest, pvLYc).value = pvPV 'EV-Planned Value (h)
            Cells(rowDest, eaLYc).value = eaPV 'Effort Actual (h)
            Cells(rowDest, blLYc).value = bePV 'Baseline Effort (h)
            Cells(rowDest, etLYc).value = etPV 'Effort Total (h)
            Cells(rowDest, stdtLYc).value = stdtPV  'Actual Start Date
            Cells(rowDest, ragLYc).value = ragPV 'Rag
            Cells(rowDest, issLYc).value = issPV  'Issues
            Cells(rowDest, rskLYc).value = rskPV 'risks
            Cells(rowDest, issRskLYc).value = rskIss(wtPV, rskPV, issPV, ragPV, alertCnt) 'Coloca comentario (RAG, Issues + Risks)
            Cells(rowDest, eddtLYc).value = eddtPV 'schedule end date
            Cells(rowDest, relByLYc).value = endBeyond(wtPV, eddtPV, relPB, alertCnt) 'End dates Beyond
            If pmPV = "" Then ' And prPB = "") Then
                Cells(rowDest, pmissLYc).value = "Sin ProjectM" 'schedule end date
                alertCnt = alertCnt + 1
            End If
            Cells(rowDest, ontckLYc).value = behindSch(sdlcPV, relPB, wtPV, mnWs, mjWs, ttlDTmnRng, ttlDTmjRng, rMnRng, rMjRng, lcMn, lcMj, lyws, alertCnt)
            Cells(rowDest, stphLYc).value = phaseDif(wtPV, sdlcPV, "", True, alertCnt)
        End If
        'De finance
        If (colocarFN = True) Then
            Cells(rowDest, e0LYc).value = e0FN 'E0 (h)
            Cells(rowDest, e1LYc).value = e1FN 'E1 (h)
            Cells(rowDest, e2LYc).value = e2FN 'E2 (h)
            Cells(rowDest, mvLYc).value = mvFN 'Marked Version (E0, E2, E2 Revisión)
            Cells(rowDest, meLYc).value = meFN 'Marked Version Estimate
            Cells(rowDest, actfnLYc).value = aeFN 'Actuals de la extracción Financiera
            Cells(rowDest, eacfnLYc).value = eacFN 'EAC de la extracción Financiera
            If (colocarPV = True) Then
                Cells(rowDest, msTmEsLYc).value = mTEstimate(e0FN, e1FN, e2FN, sdlcPV, wtPV, alertCnt) 'ver los de Capitañization exception
            End If
            Cells(rowDest, wrMMLYc).value = wtResMM(wtPV, aeFN, eacFN, meFN, sdlcPV, alertCnt)
        End If
        'De milestones
        If (colocarMS = True) Then ' tampoco colocará si no son major o minor
            If fechaNormal <> 0 Then 'IF para evitar que salga en celda "12:00:00 a.m." que es lo que pasa cuando fecha nula
                Cells(rowDest, mspiLYc).value = fechaNormal ' MS Project Implementation
            End If
            If blFecha <> 0 Then 'IF para evitar que salga en celda "12:00:00 a.m." que es lo que pasa cuando fecha nula
                Cells(rowDest, msibLYc).value = blFecha 'MS Project Implementation (Baseline)
            End If
            Cells(rowDest, riMMLYc).value = mYMatch 'Coloca "" si no hubo error y si no , la razón del error
        End If
        Cells(rowDest, noimpdtLYc).value = missImpDt(wtPV, noMsImp, sdlcPV, alertCnt)
        Cells(rowDest, cntLYc).value = alertCnt
'End-- Termina bloque para colocar valores

        contP00 = contP00 + 1
    
NEXTRow:
    Call progressBar(False, 25, LRbow, pbRow)
    Next pbRow
    
    'cerrar los archivos usados
    pbwb.Close savechanges:=False
    mlwb.Close savechanges:=False
    pvwb.Close savechanges:=False
    fnwb.Close savechanges:=False
    dtwb.Close savechanges:=False
    
    barraCarga.Hide
    Application.ScreenUpdating = True ' para volver a activar la diferenciación de cuando se trabaja en x o y página/wbk
    
End Sub


Private Sub crearDshBrdRep(wb As Workbook, ws As Worksheet, ttlRng As Range)
    Dim ttl_row As Integer, lc As Double
    ttl_row = 2
    'primero ver que no exista en layout la hoja de reportes
    wb.Activate
    For Each wkSh In Worksheets
        If wkSh.Name = "Faltantes" Then
            wkSh.Delete
        End If
    Next wkSh
    'Agregar hoja de reporte
    Set ws = Sheets.Add
    ws.Name = "Faltantes"
    ws.Select
    With ws.Tab
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
    End With
    ws.Activate
    'Crear el entorno donde vamos a poner todo en la WorkSheet
    Cells(ttl_row, 1).value = "# Proyecto"
    Cells(ttl_row, 2).value = "Extracción Proyectos Vigentes"
    Cells(ttl_row, 3).value = "Extracción Financial Navigator"
    Cells(ttl_row, 4).value = "Extracción Milestones" ' (Proyecto)"
    'Cells(ttl_row, 5).value = "Extracción Milestones (Milestone)"
    Cells(ttl_row, 5).value = "Work Status"
    Cells(ttl_row, 6).value = "SDLC Phase Change"
    'regresar valor del rango de los títulos
    lc = ws.Cells(ttl_row, Columns.COUNT).End(xlToLeft).Column
    Set ttlRng = Range(Cells(ttl_row, 1), Cells(ttl_row, lc))
    
End Sub

Private Sub toRepDsh(ws As Worksheet, rng As Range, pyCol As Integer, ttl_row As Integer, ttl As String, Miss As String, P00 As String)
    Dim col As Integer, lr As Integer, idRow As Integer
    ws.Activate
    col = busca_EnRango(rng, ttl, "columna", True)
    If (col = 0) Then
        MsgBox ("No se encontró el título: '" & ttl & "' en la hoja creada de reporte")
        Exit Sub
    End If
   'Encontrar si se está agregando un nuevo proyecto o No
    If Not IsEmpty(ws.Cells(ttl_row + 1, pyCol).value) Then
        lr = ws.Cells(ttl_row, pyCol).End(xlDown).row
        With ws.Range(Cells(ttl_row + 1, pyCol), Cells(lr, pyCol))
            Set c = .Find(P00, LookIn:=xlValues)
            'Significa que la fila ya está llenada y se pueden seguir colocando valores ahí
            If Not c Is Nothing Then
                idRow = c.row
            'Significa que no está esa fila y hay que incluir una nueva, y que ya existen filas
            Else
                idRow = lr + 1
                Cells(idRow, pyCol).value = P00
            End If
        End With
    Else
        idRow = ttl_row + 1
        Cells(idRow, pyCol).value = P00
    End If
    Cells(idRow, col).value = Miss
    'Cells(idRow, col + 1).Value = PV
    
End Sub
