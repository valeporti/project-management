Public exit_all As Boolean

Public Sub actPortBOW()

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
    Dim lista_rutas() As String, dir_arch As Variant, num_arch As Integer, verificar As Boolean, funcionalidades() As String
    ReDim funcionalidades(2)
    'la funcionalidad "neutra" es la del archivo origen
    funcionalidades(0) = "Archivo Principal"
    funcionalidades(1) = "Extracción PlanView de Proyectos"
    funcionalidades(2) = "Portafolio Proyectos BOW"
    'Ruta de donde sacará archivos
    verificar = verificar_entorno(3, num_arch, mainwb, lista_rutas(), dir_arch, funcionalidades(), False)
        If (verificar = False) Then
            Exit Sub
        End If
'**
'ANÁLISIS DE LOS ARCHIVOS
'**
    'Application.ScreenUpdating = False 'para que no se muevan las paginas/workbooks y se quede en la misma mientras trabaja
    '(idea)Por medio de un inputbox, dinámicamente preguntar cual es el archivo para X o Y acción
    Call extraccionPVyBOWPort(lista_rutas(), 3)
    If exit_all = True Then
        Exit Sub
    End If
    
    'Application.ScreenUpdating = True ' para volver a activar la diferenciación de cuando se trabaja en x o y página/wbk
    
End Sub

Private Sub extraccionPVyBOWPort(arrArch() As String, bowPort_ttl_row As Integer)
    
    'PREPARAR ENTOORNO PORTAFOLIO BOW
    ' En el portafolio del BOW, mostrar todas las columnas para que se puedan ver los cambios
    Dim PBOWwb As Workbook
    Workbooks.Open (arrArch(1))
    Set PBOWwb = ActiveWorkbook
    
    Dim LR_bow As Integer, LC_bow As Integer, Pwip As Worksheet, verq As Variant, ttlBowRng As Range
    verq = verUnaSh(PBOWwb, 1, "Favor de escribir la hoja donde está el portafolio de proyectos", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        MsgBox "No se detectó la hoja de donde obtener la información," & Chr(10) & _
            "Se saldrá de la aplicación"
        'cerrar el archivo de origen
        PBOWwb.Close savechanges:=False
        'pedir salir de todo
        exit_all = True
        Exit Sub
    End If
    'Hoja que se usará
    If (verq = True) Then
        Set Pwip = PBOWwb.Worksheets(1)
    Else ' significa que tenemos un string
        Set Pwip = PBOWwb.Worksheets(verq)
    End If
    Pwip.Activate
    Call desactivar_filtro
    
    LR_bow = Pwip.Cells(bowPort_ttl_row, 3).End(xlDown).row
    LC_bow = Pwip.Cells(bowPort_ttl_row, Columns.COUNT).End(xlToLeft).Column
    Set ttlBowRng = Pwip.Range(Cells(bowPort_ttl_row, 1), Cells(bowPort_ttl_row, LC_bow))
    
    'Mostrar todas las columnas
    Range(Columns(2), Columns(LC_bow)).Select 'Cells(bowPort_ttl_row, 2) & ":" & Cells(bowPort_ttl_row, LC_bow)
    Selection.EntireColumn.Hidden = False
    
    'PREPARAR ENTORNO DE LA EXTRACCIÓN
    Dim pvwb As Workbook, ttlPVRng As Range, LR_pv As Integer, LC_pv As Integer, pvws As Worksheet
    Workbooks.Open (arrArch(0))
    Set pvwb = ActiveWorkbook
    
    verq = verUnaSh(pvwb, 1, "Favor de escribir la hoja donde está la extracción de Proyectos de PlanView", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        MsgBox "No se detectó la hoja de donde obtener la información," & Chr(10) & _
            "Se saldrá de la aplicación"
        'cerrar el archivo de origen
        pvwb.Close savechanges:=False
        'pedir salir de todo
        exit_all = True
        Exit Sub
    End If
    'Hoja que se usará
    If (verq = True) Then
        Set pvws = pvwb.Worksheets(1)
    Else ' significa que tenemos un string
        Set pvws = pvwb.Worksheets(verq)
    End If
    pvws.Activate
    Call desactivar_filtro
    
    LR_pv = pvws.Cells(1, 1).End(xlDown).row
    LC_pv = pvws.Cells(1, Columns.COUNT).End(xlToLeft).Column
    Set ttlPVRng = pvws.Range(Cells(1, 1), Cells(1, LC_pv))
    
'EL ANÁLISIS
    'Listo, ahora de cada proyecto en la extracción y en el ProyectosWIP, verificar si hubo cambios, siendo el ProyectosWIP la referencia
    Dim bowR As Integer, pvR As Integer
    '"PB" = Portafolio BOW , "c" = Column, "r" = Row
    Dim stPBc As Integer, ragPBc As Integer, idPBc As Integer, wtPBc As Integer, sdlcPBc As Integer, cfPBc As Integer _
        , swcapPBc As Integer, faPBc As Integer, pmPBc As Integer
    stPBc = busca_EnRango(ttlBowRng, "Status", "columna", True)
    ragPBc = busca_EnRango(ttlBowRng, "RAG", "columna", True)
    idPBc = busca_EnRango(ttlBowRng, "Work Id", "columna", True)
    wtPBc = busca_EnRango(ttlBowRng, "Work Type", "columna", True)
    sdlcPBc = busca_EnRango(ttlBowRng, "SDLC Phase", "columna", True) 'SDLC Phase
    cfPBc = busca_EnRango(ttlBowRng, "Capitaliz. Flag", "columna", True) 'Capitalization Flag
    swcapPBc = busca_EnRango(ttlBowRng, "Swr Cap Qualification", "columna", True) 'SWCAP Qualification
    faPBc = busca_EnRango(ttlBowRng, "Finance Approval", "columna", True) 'Finance Approval
    pmPBc = busca_EnRango(ttlBowRng, "Project Mgr", "columna", True) 'Project Manager
    prmPBc = busca_EnRango(ttlBowRng, "Program Mgr", "columna", True) 'Program Manaeger
    Dim stPVc As Integer, idPVc As Integer, wtPVc As Integer, sdlcPVc As Integer, cfPVc As Integer, swcapPVc As Integer _
        , faPVc As Integer, pmPVc As Integer, pmrPVc As Integer
    stPVc = busca_EnRango(ttlPVRng, "Work Status", "columna", True) 'Work Status (commited, targeted, cancelled)
    idPVc = busca_EnRango(ttlPVRng, "Work ID #", "columna", True) 'P00...
    wtPVc = busca_EnRango(ttlPVRng, "Work Type", "columna", True) 'Work Type ( major, minor, maintenance..)
    sdlcPVc = busca_EnRango(ttlPVRng, "SDLC Phase", "columna", True) 'SDLC Phase (Ini, Def, Des, ...)
    cfPVc = busca_EnRango(ttlPVRng, "Capitalization Flag", "columna", True) 'Capital Flag
    swcapPVc = busca_EnRango(ttlPVRng, "SWCAP Qualification", "columna", True) 'SWCAP Qualification (qualified, not started, disqualified)
    faPVc = busca_EnRango(ttlPVRng, "Finance Approval", "columna", True) 'Finance approval (undetermined, No, Yes)
    pmPVc = busca_EnRango(ttlPVRng, "Project Manager", "columna", True) 'Project Manager
    
    'Verificar que todos los títulos tengan exitencia, sino salir
    If ((stPBc = 0) Or (ragPBc = 0) Or (idPBc = 0) Or (sdlcPBc = 0) Or (wtPBc = 0) Or (cfPBc = 0) Or (swcapPBc = 0) Or (faPBc = 0) _
        Or (stPVc = 0) Or (idPVc = 0) Or (wtPVc = 0) Or (sdlcPVc = 0) Or (cfPVc = 0) Or (swcapPVc = 0) Or (faPVc = 0) _
        Or (pmPBc = 0) Or (pmPVc = 0) Or (prmPBc = 0)) Then
        MsgBox "Se saldrá de la aplicación, hay uno o más titulos que no se encuentran"
        exit_all = True
        Exit Sub
    End If
    
    'definir el rángo de búsueda del Work ID
    Dim idPVrng As Range
    Set idPVrng = pvws.Range(Cells(2, idPVc), Cells(LR_pv, idPVc))
    
    Dim P00 As String, sdlcPB As String, cfPB As String, swcapPB As String, faPB As String, idRow As Integer, _
         wtPB As String, pmPB As String, stPV As String, wtPV As String, pmPV As String, pms As Boolean, _
         stPB As String, sdlcPV As String, cfPV As String, swcapPV As String, matchF As Boolean, faPV As String
    Dim prmPB As String, prms As Boolean, corr() As String, cnt As Integer, faltantes() As String, faltan As Integer
        cnt = 0
            
    'Crear entorno de Reporte
    Dim repWs As Worksheet, repRng As Range
    PBOWwb.Activate 'Para que la nueva hojas se cree en el de Proyectos WIP
    Call crearWsRep(repWs, repRng)
    
    'revisaremos fila por fila el contenido en BOW Portfolio
    For bowR = bowPort_ttl_row + 1 To LR_bow
        matchF = True
        Pwip.Activate
        'primero si en la columno de estatus dice completed o canceled, saltar al siguiente o en la columna del RAG está una "C".. que sería lo mismo
        stPB = Trim(Cells(bowR, stPBc).value)
        If ((UCase(stPB) = "COMPLETED") Or (UCase(stPB) = "CANCELED") Or (UCase(Cells(bowR, ragPBc).value) = "C")) Then
            GoTo NEXTRow
        End If
        'Primero guardar el ID : Work ID (núm de proyecto)
        P00 = Trim(Cells(bowR, idPBc).value)
        'Segundo, tomar valores del Portafolio BOW
        sdlcPB = Trim(Cells(bowR, sdlcPBc).value)
        cfPB = Trim(Cells(bowR, cfPBc).value)
        swcapPB = Trim(Cells(bowR, swcapPBc).value)
        faPB = Trim(Cells(bowR, faPBc).value)
        wtPB = Trim(Cells(bowR, wtPBc).value)
        pmPB = Trim(Cells(bowR, pmPBc).value)
        prmPB = Trim(Cells(bowR, prmPBc).value)
        
        'tercero, ir a extracción PV, buscar y comparar
        pvws.Activate
        idRow = busca_EnRango_falta(idPVrng, P00, "fila", True, faltantes(), faltan)
        If (idRow = 0) Then
            Call toRep(repWs, repRng, 1, 2, "Proyecto Faltante", "", P00, P00)
            GoTo NEXTRow 'ya que no existe en PV pues no se puede hacer más comparaciones
        End If
        
        'Si este es Cancelled o Completed, avisar
        stPV = Cells(idRow, stPVc).value
        If (stPV = "Cancelled" Or stPV = "Completed") Then
            Call toRep(repWs, repRng, 1, 2, "Work Status", stPV, stPB, P00)
            pvws.Activate
        End If
        
        'es dudoso pero si pasa, avisar que el Work Type es diferente en ambos documentos
        wtPV = Cells(idRow, wtPVc).value
        matchF = compararDos(wtPV, wtPB)
        If (matchF = False) Then
            Call toRep(repWs, repRng, 1, 2, "Work Type", wtPV, wtPB, P00)
            pvws.Activate
            matchF = True
        End If
        
        'comparar PMs / PrM
        pmPV = Cells(idRow, pmPVc).value
        If (pmPV <> "" And pmPB <> "") Then
            pms = compararPms(pmPV, pmPB, corr(), cnt)
            If (pms = False) Then 'Cuando los PMs no coinciden
                'Si PMS no coinciden puede ser porque ya cambiaron a lo de Program
                If (prmPB <> "") Then
                    prms = compararPms(pmPV, prmPB, corr(), cnt)
                    If (prms = False) Then
                        Call toRep(repWs, repRng, 1, 2, "Project Manager / Prgrm Mngr", pmPV, prmPB, P00)
                        pvws.Activate
                    End If
                Else
                    Call toRep(repWs, repRng, 1, 2, "Project Manager / Prgrm Mngr", pmPV, pmPB, P00)
                    pvws.Activate
                End If
            End If
        ElseIf (pmPV = "" Or pmPB = "") Then
            Call toRep(repWs, repRng, 1, 2, "Project Manager / Prgrm Mngr", pmPV, pmPB, P00)
            pvws.Activate
        End If
        
        'comparar el SDLC
        sdlcPV = Cells(idRow, sdlcPVc).value
        If (UCase(sdlcPB) <> UCase(sdlcPV)) Then
            'Colorear el cuadro en el Portafolio View, o mejor poner en un array antes y después
            Call toRep(repWs, repRng, 1, 2, "SDLC Phase", sdlcPV, sdlcPB, P00)
            pvws.Activate
        End If

        'Si es programa, sólo revisar hasta la fase SDLC
        If ((UCase(wtPV) = "PROGRAM") Or (UCase(wtPV) = "MAINTENANCE") Or (UCase(wtPV) = "NON-TRADITIONAL")) Then
            GoTo NEXTRow
        End If
        
        'cfPV
        cfPV = Cells(idRow, cfPVc).value
        matchF = compararDos(cfPV, cfPB)
        If (matchF = False) Then
            Call toRep(repWs, repRng, 1, 2, "Cap Flag", cfPV, cfPB, P00)
            pvws.Activate
            matchF = True
        End If
        
        'swcapPV
        swcapPV = Cells(idRow, swcapPVc).value
        If (UCase(swcapPV) <> UCase(swcapPB)) Then
            Call toRep(repWs, repRng, 1, 2, "SWCAP Q", swcapPV, swcapPB, P00)
            pvws.Activate
        End If
        'faPV
        faPV = Cells(idRow, faPVc).value
        If (UCase(faPV) <> UCase(faPB)) Then
            Call toRep(repWs, repRng, 1, 2, "Finance App", faPV, faPB, P00)
            pvws.Activate
        End If
        
NEXTRow:
    Next bowR
    
    PBOWwb.Activate
    
End Sub



Public Sub toRep(ws As Worksheet, rng As Range, pyCol As Integer, ttl_row As Integer, ttl As String, PV As String, PB As String, P00 As String)
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
    Cells(idRow, col).value = PB
    Cells(idRow, col + 1).value = PV
    
End Sub
Private Sub crearWsRep(ws As Worksheet, ttlRng As Range)
    Dim ttl_row As Integer
    ttl_row = 2
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
    Call juntarRango(ttl_row, 2, "Project Manager / Prgrm Mngr", "P WIP", "PV Extracc") 'PMs
    Call juntarRango(ttl_row, 4, "Work Status", "P WIP", "PV Extracc") 'Work Status
    Call juntarRango(ttl_row, 6, "Work Type", "P WIP", "PV Extracc") 'Work Type
    Call juntarRango(ttl_row, 8, "SDLC Phase", "P WIP", "PV Extracc") 'SDLC
    Call juntarRango(ttl_row, 10, "Cap Flag", "P WIP", "PV Extracc") 'Capitalization Flag
    Call juntarRango(ttl_row, 12, "SWCAP Q", "P WIP", "PV Extracc") 'SWCAP Qualification
    Call juntarRango(ttl_row, 14, "Finance App", "P WIP", "PV Extracc") 'Finance Approval
    Call juntarRango(ttl_row, 16, "Proyecto Faltante", "P WIP", "PV Extracc") 'Falta proyecto
    Set ttlRng = Range(Cells(ttl_row - 1, 2), Cells(ttl_row - 1, 17))
    
End Sub



Function compararPms(pvStr As String, pbStr As String, lista() As String, cnt As Integer) As Boolean
'http://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'https://blog.udemy.com/vba-regex/
'https://msdn.microsoft.com/en-us/library/system.text.regularexpressions.regex.split(v=vs.110).aspx
    Dim pvArr() As String, pbArr() As String, noNeed As Boolean
    Dim strPattern As String
    strPattern = "\b[a-zA-Z]+\b"
    noNeed = False
    
    'Primero splitear cada string
    Call arrMatches(pvArr(), strPattern, pvStr)
    Call arrMatches(pbArr(), strPattern, pbStr)

    'ejemplos de strings tipicos.. de PB,"Cordero, Gabriel", de PV, "Santamaria, Jorge Alberto - JS80684"
    'Los vamos a comparar, con que coincidan 2, basta
    Dim pvArrLen As Integer, pbArrLen As Integer, matches As Integer
    pvArrLen = 0
    pbArrLen = 0
    For i = LBound(pvArr) To UBound(pvArr)
        For j = LBound(pbArr) To UBound(pbArr)
            'comparar
            If (UCase(pvArr(i)) = UCase(pbArr(j))) Then
                matches = matches + 1
            End If
        Next j
    Next i
    
    '**CASO MONICA MUCIÑO**
    If (pvStr = "Mucino Zarza, M. A. - MM23847") Then
        matches = matches + 1
    End If
    
    If (matches = 0) Then 'No se parecen, pensar que es otro
        compararPms = False
    ElseIf (matches = 1) Then 'solo se parece en un elemento (nombre o apellido) verificar
        'verificar si son de los que ya se han confirmado
        If cnt > 0 Then
            For j = LBound(lista) To UBound(lista)
                If (lista(j) = pvStr) Then
                    noNeed = True
                    compararPms = True
                End If
            Next j
        End If
        If (noNeed = False) Then
        Dim valor As Integer
        valor = MsgBox("Favor de verificar que ambos nombres de PMs sean el mismo" & Chr(10) & _
            "¿Ambos nombres corresponden al mismo PM?" & Chr(10) & Chr(10) & _
            pvStr & "  :y:  " & pbStr, vbYesNo, "Verificar Nombre")
            Select Case valor
                Case 7 'No
                    compararPms = False
                Case 6 'Yes
                    'Entonces es correcto
                    ReDim Preserve lista(cnt)
                    lista(cnt) = pvStr
                    cnt = cnt + 1
                    compararPms = True
                End Select
        End If
    ElseIf (matches > 1) Then 'Suponer que el PM es el correcto
        compararPms = True
    End If

End Function



Function compararPms2(pvStr As String, pbStr As String, lista() As String, faltan As Integer, listaPmP00() As String, P00 As String) As Boolean
'http://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'https://blog.udemy.com/vba-regex/
'https://msdn.microsoft.com/en-us/library/system.text.regularexpressions.regex.split(v=vs.110).aspx
    Dim pvArr() As String, pbArr() As String
    Dim strPattern As String
    strPattern = "\b[a-zA-Z]+\b"
    
    'Primero splitear cada string
    Call arrMatches(pvArr(), strPattern, pvStr)
    Call arrMatches(pbArr(), strPattern, pbStr)

    'ejemplos de strings tipicos.. de PB,"Cordero, Gabriel", de PV, "Santamaria, Jorge Alberto - JS80684"
    'Los vamos a comparar, con que coincidan 2, basta
    Dim pvArrLen As Integer, pbArrLen As Integer, matches As Integer
    pvArrLen = 0
    pbArrLen = 0
    For i = LBound(pvArr) To UBound(pvArr)
        For j = LBound(pbArr) To UBound(pbArr)
            'comparar
            If (UCase(pvArr(i)) = UCase(pbArr(j))) Then
                matches = matches + 1
            End If
        Next j
    Next i
    
    If (matches = 0) Then 'No se parecen, pensar que es otro
        ReDim Preserve lista(faltan), listaPmP00(faltan)
        lista(faltan) = pvStr & " && " & pbStr
        listaPmP00(faltan) = P00
        faltan = faltan + 1
        compararPms2 = False
    ElseIf (matches = 1) Then 'solo se parece en un elemento (nombre o apellido) verificar
        Dim valor As Integer
        valor = MsgBox("Favor de verificar que ambos nombres de PMs sean el mismo" & Chr(10) & _
            "¿Ambos nombres corresponden al mismo PM?" & Chr(10) & Chr(10) & _
            pvStr & "  :y:  " & pbStr, vbYesNo, "Verificar Nombre")
            Select Case valor
                Case 7 'No
                    ReDim Preserve lista(faltan), listaPmP00(faltan)
                    listaPmP00(faltan) = P00
                    lista(faltan) = pvStr & " && " & pbStr
                    faltan = faltan + 1
                    compararPms2 = False
                Case 6 'Yes
                    'Entonces es correcto
                    compararPms2 = True
                End Select
    ElseIf (matches > 1) Then 'Suponer que el PM es el correcto
        compararPms2 = True
    End If

End Function




