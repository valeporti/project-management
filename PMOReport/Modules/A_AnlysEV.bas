Public exit_all As Boolean

Public Sub analisarEVyEAC()

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
    ReDim funcionalidades(1)
    'la funcionalidad "neutra" es la del archivo origen
    funcionalidades(0) = "Archivo Principal"
    funcionalidades(1) = "Extracción PlanView de Proyectos"
    'Ruta de donde sacará archivos
    verificar = verificar_entorno(2, num_arch, mainwb, lista_rutas(), dir_arch, funcionalidades(), False)
        If (verificar = False) Then
            Exit Sub
        End If
'**
'ANÁLISIS DE LOS ARCHIVOS
'**
    'Application.ScreenUpdating = False 'para que no se muevan las paginas/workbooks y se quede en la misma mientras trabaja
    '(idea)Por medio de un inputbox, dinámicamente preguntar cual es el archivo para X o Y acción
    Call reporteProyVig(lista_rutas(0), mainwb, 1)
    If exit_all = True Then
        Exit Sub
    End If
    
    'Application.ScreenUpdating = True ' para volver a activar la diferenciación de cuando se trabaja en x o y página/wbk
    
End Sub
Private Sub reporteProyVig(origen As String, destino As Workbook, ttl_row As Integer)

    Dim origwb As Workbook
    Workbooks.Open (origen)
    Set origwb = ActiveWorkbook
    origwb.Activate
    
    'Antes que nada, hacer la revisión de sanidad para entender si sí se hizo el copy paste de los nuevos archivos
    Dim verq As Variant, sh As Worksheet
    verq = verUnaSh(origwb, 1, "Favor de escribir la hoja donde está la extracción de Proyectos de PlanView", "Atención, hojas no esperadas", "Hoja1")
    If verq = False Then
        MsgBox "No se detectó la hoja de donde obtener la información," & Chr(10) & _
            "Se saldrá de la aplicación"
        'cerrar el archivo de origen
        origwb.Close savechanges:=False
        'pedir salir de todo
        exit_all = True
        Exit Sub
    End If
    'Hoja que se usará
    If (verq = True) Then
        Set sh = origwb.Worksheets(1)
    Else ' significa que tenemos un string
        Set sh = origwb.Worksheets(verq)
    End If
    sh.Activate
    Call desactivar_filtro
    
    Dim lr As Integer, lc As Integer, row As Integer, ttl_range As Range
    Dim titulos_orig(3) As String, ttl_wIDran As Range, ttl_SDLCran As Range, ttl_wTyran As Range, ttl_wStran As Range
    titulos_orig(0) = "Work ID #" 'podria servir más no inmediatamente
    titulos_orig(1) = "SDLC Phase" 'hay que discriminar los que estén en Initiation y Definition
    titulos_orig(2) = "Work Type" 'Solo Majors y Minors
    titulos_orig(3) = "Work Status" 'Sólo los que estén en Commited nos interesan
    Dim titulos_analisis(4) As String, ana_EVran As Range, ana_PVran As Range, ana_EffActran As Range, _
        ana_BLran As Range, ana_EffTotran As Range
    titulos_analisis(0) = "EV-Earned Value (h)"
    titulos_analisis(1) = "EV-Planned Value (h)"
    titulos_analisis(2) = "Effort Actual (h)"
    titulos_analisis(3) = "Baseline Effort (h)"
    titulos_analisis(4) = "Effort Total (h)"
    
    lr = sh.Cells(ttl_row, 1).End(xlDown).row
    lc = sh.Cells(ttl_row, 1).End(xlToRight).Column
    Set ttl_range = Range(Cells(ttl_row, 1), Cells(ttl_row, lc))
    Dim ttl_wID As Integer, ttl_SDLC As Integer, ttl_wTy As Integer, ttl_wSt As Integer, ana_EV As Integer, _
        ana_PV As Integer, ana_EffAct As Integer, ana_BL As Integer, ana_EffTot As Integer
    'colocar los # columna para poderlos usar posteriormente
    Set ttl_wIDran = ttl_range.Find(titulos_orig(0))
    Set ttl_SDLCran = ttl_range.Find(titulos_orig(1))
    Set ttl_wTyran = ttl_range.Find(titulos_orig(2))
    Set ttl_wStran = ttl_range.Find(titulos_orig(3))
    Set ana_EVran = ttl_range.Find(titulos_analisis(0))
    Set ana_PVran = ttl_range.Find(titulos_analisis(1))
    Set ana_EffActran = ttl_range.Find(titulos_analisis(2))
    Set ana_BLran = ttl_range.Find(titulos_analisis(3))
    Set ana_EffTotran = ttl_range.Find(titulos_analisis(4))
    'verificar que estén todos los títulos
    If (ttl_wIDran Is Nothing Or ttl_SDLCran Is Nothing Or ttl_wTyran Is Nothing Or ttl_wStran Is Nothing Or _
        ana_EVran Is Nothing Or ana_PVran Is Nothing Or ana_EffActran Is Nothing Or ana_BLran Is Nothing Or ana_EffTotran Is Nothing) Then
        MsgBox "no se encontró algún título, favor de verificar, por ejemplo, que no haya espacios de más en los títulos"
        Exit Sub
    Else
        'volver los rangos a la columna.. otras variables
        ttl_wID = ttl_wIDran.Column
        ttl_SDLC = ttl_SDLCran.Column
        ttl_wTy = ttl_wTyran.Column
        ttl_wSt = ttl_wStran.Column
        ana_EV = ana_EVran.Column
        ana_PV = ana_PVran.Column
        ana_EffAct = ana_EffActran.Column
        ana_BL = ana_BLran.Column
        ana_EffTot = ana_EffTotran.Column
    End If
    
    'Preparar entorno de inserción
    Dim calculados As Integer
    calculados = 6
    Dim ttl_extras() As String
    ReDim ttl_extras(calculados)
    ttl_extras(0) = "CPI"
    ttl_extras(1) = "SPI"
    ttl_extras(2) = "% Consumido"
    ttl_extras(3) = "Esfuerzo remanente"
    ttl_extras(4) = "EAC"
    ttl_extras(5) = "Variacion EAC"
    ttl_extras(6) = "Con Desviación"
    
    'Antes de colocar las nuevas rows, haremos espacio en el lugar donde las queremos.
    'Naaa.. solo si lo piden o de verdad es muy necesario, verificar luego
    
    For col = 0 To calculados
        Cells(ttl_row, col + lc + 1).value = ttl_extras(col) 'Colocar títulos al final
    Next
    'se copia el formato para los nuevos títulos
    Range(Cells(ttl_row, lc), Cells(ttl_row, lc)).Select
    Selection.Copy
    Range(Cells(ttl_row, lc + 1), Cells(ttl_row, lc + calculados + 1)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'si si se encontró todo, se puede ir row por row analiazando
    Dim EV As Double, PV As Double, EffAct As Double, BL As Double, EffTot As Double
    Dim cpi As Double, spi As Double, consumido As Double, remEff As Double, EAC As Double, varEAC As Double
    Dim inDef As Boolean

    For row = ttl_row + 1 To lr
        'sh.Activate
        'El work type tiene que ser Major o Minor Project
        If (Cells(row, ttl_wTy).value = "Major Project" Or Cells(row, ttl_wTy).value = "Minor Project") Then
        Else
            GoTo NEXTFor
        End If
        'Mantener track si está en Initiation y Definition para procesarlo
        If (Cells(row, ttl_SDLC).value = "Initiation" Or Cells(row, ttl_SDLC).value = "Definition") Then
            inDef = True
        Else
            inDef = False
        End If
        'Ignorar los Cancelled o Completed en WorkStatus
        If (Cells(row, ttl_wSt).value = "Cancelled" Or Cells(row, ttl_wSt).value = "Completed") Then
            GoTo NEXTFor
        End If
        
        'Ya que pasó las pruebas, podemos iniciar asignando valores y luego cálculos
        EV = Cells(row, ana_EV).value
        PV = Cells(row, ana_PV).value
        EffAct = Cells(row, ana_EffAct).value
        BL = Cells(row, ana_BL).value
        EffTot = Cells(row, ana_EffTot).value
        If (PV <> 0 And EffAct <> 0 And BL <> 0) Then
            cpi = Round(EV / EffAct, 6) 'verde +/- 5%, amarillo +/- 10% , rojo > +/-10%
            spi = Round(EV / PV, 6) 'verde +/- 5%, amarillo +/- 10% , rojo > +/-10%
            consumido = Round(EffAct / BL, 6)
            remEff = EffTot - EffAct
            EAC = EffAct + remEff '(EAC = Estimated At Completion)
            varEAC = Round((EAC - BL) / BL, 6) 'Alertar si es > 10%
            'Agregar extras
            Cells(row, lc + 1).value = cpi
            Cells(row, lc + 2).value = spi
            Cells(row, lc + 3).value = consumido
            Cells(row, lc + 4).value = remEff
            Cells(row, lc + 5).value = EAC
            Cells(row, lc + 6).value = varEAC
        Else
            MsgBox "En la línea " & row & " (anotar el proyecto) faltan valores importanes, ya sea: EV-Planned Value (h) o Effort Actual (h) o Baseline Effort (h)"
            'Agregar extras
            If (PV = 0) Then
                Cells(row, lc + 2).value = "Falta EV-Planned Value (h)"
            End If
            If (EffAct = 0) Then
                Cells(row, lc + 1).value = "Falta Effort Actual (h)"
                Cells(row, lc + 4).value = "Falta Effort Actual (h)"
                Cells(row, lc + 5).value = "Falta Effort Actual (h)"
            End If
            If (BL = 0) Then
                Cells(row, lc + 3).value = "Falta Baseline Effort (h)"
            End If
            Cells(row, lc + 6).value = "Falta Baseline Effort (h) o Effort Actual (h)"
            GoTo NEXTFor
        End If

        'Sólo nos interesa alertar aquellos que no pasen las pruebas de SPI,CPI o de varEAC
        'colorearlos
        If (((cpi < 0.952 Or cpi > 1.048) Or (spi < 0.952 Or spi > 1.048) Or (varEAC < -0.098 Or varEAC > 0.098)) And (inDef = False)) Then
            Rows(row & ":" & row).Select
            With Selection.Interior
                .pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Cells(row, lc + 7).value = "Si"
        'ElseIf si está en definition o Initiation, sólo califican la varEAC, por ello sólo nos interesa eso.
        ElseIf ((varEAC < -0.098 Or varEAC > 0.098) And (inDef = True)) Then
            Rows(row & ":" & row).Select
            With Selection.Interior
                .pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16764108
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Cells(row, lc + 7).value = "Si"
        End If
NEXTFor:
    Next
    
    'colocar formato tipo 0.00
    Range(Cells(ttl_row + 1, lc + 1), Cells(lr, lc + calculados)).Select 'Para no incluir lo de "desvicación si o no, no el % de variance"
    Selection.NumberFormat = "0.00"
    Range(Cells(ttl_row + 1, lc + 6), Cells(lr, lc + 6)).Select '% Variación EAC
    Selection.Style = "Percent"
    Range(Cells(ttl_row + 1, lc + 3), Cells(lr, lc + 3)).Select '% Consumido
    Selection.Style = "Percent"
    
    'origwb.Save
End Sub



