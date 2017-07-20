Public exit_all As Boolean

Sub Llenado_Luisa(num_proceso As Integer)

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
    Dim ruta As String, lista_rutas() As String, dir_arch As Variant, num_arch As Integer, verificar As Boolean
    'Ruta de donde sacará archivos
    If ((num_proceso = 6) Or (num_proceso = 1)) Then 'es para evitar la busqueda de archivos en caso de que no se necesite (solo para carga de)
        ruta = ActiveWorkbook.Path
        verificar = verificar_entorno(3, num_arch, ruta, lista_rutas(), dir_arch)
            If (verificar = False) Then
                Exit Sub
            End If
    End If
    
    verificar = verificar_limites(last_row_st, last_col_st)
        If (verificar = False) Then
            Exit Sub
        End If
    
'**
'ANÁLISIS DE LOS ARCHIVOS
'**
    Application.ScreenUpdating = False 'para que no se muevan las paginas/workbooks y se quede en la misma mientras trabaja
    'Por medio de un inputbox, dinámicamente preguntar cual es el archivo para X o Y acción

    'actualizar de archivo en PlanView a Reporte
    Dim reemplazar_fecha As String 'es una variable para reemplazar la fecha
    If ((num_proceso = 1) Or (num_proceso = 6)) Then
        Call registro_a_reporte(lista_rutas(2), mainwb, last_row_st, last_col_st, num_registros, reemplazar_fecha)
            If (exit_all = True) Then
                deshacer = deshacer_cambios_reporteYTD(last_row_st, last_col_st, num_registros, mainwb, exit_all)
                MsgBox ("Se saldrá del procedimiento, campos borrados")
                exit_all = False
                Exit Sub
            End If
        'actualizar los números de proyecto y estados
        Call proyectos_a_reporte(lista_rutas(1), mainwb, last_row_st, last_col_st, num_registros, reemplazar_fecha)
            If (exit_all = True) Then
                deshacer = deshacer_cambios_reporteYTD(last_row_st, last_col_st, num_registros, mainwb, exit_all)
                MsgBox ("Se saldrá del procedimiento, campos borrados")
                exit_all = False
                Exit Sub
            End If
    
        deshacer = deshacer_cambios_reporteYTD(last_row_st, last_col_st, num_registros, mainwb, False)
        If (deshacer = True) Then
            MsgBox ("Se saldrá del procedimiento, campos borrados")
            Exit Sub
        End If
        
        'dado que helpers está protegido, para no tener errores, deshabilitar captadores de eventos
        Application.EnableEvents = False
            'Colocar la fecha en cuestión en la pestaña helpers
            helpers.Cells(2, 1).Value = reemplazar_fecha
        Application.EnableEvents = True
    
        If (num_proceso = 6) Then
            Dim actualizar As Integer
            actualizar = MsgBox("Actualización del historial de registro de recursos hecha" & Chr(10) & Chr(10) & _
                "--> ¿Desea actualizar el contenido? <--", vbYesNo, "Actualización tablas")
                Select Case actualizar
                    Case 7
                        MsgBox ("No se actualizan tablas" & Chr(10) & Chr(10) & _
                        "Se saldrá de la aplicación")
                        Exit Sub
                    End Select
        End If
    End If
   
    Dim titulos_formulas(7) As String, titulos_origen(7) As String, titulos_destino(7) As String
    titulos_formulas(0) = "Tot Proyectos" '-> Suma de todos los elementos menos de OOO/Training
    titulos_formulas(1) = "Horas Capitalizadas" '->indicador (redundante) de las horas capitalizables
    titulos_formulas(2) = "% Capitalizacion " '-> % de división Hrs/Total de Proyectos
    titulos_formulas(3) = "Total general" '-> Suma de todos los elementos
    'Modificar el titulo de Meta directamente en el SUB de cambio de formatos (al principio)
    titulos_formulas(4) = "" '-> Total Proyectos * % (siendo % la meta a alcanzar)
    titulos_formulas(5) = "Cumplimiento Meta" '-> % de división Hrs/Meta (50%)
    titulos_formulas(6) = "OOO/ Training" 'Unicament se usa para un dato, no se coloca ahí fórmula
    titulos_formulas(7) = "Etapas Capitalizables" 'Unicament se usa para un dato, no se coloca ahí fórmula
    'titulos de coincidencia de la tabla dinámica donde se quiere captar una coincidencia y colocar en la tabla destino
    titulos_origen(0) = "Etiquetas de fila" '->Coincidencia para meter en destino la columna Team
    titulos_origen(1) = "Capitalization-Si" '->Coincidencia para colocar en Horas capitalizables
    titulos_origen(2) = "Capitalization-No" '->Coincidencia para colocar en Horas no capitalizables
    titulos_origen(3) = "Expense" '->Coincidencia para colocar en horas Expense
    titulos_origen(4) = "Not Defined" '->Coincidencia para colocar en horas Not Defined
    titulos_origen(5) = "Programa (No Capitalizable)" '->Coincidencia para colocar en columna de programa
    titulos_origen(6) = "Maint / Non Trad (No capitalizables)" '->Coincidencia para poner en hrs maintenance
    titulos_origen(7) = "OOO/Training" '-> Coincidencia para poner en hrs de Training
    'Titulos a matchear en la fila de titulos para ir colocando de acuerdo a coincidencais con titulos_origen en la tabla destino
    titulos_destino(0) = "Team" '->donde se pondrá lo de titulos_origen(0)
    titulos_destino(1) = "Etapas Capitalizables"
    titulos_destino(2) = "Etapas NO Capitalizables"
    titulos_destino(3) = "Proyectos: Expense"
    titulos_destino(4) = "Not Defined"
    titulos_destino(5) = "Program"
    titulos_destino(6) = "Maint / Non Traditional"
    titulos_destino(7) = "OOO/ Training"
    
    'Por si no se ejecutó la revisión completa y no se trae la fecha, puede que sea necesaria según el proceso
    If ((num_proceso <> 6)) Then
        Call fecha_en_cuestión(reemplazar_fecha)
    End If
        
    If ((num_proceso = 2) Or (num_proceso = 6) Or (num_proceso = 5)) Then
        'Actualización de la tabla team_sem, donde se ve el resulado de YTD de cada team
        Call actualizar_tablas(5, 1, 3, 2, 20, td_team, team_YTD, "Team", "Cumplimiento Meta", titulos_origen(), titulos_destino(), titulos_formulas(), False, reemplazar_fecha)
    End If
    
    If ((num_proceso = 3) Or (num_proceso = 6) Or (num_proceso = 5)) Then
        'ACTUALIZACIÓN DE LA TABLA de todos los equipos, en la semana
        Call actualizar_tablas(5, 1, 3, 2, 20, td_team, team_sem, "Team", "Cumplimiento Meta", titulos_origen(), titulos_destino(), titulos_formulas(), True, reemplazar_fecha)
    End If
    
    If ((num_proceso = 7) Or (num_proceso = 6) Or (num_proceso = 5)) Then
        titulos_destino(0) = "Recurso"
    'ACTUALIZACION de la tabla de los recursos
        Call actualizar_tablas(4, 1, 3, 2, 20, td_recursos, rec_sem, "Team", "Cumplimiento Meta", titulos_origen(), titulos_destino(), titulos_formulas(), True, reemplazar_fecha)
    End If
    
    If ((num_proceso = 4) Or (num_proceso = 6) Or (num_proceso = 5)) Then
        'ACTUALIZACION DE TABLA por periodo (sin filtros)
        titulos_origen(0) = "Bla Bla" '->Coincidencia para meter en destino la columna ->No queremos coincidencia para este
        'titulos_destino(0) = "Team" '->donde se pondrá lo de titulos_origen(0)
        Call actualizar_tablas(5, 1, 3, 2, 20, td_periodos, periodoYTD, "PERIODO", "Cumplimiento Meta", titulos_origen(), titulos_destino(), titulos_formulas(), False, reemplazar_fecha)
    End If

    Application.ScreenUpdating = True ' para volver a activar la diferenciación de cuando se trabaja en x o y página/wbk
    
    'Para verificar pechas Plaindromes (no olvidarse)
    Dim actDate As Date, newYear As Date
    actDate = Date
    newYear = #1/1/2018#
    If (Date < newYear + 30 And Date > newYear - 10) Then
        MsgBox "Favor de revisar fechas Palindromes y actualizar próximo aviso para atualización de este aviso" & Chr(10) & Chr(10) & "Gracias!!"
    End If
    
End Sub

'****/
'* Función que quita la primera parte de un string según se desee, mostrando así los nombres de los archivos unicamente junto con su relación
'****\
Function enlistar_sin_texto_inicial(texto_inicial As String, lista() As String) As String

    Dim i As Integer, depurado As String, comparar(2) As String
    
    comparar(0) = "Reporte total ---------> "
    comparar(1) = "Proyectos Vigentes ----> "
    comparar(2) = "Registro en PV --------> "
    
    For i = LBound(lista) To UBound(lista)
        enlistar_sin_texto_inicial = enlistar_sin_texto_inicial & comparar(i) & Replace(lista(i), texto_inicial, "") & vbCrLf
    Next i
    
End Function

'****/
'* Rutina que interactúa con el usuario para determinar si algunos elementos como la ruta y la lista de archivos
'* son correctos
'****\
Function verificar_entorno(max_min_archivos As Integer, num_archivos As Integer, ruta As String, lista_arch() As String, dir_arch As Variant) As Boolean
    
    'Variables para MSGBOX
    Dim lista_correcta As Integer, ruta_correcta As Integer, desplegar_num As Integer
    
    ruta_correcta = MsgBox(ruta, vbOKCancel, "Verificar si la ruta es Correcta")
    Select Case ruta_correcta
        Case 2
        MsgBox ("Ruta Incorrecta")
            verificar_entorno = False
            Exit Function
        End Select
    
'correr función para ver los archivos dentro de la carpeta
    'cuantos archivos hay
    num_archivos = CuantosArchivos(ruta) ' corre la función después descrita "cuantos archivos"
    If (num_archivos <> max_min_archivos) Then
        MsgBox ("Faltan/sobran archivos:" & Chr(10) & "deberían de ser " _
        & max_min_archivos & " pero hay " & num_archivos & Chr(10) & Chr(10) _
        & "--> Favor de colocar los archivos necesarios" & Chr(10) & Chr(10) _
        & "Por ejemplo:" & Chr(10) & " 1 Reporte Total" & Chr(10) _
        & " 2 Proyectos vigentes" & Chr(10) & " 3 Registro en PV")
        verificar_entorno = False
        Exit Function
    End If
    'desplegar_num = MsgBox("Cantidad de documentos" & Chr(10) & Chr(10) & num_archivos, , "# Documentos")
    
'almacenar archivos en un array con sus rutas
    ReDim lista_arch(num_archivos - 1) As String
    Call lista_archivos(ruta, lista_arch(), dir_arch)
    
'almacenar nombres en un array
    'desplegar lista de archivos para verificar
    lista_correcta = MsgBox("Verificar los archivos a trabajar, el orden" _
        & Chr(10) & Chr(10) & enlistar_sin_texto_inicial(ruta, lista_arch()), _
        vbOKCancel, "Lista de Archivos que se analizarán")
        Select Case lista_correcta
            Case 2
            MsgBox ("Verificar nombramiento archivos")
                verificar_entorno = False
                Exit Function
            End Select
    
    verificar_entorno = True
End Function

'****/
'* Función para interactuar con el usuario y ver si hya alguna fila vacía
'****\
Function verificar_limites(last_row_st As Integer, last_col_st As Integer) As Boolean
    
    Dim verif As Integer
    'siendo reduntante
    reporteYTD.Activate
    'antes que nada quitar filtros
    Call desactivar_filtro
    'dato curioso:
    last_col_st = Cells(1, 1).End(xlToRight).Column
    last_row_st = Cells(1, 3).End(xlDown).row
    
    If (Application.ScreenUpdating = False) Then
        Application.ScreenUpdating = True
        Cells(last_row_st, 1).Select
        Application.ScreenUpdating = False
    Else
        Cells(last_row_st, 1).Select
    End If
    
    verif = MsgBox("Favor de verificar que la última fila sea la " & last_row_st & Chr(10) & Chr(10) _
        & "Si no es correcto hacer click en NO, de otra manera, click en SI", vbYesNo, "Verificar última fila")
    Select Case verif
        Case 7
            MsgBox ("Como no es correcto, verificar que que no haya filas en blanco, si es así, eliminarlas por completo" _
            & Chr(10) & Chr(10) & "Gracias")
            verificar_limites = False
            Exit Function
        Case 7
            'colocarla en helpers para usarla si es necesario
            Application.EnableEvents = False
            helpers.Cells(6, 1).Value = last_row_st
            Application.EnableEvents = True
        End Select
    verificar_limites = True
    
End Function
'****/
'* uncamente borrar lo que se haya hecho en el archivo principal, en la pestaña de recursos
'****\
Function deshacer_cambios_reporteYTD(last_row As Integer, last_col As Integer, num_registros As Integer, wb As Workbook, exit_all As Boolean) As Boolean

    Dim deshacer As Integer, borrar_rango As Range
    
    If (num_registros = 0) Then
        'Para evitat que borre algo que no queremos, en este caso no se ha introducido nada entonces no tiene
        deshacer_cambios_reporteYTD = True
    ElseIf exit_all = True Then
        reporteYTD.Activate
        Range(Cells(last_row + 1, 1), Cells(last_row + num_registros, last_col)).Delete
        deshacer_cambios_reporteYTD = True
        'Regresar a sus valores anteriores
        Application.EnableEvents = False
        helpers.Cells(6, 1).Value = last_row 'Start Row
        helpers.Cells(8, 1).Value = 0 'Final Row
        helpers.Cells(10, 1).Value = 0 ' Total registros
        Application.EnableEvents = True
    Else
    deshacer = MsgBox("--> ¿Desea conservar los cambios? <--", vbYesNo, "Deshacer Cambios")
        Select Case deshacer
        Case 7
            reporteYTD.Activate
            Range(Cells(last_row + 1, 1), Cells(last_row + num_registros, last_col)).Delete
            deshacer_cambios_reporteYTD = True
        End Select
    End If
    
    
End Function

Public Sub elegir_proceso()

    ckbx_elegir_proceso.Show
    
End Sub
