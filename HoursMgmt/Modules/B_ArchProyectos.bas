'****/
'* Procedimiento que trabaja con el archivo donde están los proyectos de tarjetas, no tarjetas, calendario, recursos
'* y con el archivo objetivo siendo este el de reporte total
'* Lleva a cabo el llenado haciendo el match de la infomación del proyecto obtenida del archivo "source" de acuerdo
'* al proyecto que regiustró el recurso en cuestión (va linea por linea de los recursos buscando las caracerísticas del proyecto
'* Además también incluye, aunque a parte, el llenado de los proyectos y recursos no encontrados
'* y una pequeña rutina para asignación de periodo que se encuentra en la pestaña del calendario
'****\
Sub proyectos_a_reporte(origen As String, destino As Workbook, last_row As Integer, last_col As Integer, total_registros As Integer, fecha As String)

'VARIABLES
    'variable para objeto workbook
    Dim origwb As Workbook
    'variable para columnas y filas de archivos orig/dest
    Dim last_row_orig As Integer, last_column_orig As Integer, last_row_dest As Integer ', last_col_dest As Integer, last_row_rgt As Integer
    Dim key_row_dest As Integer, key_col_dest As Integer, key_row_orig As Integer, key_col_orig As Integer
    'variables para ir buscando a lo largo de los archivos los matchs
    Dim palabra_deseada As String, proyecto As String, match_proyecto As String, cards As Boolean
    'contadores generales
    Dim conteo_asignar As Integer, contar_proyectos As Integer, col As Integer, row As Integer
    'Para almacenar los proyectos no encontrados
    Dim proyectos_faltantes() As String, conteo_faltantes As Integer, recursos_faltantes() As String, num_miss_rec As Integer
    'To use worksheet codenames in other workbooks
    Dim vigentes_orig As Excel.Worksheet, faltantes_orig As Excel.Worksheet, otros_orig As Excel.Worksheet, _
        recursos_orig As Excel.Worksheet, fecha_orig As Excel.Worksheet, helpersP As Worksheet
    'para ver los valores nuevos a asignar ya que los nombre en la lista de proyectos no es la misma que en  el reporte
    Dim asignado As Boolean
    
    num_miss_rec = 0
    conteo_faltantes = 0
    contar_proyectos = 0
    
    'Buscar "Project" en la lista de títulos para así fijar esa columna de búsqueda en destino
    key_col_dest = buscar_en_fila(1, "Project", last_col, _
    "En Proyectos -> Reporte, tratando de fijar la columna de búsqueda")
    
    'Abrir archivo de donde se obtendrá la información a pegar
    Workbooks.Open (origen)
    Set origwb = ActiveWorkbook
    Set helpersP = GetWsFromCodeName(origwb, "helpers")
    'Antes que nada, hacer la revisión de sanidad para entender si sí se hizo el copy paste de los nuevos archivos
    If verificar_actualizacion_proyectos(helpersP) = False Then
        MsgBox "Se detectó que no se hizo bien la carga de proyectos," & Chr(10) & _
            "Se saldrá de la aplicación"
        'cerrar el archivo de origen
        origwb.Close savechanges:=False
        'pedir salir de todo
        exit_all = True
        Exit Sub
    End If
    
    Set recursos_orig = GetWsFromCodeName(origwb, "recursos")
    Set otros_orig = GetWsFromCodeName(origwb, "otros")
    Set faltantes_orig = GetWsFromCodeName(origwb, "faltantes")
    Set fecha_orig = GetWsFromCodeName(origwb, "periodos")
    Set vigentes_orig = GetWsFromCodeName(origwb, "vigentes")
    reporteYTD.Activate
    Call desactivar_filtro
    vigentes_orig.Activate
    Call desactivar_filtro
    'En este archivo, ver la última columna como protección de errores
    last_column_orig = Cells(1, 1).End(xlToRight).Column
    'Buscar "Nombre" en la lista de títulos para así fijar esa columna de búsqueda en origen
    key_col_orig = buscar_en_fila(1, "Name", last_column_orig, _
    "En Proyectos -> Reporte, tratando de fijar la columna de búsqueda")
    
    'antes algunaas declaraciones de variables
    'la búsqueda empezará a partir de la última fila + 1
    key_row_dest = last_row + 1
    'los arrays para los títulos de origen a destino: Array_origen(1) --> Array_destino(1)
    Dim titulos_orig(2) As String
    titulos_orig(0) = "Work Type"
    titulos_orig(1) = "SDLC Phase"
    titulos_orig(2) = "Capitalization Flag"
    Dim titulos_dest(3) As String
    titulos_dest(0) = "Project Type"
    titulos_dest(1) = "Etapa PV"
    titulos_dest(2) = "Capitalizable"
    titulos_dest(3) = "Cards/ No Cards"
        
    'Una vez teniendo la fila clave del origen, y la fila clave del destino, únicamente necesitamos
    'las columnas clave donde se irá colocando cada valor :
    '(hacer el loop Do..Loop mientras haya proyectos, o el contador iguale el total de proyectos)

    Do ' HAY UN PROBLEMA CON EL WHILE, POR ESO SE PUSO EL OTRO LIMITE PARA EL WHILE
        'Terna de valores que se asignará:
        Dim asignar(3) As String
        reporteYTD.Activate
        'entonces el proyecto es
        proyecto = Cells(key_row_dest, key_col_dest).Value
        'reinicializar el saber si está asignado o no(si es N/A o es Program o se asignó según valores)
        asignado = False
        'ver si es N/A o en realidad si es un proyecto que exista
        'de otra manera, asignar a "asignar()" los valores deseados dinámicos de acuerdo a los títulos y proyectos
        If (proyecto = "N/A") Then
            asignar(0) = ""
            asignar(1) = "N/A OOO/Training"
            asignar(2) = "OOO/Training"
            asignar(3) = ""
            key_row_orig = 0
        Else
            origwb.Activate
            key_row_orig = buscar_proyecto(proyecto, cards, origwb)
            'si no se entonctró proyecto, guardar para indicar a ususario completar más tarde
            If (key_row_orig < 0) Then
                MsgBox ("no se encontró el proyecto --> " & proyecto)
                ReDim Preserve proyectos_faltantes(conteo_faltantes)
                proyectos_faltantes(conteo_faltantes) = proyecto
                conteo_faltantes = conteo_faltantes + 1
            Else
                'si sí se encontró, entonces asignar valores a copiar en archivo destino
                palabra_deseada = ""
                conteo_asignar = 0
                col = 1
                
                'no es necesario pues el libro queda abierdo durante la búsqueda
        '        If (cards = True) Then
         '           vigentes_orig.Activate
          '      ElseIf (cards = False) Then
           '         otros_orig.Activate
            '    End If
                
                Do While ((col < last_column_orig + 1) And (conteo_asignar < 3) And asignado = False)
                    palabra_deseada = Cells(1, col).Value
                    If (palabra_deseada = titulos_orig(conteo_asignar)) Then
                        asignar(conteo_asignar) = Cells(key_row_orig, col).Value
                        conteo_asignar = conteo_asignar + 1
                    End If
                    col = col + 1
                Loop
            End If
            'Antes de pegar, cambiar unos valores si es necesario
            Call verificar_cambio(asignar())
        End If
        
        'DESTINO
        If (key_row_orig >= 0) Then
            'para que asigne la parte de cards/no cards
            Call asign_cards(cards, asignar(3), proyecto)
            'Una vez obtenida la terna de valores necesitada del origen, se asignan al destino
            reporteYTD.Activate '-> hoja donde se pondrán los valores
            palabra_deseada = ""
            conteo_asignar = 0
            col = 1
            Do While ((col < last_col + 1) And (conteo_asignar < 3))
                palabra_deseada = Cells(1, col).Value
                If (palabra_deseada = titulos_dest(conteo_asignar)) Then
                    Cells(key_row_dest, col).Value = asignar(conteo_asignar)
                    conteo_asignar = conteo_asignar + 1
                ElseIf (palabra_deseada = "Cards/ No Cards") Then
                    Cells(key_row_dest, col).Value = asignar(3)
                End If
                col = col + 1
            Loop
        End If

        key_row_dest = key_row_dest + 1
        contar_proyectos = contar_proyectos + 1
        
    Loop While ((proyecto <> "") And (contar_proyectos < total_registros))
    
    'borrar el contenido en faltantes antes de asignar nuevos faltantes o decir que no hay faltantes
    Call borrar_faltantes(1, 5, faltantes_orig)
    
    If (conteo_faltantes <> 0) Then
        Call colocar_faltantes(proyectos_faltantes(), conteo_faltantes, faltantes_orig, 1)
    End If
    
    Call asignar_recursos(origwb, destino, last_row, last_col, last_column_orig, total_registros)
    
    Call ubicar_periodo(fecha, fecha_orig, reporteYTD, last_row, last_col, total_registros)
        
    'Unicamente para mostrar el llenado
    If Not Application.ScreenUpdating Then
        Application.ScreenUpdating = True
        Cells(key_row_dest - 1, 1).Select
        Application.ScreenUpdating = False
    Else
        Cells(key_row_dest - 1, 1).Select
    End If
    
    MsgBox ("Se tuvieron " & total_registros & " registros." & Chr(10) & Chr(10) _
        & "Además la ultima fila ->" & (key_row_dest - 1))
        
        'Colocar en helpers para posible uso
        Application.EnableEvents = False
        helpers.Cells(10, 1).Value = total_registros
        helpers.Cells(8, 1).Value = key_row_dest - 1
        Application.EnableEvents = True
        
    'cerrar el archivo de origen
    origwb.Close savechanges:=True
        
End Sub

Sub asignar_recursos(origen As Workbook, destino As Workbook, ByVal last_row_dest As Integer, ByVal last_col_dest As Integer, ByVal last_col_orig As Integer, ByVal total_registros As Integer)
    
    Dim key_row_dest As Integer, key_col_dest_resource As Integer, key_col_dest_team As Integer
    Dim row As Integer
    Dim resource As String, resource_match As String, match As Boolean, team As String, recursos_faltantes() As String
    Dim num_recurso As Integer, conteo_faltantes As Integer
    Dim faltantes_orig As Excel.Worksheet, recursos_orig As Excel.Worksheet
    
    destino.Activate
    reporteYTD.Activate
    key_col_dest_resource = buscar_en_fila(1, "Resource", last_col_dest, _
        "Está en la asignación de recursos, detectando una columna clave")
    key_col_dest_team = buscar_en_fila(1, "Team", last_col_dest, _
        "Está en la asignación de recursos, detectando una columna clave")
    key_row_dest = last_row_dest + 1
    conteo_faltantes = 0
    origen.Activate
    Set recursos_orig = GetWsFromCodeName(origen, "recursos")
    Set faltantes_orig = GetWsFromCodeName(origen, "faltantes")
    recursos_orig.Activate
        'suponiendo que se quedan sólo dos columnas, no nos interesa si el last_column de esa hoja
        last_row = Cells(1, 1).End(xlDown).row
        
    For num_recurso = 0 To total_registros - 1
        destino.Activate
        resource = Cells(key_row_dest, key_col_dest_resource).Value
        recursos_orig.Activate
        match = True
            'buscar coincidencia en row
            resource_match = ""
            row = 0
            Do While ((resource <> resource_match) And (last_row + 1 > row))
            row = row + 1
            resource_match = Cells(row, 2).Value
                If (row > last_row) Then
                    MsgBox ("No se encontró el recurso con el nombre: " & resource)
                    match = False
                    Exit Do
                End If
            Loop
            
            If (match = True) Then
                'copiar el valor
                team = Cells(row, 1).Value
                'pegar el valor de team
                destino.Activate
                reporteYTD.Activate
                Cells(key_row_dest, key_col_dest_team).Value = team
            ElseIf (match = False) Then
                ReDim Preserve recursos_faltantes(conteo_faltantes)
                recursos_faltantes(conteo_faltantes) = resource
                conteo_faltantes = conteo_faltantes + 1
            End If
            key_row_dest = key_row_dest + 1
    Next num_recurso
    
    If (conteo_faltantes <> 0) Then
        Call colocar_faltantes(recursos_faltantes(), conteo_faltantes, faltantes_orig, 5)
    End If

End Sub

'****/
'* Pequeña rutina que compara las fechas en la hoja de calendario para determinar el periodo y pegarlo en su respectiva columna
'****\

Private Sub ubicar_periodo(fecha As String, fws As Excel.Worksheet, dest As Excel.Worksheet, last_row_st As Integer, last_col_st As Integer, num_registros As Integer)

    Dim f_a As Date, f_d As Date, f As Date, last_row As Integer, row As Integer, periodo As Integer, key_col As Integer

    fws.Activate
    'de aceurdo a la sheet donde está el show, la primera fila es 2 y la columna de inferior es la 2 y sup la 3
    f = CDate(fecha)
    'comparar f con fecha anterior y después y en caso de que esté entre las dos, regresar el valor del periodo
    last_row = Cells(2, 2).End(xlDown).row
    
    For row = 2 To last_row
        f_a = Cells(row, 2).Value
        f_d = Cells(row, 3).Value
        If ((f_a <= f) And (f <= f_d)) Then
            periodo = Cells(row, 6).Value
            Exit For
        End If
    Next row
    
    dest.Activate
    key_col = buscar_en_fila(1, "Periodo", last_col_st, _
        "Está en ubicación de periodo (también revisar la hoja fuente)")
    
    For row = last_row_st + 1 To last_row_st + num_registros
        Cells(row, key_col).Value = periodo
    Next row

End Sub

'****/
'* Manda mensaje de la lista de lo que hace falta y
'* Enlista en la hoja sleccionada y columna seleccionada esta lista, un elemento por fila
'****\
Private Sub colocar_faltantes(array_faltantes() As String, num_faltantes As Integer, hoja As Excel.Worksheet, columna As Integer)

        MsgBox ("Faltaron los siguientes elementos:" & Chr(10) & enlistar(array_faltantes()) & Chr(10) & _
                "  --> Favor de completarlos a mano y luego completar lista en archivo PROYECTOS VIGENTES" & _
                Chr(10) & Chr(10) & "el listado está disponible en la pestaña del archivo PROYECTOS VIGENTES -> Faltantes")
                
        Dim row As Integer
                
        hoja.Activate
        row = 0
        While (num_faltantes <> 0)
            row = row + 1
            Cells(row, columna).Value = array_faltantes(row - 1)
            num_faltantes = num_faltantes - 1
        Wend
End Sub


'****/
'* pequeña rutina que borra el registro de los faltantes anteriores sea de proyectos o de recursos
'****\
Private Sub borrar_faltantes(col_proyectos As Integer, col_recursos As Integer, hoja As Excel.Worksheet)
    
    Dim last_row As Integer
    hoja.Activate
    
    If (IsEmpty(Cells(2, col_proyectos))) Then
        Range(Cells(1, col_proyectos), Cells(1, col_proyectos)).ClearContents
    Else
        last_row = Cells(1, col_proyectos).End(xlDown).row
        If (last_row > 0) Then
            Range(Cells(1, col_proyectos), Cells(last_row, col_proyectos)).ClearContents
        End If
    End If
    If (IsEmpty(Cells(2, col_recursos))) Then
        Range(Cells(1, col_recursos), Cells(1, col_recursos)).ClearContents
    Else
        last_row = Cells(1, col_recursos).End(xlDown).row
        If (last_row > 0) Then
            Range(Cells(1, col_recursos), Cells(last_row, col_recursos)).ClearContents
        End If
    End If
End Sub

'****/
'* Cambiar el estado de cards para entender si es o no un poryecto de tarjetas u de otro dominio
'****\
Private Sub asign_cards(cards As Boolean, asign As String, proyecto)

    If (proyecto <> "N/A") Then
        If (cards = True) Then
            asign = "Cards"
        ElseIf (cards = False) Then
            asign = "No cards"
        End If
    End If

End Sub
'****/
'* La intención es cambiar los valores de la asignación según se necesite
'****\
Private Sub verificar_cambio(asignar() As String)

    If (asignar(0) = "Program") Then
        asignar(2) = "Programa (No Capitalizable)"
    ElseIf (asignar(0) = "Maintenance") Then
        asignar(2) = "Maint / Non Trad (No capitalizables)"
    ElseIf (asignar(1) = "Post Implementation" And asignar(2) <> "Expense") Then
        asignar(2) = "Capitalization-No"
    ElseIf (asignar(2) = "Undetermined") Then
        asignar(2) = "Not Defined"
    ElseIf (asignar(2) = "Capital") Then
        asignar(2) = "Capitalization-Si"
    End If
    
    'Para corregir el estado de Definition E Initiation
    'Generalmente estos dos salen con Not Defined (pues aún no se sabe si seran capitalizables o no)
    'Entonces cuando están en Definition e Initiation y tienen flag de Capitalizables, esas horas se tienen que ir a Non Capitalizables
    If ((asignar(1) = "Definition") And (asignar(2) = "Capitalization-Si")) Then
        asignar(2) = "Capitalization-No"
    ElseIf ((asignar(1) = "Initiation") And (asignar(2) = "Capitalization-Si")) Then
        asignar(2) = "Capitalization-No"
    End If
    

End Sub

'***/
'* Función/sub que busca el proyecto deseado en la hoja deseada (en dos hojas, sea la de cards y la de no cards)
'* SE PUEDE MEJORAR / -> Haciéndolo Sub+función, haciendolo reiterativo (si es posible)
'***\
Function buscar_proyecto(proyecto As String, cards As Boolean, archivo As Workbook) As Integer

    Dim last_row As Integer, last_col As Integer, row As Integer
    Dim proyecto_match As String
    Dim wsPivot As Excel.Worksheet
    
    'BUSCAR EN HOJA DE PROYECTOS VIGENTES DE TARJETAS
    Set wsPivot = GetWsFromCodeName(archivo, "vigentes")
    wsPivot.Activate

    last_row = Cells(1, 1).End(xlDown).row
    last_col = Cells(1, 1).End(xlToRight).Column

    'buscar coincidencia en row
    row = 1
    Do While (proyecto <> proyecto_match)
    row = row + 1
    proyecto_match = Cells(row, 1).Value
        If (row > last_row) Then
            'MsgBox ("No se encontró la palabra")
            row = -1
            Exit Do
        End If
    Loop
    
    'BUSCAR EN HOJA DE PROYECTOS VIGENTES DE "NO CARDS" U "OTROS"
    'buscar coincidencia en row
    If (row < 0) Then
        Set wsPivot = GetWsFromCodeName(archivo, "otros")
        wsPivot.Activate
        
        last_row = Cells(1, 1).End(xlDown).row
        last_col = Cells(1, 1).End(xlToRight).Column
        
        row = 1
        Do While (proyecto <> proyecto_match)
        row = row + 1
        proyecto_match = Cells(row, 1).Value
            If (row > last_row) Then
                'MsgBox ("No se encontró la palabra")
                row = -1
                Exit Do
            End If
        Loop
        cards = False
    ElseIf (row > 0) Then
        cards = True
    End If
    'si aún buscando en ambas hojas no se encuentra, significa que no está y que hay que agregarlo, se manda el valor negativo
    buscar_proyecto = row
    
End Function
