'****/
'* Función que obtiene el entorno de las hojas de fechas para minor/major y maintenance
'* Regresa error si hay algún elemento que no se encuentre sea de majory minor como de maintenance
'****\
Function entornoDT(wb As Workbook, mjWs As Worksheet, mnWs As Worksheet, ttlMjRng As Range, ttlMnRng As Range, _
    ttlrow As Integer, lrMn As Integer, lrMj As Integer, lcMn As Integer, lcMj As Integer, _
    colATomarEnCuentaParaLastRow As Integer, rMnRng As Range, rMjRng As Range, releaseCol As Integer)
    
    wb.Activate
    Dim ws As Worksheet, mnFnd As Boolean, mjFnd As Boolean
    mnFnd = False
    mjFnd = False
    For Each ws In Worksheets
        If ws.Name = "Release - Major" Then
            Set mjWs = ws
            mjFnd = True
        ElseIf ws.Name = "Release - Maintenance" Then
            Set mnWs = ws
            mnFnd = True
        End If
    Next ws
    'verificado que se hayan encontrado ambas hojas
    'si no avisar y salir
    If (mnFnd = False Or mjFnd = False) Then
        MsgBox ("No se encontró página, favor de verificar que se tenga los siguientes títulos: " & Chr(10) & _
        Chr(10) & "Para Major/minor: 'Release - Major'" & Chr(10) & "Para Maintenanance: 'Release - Maintenance'")
        entornoDT = False
        Exit Function
    End If
    'si continua, asignar lo que se espera de los rngos, para poder leer los títulos
    'primero hoja de major
    mjWs.Activate
    lcMj = mjWs.Cells(ttlrow, Columns.COUNT).End(xlToLeft).Column
    lrMj = mjWs.Cells(Rows.COUNT, colATomarEnCuentaParaLastRow).End(xlUp).row
    Set ttlMjRng = mjWs.Range(Cells(ttlrow, 1), Cells(ttlrow, lcMj + 1)) 'Pongo el "+1" pues por alguna razón a pesar de que la palabra buscada esté en el último elemento a buscar, no lo encuentra a menos que se alargue la busqueda por uno más, aunque esté vacío
    Set rMjRng = mjWs.Range(Cells(1, releaseCol), Cells(lrMj, releaseCol))
    mnWs.Activate
    lcMn = mnWs.Cells(ttlrow, Columns.COUNT).End(xlToLeft).Column
    lrMn = mnWs.Cells(Rows.COUNT, colATomarEnCuentaParaLastRow).End(xlUp).row
    Set ttlMnRng = mnWs.Range(Cells(ttlrow, 1), Cells(ttlrow, lcMn + 1)) 'Pongo el "+1" pues por alguna razón a pesar de que la palabra buscada esté en el último elemento a buscar, no lo encuentra a menos que se alargue la busqueda por uno más, aunque esté vacío
    Set rMnRng = mnWs.Range(Cells(1, releaseCol), Cells(lrMn, releaseCol))
    entornoDT = True
    
    If ((rMnRng.Find("Release") Is Nothing) Or (rMjRng.Find("Release") Is Nothing)) Then
        MsgBox "No se encontró el título 'Release' en el archivo de fechas de release (ej: revisar espacios)"
        entornoDT = False
        Exit Function
    End If
        
End Function


'****/
'* Función que analiza fecha (en string) y regresa fecha en Date
'****\
Function strToDate(f As String) As Date
    Dim año As Integer, dia As Integer, mes As Integer
    Dim primer_signo As Integer
    
    año = CInt(Right(f, 4))
    primer_signo = InStr(1, f, "/")
    dia = CInt(Left(f, primer_signo - 1))
    mes = CInt(Mid(f, primer_signo + 1, 2))
    
    strToDate = Format(DateSerial(año, mes, dia), "dd/mm/yyyy")

End Function

'****/
'* Función que analiza fecha (en string) y Release
'* Regresa falso si no se logró match y verdadero si hay match
'****\
Function analizaRyFecha(f As String, r As String, COUNT As Integer) As String
    
    If (f = "" Or r = "") Then
        analizaRyFecha = "Falta el Release o una fecha para análisis"
        Exit Function
    End If

    Dim añof As String, diaf As String, mesf As String, añoR As String, mesR As String
    Dim primer_signo As Integer
    
    añof = Right(f, 4)
    primer_signo = InStr(1, f, "/")
    'diaf = Left(f, primer_signo - 1)
    mesf = Mid(f, primer_signo + 1, 2)
    
    añoR = Left(r, 4)
    primer_signo = InStr(1, r, ".")
    If (Mid(r, 6, 1) = "1" And Len(r) = 6) Then
            mesR = "10"
    Else
        mesR = Mid(r, 6, Len(r) - primer_signo + 1)
    End If
           
    If (añof <> añoR Or mesf <> mesR) Then
        analizaRyFecha = "R: " & añoR & "." & mesR & " y Milestone: " & añof & "." & mesf
        COUNT = COUNT + 1
    Else
        analizaRyFecha = ""
    End If

End Function

'****/
'* Función que, regresa un string, después de extraer el el dia mes y año de un string con formato: mm/dd/aaaa
'****\
Function dd_mm_aa(fecha_a_convertir As String, formatoddmmaaa As Boolean) As String
    
    Dim año As String, dia As String, mes As String
    Dim primer_signo As Integer, segundo_signo As Integer
    
    'verificar si es fecha
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}"
    End With
    If (Not (regEx.Test(fecha_a_convertir))) Then
        dd_mm_aa = ""
        Exit Function
    End If
    
    
    año = Right(fecha_a_convertir, 4)
    primer_signo = InStr(1, fecha_a_convertir, "/")
    segundo_signo = InStr(primer_signo + 1, fecha_a_convertir, "/")
    If (formatoddmmaaa = True) Then
        dia = Left(fecha_a_convertir, primer_signo - 1)
        mes = Mid(fecha_a_convertir, primer_signo + 1, segundo_signo - primer_signo - 1)
    Else
        mes = Left(fecha_a_convertir, primer_signo - 1)
        dia = Mid(fecha_a_convertir, primer_signo + 1, segundo_signo - primer_signo - 1)
    End If
    'Call pruebaLogicaFecha(año, mes, dia) ' no lo necesitamos pues abemos que siempre será mes/dia/año
    If ((primer_signo - 1) < 2) Then
        mes = "0" & mes
    End If
    If ((segundo_signo - primer_signo - 1) < 2) Then
        dia = "0" & dia
    End If
    dd_mm_aa = dia & "/" & mes & "/" & año
    'dd_mm_aa = Format(convertir_fecha_texto, "yyyy-mm-dd")
    
End Function

'****/
'* Rutina que interactúa con el usuario para determinar si algunos elementos como la ruta y la lista de archivos
'* son correctos
'****\
Function verificar_entorno(max_min_archivos As Integer, num_archivos As Integer, wb As Workbook, lista_arch() As String, dir_arch As Variant, descFuncArch() As String, verificarPorUsuario As Boolean) As Boolean
    
    'Variables para MSGBOX
    Dim lista_correcta As Integer, ruta_correcta As Integer, desplegar_num As Integer, ruta As String
    ruta = wb.Path
    
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
    'almacenar archivos en un array con sus rutas
    ReDim lista_arch(num_archivos - 1 - 1) As String 'Es la lista de archivos SIN el archivo principal
    'ver si existen cosas a comparar en descFuncArch()
    Dim descFuncArchExistentes As Integer, i As Integer, exito As Boolean
    exito = False
    descFuncArchExistentes = 0
    'Si se tienen más de dos archivos en la carpeta, es posible que se necesite que se defina para qué se usará cada uno
    If num_archivos - 1 - 1 > 0 Then
        For i = LBound(descFuncArch) To UBound(descFuncArch) 'SIEMPRE se debe entregar funcionalidades
            descFuncArchExistentes = descFuncArchExistentes + 1
        Next i
    End If
    
    If (num_archivos < max_min_archivos) Then
        MsgBox ("Faltan archivos:" & Chr(10) & "deberían de ser " _
        & max_min_archivos & " pero hay " & num_archivos & Chr(10) & Chr(10) _
        & "--> Favor de colocar los archivos necesarios" & Chr(10) & Chr(10) _
        & "Por ejemplo:" & Chr(10) & " 1 Reporte Total" & Chr(10) _
        & " 2 Proyectos vigentes" & Chr(10) & " 3 Registro en PV")
        verificar_entorno = False
        Exit Function
    ElseIf num_archivos > max_min_archivos Or verificarPorUsuario = True Then 'verificar cuál archivo es para qué
        Call lista_archivos(lista_arch(), dir_arch, wb)
        exito = relacionarFuncionalidadArchivo(descFuncArch(), lista_arch(), ruta)
    ElseIf num_archivos = max_min_archivos Or verificarPorUsuario = False Then 'verificar que la asignación de funcionalidades sea la correcta
        Call lista_archivos(lista_arch(), dir_arch, wb)
        lista_correcta = MsgBox("Verificar los archivos a trabajar, el orden" _
            & Chr(10) & Chr(10) & enlistar_sin_texto_inicial(ruta, lista_arch(), wb.FullName, descFuncArch()), _
            vbOKCancel, "Lista de Archivos que se analizarán")
            Select Case lista_correcta
                Case 1
                    exito = True
                    verificar_entorno = True
                    Exit Function
                Case 2
                    If (descFuncArchExistentes > 0) Then
                        exito = relacionarFuncionalidadArchivo(descFuncArch(), lista_arch(), ruta)
                    End If
                    If exito = False Then
                        MsgBox ("Verificar nombramiento archivos")
                        verificar_entorno = False
                        Exit Function
                    End If
                End Select
        
    End If
    'Si algo sale mal, sólo se modifica "exit_all" y con ello se sale de la función y se entrega false
    If exito = False Then
        verificar_entorno = False
        Exit Function
    End If
    
'almacenar nombres en un array
    'desplegar lista de archivos para verificar
    If (num_archivos > 2) Then 'Ya es necesario analizar más si no simplemente sería el arhcivo maestro la base
        lista_correcta = MsgBox("Verificar los archivos a trabajar, el orden" _
            & Chr(10) & Chr(10) & enlistar_sin_texto_inicial(ruta, lista_arch(), wb.FullName, descFuncArch()), _
            vbOKCancel, "Lista de Archivos que se analizarán")
            Select Case lista_correcta
                Case 2
                    MsgBox ("Verificar el nombre de archivos, o la asignación." & Chr(10) & "Volver a correr programa")
                    verificar_entorno = False
                    Exit Function
                End Select
    End If
    
    verificar_entorno = True
End Function

'***/
'* Función para contar los archivos hay en la ruta especificada
'***\
Function CuantosArchivos(ruta As String) As Integer

    'se obiene el primer archivo buscado en la ruta seleccionada
    dirarchivos = Dir(ruta & "\")
    'MsgBox (lista1)
    CuantosArchivos = 0
    'Función que mientras haya archivos los cuenta
    Do While dirarchivos <> ""
        CuantosArchivos = CuantosArchivos + 1
        'declarar "dir()" vacía llama al siguiente miembro de la lista
        dirarchivos = Dir()
    Loop
End Function

'****/
'* Función que quita la primera parte de un string según se desee, mostrando así los nombres de los archivos unicamente junto con su relación
'****\
Function enlistar_sin_texto_inicial(texto_inicial As String, lista() As String, principal As String, funcionalidades() As String) As String

    Dim i As Integer, depurado As String, newlista() As String
    ReDim newlista(0)
    newlista(0) = principal
    For i = LBound(lista) To UBound(lista)
        ReDim Preserve newlista(i + 1)
        newlista(i + 1) = lista(i)
    Next i
    
    enlistar_sin_texto_inicial = ""
    For i = LBound(newlista) To UBound(newlista)
        enlistar_sin_texto_inicial = enlistar_sin_texto_inicial & funcionalidades(i) & " ---------> " & Replace(newlista(i), texto_inicial, "") & vbCrLf
    Next i
    
End Function

'***/
'* Función que permite usar code names de otros workbooks ya que en realidad, la única manera de usar un
'* coddename es en el mismo workbook donde está escrito el código de vba
'* Return Value --> worksheet where to work in
'***\
Function GetWsFromCodeName(wb As Workbook, codename As String) As Excel.Worksheet

    Dim ws As Excel.Worksheet
    
    For Each ws In wb.Worksheets
        If ws.codename = codename Then
            Set GetWsFromCodeName = ws
            Exit For
        End If
    Next ws
    
End Function

'****/
'* para cada funcionalidad de archivo esperada (previamente asignada de acuerdo a actividades a ejecutar)
'* se relacionará cada archivo, para tener así claro en dónde actuará la macro
'****\
Function relacionarFuncionalidadArchivo(funcionalidades() As String, listaArch() As String, texto_inicial As String) As Boolean

    Dim i As Integer, numFunc As Integer, j As Integer, numarch As Integer
    
    numFunc = 0
    numarch = 1 '1 para poder tomar en cuenta el archivo principal
    For i = LBound(funcionalidades) To UBound(funcionalidades)
        numFunc = numFunc + 1
    Next i
    For j = LBound(listaArch) To UBound(listaArch)
        numarch = numarch + 1
    Next j

    'relacionar cada archivo
    Dim numArchYaEnlistados() As Integer, listaDesplegar As String, k As Integer, coincidencia As Boolean, first As Boolean
    Dim countArch As Integer, coin As Boolean ', archYatempo() As Integer, temp As Integer
        '1 - colocar cada archivo en una lista para presentar y el usuario sepa que número colocar
        
        'for para cada funcionalidad
        first = True
        countArch = 0
        For i = 1 To numFunc - 1 'empieza en uno para no tomar la funcionalidad del archivo principal
            listaDesplegar = ""
            coincidencia = False
            'for para cada archivo: hacer nueva lista cada vez, con los archivos restantes
            For k = 1 To numarch - 1 'empieza en uno para no tomar el archivo principal
                'for para comparación de archivos ya elegidos y evitar mostrarlos
                coin = False
                If first = False Then
                    For j = LBound(numArchYaEnlistados) To UBound(numArchYaEnlistados)
                        If numArchYaEnlistados(j) = k Then
                            coin = True
                        End If
                    Next j
                    If coin = False Then
                        listaDesplegar = listaDesplegar & "(" & k & ") " & Replace(listaArch(k - 1), texto_inicial, "") & vbCrLf
                    End If
                Else
                    listaDesplegar = listaDesplegar & "(" & k & ") " & Replace(listaArch(k - 1), texto_inicial, "") & vbCrLf
                End If
                
            Next k
            'Pedir al usuario poner el número del archivo correspondiente a la funcionalidad
            relacion = InputBox("Los archivos son:" & vbCrLf & vbCrLf & listaDesplegar & vbCrLf & "Favor de colocar el número correspondiente al uso." & _
                Chr(10) & "USO: " & funcionalidades(i), "Relacionar Archivos con Su Uso", "00")
            'almacenar el número obtenido para asignación de funcionalidades con archivos o volver a preguntar.
            If IsNumeric(relacion) And (relacion > 0 And relacion < numarch) Then
                ReDim Preserve numArchYaEnlistados(countArch)
                numArchYaEnlistados(countArch) = relacion
                first = False
                countArch = countArch + 1
            ElseIf relacion = vbNullString Then
                'significa que el usuario puso "cancelar"
                relacionarFuncionalidadArchivo = False
                Exit Function
            Else
                i = i - 1 'volver a la funcionalidad (para volver a preguntar hasta que coloque un número
                MsgBox "ERROR" & Chr(10) & Chr(10) & "Favor de escribir un número entre 1 y " & numarch - 1 & " incluidos"
            End If
        Next i
    
    'asignar a nuevo array el conenido de listaArch() pues se modificará
    Dim newList() As String
    For i = 0 To numarch - 1 - 1
        ReDim Preserve newList(i)
        newList(i) = listaArch(i)
    Next i
    'Reordenar archivos para su asignación de acuerdo a usuario.
    ReDim listaArch(numFunc - 1 - 1)
    For i = 0 To numFunc - 1 - 1
        ReDim Preserve listaArch(i)
        'If numArchYaEnlistados(i) <> i + 1 Then
            listaArch(i) = newList(numArchYaEnlistados(i) - 1)
        'End If
    Next i
    
    relacionarFuncionalidadArchivo = True

End Function

'****/
'* Busca si sólo hay una sheet en la que se vaya a trabajar o si hay más preguntar con cuál se trabaja
'****\
Function verUnaSh(wb As Workbook, hojas_esperadas As Integer, mensaje As String, titulo As String, ejemploValEsperado As String)
    Dim nsh As Integer, inp As String, shO As Object, repeat As Integer, found As Boolean
    nsh = wb.Worksheets.COUNT
    verUnaSh = False
    
    If nsh > hojas_esperadas Then
        For repeat = 0 To 1 'solo volver a preguntar una vez más
            inp = InputBox(mensaje, titulo, ejemploValEsperado)
            For Each shO In wb.Worksheets
                If shO.Name = inp Then
                    verUnaSh = inp
                    Exit Function
                End If
            Next shO
        Next repeat
    ElseIf nsh = hojas_esperadas Then
        verUnaSh = True
        Exit Function
    End If
End Function

'****/
'* Busca si sólo hay una sheet en la que se vaya a trabajar o si hay más preguntar con cuál se trabaja
'****\
Function verUnaShExt(wb As Workbook, ws As Worksheet, lr As Integer, lc As Integer, ttlRng As Range, ttlrow As Integer, hojas_esperadas As Integer, mensaje As String, titulo As String, ejemploValEsperado As String)
    Dim nsh As Integer, inp As String, shO As Object, repeat As Integer, found As Boolean
    nsh = wb.Worksheets.COUNT
    verUnaShExt = False
    
    If nsh > hojas_esperadas Then
        For repeat = 0 To 1 'solo volver a preguntar una vez más
            inp = InputBox(mensaje, titulo, ejemploValEsperado)
            For Each shO In wb.Worksheets
                If shO.Name = inp Then
                    verUnaShExt = inp
                    'Exit Function
                    GoTo nextStep
                End If
            Next shO
        Next repeat
    ElseIf nsh = hojas_esperadas Then
        verUnaShExt = True
        'Exit Function
        GoTo nextStep
    End If

nextStep:
    If verUnaShExt = False Then
        MsgBox "No se detectó la hoja de donde obtener la información," & Chr(10) & _
            "Se saldrá de la aplicación"
        'cerrar el archivo de origen
        wb.Close savechanges:=False
        'pedir salir de todo
        verUnaShExt = False
        Exit Function
    End If
    'Hoja que se usará
    If (verUnaShExt = True) Then
        Set ws = wb.Worksheets(1)
    Else ' significa que tenemos un string
        Set ws = wb.Worksheets(verUnaShExt)
    End If
    ws.Activate
    Call desactivar_filtro
    
    Dim temprow As Integer
    If (ws.Cells(ttlrow + 1, 3).value = "") Then
        temprow = ttlrow - 1
    Else
        temprow = ttlrow
    End If
    lr = ws.Cells(temprow, 3).End(xlDown).row
    lc = ws.Cells(ttlrow, Columns.COUNT).End(xlToLeft).Column
    Set ttlRng = ws.Range(Cells(ttlrow, 1), Cells(ttlrow, lc))
    
End Function

'****/
'* Busca en un rango, sea columna o fila, y agrega o acumula en un array lo no encontrado que se buscó
'* si encuentra regresa el número de la fila o columna,
'* si no encuentra nada regresa 0
'****\
Function busca_EnRango_falta(rng As Range, palabra As String, columnaOfila As String, exactCoin As Boolean, faltantes() As String, faltan As Integer) As Integer
        With rng
            If (exactCoin = True) Then
                Set c = .Find(palabra, LookIn:=xlValues, MatchCase:=True, LookAt:=xlWhole) 'MatchCase:= true para que busque el caso exacto
            Else
                Set c = .Find(palabra, LookIn:=xlValues) 'MatchCase:= true para que busque el caso exacto
            End If
            'Significa que la fila ya está llenada y se pueden seguir colocando valores ahí
            If c Is Nothing Then
                'significa que no encontró la palabra
                ReDim Preserve faltantes(faltan)
                faltantes(faltan) = palabra
                faltan = faltan + 1
                busca_EnRango_falta = 0
            Else
                'pasar el número de columna/fila
                If (columnaOfila = "columna") Then
                    busca_EnRango_falta = c.Column
                ElseIf (columnaOfila = "fila") Then
                    busca_EnRango_falta = c.row
                End If
            End If
        End With
End Function

'****/
'* simplemente busca en un rango y si encuentra regresa el número de la fila o columna,
'* si no encuentra nada regresa 0
'****\
Function busca_EnRango(rng As Range, palabra As String, columnaOfila As String, exactCoin As Boolean) As Integer
    Dim ultimaOp As Boolean, cnt As Integer, empiezaEn As Integer
    ultimaOp = False
    cnt = 0
        With rng
            If (exactCoin = True) Then
                Set c = .Find(palabra, LookIn:=xlValues, MatchCase:=True, LookAt:=xlWhole) 'MatchCase:= true para que busque el caso exacto
            Else
                Set c = .Find(palabra, LookIn:=xlValues)
            End If
            'Significa que la fila ya está llenada y se pueden seguir colocando valores ahí
            If c Is Nothing Then
                'primero tratar con un loop (por si merged cells)
                If (columnaOfila = "columna") Then
                    empiezaEn = rng.Column
                ElseIf (columnaOfila = "fila") Then
                    empiezaEn = rng.row
                End If
                For Each srVal In rng.Value2
                    If (srVal = palabra) Then
                        ultimaOp = True
                        busca_EnRango = cnt + empiezaEn ' empiezaEn ya es por default 1
                    End If
                    cnt = cnt + 1
                Next srVal
                If (ultimaOp = False) Then
                    'significa que no encontró la palabra
                    MsgBox ("No se encontró la " & columnaOfila & " con título: " & palabra)
                    busca_EnRango = 0
                End If
            Else
                'pasar el número de columna/fila
                If (columnaOfila = "columna") Then
                    busca_EnRango = c.Column
                ElseIf (columnaOfila = "fila") Then
                    busca_EnRango = c.row
                End If
            End If
        End With
End Function

'****/
'* compara 2 strings de varias palabras, separa las palabras y las coloca en un array
'* luego compara los carácteres de letra únicamente y ve si hay coincidencias y decide al respecto
'****\
Function compararDos(pvStr As String, pbStr As String) As Boolean
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
        compararDos = False
    ElseIf (matches > 0) Then ' es lo unico que nos interesa
        compararDos = True
    End If

End Function

'****/
'* Función que busca en rango una variable, si no existe el rango lo asinga, si sí, pasa por alto la asignación
'****\
Function busca_EnRangoV2(colRowConstant As Integer, firstRorC As Integer, lastRorC As Integer, rng As Range, palabra As String, columnaOfila As String, exactCoin As Boolean, Optional scndCoin, Optional rowColScndCoin As Integer) As Integer
    Dim tmp As String
        If rng Is Nothing Then
            If (columnaOfila = "columna") Then
                Set rng = Range(Cells(colRowConstant, firstRorC), Cells(colRowConstant, lastRorC))
            ElseIf (columnaOfila = "fila") Then
                Set rng = Range(Cells(firstRorC, colRowConstant), Cells(lastRorC, colRowConstant))
            End If
        End If
        With rng
            If (exactCoin = True) Then
                Set c = .Find(palabra, LookIn:=xlValues, MatchCase:=True, LookAt:=xlWhole) 'MatchCase:= true para que busque el caso exacto
            Else
                Set c = .Find(palabra, LookIn:=xlValues)
            End If
            'Significa que la fila ya está llenada y se pueden seguir colocando valores ahí
            If c Is Nothing Then
                'significa que no encontró la palabra
                'MsgBox ("No se encontró la " & columnaOfila & " con título: " & palabra)
                busca_EnRangoV2 = -1
            Else
                'Buscar dentro de las varias coincidencias (SI Aplica) los distintos valores
                If Not IsMissing(scndCoin) Then
                    firstaddress = c.Address 'Para no repetir infinitamente la búsqueda (ue encuentre dónde empezó)
                    Do
                        If (columnaOfila = "columna") Then
                            busca_EnRangoV2 = c.Column
                        ElseIf (columnaOfila = "fila") Then
                            busca_EnRangoV2 = c.row
                        End If
                        'intermcambiar columna o row para optimizar búsqueda
                        If (columnaOfila = "columna") Then
                            tmp = Cells(rowColScndCoin, busca_EnRangoV2).value
                        ElseIf (columnaOfila = "fila") Then
                            tmp = Cells(busca_EnRangoV2, rowColScndCoin).value
                        End If
                        'Ver si el valor es el esperado
                        If (tmp = scndCoin) Then
                            Exit Function
                        End If
                        'Si no es el valor esperado, podemos seguir buscando hasta el final
                        'Sea sub, sea variable para después llenar, esto permite colocar en el row encontrado
                        'el dato en su columna correspondiente
                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstaddress '*esta última condición es pués cuando termina de encontrar todos los mathc, vuelve a empezar del primero
                    'Si llega acá significa que no encontró nada
                    busca_EnRangoV2 = 0
                Else
                    If (columnaOfila = "columna") Then
                        busca_EnRangoV2 = c.Column
                    ElseIf (columnaOfila = "fila") Then
                        busca_EnRangoV2 = c.row
                    End If
                End If
            End If
        End With
End Function
