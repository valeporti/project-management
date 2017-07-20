
'****/
'* pequeña rutina que borra el contenido del layout
'****\
Public Sub borrar_entorno(ttl_row As Integer, colToCeckLastRow As Integer)
    Dim last_row As Integer, last_col As Integer
    Call desactivar_filtro
    last_row = Cells(Rows.COUNT, colToCeckLastRow).End(xlUp).row
    If (last_row > ttl_row) Then
        last_col = Cells(ttl_row, Columns.COUNT).End(xlToLeft).Column
        'hoja.Rows(ttl_row + 1 & ":" & last_row).Select
        'Selection.Delete Shift:=xlUp 'primero borrar la fila existente
        Range(Cells(ttl_row + 1, 1), Cells(last_row, last_col)).ClearContents
    
    End If
End Sub

'****/
'* Enlistar en un array los archivos existentes en la carpeta seleccionada
'****\
Public Sub lista_archivos(lista() As String, dir_arch As Variant, wb As Workbook)
    
    Dim l As Integer, ruta As String, principalName As String, exito As Boolean
    ruta = wb.Path
    principalName = wb.Name
    
    l = 0
    dir_arch = Dir(ruta & "\")
    Do While dir_arch <> ""
        If (principalName <> dir_arch) Then
            lista(l) = ruta & "\" & dir_arch
            l = l + 1
        End If
        dir_arch = Dir()
    Loop
    
End Sub


'****/
'* Unicamente comprueba que haya un filtro y en caso positivo lo quita
'****\
Public Sub desactivar_filtro()

    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If

End Sub


Sub arrMatches(arr() As String, pattern As String, fromStr As String)

    Dim regEx As New RegExp, numMatches As Integer
        
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    numMatches = regEx.Execute(fromStr).COUNT
    
    For i = 0 To numMatches - 1
        ReDim Preserve arr(i)
        arr(i) = regEx.Execute(fromStr).Item(i)
    Next
    
End Sub

Public Sub juntarRango(ttl_row As Integer, fst As Integer, ttl As String, txtBef As String, txtNow As String)
    
    Cells(ttl_row - 1, fst).value = ttl
    Range(Cells(ttl_row - 1, fst), Cells(ttl_row - 1, fst + 1)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Cells(ttl_row, fst).value = txtBef
    Cells(ttl_row, fst + 1).value = txtNow
End Sub

Public Sub buscaYcoloca(rng As Range, tit As String, dato, row As Double)
    Dim col As Integer
    With rng
        Set c = .Find(palabra, LookIn:=xlValues, MatchCase:=True) 'MatchCase:= true para que busque el caso exacto
        'Significa que la fila ya está llenada y se pueden seguir colocando valores ahí
        If c Is Nothing Then
            'significa que no encontró la palabra
            MsgBox "no se encontro nada en busca y coloca ?????!!!"
        Else
            col = c.Column
        End If
    End With
    Cells(row, col).value = dato
End Sub


