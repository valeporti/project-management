Private Sub inicio_Click()
    Dim act_reporte As Boolean, act_equipos As Boolean, act_equipos_sem As Boolean, act_periodos As Boolean, _
        act_tablas As Boolean, act_todos As Boolean, act_rec As Boolean
    Dim a_ejecutar As Integer
    
    act_reporte = reporte.Value
    act_equipos = equipos.Value
    act_equipos_sem = equipos_sem.Value
    act_periodos = periodos.Value
    act_rec = rec.Value
    act_tablas = tablas.Value
    act_todos = todos.Value
    determinar = 0
    
    If (act_reporte = True) Then
        a_ejecutar = 1
    ElseIf (act_equipos = True) Then
        a_ejecutar = 2
    ElseIf (act_equipos_sem = True) Then
        a_ejecutar = 3
    ElseIf (act_periodos = True) Then
        a_ejecutar = 4
    ElseIf (act_tablas = True) Then
        a_ejecutar = 5
    ElseIf (act_todos = True) Then
        a_ejecutar = 6
    ElseIf (act_rec = True) Then
        a_ejecutar = 7
    End If
       
    ckbx_elegir_proceso.Hide
    'MsgBox (a_ejecutar)
    Call Llenado_Luisa(a_ejecutar)

    End
End Sub

Private Sub salir_Click()
    ckbx_elegir_proceso.Hide
End Sub