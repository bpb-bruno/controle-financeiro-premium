Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range
    Set rng = Me.Range("J9")

    ' Verifica se a alteração ocorreu na célula monitorada
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False
        
        ' Se o usuário deletar o valor, o sistema redefine para 0
        If Target.Value = "" Then
            Target.Value = 0
        End If
        
        Application.EnableEvents = True
    End If
End Sub