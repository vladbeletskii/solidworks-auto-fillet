Dim swApp As Object
Dim swModel As Object
Dim swFeatureManager As Object
Dim swSelectionMgr As Object
Dim swEdge As Object
Dim swFeature As Object
Dim radius As Double
Dim i As Integer

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "Откройте деталь перед запуском макроса!", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' Запрашиваем радиус скругления
    radius = InputBox("Введите радиус скругления (мм):", "АвтоСкругление", 1)
    
    ' Получаем менеджер фич и селектор
    Set swFeatureManager = swModel.FeatureManager
    Set swSelectionMgr = swModel.SelectionManager
    
    ' Выбираем все рёбра
    Dim swEdges As Variant
    swEdges = swModel.GetEntitiesByType(swSelEDGES)
    
    If Not IsEmpty(swEdges) Then
        For i = 0 To UBound(swEdges)
            Set swEdge = swEdges(i)
            swEdge.Select4 True, Nothing
        Next i
        
        ' Применяем скругление
        Set swFeature = swFeatureManager.FeatureFillet3(2, radius, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        
        If Not swFeature Is Nothing Then
            MsgBox "Скругления добавлены!", vbInformation, "Готово"
        Else
            MsgBox "Ошибка: не удалось создать скругления.", vbExclamation, "Ошибка"
        End If
    Else
        MsgBox "Рёбра не найдены!", vbExclamation, "Ошибка"
    End If
End Sub
