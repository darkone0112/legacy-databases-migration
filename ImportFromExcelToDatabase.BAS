Option Compare Database

'------------------------------------------------------------
' Importa_Datos_1ª_Entrega
'
'------------------------------------------------------------
Function Importa_Datos_1ª_Entrega()
On Error GoTo Importa_Datos_1ª_Entrega_Err

    DoCmd.TransferSpreadsheet acImport, 10, "Importa Datos ONTIGOLA - Completed Deliveries", "C:\Users\velar\Desktop\Distribución CEVA Centro\Importa Salidas ONTIGOLA - Completed Deliveries.xlsx", True, ""
    DoCmd.TransferSpreadsheet acImport, 10, "Importa Datos ONTIGOLA - Envíos", "C:\Users\velar\Desktop\Distribución CEVA Centro\Importa Salidas ONTIGOLA - Envíos.xlsx", True, ""
    DoCmd.OpenQuery "Elimina Importa Datos ONTÍGOLA - Completed Deliveries", acViewNormal, acEdit
    DoCmd.OpenQuery "Elimina Importa Datos ONTÍGOLA - Envíos", acViewNormal, acEdit
    DoCmd.TransferSpreadsheet acImport, 10, "Importa Datos SUBIRATS - Completed Deliveries", "C:\Users\velar\Desktop\Distribución CEVA Centro\Importa Salidas SUBIRATS - Completed Deliveries.xlsx", True, ""
    DoCmd.TransferSpreadsheet acImport, 10, "Importa Datos SUBIRATS - Envíos", "C:\Users\velar\Desktop\Distribución CEVA Centro\Importa Salidas SUBIRATS - Envíos", True, ""
    DoCmd.OpenQuery "Elimina Importa Datos SUBIRATS - Completed Deliveries", acViewNormal, acEdit
    DoCmd.OpenQuery "Elimina Importa Datos SUBIRATS - Envíos", acViewNormal, acEdit


Importa_Datos_1ª_Entrega_Exit:
    Exit Function

Importa_Datos_1ª_Entrega_Err:
    MsgBox Error$
    Resume Importa_Datos_1ª_Entrega_Exit

End Function