Public Class ArenasMap
    'usamos un diccionario. La key es el nombre del campo en la tabla importada. El value es el campo de destino
    Public ReadOnly Property FieldMap As New Dictionary(Of String, String)

    Public Sub New()
        FieldMap.Add("Temperatura", "Temperatura")
        FieldMap.Add("Compact_Baja", "Compactibilidad")
        FieldMap.Add("Compact_Alta", "Compactibilidad")
        FieldMap.Add("Humedad", "HumedadPCT")
        FieldMap.Add("AcrillaActivada", "ArcillaActiva")
        FieldMap.Add("Friabilidad", "Friabilidad")
        FieldMap.Add("ResistenciaEnVerde", "RCV")
        FieldMap.Add("PesoProbetaCompresion", "PesoMuestra")
        FieldMap.Add("LOI", "LOI")
        FieldMap.Add("MaterialVolatil", "MateriaVol")
        FieldMap.Add("TiempoDeMezclado", "TiempoDeMezclado")
        FieldMap.Add("ResitenciaVerdeT", "ResistenciaVerdeT")
        FieldMap.Add("RCS", "ResistenciaSeco")
    End Sub
End Class
