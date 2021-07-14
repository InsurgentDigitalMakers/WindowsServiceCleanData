Imports System.IO
Imports System.Reflection

Public Class DataCleaner
    Public Sub CleanData(ByVal file As String)
        Dim params As DataParameter = getDataParameters(file)
        Dim data As New DataAccess
        Dim tipo As Type = Type.GetType(getAssemblyQualifiedName(params.ClaseMapeo))
        Dim claseMap = tipo.GetConstructor(Type.EmptyTypes).Invoke(Nothing)

        MapeaCampos(tipo, claseMap, params)
        data.CopyData(params.SPCopy)
    End Sub

    Private Function getAssemblyQualifiedName(ByVal clase As String)
        Dim dataParam As New DataParameter
        Dim ns = dataParam.GetType().Namespace
        Dim assembly = GetType(DataParameter).Assembly.GetName.FullName

        'Dim objectType As Type = (From asm In AppDomain.CurrentDomain.GetAssemblies()
        '                          From type In asm.GetTypes()
        '                          Where type.IsClass And type.Name = clase
        '                          Select type).Single()
        'Dim obj = Activator.CreateInstance(objectType)


        'Dim appAssemblies = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(Function(x) x.GetTypes.FirstOrDefault(Function(y) y.Name = clase) IsNot Nothing)
        'Dim nameAssembly = appAssemblies.FullName

        Return ns & "." & clase & ", " & assembly
    End Function


    Private Sub MapeaCampos(Of T)(ByVal tipo As Type, claseMap As T, params As DataParameter)
        Dim dataRead As New DataAccess
        Dim diccionario = getDictionary(tipo, claseMap)

        Dim dt = dataRead.GetRowsNews(params.TablaOrigen, params.TablaDestino, params.FieldKey)
        Dim keyField = params.FieldKey
        Dim tablaDestino = params.TablaDestino
        Dim tablaOrigen = params.TablaOrigen
        For Each row As DataRow In dt.Rows
            For Each item As KeyValuePair(Of String, String) In diccionario
                Dim fieldDestino = item.Value
                Dim fieldOrigen = item.Key
                Dim Key = row(0)
                'obtenemos tipo de dato 
                Dim type = dataRead.GetTypeField(tablaDestino, fieldDestino)
                'obtenemos el valor en la tabla origen
                Dim value = dataRead.getValue(tablaOrigen, fieldOrigen, keyField, Key)
                Select Case type
                    Case "System.Decimal"
                        If Len(value) > 0 Then
                            If validaDecimal(value) = False Then
                                value = "0"
                                dataRead.SetCleanValue(value, tablaOrigen, keyField, Key, fieldOrigen)
                            End If
                        End If
                    Case "System.Int32"
                        If Len(value) > 0 Then
                            If ValidaInteger(value) = False Then

                                Dim number = TransformaInteger(value)
                                dataRead.SetCleanValue(number, tablaOrigen, keyField, Key, fieldOrigen)
                            End If
                        End If
                End Select

            Next
        Next

    End Sub


    Private Function ValidaInteger(ByVal valor As String) As Boolean
        If Not (IsNumeric(valor)) Then

            Return False
        End If

        If valor.Contains(".") Then
            Return False
        End If

        Return True
    End Function

    Private Function TransformaInteger(ByVal valor As String) As String
        If Not (IsNumeric(valor)) Then

            Return "0"
        End If

        If valor.Contains(".") Then
            Dim array = valor.Split(".")
            Dim number = array(0)
            Return number
        End If

        Return valor
    End Function
    ''' <summary>
    ''' Esta Función valida un campo decimal. Básicamente validamos que sea numérico
    ''' </summary>
    Private Function validaDecimal(ByVal valor As String) As Boolean
        If IsNumeric(valor) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Esta Función permite obtener el diccionario de mapeo.
    ''' Para esto usamos la clase genérica leída desde el archivo de texto.
    ''' </summary>
    Private Function getDictionary(Of T)(ByVal tipo As Type, claseMap As T) As Dictionary(Of String, String)
        Dim propiedades As PropertyInfo() = tipo.GetProperties
        Dim diccionario As Dictionary(Of String, String)
        For Each propiedad As PropertyInfo In propiedades

            diccionario = CallByName(claseMap, propiedad.Name, CallType.Get)
        Next
        Return diccionario
    End Function

    ''' <summary>
    ''' Esta Función permite obtener los parametros para trabajar.
    ''' Se leen desde un arhivo de texto, incluyen tabla de origen, tabla de destino, clase de mapeo y nombre del campo llave
    ''' </summary>
    Private Function getDataParameters(ByVal file As String) As DataParameter

        Dim data As New DataParameter
        Dim propiedades As PropertyInfo() = GetType(DataParameter).GetProperties

        For Each propiedad As PropertyInfo In propiedades
            Dim param = getParam(file, propiedad.Name)
            CallByName(data, propiedad.Name, CallType.Set, param)
        Next
        Return data
    End Function

    Public Function getParam(ByVal file As String, NombreParam As String) As String

        Dim objReader As New StreamReader(file)
        Dim sLine As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not (sLine Is Nothing) Then
                Dim vectorData = Split(sLine, ":")
                Dim key = vectorData(0)
                Dim value = vectorData(1)
                If key = NombreParam Then
                    Return value
                End If
            End If
        Loop Until sLine Is Nothing
        objReader.Close()
        Return ""
    End Function
End Class

