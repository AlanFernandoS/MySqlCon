Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.ComponentModel
Imports System.Reflection
Imports System.Text.RegularExpressions

''' <summary>
''' Creted by Alan Fernando Santacruz Rodríguez 2020
''' For the community with love and passion
''' </summary>
Public Class AdmSQL
    Protected ConnectionString As String = ""

    Sub New(ByVal MyIp As String, ByVal BaseD As String, ByVal Usu As String, ByVal Pwd As String)
        ConnectionString = "Data Source = " + MyIp + ";Initial Catalog=" + BaseD + ";Persist Security Info=True;User ID=" + Usu + ";Password=" + Pwd
    End Sub

    Sub New(ByVal InOneLineSpScp As String, ByRef CompError As Boolean)
        Dim MyData() As String = InOneLineSpScp.Split(vbNewLine)
        If MyData.Count = 4 Then
            ConnectionString = "Data Source = " + MyData(0) + ";Initial Catalog=" + MyData(1) + ";Persist Security Info=True;User ID=" + MyData(2) + ";Password=" + MyData(3)
        Else
            CompError = True
        End If
    End Sub

    ''' <summary>
    ''' Permite retornar el string para la conexión creado en base al constructor de la clase.
    ''' </summary>
    ''' <returns></returns>
    Function RetornaElConnectionString()
        Return ConnectionString
    End Function

    ''' <summary>
    ''' Permite de una manera tener las comparaciones necesarias para relizar el cambio o actualizacion
    ''' regresa una lista de manera que quedan las columnas igualadas de esta manera Columna(n)='Lis(n)'
    ''' </summary>
    ''' <param name="Columna"></param>
    ''' <param name="Lis2"></param>
    ''' <returns></returns>
    Function RetornaIgualdades(ByVal Columna As List(Of String), ByVal Lis2 As List(Of String)) As List(Of String)
        Dim Devuelta As List(Of String) = New List(Of String)
        If Columna.Count = Lis2.Count Then
            For index As Integer = 0 To Columna.Count - 1
                Devuelta.Add(Columna(index) + " = '" + Lis2(index) + "'")
            Next
        Else
            MsgBox("Error las listas no son del mismo tamaño")
        End If
        Return Devuelta
    End Function


    Function RetornaIgualdadesV2(ByVal Columna As List(Of String), ByVal Lis2 As List(Of String), ByRef IndicadorDeError As Boolean) As List(Of String)
        IndicadorDeError = False
        Dim Devuelta As List(Of String) = New List(Of String)
        If Columna.Count = Lis2.Count Then
            For index As Integer = 0 To Columna.Count - 1
                Devuelta.Add(Columna(index) + " = '" + Lis2(index) + "'")
            Next
        Else
            IndicadorDeError = True
            'MsgBox("Error las listas no son del mismo tamaño")
        End If
        Return Devuelta
    End Function
    ''' <summary>
    ''' Columa = 'DatoToCompare' Nota: No Añade un espacio en blanco al final
    ''' </summary>
    ''' <param name="Columna"></param>
    ''' <param name="DatoToCompare"></param>
    ''' <returns></returns>
    Function RetornaIgualdadesV2(ByVal Columna As String, ByVal DatoToCompare As String) As String
        Return Columna + " = " + InsertComillas(DatoToCompare)
    End Function

    Function RetornaIgualdadesSinComillas(ByVal Columna As List(Of String), ByVal Lis2 As List(Of String))
        Dim Devuelta As List(Of String) = New List(Of String)
        If Columna.Count = Lis2.Count Then
            For index As Integer = 0 To Columna.Count - 1
                Devuelta.Add(Columna(index) + " = " + Lis2(index) + " ")
            Next
        Else
            MsgBox("Error las listas no son del mismo tamaño")
        End If
        Return Devuelta
    End Function

    ''' <summary>
    ''' Retorna el string de consulta basado en
    ''' Select C1,C2,...,Cn From TablaSQL Where Cond1 Or/And Cond2 ...Or/And CondN
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="lColum"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <param name="Condicionante"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal lColum As List(Of String), ByVal CondBusqueda As List(Of String), ByVal Condicionante As List(Of String))
        Dim StrConsulta = "Select "
        Dim MaxIndex = lColum.Count - 1
        Try
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + lColum(index)
                Else
                    StrConsulta = StrConsulta + lColum(index) + ", "
                End If
            Next
            StrConsulta = StrConsulta + " From " + TablaSQL + " Where "

            MaxIndex = CondBusqueda.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + CondBusqueda(index)
                Else
                    StrConsulta = StrConsulta + CondBusqueda(index) + Condicionante(index) + " "
                End If
            Next
        Catch er As System.Exception
            MsgBox(er.Message,, "Módulo de armado de consulta condicionada")
            Return ""
        End Try
        Return StrConsulta.Trim
    End Function

    ''' <summary>
    ''' Permite estructurar una consulta condicionada
    ''' Select colum1,colum2,...,columnN form TablaSQL where Con1 and Cond2 ... and CondN
    ''' </summary>
    ''' <param name="lColum"></param>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal lColum As List(Of String), ByVal CondBusqueda As List(Of String))
        Dim StrConsulta = "Select "
        Dim MaxIndex = lColum.Count - 1
        Try
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + lColum(index)
                Else
                    StrConsulta = StrConsulta + lColum(index) + ", "
                End If
            Next
            StrConsulta = StrConsulta + " From " + TablaSQL + " Where "

            MaxIndex = CondBusqueda.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + CondBusqueda(index)
                Else
                    StrConsulta = StrConsulta + CondBusqueda(index) + "And "
                End If
            Next
        Catch er As System.Exception
            MsgBox(er.Message,, "Módulo de armado de consulta condicionada")
        End Try
        Return StrConsulta.Trim
    End Function

    ''' <summary>
    ''' Estructura una busqueda de manera que select * from TablaSQL where Cond1 and Cond2 ... and ConN
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal CondBusqueda As List(Of String))
        Dim StrConsulta = "Select * From " + TablaSQL + " Where "
        Dim MaxIndex = CondBusqueda.Count - 1
        Try
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + CondBusqueda(index)
                Else
                    StrConsulta = StrConsulta + CondBusqueda(index) + " And "
                End If
            Next
        Catch er As System.Exception
            MsgBox("Ha habido un error en la busqueda condicionada" + vbCrLf + er.Message + vbCrLf + StrConsulta,, "Error en el módulo armaSQL")
        End Try
        Return StrConsulta.Trim
    End Function

    ''' <summary>
    ''' Retorna Select * from TablaSQL
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <returns></returns>
    Function ArmaConSQL(ByVal TablaSQL As String)
        Return "Select * From " + TablaSQL
    End Function

    ''' <summary>
    ''' Simplificada para el uso de simplemente el String de condicion
    ''' Select * from TablaSQL where CondBusqueda
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal CondBusqueda As String)
        Dim StrConsulta = "Select * From " + TablaSQL + " Where "
        Try
            StrConsulta = StrConsulta + CondBusqueda
        Catch er As System.Exception
            MsgBox("Ha habido un error en ArmaSql" + vbCrLf + er.Message,, "Error en el módulo armaSQL")
        End Try
        Return StrConsulta.Trim
    End Function
    Function ConcatenaDataSortedList2Update(ByVal Data As SortedList(Of String, String)) As String
        Dim MaxIndexRegister = Data.Keys.Count - 1
        Dim ConcatenacionData = ""
        'Dim BaseOfChange As String = " @Columna='@Data' "
        For IndexRegister As Integer = 0 To MaxIndexRegister
            If IndexRegister = MaxIndexRegister Then
                ConcatenacionData += " " + Data.Keys(IndexRegister) + "= '" + Data.Values(IndexRegister) + "' "
            Else
                ConcatenacionData += " " + Data.Keys(IndexRegister) + "= '" + Data.Values(IndexRegister) + "', "
            End If
        Next
        Return ConcatenacionData
    End Function

    ''' <summary>
    ''' Dim ComandoSql = "Update @TableSql  Set @NewValues Where @Conditions"
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="Condiciones"></param>
    ''' <param name="NewData"></param>
    ''' <returns></returns>
    Function UpdateOnSQL(ByVal TablaSQL As String, ByVal Condiciones As List(Of String), ByVal NewData As SortedList(Of String, String)) As String
        Dim ComandoSql = "Update @TableSql  Set @NewValues Where @Conditions"
        Dim ErrorSql As String = ""
        ComandoSql = ComandoSql.Replace("@TableSql", TablaSQL)

        Dim SqlCondition As String = JoinTheConditionsClauses(Condiciones)
        ComandoSql = ComandoSql.Replace("@Conditions", SqlCondition)

        Dim ConcatenacionData = ConcatenaDataSortedList2Update(NewData)
        ComandoSql = ComandoSql.Replace("@NewValues", ConcatenacionData)

        ComandoSql = ComandoSql.Replace("'@NULL'", "NULL")

        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ComandoSql, con)
                Dim Respuesta As Integer = cmd.ExecuteNonQuery()
                If Respuesta = 0 Then
                    ErrorSql = "Error: No se ha logrado registrar la informacion en la base de datos: " + ComandoSql
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            ErrorSql = "Error:" + vbNewLine + ex.Message + vbNewLine + "ComandoSql: " + ComandoSql
        End Try
        Return ErrorSql
    End Function

    '-------------Funciones de UPDATE de la informacion
    'UPDATE table_name
    'Set column1 = value1, column2 = value2, ...
    'WHERE condition;
    ''' <summary>
    ''' Realizara una instrucción basada en
    ''' Update TablaSQL set C1 ='CambiosSQL1', C2 ='CambiosSQL2',..., CN ='CambiosSQLN' Where Con1 And Con2 And Con3.
    ''' Devuelve un true en caso de ser exitosa la actualización
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CambiosSQL"></param>
    ''' <param name="Condiciones"></param>
    Function UpdateOnSQL(ByVal TablaSQL As String, ByVal CambiosSQL As List(Of String), ByVal Condiciones As List(Of String)) As List(Of String)
        Dim Comando As String = "Update " + TablaSQL + " set "
        Dim ListaDeErrores As New List(Of String)
        Try
            Dim MaxIndex = CambiosSQL.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    Comando = Comando + CambiosSQL(index)
                Else
                    Comando = Comando + CambiosSQL(index) + ", "
                End If
            Next
            Comando = Comando + " where "
            MaxIndex = Condiciones.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    Comando = Comando + Condiciones(index)
                Else
                    Comando = Comando + Condiciones(index) + " and "
                End If
            Next
            'Until here are the preparations------------------------
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim respuesta = cmd.ExecuteNonQuery()
                If respuesta = 0 Then
                    'Exito = False
                    ListaDeErrores.Add("No se ha modificado ningun registro por alguna extraña razón, nota del programador." + vbNewLine + Comando)
                    'MsgBox("")
                End If
                con.Close()
            End Using
        Catch er As System.Exception
            ListaDeErrores.Add(er.Message)
            'Exito = False
        End Try
        Return ListaDeErrores
    End Function
    Public Async Function UpdateOnSQLAsync(ByVal TablaSQL As String, ByVal CambiosSQL As List(Of String), ByVal Condiciones As List(Of String)) As Task(Of List(Of String))
        Dim Comando As String = "Update " + TablaSQL + " set "
        Dim ListaDeErrores As New List(Of String)
        Try
            Dim ComandoA = Await JoinNewDataInUpdateClauseAsync(CambiosSQL)
            Dim ComandoB = Await JoinTheConditionsClausesAsync(Condiciones)

            Comando = Comando + ComandoA
            Comando = Comando + ComandoB + ";"
            'Until here are the preparations------------------------
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim respuesta = cmd.ExecuteNonQuery()
                If respuesta = 0 Then
                    'Exito = False
                    ListaDeErrores.Add("No se ha modificado ningun registro por alguna extraña razón, nota del programador." + vbNewLine + Comando)
                    'MsgBox("")
                End If
                con.Close()
            End Using
        Catch er As System.Exception
            ListaDeErrores.Add(er.Message)
            'Exito = False
        End Try
        Return ListaDeErrores
    End Function
    Public Async Function JoinNewDataInUpdateClauseAsync(ByVal CambiosSQL As List(Of String)) As Task(Of String)
        Dim Comando As String = ""
        Dim MaxIndex = CambiosSQL.Count - 1
        For index As Integer = 0 To MaxIndex
            If index = MaxIndex Then
                Comando = Comando + CambiosSQL(index)
            Else
                Comando = Comando + CambiosSQL(index) + ", "
            End If
        Next
        Return Comando
    End Function
    Public Async Function JoinTheConditionsClausesAsync(ByVal Condiciones As List(Of String)) As Task(Of String)
        Dim Comando As String = ""
        Dim MaxIndex = Condiciones.Count - 1
        For index As Integer = 0 To MaxIndex
            If index = MaxIndex Then
                Comando = Comando + Condiciones(index)
            Else
                Comando = Comando + Condiciones(index) + " and "
            End If
        Next
        Return Comando
    End Function
    ''' <summary>
    ''' Cond1 and Cond2 ...and CondN
    ''' </summary>
    ''' <param name="Condiciones"></param>
    ''' <returns></returns>
    Public Function JoinTheConditionsClauses(ByVal Condiciones As List(Of String)) As String
        Dim Comando As String = ""
        Dim MaxIndex = Condiciones.Count - 1
        For index As Integer = 0 To MaxIndex
            If index = MaxIndex Then
                Comando = Comando + Condiciones(index)
            Else
                Comando = Comando + Condiciones(index) + " and "
            End If
        Next
        Return Comando
    End Function
    Public Function JoinNewDataInUpdateClause(ByVal CambiosSQL As List(Of String)) As String
        Dim Comando As String = ""
        Dim MaxIndex = CambiosSQL.Count - 1
        For index As Integer = 0 To MaxIndex
            If index = MaxIndex Then
                Comando = Comando + CambiosSQL(index)
            Else
                Comando = Comando + CambiosSQL(index) + ", "
            End If
        Next
        Return Comando
    End Function
    ''' <summary>
    ''' Examina en la Tabla SQL en base a la consulta y para acelerar el proceso solo se toma una columna, si esta existe existe el resto
    ''' para ello es necesario recuperar el contenido de la columna en base a un tipo propuesto.
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="Consulta"></param>
    ''' <param name="MyColumn"></param>
    ''' <param name="MyType"></param>
    ''' <returns></returns>
    Function ExisteLaconsulta(ByVal TablaSQL As String, ByVal Consulta As String, ByVal MyColumn As String, ByVal MyType As Object)
        Dim existe As Boolean = False
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                cmd.Parameters.AddWithValue("@" + MyColumn, MyType)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    existe = True
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Return existe
    End Function
    Function ExisteLaconsultaV2(ByVal TablaSQL As String, ByVal Consulta As String, ByRef AreSomeErrorInside As String) As Boolean
        AreSomeErrorInside = ""
        Dim existe As Boolean = False
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    existe = True
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            AreSomeErrorInside = ex.Message
        End Try
        Return existe
    End Function
    Function ExisteLaconsultaV2(ByVal Consulta As String, ByRef AreSomeErrorInside As String) As Boolean
        AreSomeErrorInside = ""
        Dim existe As Boolean = False
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    existe = True
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            AreSomeErrorInside = ex.Message
        End Try
        Return existe
    End Function
    ''' <summary>
    ''' Regresa un string que dice
    ''' Existe si la clave fue encontrada
    ''' No existe, si la clave no fue encontrada
    ''' Error: @Error, si hubo algun error por parte de la instruccion, el @Error indica el error que sucedio
    ''' </summary>
    ''' <param name="Tabla"></param>
    ''' <param name="Columna"></param>
    ''' <param name="Dato"></param>
    ''' <returns> Existe, No Existe, Error: @Error </returns>
    Function ExisteLaClave(ByVal Tabla As String, ByVal Columna As String, ByVal Dato As String) As String
        Dim ConsultaClave As String = "Select * From @Tabla where @Columna='@Dato'"
        ConsultaClave = ConsultaClave.Replace("@Tabla", Tabla).Replace("@Columna", Columna).Replace("@Dato", Dato)
        Dim ResultadoDeLaConsulta As String = "No existe"
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaClave, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    ResultadoDeLaConsulta = "Existe"
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            ResultadoDeLaConsulta = "Error : " + ex.Message
        End Try
        Return ResultadoDeLaConsulta
    End Function
    ''' <summary>
    ''' Retorna un string con 'Existe' o 'No existe', o en caso de error, devuelve del error
    ''' </summary>
    ''' <param name="Tabla"></param>
    ''' <param name="Columna1"></param>
    ''' <param name="Dato1"></param>
    ''' <param name="Columna2"></param>
    ''' <param name="Dato2"></param>
    ''' <returns></returns>
    Function ExistsTwoColumnKey(ByVal Tabla As String, ByVal Columna1 As String, ByVal Dato1 As String, ByVal Columna2 As String, ByVal Dato2 As String) As String
        Dim ConsultaClave As String = "Select TOP 1 PERCENT @Col1 From @Tabla where @Col1='@Dato1' and @Colum2='@Dato2'"
        ConsultaClave = ConsultaClave.Replace("@Tabla", Tabla).Replace("@Col1", Columna1).Replace("@Dato1", Dato1)
        ConsultaClave = ConsultaClave.Replace("@Colum2", Columna2).Replace("@Dato2", Dato2)
        Dim ResultadoDeLaConsulta As String = "No existe"
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaClave, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    ResultadoDeLaConsulta = "Existe"
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            ResultadoDeLaConsulta = "Error : " + ex.Message
        End Try
        Return ResultadoDeLaConsulta
    End Function

    ''' <summary>
    ''' Retorna un string con Existe o No existe, o en caso de error, devuelve del error
    ''' "Select TOP 1 PERCENT @Columna From @Tabla where @Columna='@Dato' and @Colum2='@Dato2' and @ExtraCondition"
    ''' </summary>
    ''' <param name="Tabla"></param>
    ''' <param name="Columna1"></param>
    ''' <param name="Dato1"></param>
    ''' <param name="Columna2"></param>
    ''' <param name="Dato2"></param>
    ''' <param name="ExtraCondition"></param>
    ''' <returns></returns>
    Function ExistsTwoColumnKey(ByVal Tabla As String, ByVal Columna1 As String, ByVal Dato1 As String, ByVal Columna2 As String, ByVal Dato2 As String, ByVal ExtraCondition As String) As String
        Dim ConsultaClave As String = "Select TOP 1 PERCENT @Columna From @Tabla where @Columna='@Dato1' and @Colum2='@Dato2' and @ExtraCondition"
        ConsultaClave = ConsultaClave.Replace("@Tabla", Tabla).Replace("@Columna", Columna1).Replace("@Dato1", Dato1)
        ConsultaClave = ConsultaClave.Replace("@Colum2", Columna2).Replace("@Dato2", Dato2)
        ConsultaClave = ConsultaClave.Replace("@ExtraCondition", ExtraCondition)
        Dim ResultadoDeLaConsulta As String = "No existe"
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaClave, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    ResultadoDeLaConsulta = "Existe"
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            ResultadoDeLaConsulta = "Error : " + ex.Message
        End Try
        Return ResultadoDeLaConsulta
    End Function
    ''' <summary>
    ''' Se regresa la informacion de la consulta en una lista con el orden previsto según el orden de la otras dos listas
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <param name="MyColumns"></param>
    ''' <returns></returns>
    Function SqlReader2List(ByVal Consulta As String, ByVal MyColumns As List(Of String))
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then
                    For Each str As String In MyColumns
                        If dr(str) Is DBNull.Value Then
                            listaC.Add("")
                        Else
                            'listaC.Add(dr(str).ToString)
                            listaC.Add(Convert.ToString(dr(str)))
                        End If

                    Next
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", Consulta)
            MsgBox(TemporalConcat, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
        End Try
        Return listaC
    End Function

    'Function SqlReader2ListAllRow(ByVal Consulta As String, ByVal MyErrors As List(Of String))
    '    Dim listaC As List(Of String) = New List(Of String)

    '    Try
    '        Using con As New SqlConnection(ConnectionString)
    '            con.Open()
    '            Dim cmd As New SqlCommand(Consulta, con)
    '            Dim dr As SqlDataReader = cmd.ExecuteReader()

    '            End If
    '            con.Close()
    '        End Using
    '    Catch ex As System.Exception
    '        MsgBox("Se ha detectado un error:  " + vbCrLf + ex.Message + vbCrLf + Consulta, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
    '    End Try
    '    Return listaC
    'End Function
    ''' <summary>
    ''' Se regresa la informacion de la consulta en una lista con el orden previsto según el orden de la otras dos listas
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <param name="MyColumns"></param>
    ''' <returns></returns>
    Function SqlReader2List(ByVal Consulta As String, ByVal MyColumns As List(Of String), ByRef MyErrors As List(Of String)) As List(Of String)
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then
                    For Each str As String In MyColumns
                        If dr(str) Is DBNull.Value Then
                            listaC.Add("")
                        Else
                            'listaC.Add(dr(str).ToString)
                            listaC.Add(Convert.ToString(dr(str)))
                        End If

                    Next
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim FormaError As String = "" +
                "Se ha detectado el error: @Error:" + vbNewLine +
                "Debido a la consulta: @Consulta"
            FormaError = FormaError.Replace("@Error", ex.Message.ToString()).Replace("@Consulta", Consulta)
            MyErrors.Add(FormaError)
        End Try
        Return listaC
    End Function
    ''' <summary>
    ''' Nueva Version de esta funcion, tabla
    ''' "Select " + ConcatenaColumnas(MyColumsInSearch) + " from " + Table + " where " + Condition
    ''' Only Obtain the fist value
    ''' </summary>
    ''' <param name="Table"></param>
    ''' <param name="Condition"></param>
    ''' <param name="MyColumsInSearch"></param>
    ''' <returns></returns>
    Function SqlReader2List(ByVal Table As String, ByVal Condition As String, ByVal MyColumsInSearch As List(Of String))
        Dim listaC As List(Of String) = New List(Of String)
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        Dim Consulta As String = "Select " + ConcatenaColumnas(MyColumsInSearch) + " from " + Table + " where " + Condition
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then
                    For Each str As String In MyColumsInSearch
                        If dr(str) Is DBNull.Value Then
                            listaC.Add("")
                        Else
                            'listaC.Add(dr(str).ToString)
                            listaC.Add(Convert.ToString(dr(str)))
                        End If

                    Next
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", Consulta)
            MsgBox(TemporalConcat, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
        End Try
        Return listaC
    End Function

    Function SqlReader2OneString(ByVal Table As String, ByVal Condition As String, ByVal Column As String)
        Dim listaC As String = ""
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        Dim Consulta As String = "Select " + Column + " from " + Table + " where " + Condition
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then

                    If dr(Column) Is DBNull.Value Then
                        listaC = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        listaC = Convert.ToString(dr(Column))
                    End If
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", Consulta)
            MsgBox(TemporalConcat, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
        End Try
        Return listaC
    End Function
    Function SqlReader2OneString(ByVal Consulta As String, ByVal Columna As String)
        Dim listaC As String = ""
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then

                    If dr(Columna) Is DBNull.Value Then
                        listaC = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        listaC = Convert.ToString(dr(Columna))
                    End If
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", Consulta)
            MsgBox(TemporalConcat, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
        End Try
        Return listaC
    End Function

    Function SqlReader2OneStringV2(ByVal Table As String, ByVal Condition As String, ByVal Column As String, ByRef IndicadorDeError As Boolean) As String
        Dim StringConsultado As String = ""
        IndicadorDeError = False
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        Dim Consulta As String = "Select " + Column + " from " + Table + " where " + Condition
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then

                    If dr(Column) Is DBNull.Value Then
                        StringConsultado = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        StringConsultado = Convert.ToString(dr(Column))
                    End If
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            IndicadorDeError = True
        End Try
        Return StringConsultado
    End Function
    Function SqlReader2OneStringV2(ByVal Table As String, ByVal Condition As String, ByVal Column As String, ByRef ListaDeErrores As List(Of String)) As String
        Dim StringConsultado As String = ""
        'IndicadorDeError = False
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        Dim Consulta As String = "Select " + Column + " from " + Table + " where " + Condition
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then

                    If dr(Column) Is DBNull.Value Then
                        StringConsultado = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        StringConsultado = Convert.ToString(dr(Column))
                    End If
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            ListaDeErrores.Add(ex.Message.ToString)
        End Try
        Return StringConsultado
    End Function
    Function SqlReaderDownOnlyOneColumn(ByVal Table As String, ByVal Column As String, ByRef IndicadorDeError As Boolean, Optional Condition As String = "", Optional ClauseOrder As String = "") As List(Of String)
        Dim StringConsultado As New List(Of String)
        IndicadorDeError = False
        Dim Consulta As String
        If Condition.Length = 0 Then
            Consulta = "Select " + Column + " from " + Table
        Else
            Consulta = "Select " + Column + " from " + Table + " where " + Condition
        End If
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        If ClauseOrder.Length > 0 Then
            Consulta = Consulta + " Order By " + ClauseOrder
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If dr(Column) Is DBNull.Value Then
                        StringConsultado.Add("")
                    Else
                        'listaC.Add(dr(str).ToString)
                        StringConsultado.Add(Convert.ToString(dr(Column)))
                    End If
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            IndicadorDeError = True
        End Try
        Return StringConsultado
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Table">Tabla a consultar</param>
    ''' <param name="Column">Columna a leer</param>
    ''' <param name="MyListaDeErrores">Son los errores que se generan</param>
    ''' <param name="Condition">Son las clausulas de condiciones</param>
    ''' <param name="ClauseOrder">Aplica la clausula de order</param>
    ''' <param name="IsUnique">Aplica para el primer elemento de la lista</param>
    ''' <returns></returns>
    Function SqlReaderDownOnlyOneColumn(ByVal Table As String, ByVal Column As String, ByRef MyListaDeErrores As List(Of String), Optional Condition As String = "", Optional ClauseOrder As String = "", Optional ByVal IsUnique As Boolean = False) As List(Of String)
        Dim StringConsultado As New List(Of String)
        Dim Consulta As String
        If IsUnique Then
            Consulta = "Select " + Column + " Distinct "
        Else
            Consulta = "Select " + Column
        End If
        If Condition.Length = 0 Then
            Consulta = Consulta + " From " + Table
        Else
            Consulta = Consulta + " From " + Table + " where " + Condition
        End If
        'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
        If ClauseOrder.Length > 0 Then
            Consulta = Consulta + " Order By " + ClauseOrder
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If dr(Column) Is DBNull.Value Then
                        StringConsultado.Add("")
                    Else
                        'listaC.Add(dr(str).ToString)
                        StringConsultado.Add(Convert.ToString(dr(Column)))
                    End If
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            MyListaDeErrores.Add(ex.Message)
        End Try
        Return StringConsultado
    End Function
    'Function SqlReader2ListOfList(ByVal Table As String, ByVal Condition As String, ByVal MyColumsInSearch As List(Of String)) As List(Of List(Of String))
    '    Dim ListaC As List(Of String) = New List(Of String)
    '    Dim ListaR As List(Of List(Of String)) = New List(Of List(Of String))
    '    'Dim MyColums As List(Of String) = ObtenerColumnasDeTabla(Table)
    '    Dim Consulta As String = "Select " + ConcatenaColumnas(MyColumsInSearch) + " from " + Table + " where " + Condition
    '    Try
    '        Using con As New SqlConnection(ConnectionString)
    '            con.Open()
    '            Dim cmd As New SqlCommand(Consulta, con)
    '            Dim dr As SqlDataReader = cmd.ExecuteReader()
    '            While (dr.Read())
    '                ListaR.Add(New List(Of String))
    '                Dim MyActualList As List(Of String) = ReturnLastChildren(ListaR)
    '                For Each str As String In MyColumsInSearch
    '                    If dr(str) Is DBNull.Value Then
    '                        MyActualList.Add("")
    '                    Else
    '                        'listaC.Add(dr(str).ToString)
    '                        MyActualList.Add(Convert.ToString(dr(str)))
    '                    End If

    '                Next
    '            End While
    '            con.Close()
    '        End Using
    '    Catch ex As System.Exception
    '        Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
    '        "Consulta: @Co1"
    '        TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", Consulta)
    '        MsgBox(TemporalConcat, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
    '    End Try
    '    Return ListaR
    'End Function

    ''' <summary>
    ''' Esta función hace uso de
    ''' Insert Inot TablaSQL Values(MyValues1,MyValues2,...MyValuesN)
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="MyValues"></param>
    ''' <returns></returns>
    Function InsertaEnSql(ByVal TablaSQL As String, ByVal MyValues As List(Of String))
        Dim MyError = False
        Try
            Dim Comando As String = "Insert Into " + TablaSQL + " Values ("
            Dim Limite = MyValues.Count - 1
            For Index As Integer = 0 To Limite
                If Index = Limite Then
                    Comando = Comando + MyValues(Index) + ")"
                Else
                    Comando = Comando + MyValues(Index) + ", "
                End If
            Next
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim Respuesta = cmd.ExecuteNonQuery()
                con.Close()
            End Using
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo InsertaEnSQL")
            MyError = True
        End Try
        Return MyError
    End Function

    Function InsertaEnSqlV2(ByVal TablaSQL As String, ByVal MyValues As List(Of String)) As List(Of String)
        Dim MyError As New List(Of String)
        Try
            Dim Comando As String = "Insert Into " + TablaSQL + " Values ("
            Dim Limite = MyValues.Count - 1
            For Index As Integer = 0 To Limite
                If Index = Limite Then
                    Comando = Comando + MyValues(Index) + ")"
                Else
                    Comando = Comando + MyValues(Index) + ", "
                End If
            Next
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim Respuesta = cmd.ExecuteNonQuery()
                con.Close()
            End Using
        Catch ex As System.Exception
            MyError.Add("->InsertaEnSqlV2")
            MyError.Add(ex.Message)
            'MyError = True
        End Try
        Return MyError
    End Function

    Function InsertaSQLSpecificColums(ByVal TablaSQL As String, ByVal MyValues As List(Of String), ByVal ToColumns As List(Of String)) As List(Of String)
        Dim MyError As New List(Of String)
        If MyValues.Count = ToColumns.Count Then
            Try
                Dim Comando As String = "Insert Into " + TablaSQL + " ("
                Dim MyColumnsCount As Integer = ToColumns.Count - 1
                For Index As Integer = 0 To MyColumnsCount
                    If Index = MyColumnsCount Then
                        Comando = Comando + ToColumns(Index) + ")"
                    Else
                        Comando = Comando + ToColumns(Index) + ", "
                    End If
                Next
                Comando = Comando + " Values ("
                Dim Limite = MyValues.Count - 1
                For Index As Integer = 0 To Limite
                    If Index = Limite Then
                        Comando = Comando + MyValues(Index) + ")"
                    Else
                        Comando = Comando + MyValues(Index) + ", "
                    End If
                Next
                Using con As New SqlConnection(ConnectionString)
                    con.Open()
                    Dim cmd As New SqlCommand(Comando, con)
                    Dim Respuesta As Integer = cmd.ExecuteNonQuery()
                    If Respuesta = 0 Then
                        MyError.Add("No se ha efectuar el registro en la base de datos")
                    End If
                    con.Close()
                End Using
            Catch ex As System.Exception
                MyError.Add("->InsertaEnSqlV2")
                MyError.Add(ex.Message)
                'MyError = True
            End Try
        Else
            MyError.Add("Las listas no son del mismo tamaño, nota del programador, check InsertaSQLSpecificColums")
        End If
        Return MyError
    End Function
    ''' <summary>
    ''' Los valores como @NULL, pueden ser remplazados como NULL, sin comillas para eliminar los valores que no tienen validad
    ''' Dentro del comando '@NULL'->NULL
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="Data2Insert"></param>
    ''' <returns></returns>
    Function InsertaSQLSpecificColums(ByVal TablaSQL As String, ByVal Data2Insert As SortedList(Of String, String)) As String
        Dim ErrorSql As String = ""
        Dim ComandoSql As String = "Insert Into @Table (@Columnas) Values (@Data)"
        Dim ConcatenacionColumnas = ""
        Dim MaxIndexRegister = Data2Insert.Keys.Count - 1
        Dim ConcatenacionData = ""
        For IndexRegister As Integer = 0 To MaxIndexRegister
            If IndexRegister = MaxIndexRegister Then
                ConcatenacionColumnas += Data2Insert.Keys(IndexRegister)
                ConcatenacionData += "'" + Data2Insert.Values(IndexRegister) + "'"
            Else
                ConcatenacionColumnas += Data2Insert.Keys(IndexRegister) + ", "
                ConcatenacionData += " '" + Data2Insert.Values(IndexRegister) + "'" + ","
            End If
        Next
        ComandoSql = ComandoSql.Replace("@Table", TablaSQL)
        ComandoSql = ComandoSql.Replace("@Columnas", ConcatenacionColumnas)
        ComandoSql = ComandoSql.Replace("@Data", ConcatenacionData)
        ComandoSql = ComandoSql.Replace("'@NULL'", "NULL")

        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ComandoSql, con)
                Dim Respuesta As Integer = cmd.ExecuteNonQuery()
                If Respuesta = 0 Then
                    ErrorSql = "No se ha logrado registrar la informacion en la base de datos: " + ComandoSql
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            ErrorSql = ex.Message + vbNewLine + "ComandoSql: " + ComandoSql
        End Try
        Return ErrorSql
    End Function
    ''' <summary>
    ''' With Column Of Hour, The data that you put in ToReplaceWithHourServer is replaced by  SYSDATETIME()
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="Data2Insert"></param>
    ''' <returns></returns>
    Function InsertaSQLSpecificColums(ByVal TablaSQL As String, ByVal Data2Insert As SortedList(Of String, String), ByVal ToReplaceWithHourServer As String) As String
        Dim ErrorSql As String = ""
        Dim ComandoSql As String = "Insert Into @Table (@Columnas) Values (@Data)"
        Dim ConcatenacionColumnas = ""
        Dim MaxIndexRegister = Data2Insert.Keys.Count - 1
        Dim ConcatenacionData = ""
        For IndexRegister As Integer = 0 To MaxIndexRegister
            If IndexRegister = MaxIndexRegister Then
                ConcatenacionColumnas += Data2Insert.Keys(IndexRegister)
                ConcatenacionData += "'" + Data2Insert.Values(IndexRegister) + "'"
            Else
                ConcatenacionColumnas += Data2Insert.Keys(IndexRegister) + ", "
                ConcatenacionData += " '" + Data2Insert.Values(IndexRegister) + "'" + ","
            End If
        Next
        ComandoSql = ComandoSql.Replace("@Table", TablaSQL)
        ComandoSql = ComandoSql.Replace("@Columnas", ConcatenacionColumnas)
        ComandoSql = ComandoSql.Replace("@Data", ConcatenacionData)
        ComandoSql = ComandoSql.Replace("'" + ToReplaceWithHourServer + "'", "SYSDATETIME()")
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ComandoSql, con)
                Dim Respuesta As Integer = cmd.ExecuteNonQuery()
                If Respuesta = 0 Then
                    ErrorSql = "No se ha logrado registrar la informacion en la base de datos: " + ComandoSql
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            ErrorSql = ex.Message + vbNewLine + "ComandoSql: " + ComandoSql
        End Try
        Return ErrorSql
    End Function
    Function GetRegistersOfClv(ByVal TablaSql As String, ByVal Columna As String, ByVal Clave As String) As SortedList(Of String, String)
        Dim ListaDeColumnas As New List(Of String)
        Dim DataRegisters As New SortedList(Of String, String)
        ListaDeColumnas = ObtenerColumnasDeTabla(TablaSql)
        Dim Consulta As String = "Select " + ConcatenaColumnas(ListaDeColumnas) + " from " + TablaSql + " where " + Columna + "='" + Clave + "';"
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then
                    For Each str As String In ListaDeColumnas
                        If dr(str) Is DBNull.Value Then
                            DataRegisters.Add(str.Trim, " ")
                        Else
                            'listaC.Add(dr(str).ToString)
                            DataRegisters.Add(str.Trim, Convert.ToString(dr(str)).Trim)
                        End If
                    Next
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", Consulta)
            MsgBox(TemporalConcat, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
        End Try
        Return DataRegisters
    End Function
    ''' <summary>
    ''' Permite pasar de
    ''' Dato-> 'Dato'
    ''' Pero con elementos de una lista
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <returns></returns>
    Function InsertComillas(ByVal MyList As List(Of String))
        Dim Retorno As List(Of String) = New List(Of String)
        For Each str As String In MyList
            Retorno.Add("'" + str + "'")
        Next
        Return Retorno
    End Function

    Function InsertComillas(ByVal MyDato As String) As String
        Return " '" + MyDato + "' "
    End Function

    Function InsertComillas(ByVal MyList As List(Of String), ByVal Pini As Integer, Pfin As Integer)
        Dim Retorno As List(Of String) = New List(Of String)
        Try
            For Index As Integer = 0 To MyList.Count - 1
                If (Pini <= Index And Index <= Pfin) Then
                    Retorno.Add(" '" + MyList(Index) + "' ")
                Else
                    Retorno.Add(MyList(Index))
                End If
            Next
        Catch er As System.Exception
            MsgBox(er.Message)
        End Try
        Return Retorno
    End Function

    ''' <summary>
    ''' Permite ejecutar la consulta definida, siendo la TablaSQL unicamente como dato informativo.
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <returns></returns>
    Function SqlReaderDown2List(ByVal Consulta As String, Optional ByRef ErrorOut As String = "") As List(Of String)
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If Not dr(0) Is DBNull.Value Then
                        listaC.Add(Convert.ToString(dr(0)))
                    End If
                    'listaC.Add(dr(0))
                End While
                'If listaC.Count = 0 Then
                '    listaC.Add("V")
                'End IfBu
                con.Close()
            End Using
        Catch ex As System.Exception
            listaC.Clear()
            ErrorOut = ex.Message
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo SqlReaderDown2List")
        End Try
        Return listaC
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <returns></returns>
    Function SqlReaderDown2List(ByVal Consulta As String, ByRef ListaDeErrores As List(Of String)) As List(Of String)
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If Not dr(0) Is DBNull.Value Then
                        listaC.Add(Convert.ToString(dr(0)))
                    End If
                    'listaC.Add(dr(0))
                End While
                'If listaC.Count = 0 Then
                '    listaC.Add("V")
                'End If
                con.Close()
            End Using
        Catch ex As System.Exception
            listaC.Clear()
            Dim TemporalConcat As String = "Se ha presentado el error: " + ex.Message + vbNewLine + "al ejecutar la consulta: " + Consulta
            ListaDeErrores.Add(TemporalConcat)
        End Try
        Return listaC
    End Function

    ''' <summary>
    ''' Permite ejecutar la consulta definida, siendo la TablaSQL unicamente como dato informativo.
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <returns></returns>
    Function SqlReaderDown2List(ByVal Consulta As String, ByRef IndicadorErr As Boolean)
        IndicadorErr = False
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If Not dr(0) Is DBNull.Value Then
                        listaC.Add(Convert.ToString(dr(0)))
                    End If
                    'listaC.Add(dr(0))
                End While
                'If listaC.Count = 0 Then
                '    listaC.Add("V")
                'End If
                con.Close()
            End Using
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo SqlReaderDown2List")
            IndicadorErr = True
        End Try
        Return listaC
    End Function
    ''' <summary>
    ''' Permite ejecutar la consulta definida, siendo la TablaSQL unicamente como dato informativo.
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <returns></returns>
    Function SqlReaderDown2ListWithDetailError(ByVal Consulta As String, ByRef IndicadorErr As List(Of String))
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If Not dr(0) Is DBNull.Value Then
                        listaC.Add(Convert.ToString(dr(0)))
                    End If
                    'listaC.Add(dr(0))
                End While
                'If listaC.Count = 0 Then
                '    listaC.Add("V")
                'End If
                con.Close()
            End Using
        Catch ex As System.Exception
            ''MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo SqlReaderDown2List")
            IndicadorErr.Add(ex.Message)
        End Try
        Return listaC
    End Function
    ''' <summary>
    ''' Permite ejecutar la consulta definida, siendo la TablaSQL unicamente como dato informativo.
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <returns></returns>
    Public Async Function SqlReaderDown2ListAsync(ByVal Consulta As String) As Task(Of List(Of String))

        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If Not dr(0) Is DBNull.Value Then
                        listaC.Add(Convert.ToString(dr(0)))
                    End If
                    'listaC.Add(dr(0))
                End While
                'If listaC.Count = 0 Then
                '    listaC.Add("V")
                'End If
                con.Close()
            End Using
        Catch ex As System.Exception
            ''MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo SqlReaderDown2List")
            listaC.Clear()
            listaC.Add(ex.Message)
        End Try
        Return listaC
    End Function
    ''' <summary>
    ''' Permite realizar la consulta de esta manera
    ''' Select MyColR from MyTable where MyColB='MyStr'
    ''' </summary>
    ''' <param name="Mytable"></param>
    ''' <param name="MyColR"></param>
    ''' <param name="MyColB"></param>
    ''' <param name="MyStr"></param>
    ''' <returns></returns>
    Function BusquedaDeCambioEnSql(ByVal Mytable As String, ByVal MyColR As String, ByVal MyColB As String, ByVal MyStr As String, Optional ByRef OutputError As String = "")
        Dim MyError As Boolean = True
        Dim Consulta As String = "Select " + MyColR + " from " + Mytable + " where " + MyColB + " ='" + MyStr + "'"
        OutputError = ""
        Dim MyR As List(Of String) = SqlReaderDown2List(Consulta, OutputError)
        If MyR.Count = 0 Then
            Return "" 'e de error
        Else
            Return MyR(0)
        End If
    End Function

    ''' <summary>
    ''' Aquí se cargan los registros que consideramos valiosos, recuerda encomillar aquellos datos que de verdad
    ''' lo necesiten
    ''' </summary>
    ''' <returns></returns>
    Function InsertaEnSql(ByVal MyTable As String, ByVal MyColumns As List(Of String), ByVal MyData As List(Of String))
        Dim MyError = False
        Try
            If MyColumns.Count = MyData.Count Then
                Dim Comando As String = "Insert Into " + MyTable + "("
                Dim Index As Integer = 0
                Dim UltIn As Integer = MyColumns.Count - 1
                For Index = 0 To UltIn
                    If Index = UltIn Then
                        Comando = Comando + MyColumns(Index) + ")"
                    Else
                        Comando = Comando + MyColumns(Index) + ", "
                    End If
                Next
                Comando = Comando + " values("
                For Index = 0 To UltIn
                    If Index = UltIn Then
                        Comando = Comando + MyTable(Index) + ")"
                    Else
                        Comando = Comando + MyTable(Index) + ", "
                    End If
                Next
                Using con As New SqlConnection(ConnectionString)
                    con.Open()
                    Dim cmd As New SqlCommand(Comando, con)
                    Dim Respuesta = cmd.ExecuteNonQuery()
                    con.Close()
                End Using
            Else
                MsgBox("Las listas no tienen el mismo tamaño en InsertaEnSql",, "AdminSQL")
            End If
        Catch er As System.Exception
            MsgBox(er.Message,, "AdminSQL")
        End Try
        Return MyError
    End Function

    ''' <summary>
    ''' Return -1 en caso de un error con SQL
    ''' Return 0 si no esta vacia
    ''' Return 1 si esta vacia
    ''' </summary>
    ''' <param name="MyTable"></param>
    ''' <returns></returns>
    Function IsEmplyTheTable(ByVal MyTable As String)
        Dim IndicaError As Boolean = False
        Dim MyColumns As List(Of String) = ObtenerColumnasDeTabla(MyTable)
        If MyColumns.Count > 0 Then
            Dim Data As List(Of String) = SqlReaderDown2List(ArmaConSQLColumnas(MyTable, MyColumns(0), 1), IndicaError)
            If IndicaError Then
                Return -1
            End If
            If Data.Count > 0 Then
                Return 0
            Else
                Return 1
            End If
        Else
            Return -2
        End If
    End Function

    Function ArmaConSQLColumnas(ByVal MyTable As String, ByVal Columna As String)
        Return "Select " + Columna + " From " + MyTable
    End Function

    Function ArmaConSQLColumnas(ByVal MyTable As String, ByVal Columnas As List(Of String))
        Dim MyConsulta As String = "Select "
        If Columnas.Count > 0 Then
            Dim MaximoI = Columnas.Count - 1
            For Index As Integer = 0 To MaximoI
                If Index = MaximoI Then
                    MyConsulta = MyConsulta + Columnas(Index)
                Else
                    MyConsulta = MyConsulta + Columnas(Index) + ", "
                End If
            Next
            MyConsulta = MyConsulta + " from " + MyTable
        End If
        Return MyConsulta
    End Function

    Function ArmaConSQLColumnas(ByVal MyTable As String, ByVal Columna As String, ByVal LimitAt As Integer)
        Return "Select Top " + LimitAt.ToString + Columna + " From " + MyTable
    End Function

    Function ObtenerColumnasDeTabla(ByVal MyTable As String) As List(Of String)
        Dim ConsultaClv = "Select Column_name from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + MyTable + "'"
        Dim MyColumns As List(Of String) = SqlReaderDown2List(ConsultaClv)
        Return MyColumns
    End Function
    ''' <summary>
    ''' Recuerda que este método conserva todos los metodos y clases derivadas de cada uno de los campos, es decir
    ''' Que si en la base de datos es una fecha, aquí tambien es una fecha, lo que no he determinado es si mantiene las relaciones
    ''' de correspondencia, es decir, si se trata de una Foregin Key si conserva esa relacion directamente, por que si fuera así nos
    ''' ahorraria mucho trabajo de aquí en delante.
    ''' </summary>
    ''' <param name="MyConsulta"></param>
    ''' <returns></returns>
    Function DataParaTable(ByVal MyConsulta As String, Optional ByRef StringErrorOut As String = "") As DataTable
        Dim MyTable As DataTable = New DataTable()
        Dim MyData As SqlDataAdapter
        MyConsulta = MyConsulta.Trim
        Try
            Using MyCon As New SqlConnection(RetornaElConnectionString)
                MyCon.Open()
                MyData = New SqlDataAdapter(MyConsulta, MyCon)
                MyData.Fill(MyTable)
                MyCon.Close()
            End Using
            Return MyTable
        Catch ex As Exception
            StringErrorOut = ex.Message.ToString()
            Return MyTable
        End Try
    End Function


    ''' <summary>
    ''' Esta funcion retorna un uno en caso de exito, -3 en caso de que ya exista
    ''' -2 En caso de que exista una tabla con el número de columnas incompletas
    ''' -1 En el caso de un error en SQL y no se haya creado
    ''' </summary>
    ''' <param name="NombreTb"></param>
    ''' <param name="Columnas"></param>
    ''' <returns></returns>
    Function CreaTb(ByVal NombreTb As String, ByVal Columnas As List(Of String))
        'Inicial setting of variables'
        Dim Existe = IsTableExists(NombreTb, Columnas.Count)
        If Existe = -1 Then
            Dim MyStartStr = "Create table " + NombreTb + " ("
            Dim MaxiIndex As Integer = Columnas.Count - 1
            For Index As Integer = 0 To MaxiIndex
                If Index <> MaxiIndex Then
                    MyStartStr = MyStartStr + Columnas(Index) + ", "
                Else
                    MyStartStr = MyStartStr + Columnas(Index) + ") "
                End If
            Next
            Dim Comando = MyStartStr
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim Respuesta = cmd.ExecuteNonQuery()
                If Respuesta > 0 Then
                    Return 1
                Else
                    Return 0
                End If
                con.Close()
            End Using
            Return -1
        ElseIf Existe = 0 Then
            Return -2
        Else
            Return -3
        End If
    End Function

    ''' <summary>
    ''' Permite conocer si una tabla en Sql existe de manera contraria, regresa un -1
    ''' El caso 0 es si no contiene el número esperado de columnas la tabla
    ''' </summary>
    ''' <param name="NombreTb"></param>
    ''' <param name="NumberOfColums"></param>
    ''' <returns></returns>
    Function IsTableExists(ByVal NombreTb As String, ByVal NumberOfColums As Integer)
        Dim BusquedaDeTablaExistente = "Select Column_name From INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = '" + NombreTb + "'"
        Dim MyExistence As List(Of String) = SqlReaderDown2List(BusquedaDeTablaExistente)
        If MyExistence.Count = NumberOfColums Then
            Return 1
        ElseIf MyExistence.Count > 0 Then
            Return 0
        Else
            Return -1
        End If
    End Function

    Function ConcatenaColumnas(ByVal MyColumnas As List(Of String)) As String
        Dim Concatenacion As String = ""
        For Each _Str In MyColumnas
            Concatenacion = Concatenacion + " " + _Str + ","
        Next
        Return Concatenacion.Remove(Concatenacion.Length - 1)
    End Function

    Sub DeleteAnRegister(ByVal Table As String, ByVal ListOfConditiones As List(Of String), ByRef ListaDeErrores As List(Of String))
        Dim Comando As String = "Delete From @Table where @Conditions"
        If ListOfConditiones.Count > 0 Then
            Dim Condicional As String = JoinTheConditionsClauses(ListOfConditiones)
            Comando = Comando.Replace("@Table", Table)
            Comando = Comando.Replace("@Conditions", Condicional)
            Try
                Using con As New SqlConnection(ConnectionString)
                    con.Open()
                    Dim cmd As New SqlCommand(Comando, con)
                    Dim Respuesta = cmd.ExecuteNonQuery()
                    con.Close()
                End Using
            Catch ex As System.Exception
                ListaDeErrores.Add(ex.Message)
            End Try
        Else
            ListaDeErrores.Add("No se han indicado los parametros para eliminar un registro")
        End If
    End Sub
    Function DeleteAnRegister(ByVal Table As String, ByVal ListOfConditiones As List(Of String)) As String
        Dim Comando As String = "Delete From @Table where @Conditions"
        If ListOfConditiones.Count > 0 Then
            Dim Condicional As String = JoinTheConditionsClauses(ListOfConditiones)
            Comando = Comando.Replace("@Table", Table)
            Comando = Comando.Replace("@Conditions", Condicional)
            Try
                Using con As New SqlConnection(ConnectionString)
                    con.Open()
                    Dim cmd As New SqlCommand(Comando, con)
                    Dim Respuesta = cmd.ExecuteNonQuery()
                    con.Close()
                End Using
            Catch ex As System.Exception
                Return ex.Message
            End Try
        Else
            Return "No se han indicado los parametros para eliminar un registro"
        End If
        Return ""
    End Function

    Function ExecuteOnSqlImportingTheNumberOfAffectedRows(ByVal SqlComand As String) As List(Of String)
        Dim ListaDeErrores As New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(SqlComand, con)
                Dim respuesta = cmd.ExecuteNonQuery()
                If respuesta = 0 Then
                    'Exito = False
                    ListaDeErrores.Add("Respuesta :" + vbTab + SqlComand)
                    'MsgBox("")
                End If
                con.Close()
            End Using
        Catch er As System.Exception
            ListaDeErrores.Add(er.Message)
            'Exito = False
        End Try
        Return ListaDeErrores
    End Function
    Function ExecuteOnSqlImportingTheNumberOfAffectedRowsKV(ByVal SqlComand As String) As KeyValuePair(Of String, String)
        Dim ARetornar = New KeyValuePair(Of String, String)("", "")
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(SqlComand, con)
                Dim respuesta = cmd.ExecuteNonQuery()
                ARetornar = New KeyValuePair(Of String, String)("Respuesta SQL OK : ", respuesta.ToString)
                con.Close()
            End Using
        Catch er As System.Exception
            ARetornar = New KeyValuePair(Of String, String)("Error ", er.Message.ToString)
            'Exito = False
        End Try
    End Function
    Function GetkeyPairs2ComboBoxWithOutDescription(ByVal TableSql As String, ByVal ColumnValue As String, ByVal ColumnStatus As String, ByRef ErrorSql As String, Optional ByVal Condition As String = "") As List(Of KeyValuePair(Of String, String))
        Dim NewListOfkyes As New List(Of KeyValuePair(Of String, String))
        Dim ConsultaSql As String = "Select DISTINCT @Values, @Status from @Table @WhereClause Order by @Values;"
        ConsultaSql = ConsultaSql.Replace("@Table", TableSql)
        ConsultaSql = ConsultaSql.Replace("@Values", ColumnValue)
        ConsultaSql = ConsultaSql.Replace("@Status", ColumnStatus)
        If Condition.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "Where " + Condition)
        Else
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "")
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaSql, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    Dim AMostrar As String = ""
                    Dim AStatus As String
                    Dim AValor As String
                    If dr(ColumnStatus) Is DBNull.Value Then
                        AStatus = ""
                    Else
                        AStatus = Convert.ToString(dr(ColumnStatus))
                    End If
                    If dr(ColumnValue) Is DBNull.Value Then
                        AValor = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        AValor = Convert.ToString(dr(ColumnValue))
                    End If
                    NewListOfkyes.Add(New KeyValuePair(Of String, String)(AValor + ": " + vbTab + AStatus, AValor))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", ConsultaSql)
            ErrorSql = TemporalConcat
        End Try
        Return NewListOfkyes
    End Function
    Function GetkeyPairs2ComboBox(ByVal TableSql As String, ByVal ColumnaDisplayMember As String, ByVal ColumnStatus As String, ByVal ColumnValue As String, ByRef ErrorSql As String, Optional ByVal Condition As String = "") As List(Of KeyValuePair(Of String, String))
        Dim NewListOfkyes As New List(Of KeyValuePair(Of String, String))
        Dim ConsultaSql As String = "Select DISTINCT @Values, @Display, @Status from @Table @WhereClause Order by @Display;"
        ConsultaSql = ConsultaSql.Replace("@Table", TableSql)
        ConsultaSql = ConsultaSql.Replace("@Values", ColumnValue)
        ConsultaSql = ConsultaSql.Replace("@Display", ColumnaDisplayMember)
        ConsultaSql = ConsultaSql.Replace("@Status", ColumnStatus)
        If Condition.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "Where " + Condition)
        Else
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "")
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaSql, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    Dim AMostrar As String
                    Dim AStatus As String
                    Dim AValor As String
                    If dr(ColumnaDisplayMember) Is DBNull.Value Then
                        AMostrar = ""
                    Else
                        AMostrar = Convert.ToString(dr(ColumnaDisplayMember))
                    End If
                    If dr(ColumnStatus) Is DBNull.Value Then
                        AStatus = ""
                    Else
                        AStatus = Convert.ToString(dr(ColumnStatus))
                    End If
                    If dr(ColumnValue) Is DBNull.Value Then
                        AValor = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        AValor = Convert.ToString(dr(ColumnValue))
                    End If
                    NewListOfkyes.Add(New KeyValuePair(Of String, String)(AMostrar + ": " + vbTab + AStatus, AValor))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", ConsultaSql)
            ErrorSql = TemporalConcat
        End Try
        Return NewListOfkyes
    End Function

    Function GetkeyPairs2ComboBoxCustomDescription(ByVal TableSql As String, ByVal ExpresionDescription As String, ByVal AliasDescription As String, ByVal ColumnValue As String, ByRef ErrorSql As String, Optional ByVal Condition As String = "") As List(Of KeyValuePair(Of String, String))
        Dim NewListOfkyes As New List(Of KeyValuePair(Of String, String))
        Dim ConsultaSql As String = "Select @Values, @Expresion from @Table @WhereClause Order by @Alias;"
        ConsultaSql = ConsultaSql.Replace("@Table", TableSql)
        ConsultaSql = ConsultaSql.Replace("@Values", ColumnValue)
        ConsultaSql = ConsultaSql.Replace("@Alias", AliasDescription)
        If ExpresionDescription.Contains(" as ") Then
            ConsultaSql = ConsultaSql.Replace("@Expresion", ExpresionDescription)
        Else
            ConsultaSql = ConsultaSql.Replace("@Expresion", ExpresionDescription + " as [" + AliasDescription + "]")
        End If
        If Condition.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "Where " + Condition)
        Else
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "")
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaSql, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    Dim AMostrar As String
                    Dim AStatus As String
                    Dim AValor As String
                    If dr(AliasDescription) Is DBNull.Value Then
                        AMostrar = ""
                    Else
                        AMostrar = Convert.ToString(dr(AliasDescription))
                    End If

                    If dr(ColumnValue) Is DBNull.Value Then
                        AValor = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        AValor = Convert.ToString(dr(ColumnValue))
                    End If

                    NewListOfkyes.Add(New KeyValuePair(Of String, String)(AMostrar, AValor))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", ConsultaSql)
            ErrorSql = TemporalConcat
        End Try
        Return NewListOfkyes
    End Function

    Function GetkeyPairs2ComboBoxValueDescriptionPlusState(ByVal TableSql As String, ByVal Description As String, ByVal StatusExpresion As String, ByVal StatusAlias As String, ByVal ColumnValue As String, ByRef ErrorSql As String, Optional ByVal Condition As String = "") As List(Of KeyValuePair(Of String, String))
        Dim NewListOfkyes As New List(Of KeyValuePair(Of String, String))
        Dim ConsultaSql As String = "Select @Values, @Description, @Expresion from @Table @WhereClause Order by @Description;"
        ConsultaSql = ConsultaSql.Replace("@Table", TableSql)
        ConsultaSql = ConsultaSql.Replace("@Values", ColumnValue)
        ConsultaSql = ConsultaSql.Replace("@Description", Description)
        If StatusExpresion.Contains(" as ") Then
            ConsultaSql = ConsultaSql.Replace("@Expresion", StatusExpresion)
        Else
            ConsultaSql = ConsultaSql.Replace("@Expresion", StatusExpresion + " as [" + StatusAlias + "]")
        End If
        If Condition.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "Where " + Condition)
        Else
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "")
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaSql, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    Dim AMostrar As String
                    Dim AStatus As String
                    Dim ADescription As String
                    Dim AValor As String
                    If dr(StatusAlias) Is DBNull.Value Then
                        AStatus = ""
                    Else
                        AStatus = Convert.ToString(dr(StatusAlias))
                    End If
                    If dr(Description) Is DBNull.Value Then
                        ADescription = ""
                    Else
                        ADescription = Convert.ToString(dr(Description))
                    End If

                    If dr(ColumnValue) Is DBNull.Value Then
                        AValor = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        AValor = Convert.ToString(dr(ColumnValue))
                    End If
                    AMostrar = ADescription + ": " + AStatus
                    NewListOfkyes.Add(New KeyValuePair(Of String, String)(AMostrar, AValor))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", ConsultaSql)
            ErrorSql = TemporalConcat
        End Try
        Return NewListOfkyes
    End Function

    Function GetkeyPairs2ComboBoxWithoutOrder(ByVal TableSql As String, ByVal ColumnaDisplayMember As String, ByVal ColumnStatus As String, ByVal ColumnValue As String, ByRef ErrorSql As String, Optional ByVal DefineOrder As String = "", Optional ByVal Condition As String = "") As List(Of KeyValuePair(Of String, String))
        Dim NewListOfkyes As New List(Of KeyValuePair(Of String, String))
        Dim ConsultaSql As String = "Select DISTINCT @Values, @Display, @Status from @Table @WhereClause @OrderClause;"
        ConsultaSql = ConsultaSql.Replace("@Table", TableSql)
        ConsultaSql = ConsultaSql.Replace("@Values", ColumnValue)
        ConsultaSql = ConsultaSql.Replace("@Display", ColumnaDisplayMember)
        ConsultaSql = ConsultaSql.Replace("@Status", ColumnStatus)
        If DefineOrder.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@OrderClause;", "Order by " + DefineOrder)
        Else
            ConsultaSql = ConsultaSql.Replace("@OrderClause;", ";")
        End If
        If Condition.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "Where " + Condition)
        Else
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "")
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaSql, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    Dim AMostrar As String
                    Dim AStatus As String
                    Dim AValor As String
                    If dr(ColumnaDisplayMember) Is DBNull.Value Then
                        AMostrar = ""
                    Else
                        AMostrar = Convert.ToString(dr(ColumnaDisplayMember))
                    End If
                    If dr(ColumnStatus) Is DBNull.Value Then
                        AStatus = ""
                    Else
                        AStatus = Convert.ToString(dr(ColumnStatus))
                    End If
                    If dr(ColumnValue) Is DBNull.Value Then
                        AValor = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        AValor = Convert.ToString(dr(ColumnValue))
                    End If
                    NewListOfkyes.Add(New KeyValuePair(Of String, String)(AMostrar + ": " + vbTab + AStatus, AValor))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", ConsultaSql)
            ErrorSql = TemporalConcat
        End Try
        Return NewListOfkyes
    End Function

    Function GetkeyPairs2ComboBoxWithOutStatus(ByVal TableSql As String, ByVal ColumnaDisplayMember As String, ByVal ColumnValue As String, ByRef ErrorSql As String, Optional ByVal Condition As String = "") As List(Of KeyValuePair(Of String, String))
        Dim NewListOfkyes As New List(Of KeyValuePair(Of String, String))
        Dim ConsultaSql As String = "Select DISTINCT @Values, @Display from @Table @WhereClause Order by @Display;"
        ConsultaSql = ConsultaSql.Replace("@Table", TableSql)
        ConsultaSql = ConsultaSql.Replace("@Values", ColumnValue)
        ConsultaSql = ConsultaSql.Replace("@Display", ColumnaDisplayMember)

        If Condition.Length > 0 Then
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "Where " + Condition)
        Else
            ConsultaSql = ConsultaSql.Replace("@WhereClause", "")
        End If
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(ConsultaSql, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    Dim AMostrar As String
                    Dim AValor As String
                    If dr(ColumnaDisplayMember) Is DBNull.Value Then
                        AMostrar = ""
                    Else
                        AMostrar = Convert.ToString(dr(ColumnaDisplayMember))
                    End If

                    If dr(ColumnValue) Is DBNull.Value Then
                        AValor = ""
                    Else
                        'listaC.Add(dr(str).ToString)
                        AValor = Convert.ToString(dr(ColumnValue))
                    End If
                    NewListOfkyes.Add(New KeyValuePair(Of String, String)(AMostrar, AValor))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            Dim TemporalConcat As String = "Se ha detectado un error: @Error" + vbNewLine +
            "Consulta: @Co1"
            TemporalConcat = TemporalConcat.Replace("@Error", ex.Message.ToString()).Replace("@Co1", ConsultaSql)
            ErrorSql = TemporalConcat
        End Try
        Return NewListOfkyes
    End Function

    'Function ReturnLastChildren(ByRef MyStackPanel As StackPanel) As Object
    '    Return MyStackPanel.Children(MyStackPanel.Children.Count - 1)
    'End Function

    'Function ReturnLastChildren(ByRef MyStackPanel As Object) As Object
    '    Return MyStackPanel.Children(MyStackPanel.Children.Count - 1)
    'End Function
    'Function RegresaHoraServidor() As DateTime
    '    '-------------------------Generación de la fecha del servidor
    '    Dim MySqlCon = New GeneradoresDeConexion().GeneraSql2Productividad
    '    Dim ConsultaSql = "Select SYSDATETIME()"
    '    Dim ResponseSql = ""
    '    Dim _Table = MySqlCon.DataParaTable(ConsultaSql, ResponseSql)
    '    'Se verifica que el tipo de incidencia sea registrable, es decir que sea del tipo prevista
    '    If ResponseSql.Length > 0 Then
    '        Throw New Exception("No se ha logrado obtener la fecha del servidor")
    '    End If
    '    Return Convert.ToDateTime(_Table.Rows(0)(0))
    'End Function

    Private Shared Function ConvertDataTable2(Of T)(ByVal dt As DataTable) As List(Of T)
        Dim data As List(Of T) = New List(Of T)()
        For Each row As DataRow In dt.Rows
            Dim item As T = GetItem(Of T)(row)
            data.Add(item)
        Next
        Return data
    End Function

    Private Shared Function GetItem(Of T)(ByVal dr As DataRow) As T
        Dim temp As Type = GetType(T)
        Dim obj As T = Activator.CreateInstance(Of T)()
        For Each column As DataColumn In dr.Table.Columns
            For Each pro As PropertyInfo In temp.GetProperties()
                If pro.Name = column.ColumnName Then
                    pro.SetValue(obj, dr(column.ColumnName), Nothing)
                Else
                    Continue For
                End If
            Next
        Next

        Return obj
    End Function

    Public Function ToLabelFormat(ByVal Str As String) As String
        Dim newStr = Regex.Replace(Str, "(?<=[A-Z])(?=[A-Z][a-z])", " ")
        newStr = Regex.Replace(newStr, "(?<=[^A-Z])(?=[A-Z])", " ")
        newStr = Regex.Replace(newStr, "(?<=[A-Za-z])(?=[^A-Za-z])", " ")
        Return newStr
    End Function

    Function InsertaEnForma(ByRef MySQLCon As AdmSQL, ByVal Tabla As String, ByVal NewData As List(Of String)) As List(Of String)
        Dim MyListaDeErrores As New List(Of String)
        MyListaDeErrores = MySQLCon.InsertaEnSqlV2(Tabla, MySQLCon.InsertComillas(NewData))
        Return MyListaDeErrores
    End Function

    Function ActualizaEnFormaSinLimitante(ByRef MySQLCon As AdmSQL, ByVal Tabla As String, ByVal MyOldData As List(Of String), ByVal MyNewData As List(Of String), ByVal MyColumns As List(Of String)) As List(Of String)

        Dim ListaDeErrores As New List(Of String)

        Dim FiltroOld = FiltroInteligente(MyOldData)
        Dim FiltroNew = FiltroInteligente(MyNewData)

        Dim ColumnasWithOld As List(Of String) = ArmaNuevaLista(MyColumns, FiltroOld)
        Dim ColumnasWithNew As List(Of String) = ArmaNuevaLista(MyColumns, FiltroNew)
        Dim IndicadorA As Boolean = False

        Dim Cond = MySQLCon.RetornaIgualdadesV2(ColumnasWithOld, ArmaNuevaLista(MyOldData, FiltroOld), IndicadorA)
        If IndicadorA Then
            ListaDeErrores.Add("Se ha encontrado un error en crear las condiciones de cambio, nota del programador")
        End If

        Dim IndicadorB As Boolean = False
        Dim Nuev = MySQLCon.RetornaIgualdadesV2(ColumnasWithNew, ArmaNuevaLista(MyNewData, FiltroNew), IndicadorB)
        If IndicadorA Then
            ListaDeErrores.Add("Se ha encontrado un error en crear las igualdades de cambio, nota del programador")
        End If

        Dim MyListOfTheError As List(Of String) = MySQLCon.UpdateOnSQL(Tabla, Nuev, Cond)
        If MyListOfTheError.Count > 0 Then
            ListaDeErrores.AddRange(MyListOfTheError)
        End If
        Return ListaDeErrores
    End Function


    ''' <summary>
    ''' Retorna true si el unico caracter existente es el antes propuesto.
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <param name="MyStrCmp"></param>
    ''' <returns></returns>
    Function IsOnlyWith(ByVal MyList As List(Of String), ByVal MyStrCmp As String)
            Dim IsOnl As Boolean = True
            For Each str As String In MyList
                If Not String.Compare(str.Trim, MyStrCmp) = 0 Then
                    IsOnl = False
                End If
            Next
            Return IsOnl
        End Function

        ''' <summary>
        ''' Retorna true si el unico caracter existente es el antes propuesto.
        ''' </summary>
        ''' <returns></returns>
        Function IsOnlyWith(ByVal MyList As List(Of Boolean), ByVal MyBool As Boolean)
            Dim IsOnl As Boolean = True
            For Each Bol As Boolean In MyList
                If Bol <> MyBool Then
                    IsOnl = False
                End If
            Next
            Return IsOnl
        End Function

        ''' <summary>
        ''' Devuelve una lista con las posiciones de MyList que no sean NULL, o con simplemente "".
        ''' </summary>
        ''' <param name="MyList"></param>
        ''' <returns></returns>
        Function FiltroInteligente(ByRef MyList As List(Of String)) As List(Of Integer)
            Dim MyRet As List(Of Integer) = New List(Of Integer)
            For Index As Integer = 0 To MyList.Count - 1
                If String.Compare(MyList(Index).Replace("'", " ").Replace("NULL", " ").Trim, "") <> 0 Then
                    MyRet.Add(Index)
                End If
            Next
            Return MyRet
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="MyList"></param>
        ''' <returns></returns>
        Function FiltroInteligenteV2(ByRef MyList As List(Of String)) As List(Of Integer)
            Dim MyRet As List(Of Integer) = New List(Of Integer)
            For Index As Integer = 0 To MyList.Count - 1
                Dim MyStr As String = MyList(Index).Replace("'", " ").Replace("NULL", " ").Trim
                If MyStr.Count > 0 Then
                    MyRet.Add(Index)
                End If
            Next
            Return MyRet
        End Function

        ''' <summary>
        ''' En base a una lista de posiciones regresa otra con solo las posiciones especificadas de MyList.
        ''' </summary>
        ''' <param name="MyList"></param>
        ''' <param name="MyPos"></param>
        ''' <returns></returns>
        Function ArmaNuevaLista(ByRef MyList As List(Of String), ByVal MyPos As List(Of Integer))
            Dim MyRet As List(Of String) = New List(Of String)
            For Each int As Integer In MyPos
                MyRet.Add(MyList(int))
            Next
            Return MyRet
        End Function







    Sub VerificaString2Double(ByRef MyText As String, ByRef ListOfErrors As List(Of String))
            Dim MyD As Double
            Dim MyTxt As String = MyText.Trim
            If Not Double.TryParse(MyTxt, MyD) Then
                'MyText.BackColor = ColorI
                ListOfErrors.Add("No se ha logrado convertir " + MyTxt + " a un número")
            End If
        End Sub
        Function VerificaString2DoubleSiEsCorrectoIsOn(ByRef MyText As String, ByRef Data As List(Of String), ByRef ListOfErrors As List(Of String))
            Dim MyD As Double
            Dim MyTxt As String = MyText.Trim
            If Not Double.TryParse(MyTxt, MyD) Then
                'MyText.BackColor = ColorI
                ListOfErrors.Add("No se ha logrado convertir " + MyTxt + " a un número")
                Return False
            Else
                'MyText.BackColor = colorC
                Data.Add(MyD.ToString)
            End If
            Return True
        End Function


        Function RetornaDiferencias(ByVal MyListA As List(Of String), ByVal MyListB As List(Of String))
            Dim MyDiferencias As List(Of Integer) = New List(Of Integer)
            If MyListA.Count = MyListB.Count Then
                For Index As Integer = 0 To MyListB.Count - 1
                    If MyListA(Index) <> MyListB(Index) Then
                        MyDiferencias.Add(1)
                    Else
                        MyDiferencias.Add(0)
                    End If
                Next
            Else
                MsgBox("Las listas son de diferentes tamaños",, "Favor de notificar a sistemas")
            End If
            Return MyDiferencias
        End Function

        Function RetornaDiferenciasBoolean(ByVal MyListA As List(Of String), ByVal MyListB As List(Of String))
            Dim MyDiferencias As List(Of Boolean) = New List(Of Boolean)
            If MyListA.Count = MyListB.Count Then
                For Index As Integer = 0 To MyListB.Count - 1
                    If MyListA(Index) <> MyListB(Index) Then
                        MyDiferencias.Add(True)
                    Else
                        MyDiferencias.Add(False)
                    End If
                Next
            Else
                MsgBox("Las listas son de diferentes tamaños",, "Favor de notificar a sistemas")
            End If
            Return MyDiferencias
        End Function

        ''' <summary>
        ''' This functions is used to return true if any of the list contains a true; this was created for work into Empleados
        ''' </summary>
        ''' <returns></returns>
        Function FindAndReduceBoolean(ByVal Mylist As List(Of Boolean))
            For Each MyDato As Boolean In Mylist
                If MyDato = True Then
                    Return True
                    Exit Function
                End If
            Next
            Return False
        End Function




        Function ListOfNumbersBetween(ByVal MyList As List(Of Integer), ByVal Inferior As Integer, ByVal Superior As Integer)
            Dim IsValid As Boolean = True
            For Each MyInteger As Integer In MyList
                If MyInteger < Inferior OrElse MyInteger > Superior Then
                    IsValid = False
                    Exit For
                End If
            Next
            Return IsValid
        End Function

        ''' <summary>
        ''' Permite verificar un texto de manera que si contiene información se anida a la lista
        ''' de otra manera se anida el texto propuesto al error
        ''' </summary>
        '''<param name="ListData">"Lista de datos"</param>
        Sub VerifiTextConLE(ByVal MyTxt As String, ByRef ListData As List(Of String), ByRef MyError As String, ByVal TextoDeError As String)
            MyTxt = MyTxt.Trim
            If MyTxt.Length > 0 Then
                ListData.Add(MyTxt)
            Else
                MyError = MyError + TextoDeError + vbCrLf
            End If
        End Sub

        ''' <summary>
        ''' Permite determinar si el texto es un decimal válido, en caso de que pueda estar vacio el campo se agrega a la lista un "0"
        ''' Si el decimal es válido se redondea a los decimales especificados
        '''
        ''' </summary>
        ''' <param name="MyTxt"></param>
        ''' <param name="ListData"></param>
        ''' <param name="MyError"></param>
        ''' <param name="TextoDeErrorVacio"></param>
        ''' <param name="TextoDeErrorMalaConversion"></param>
        ''' <param name="CanBeEmply"></param>
        ''' <param name="NumeroDecimales"></param>
        Sub VerifiTextConLEDecimal(ByVal MyTxt As String, ByRef ListData As List(Of String), ByRef MyError As String, ByVal TextoDeErrorVacio As String, ByVal TextoDeErrorMalaConversion As String, ByVal CanBeEmply As Boolean, ByVal NumeroDecimales As Integer)
            Dim TextoT As String
            Dim NumerT As Decimal
            TextoT = MyTxt.Trim         '9 Horas rec RPM
            If TextoT.Length > 0 Then
                If Decimal.TryParse(TextoT, NumerT) Then
                    'ListData.Add(RedondeaDecimal(NumerT, NumeroDecimales).ToString)
                Else
                    MyError = MyError + TextoDeErrorMalaConversion + vbCrLf
                End If
            Else
                If CanBeEmply Then
                    ListData.Add("0")
                Else
                    MyError = MyError + TextoDeErrorVacio + vbCrLf
                End If
            End If
        End Sub
        ''' <summary>
        ''' Retorna 0 para 
        ''' </summary>
        ''' <param name="MyTxtToCheck"></param>
        ''' <returns></returns>
        Function VerificaBooleano(ByVal MyTxtToCheck As String, ByRef NotificadorDeError As Boolean) As Boolean
            Dim ValorDeRetorno As Boolean = False
            NotificadorDeError = False
            If Not Boolean.TryParse(MyTxtToCheck, ValorDeRetorno) Then
                NotificadorDeError = True
            End If
            Return ValorDeRetorno
        End Function
        Function VerificaBooleano(ByVal MyTxtToCheck As String, ByRef NotificadorDeError As String) As Boolean
            NotificadorDeError = ""
            Dim ValorDeRetorno As Boolean
            If Not Boolean.TryParse(MyTxtToCheck, ValorDeRetorno) Then
                NotificadorDeError = "No se ha logrado convertir el String '" + MyTxtToCheck + "'"
            End If
            Return ValorDeRetorno
        End Function
        Function VerificaHoraValida(ByVal MyStrHora As String) As Boolean
            Dim IsValid As Boolean = False
            Dim TemporalDate As DateTime
            Return DateTime.TryParse(MyStrHora, TemporalDate)
        End Function

        Function ThisListHasTheEnoughtData(ByVal MyListOfData As List(Of String), Optional ByVal MinOfData As Integer = 1) As Boolean
            Return FiltroInteligenteV2(MyListOfData).Count >= MinOfData
        End Function
        Function HowMuchDataHasThisList(ByVal MyListOfData As List(Of String)) As Integer
            Return FiltroInteligenteV2(MyListOfData).Count
        End Function
        Function ListStr2ListInt(ByVal MyListOfStrings As List(Of String), ByRef MyErrorsOut As List(Of String)) As List(Of Integer)
            Dim MyInts As New List(Of Integer)
            MyErrorsOut.Clear()
            Dim Temporal As Integer = 0
            For Each My_Str As String In MyListOfStrings
                If Integer.TryParse(My_Str, Temporal) Then
                    MyInts.Add(Temporal)
                Else
                    MyErrorsOut.Add("Error, trying to pass: " + My_Str + "to an Int")
                End If
            Next
            Return MyInts
        End Function
        Sub DontAddEmplyToList(ByVal MyListaOfData As List(Of String), ByVal StrToAdd As String)
            If StrToAdd.Count > 0 Then
                MyListaOfData.Add(StrToAdd)
            End If
            'Return MyListaOfData
        End Sub
        Function IsTextADoubleValid(ByVal _Str As String) As Boolean
            Dim Temporal As Double
            Dim _2ret As Boolean = False
            If Double.TryParse(_Str, Temporal) Then
                _2ret = True
            End If
            Return _2ret
        End Function
        Function CheckStringWithList(ByVal String2Compare As String, ByVal ListaDeOpciones As List(Of String))
            String2Compare = String2Compare.Trim
            If ListaDeOpciones.Contains(String2Compare) Then
                Return True
            Else
                Return False
            End If
        End Function
        Function CheckString1OrTrue(ByVal _String2Compare As String)
            Dim ListaDeOpcion As New List(Of String) From {
                "1", "True", "true", "TRUE"
            }
            Return CheckStringWithList(_String2Compare, ListaDeOpcion)
        End Function
    Function CheckString0OrFalse(ByVal _String2Compare As String)
        Dim ListaDeOpcion As New List(Of String) From {
                "0", "False", "false", "FALSE"
            }
        Return CheckStringWithList(_String2Compare, ListaDeOpcion)
    End Function

End Class
