Imports System.IO
Imports System.Text.RegularExpressions

Public Class frmPaginaPrincipal
    Private _dtRaices As DataTable
    Private _dtPlataformas As DataTable
    Private _dtDirectorios As DataTable
    Private _dtArchivos As DataTable
    Private _dtBls As DataTable
    Private _dtMappers As DataTable
    Private _dtScripts As DataTable
    Private _dtFuncionesBL As DataTable
    Private _dtFuncionesMapper As DataTable
    Private _rutaPlataforma As String
    Private _lstExcluidos As List(Of String)
    Private _htExcepcionesBL As Hashtable
    Private _htExcepcionesMapper As Hashtable
    Private _Solucion As String

    Private Sub frmPaginaPrincipal_Load(sender As Object, e As EventArgs) Handles Me.Load

        _dtPlataformas = New DataTable
        With _dtPlataformas
            .Columns.Add("idPlataforma", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Ruta", GetType(String))

            .Rows.Add(1, "RecFaecys2", "C:\Users\Fernando\Dropbox\Fernando.viu\Programacion\RecFaecys2\Sistema")
            .Rows.Add(2, "Filiales", "C:\Users\Fernando\Dropbox\Fernando.viu\Programacion\Filiales")
            .Rows.Add(3, "Filiales2", "C:\Users\Fernando\Dropbox\Fernando.viu\Programacion\Filiales2")
        End With

        _dtDirectorios = New DataTable
        With _dtDirectorios
            .Columns.Add("idDirectorio", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Ruta", GetType(String))
        End With

        _dtArchivos = New DataTable
        With _dtArchivos
            .Columns.Add("idArchivo", GetType(Integer))
            .Columns.Add("idDirectorio", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Ruta", GetType(String))
        End With

        _dtBls = New DataTable
        With _dtBls
            .Columns.Add("idBL", GetType(Integer))
            .Columns.Add("idArchivo", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Ruta", GetType(String))
        End With

        _dtFuncionesBL = New DataTable
        With _dtFuncionesBL
            .Columns.Add("idFuncionBL", GetType(Integer))
            .Columns.Add("idBL", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Parametros", GetType(String))
            .Columns.Add("Retorno", GetType(String))
        End With

        _dtMappers = New DataTable
        With _dtMappers
            .Columns.Add("idMapper", GetType(Integer))
            .Columns.Add("idBL", GetType(Integer))
            .Columns.Add("idFuncionBL", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Ruta", GetType(String))
        End With

        _dtFuncionesMapper = New DataTable
        With _dtFuncionesMapper
            .Columns.Add("idFuncionMapper", GetType(Integer))
            .Columns.Add("idMapper", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Parametros", GetType(String))
            .Columns.Add("Retorno", GetType(String))
        End With

        _dtScripts = New DataTable
        With _dtScripts
            .Columns.Add("idScript", GetType(Integer))
            .Columns.Add("idMapper", GetType(Integer))
            .Columns.Add("idFuncionMapper", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
        End With

        _lstExcluidos = New List(Of String)
        _lstExcluidos.Add("BL")
        _lstExcluidos.Add("Mapper")
        _lstExcluidos.Add("Scripts")

        cmbPlataformas.DisplayMember = "Nombre"
        cmbPlataformas.ValueMember = "idPlataforma"
        cmbPlataformas.DataSource = _dtPlataformas

        _htExcepcionesBL = New Hashtable
        _htExcepcionesBL("frmVerificacionPagos.vb") = "CorreccionPagos"
        _htExcepcionesBL("frmPrincipal.vb") = "Alertas"

        _htExcepcionesMapper = New Hashtable
        _htExcepcionesMapper("CorreccionPagosBL.vb") = "CorreccionPagos"
        _htExcepcionesMapper("RecordatorioProcesoLineasBL.vb") = "Alertas"
        _htExcepcionesMapper("RecordatorioProcesoBL.vb") = "Alertas"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim drPlataforma As DataRow = _dtPlataformas.Select("idPlataforma = " & cmbPlataformas.SelectedValue.ToString).First
        'Listas
        Dim lstDirectorios As List(Of String) = Nothing
        Dim lstArchivos As List(Of String) = Nothing
        Dim lstBLs As List(Of String) = Nothing
        Dim lstFuncionesBL As List(Of String) = Nothing
        Dim lstMappers As List(Of String) = Nothing
        Dim lstFuncionesMapper As List(Of String) = Nothing
        Dim lstScripts As List(Of String) = Nothing
        'Ids
        Dim idDirectorio As Integer = 0
        Dim idArchivo As Integer = 0
        Dim idBL As Integer = 0
        Dim idFuncionBL As Integer = 0
        Dim idMapper As Integer = 0
        Dim idFuncionMapper As Integer = 0
        Dim idScript As Integer = 0

        _rutaPlataforma = drPlataforma.Item("Ruta")
        _Solucion = File.ReadAllText(drPlataforma.Item("Ruta") & "\" & drPlataforma("Nombre") & ".sln")
        lstDirectorios = Directory.GetDirectories(_rutaPlataforma).ToList

        'Cicla entre todos los directorios de la raiz de la plataforma.
        For Each rutaDirectorio As String In lstDirectorios
            Dim direcotorioInfo As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(rutaDirectorio)

            If _lstExcluidos.Find(Function(p As String) p = direcotorioInfo.Name) = "" And _Solucion.Contains(direcotorioInfo.Name & ".vbproj") Then
                'Busca los archivos en el directorio
                If lstArchivos IsNot Nothing Then lstArchivos.Clear()
                lstArchivos = BuscarTextoEnArchivosVB(rutaDirectorio, "BL.")

                'Si encontró uno o mas archivos con el texto va a ciclarlo/s
                If lstArchivos IsNot Nothing AndAlso lstArchivos.Count > 0 Then
                    Dim archivoInfo As FileInfo

                    idDirectorio += 1
                    _dtDirectorios.Rows.Add(idDirectorio, direcotorioInfo.Name)

                    For Each rutaArchivo In lstArchivos
                        archivoInfo = My.Computer.FileSystem.GetFileInfo(rutaArchivo)

                        'se entiende que todos los archivos acá tienen un BL así que no tengo que esperar a que los busque para agregar el registro
                        idArchivo += 1
                        _dtArchivos.Rows.Add(idArchivo, idDirectorio, archivoInfo.Name, rutaArchivo)

                        If lstBLs IsNot Nothing Then lstBLs.Clear()
                        lstBLs = BuscarBLs(rutaArchivo, archivoInfo.Name)

                        If lstBLs IsNot Nothing AndAlso lstBLs.Count > 0 Then
                            For Each archivoBL As String In lstBLs
                                Dim rutaBL As String = _rutaPlataforma & "\BL\" & archivoBL & ".vb"
                                Dim blInfo As FileInfo

                                If Not File.Exists(rutaBL) Then
                                    Dim alternativa As String = _htExcepcionesBL(archivoInfo.Name)
                                    rutaBL = _rutaPlataforma & "\" & alternativa & "\" & archivoBL & ".vb"
                                End If

                                blInfo = My.Computer.FileSystem.GetFileInfo(rutaBL)

                                idBL += 1

                                _dtBls.Rows.Add(idBL, idArchivo, archivoBL, rutaBL)

                                If blInfo.Name = "RequerimientosBL.vb" Then
                                    Dim kk As Integer = 0
                                End If

                                lstFuncionesBL = BuscarFuncionesBL(rutaBL, archivoInfo.Name)
                                lstMappers = BuscarMappers(rutaBL, archivoInfo.Name)

                                If lstFuncionesBL IsNot Nothing AndAlso lstFuncionesBL.Count > 0 Then
                                    For Each funcionBL As String In lstFuncionesBL
                                        Dim parametros() As String = funcionBL.Split(";")

                                        idFuncionBL += 1
                                        _dtFuncionesBL.Rows.Add(idFuncionBL, idBL, parametros(0), parametros(1), parametros(2))
                                    Next
                                End If

                                If lstMappers IsNot Nothing AndAlso lstMappers.Count > 0 Then
                                    For Each archivoMapper In lstMappers
                                        Dim rutaMapper As String = _rutaPlataforma & "\Mapper\" & archivoMapper & ".vb"
                                        Dim mapperInfo As FileInfo

                                        If Not File.Exists(rutaMapper) Then
                                            Dim alternativa As String = _htExcepcionesMapper(blInfo.Name)
                                            rutaMapper = _rutaPlataforma & "\" & alternativa & "\" & archivoMapper & ".vb"
                                        End If

                                        mapperInfo = My.Computer.FileSystem.GetFileInfo(rutaMapper)

                                        idMapper += 1

                                        _dtMappers.Rows.Add(idMapper, idBL, Nothing, archivoMapper, rutaMapper)

                                        lstFuncionesMapper = BuscarFuncionesMapper(rutaMapper, blInfo.Name)
                                        lstScripts = BuscarScripts(rutaMapper, blInfo.Name)

                                        If lstFuncionesMapper IsNot Nothing AndAlso lstFuncionesMapper.Count > 0 Then
                                            For Each funcionMapper As String In lstFuncionesMapper
                                                Dim parametros() As String = funcionMapper.Split(";")

                                                idFuncionMapper += 1
                                                _dtFuncionesMapper.Rows.Add(idFuncionMapper, idMapper, parametros(0), parametros(1), parametros(2))
                                            Next
                                        End If

                                        If lstScripts IsNot Nothing AndAlso lstScripts.Count > 0 Then
                                            For Each script In lstScripts
                                                idScript += 1

                                                _dtScripts.Rows.Add(idScript, idMapper, Nothing, script)
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next

        GenerarArbol()
    End Sub

    Private Sub GenerarArbol()
        'Dim nodoRaiz As New TreeNode
        'idArchivo += 1
        'nodoRaiz.Text = archivo.Name
        'tvEstructura.Nodes.Add(nodoRaiz)

        For Each drDirectorio As DataRow In _dtDirectorios.Rows
            Dim nodoRaiz As New TreeNode
            Dim drArchivos() As DataRow = _dtArchivos.Select("idDirectorio = " & drDirectorio("idDirectorio"))

            If drArchivos IsNot Nothing AndAlso drArchivos.Count > 0 Then
                nodoRaiz.Text = drDirectorio("Nombre")
                nodoRaiz.Name = drDirectorio("idDirectorio").ToString & "-"

                For Each drArchivo As DataRow In drArchivos
                    Dim nodoArchivo As New TreeNode
                    Dim drBLs() As DataRow = _dtBls.Select("idArchivo = " & drArchivo("idArchivo"), "Nombre")

                    nodoArchivo.Text = drArchivo("Nombre")
                    nodoArchivo.Name = drDirectorio("idDirectorio").ToString & "-" & drArchivo("idArchivo").ToString & "-"

                    nodoRaiz.Nodes.Add(nodoArchivo)

                    If drBLs IsNot Nothing AndAlso drBLs.Count > 0 Then
                        For Each drBL As DataRow In drBLs
                            Dim nodoBL As New TreeNode
                            Dim drFuncionesBL() As DataRow = _dtFuncionesBL.Select("idBL = " & drBL("idBL").ToString, "Nombre")
                            Dim drMappers() As DataRow = _dtMappers.Select("idBL = " & drBL("idBL").ToString, "Nombre")

                            nodoBL.Text = drBL("Nombre")
                            nodoBL.Name = drDirectorio("idDirectorio").ToString & "-" & drArchivo("idArchivo").ToString & "-" & drBL("idBL").ToString & "-"

                            If drFuncionesBL IsNot Nothing AndAlso drFuncionesBL.Count > 0 Then
                                Dim nodoFuncionesBL As New TreeNode
                                nodoFuncionesBL.Text = "Funciones"

                                For Each drFuncionBL As DataRow In drFuncionesBL
                                    Dim nodoFuncionBL As New TreeNode
                                    nodoFuncionBL.Text = drFuncionBL("Nombre")
                                    nodoFuncionBL.Name = drDirectorio("idDirectorio").ToString & "-" & drArchivo("idArchivo").ToString & "-" & drBL("idBL").ToString & "-" & drFuncionBL("idFuncionBL").ToString & "-"
                                    nodoFuncionesBL.Nodes.Add(nodoFuncionBL)
                                Next

                                nodoBL.Nodes.Add(nodoFuncionesBL)
                            End If

                            If drMappers IsNot Nothing AndAlso drMappers.Count > 0 Then
                                Dim nodoMappers As New TreeNode
                                nodoMappers.Text = "Mappers"

                                For Each drMapper As DataRow In drMappers
                                    Dim nodoMapper As New TreeNode
                                    Dim drFuncionesMapper() As DataRow = _dtFuncionesMapper.Select("idMapper = " & drMapper("idMapper").ToString, "Nombre")
                                    Dim drScripts() As DataRow = _dtScripts.Select("idMapper = " & drMapper("idMapper").ToString, "Nombre")

                                    nodoMapper.Text = drMapper("Nombre")
                                    nodoMapper.Name = drDirectorio("idDirectorio").ToString & "-" & drArchivo("idArchivo").ToString & "-" & drBL("idBL").ToString & "-" & drMapper("idMapper").ToString & "-"

                                    If drFuncionesMapper IsNot Nothing AndAlso drFuncionesMapper.Count > 0 Then
                                        Dim nodoFuncionesMapper As New TreeNode
                                        nodoFuncionesMapper.Text = "Funciones"

                                        For Each drFuncionMapper In drFuncionesMapper
                                            Dim nodoFuncionMapper As New TreeNode
                                            nodoFuncionMapper.Text = drFuncionMapper("Nombre")
                                            nodoFuncionMapper.Name = drDirectorio("idDirectorio").ToString & "-" & drArchivo("idArchivo").ToString & "-" & drBL("idBL").ToString & "-" & drMapper("idMapper").ToString & "-" & drFuncionMapper("idFuncionMapper").ToString & "-"
                                            nodoFuncionesMapper.Nodes.Add(nodoFuncionMapper)
                                        Next

                                        nodoMapper.Nodes.Add(nodoFuncionesMapper)
                                    End If

                                    If drScripts IsNot Nothing AndAlso drScripts.Count > 0 Then
                                        Dim nodoScripts As New TreeNode
                                        nodoScripts.Text = "Scripts"

                                        For Each drScript As DataRow In drScripts
                                            Dim nodoScript As New TreeNode
                                            nodoScript.Text = drScript("Nombre")
                                            nodoScript.Name = drDirectorio("idDirectorio").ToString & "-" & drArchivo("idArchivo").ToString & "-" & drBL("idBL").ToString & "-" & drMapper("idMapper").ToString & "-" & drScript("idScript").ToString & "-"
                                            nodoScripts.Nodes.Add(nodoScript)
                                        Next

                                        nodoMapper.Nodes.Add(nodoScripts)
                                    End If

                                    nodoMappers.Nodes.Add(nodoMapper)
                                Next

                                nodoBL.Nodes.Add(nodoMappers)
                            End If

                            nodoArchivo.Nodes.Add(nodoBL)
                        Next

                    End If

                Next

                tvEstructura.Nodes.Add(nodoRaiz)
            End If
        Next

    End Sub


    Private Function BuscarTextoEnArchivosVB(carpeta As String, textoBuscado As String) As List(Of String)
        Dim idArchivos As Integer = 0
        Dim idBls As Integer = 0
        Dim idMappers As Integer = 0
        Dim contenido As String

        Dim lstArchivos As New List(Of String)

        For Each RutaDeArchivo As String In Directory.GetFiles(carpeta, "*.vb", SearchOption.AllDirectories)
            contenido = File.ReadAllText(RutaDeArchivo)

            'Va a leer en todo el contenido del archivo y va a buscar cualquier ocurrencia del texto que necesito.
            If contenido.Contains(textoBuscado) Then
                idArchivos += 1
                lstArchivos.Add(RutaDeArchivo)
            End If
        Next

        'Devuelve la lista con las rutas de archivos 
        Return lstArchivos
    End Function

    Private Function BuscarBLs(filePath As String, fileCaller As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstBLs As New List(Of String)

        For Each linea As String In lstLineas
            If Not linea.Contains("'") Then
                Dim m As Match = Regex.Match(linea, "(\w+)BL\.")

                If m.Success Then
                    If lstBLs.Find(Function(p As String) p = m.Groups(1).Value & "BL") = "" Then

                        lstBLs.Add(m.Groups(1).Value & "BL")
                    End If
                End If
            End If
        Next

        Return lstBLs
    End Function

    Private Function BuscarFuncionesBL(filePath As String, fileCaller As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstFuncionesBLs As New List(Of String)

        For Each linea As String In lstLineas
            If Not linea.Contains("'") And (linea.Contains("Function ") Or linea.Contains("Sub ")) Then
                Dim m As Match = Regex.Match(linea, "(\w+)\(")

                If m.Success Then
                    If lstFuncionesBLs.Find(Function(p As String) p = m.Groups(1).Value & "BL") = "" Then
                        Dim p As Match = Regex.Match(linea, "\((.*?)\)")
                        Dim retorno As String = ""

                        If linea.Contains("Function") Then
                            Dim q As Match = Regex.Match(linea, "\)\s+As\s+(\w+)")
                            retorno = q.Groups(1).Value
                        End If

                        lstFuncionesBLs.Add(m.Groups(1).Value & ";" & p.Groups(1).Value & ";" & retorno)
                    End If
                End If
            End If
        Next

        Return lstFuncionesBLs
    End Function

    Private Function BuscarMappers(filePath As String, fileCaller As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstMappers As New List(Of String)

        For Each linea As String In lstLineas
            Dim m As Match = Regex.Match(linea, "\bNew\s+(\w*Mapper)\b", RegexOptions.IgnoreCase)

            If m.Success Then
                If lstMappers.Find(Function(p As String) p = m.Groups(1).Value) = "" Then
                    lstMappers.Add(m.Groups(1).Value)
                End If
            End If
        Next

        Return lstMappers
    End Function

    Private Function BuscarFuncionesMapper(filePath As String, fileCaller As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstFuncionesMappers As New List(Of String)

        For Each linea As String In lstLineas
            If Not linea.Contains("'") And (linea.Contains("Function ") Or linea.Contains("Sub ")) Then
                Dim m As Match = Regex.Match(linea, "(\w+)\(")

                If m.Success Then
                    If lstFuncionesMappers.Find(Function(p As String) p = m.Groups(1).Value & "Mapper") = "" Then
                        Dim p As Match = Regex.Match(linea, "\((.*?)\)")
                        Dim retorno As String = ""

                        If linea.Contains("Function") Then
                            Dim q As Match = Regex.Match(linea, "\)\s+As\s+(\w+)")
                            retorno = q.Groups(1).Value
                        End If

                        lstFuncionesMappers.Add(m.Groups(1).Value & ";" & p.Groups(1).Value & ";" & retorno)
                    End If
                End If
            End If
        Next

        Return lstFuncionesMappers
    End Function

    Private Function BuscarScripts(filePath As String, fileCaller As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstScripts As New List(Of String)

        For Each linea As String In lstLineas
            Dim m As Match = Regex.Match(linea, "=\s*""([^""]+)""")


            If m.Success Then
                If lstScripts.Find(Function(p As String) p = m.Groups(1).Value) = "" Then
                    lstScripts.Add(m.Groups(1).Value)
                End If
            End If
        Next

        Return lstScripts
    End Function
End Class
