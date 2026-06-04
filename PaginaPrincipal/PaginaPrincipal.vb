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

    Private Sub frmPaginaPrincipal_Load(sender As Object, e As EventArgs) Handles Me.Load

        _dtPlataformas = New DataTable
        With _dtPlataformas
            .Columns.Add("idPlataforma", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
            .Columns.Add("Ruta", GetType(String))

            .Rows.Add(1, "Faecys2", "C:\Users\Fernando\Dropbox\Fernando.viu\Programacion\RecFaecys2\Sistema")
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
        lstDirectorios = Directory.GetDirectories(_rutaPlataforma).ToList

        'Cicla entre todos los directorios de la raiz de la plataforma.
        For Each rutaDirectorio As String In lstDirectorios
            Dim direcotorioInfo As DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(rutaDirectorio)

            If _lstExcluidos.Find(Function(p As String) p = direcotorioInfo.Name) = "" Then
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
                        lstBLs = BuscarBLs(rutaArchivo)

                        If lstBLs IsNot Nothing AndAlso lstBLs.Count > 0 Then
                            For Each archivoBL As String In lstBLs
                                Dim rutaBL As String = _rutaPlataforma & "\BL\" & archivoBL & ".vb"
                                idBL += 1

                                _dtBls.Rows.Add(idBL, idArchivo, archivoBL, rutaBL)

                                lstFuncionesBL = BuscarFuncionesBL(rutaBL)
                                lstMappers = BuscarMappers(rutaBL)

                                If lstFuncionesBL IsNot Nothing AndAlso lstFuncionesBL.Count > 0 Then
                                    For Each funcionBL As String In lstFuncionesBL
                                        Dim parametros() As String = funcionBL.Split(";")

                                        idFuncionBL += 1
                                        _dtFuncionesBL.Rows.Add(idFuncionBL, idBL, parametros(0), parametros(1), parametros(2))
                                    Next
                                End If

                                If lstMappers IsNot Nothing AndAlso lstMappers.Count > 0 Then
                                    For Each mapper In lstMappers
                                        Dim rutaMapper As String = _rutaPlataforma & "\Mapper\" & mapper & ".vb"

                                        idMapper += 1

                                        _dtMappers.Rows.Add(idMapper, idBL, Nothing, mapper, rutaMapper)

                                        lstFuncionesMapper = BuscarFuncionesMapper(rutaMapper)
                                        lstScripts = BuscarScripts(rutaMapper)

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

                                                _dtScripts.Rows.Add(idScript, idMapper, Nothing, lstScripts)
                                            Next
                                        End If
                                    Next
                                End If
                            Next

                        End If
                    Next

                    'Dim nodoRaiz As New TreeNode
                    'idArchivo += 1
                    'nodoRaiz.Text = archivo.Name
                    'tvEstructura.Nodes.Add(nodoRaiz)
                End If
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

    Private Function BuscarBLs(filePath As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstBLs As New List(Of String)

        For Each linea As String In lstLineas
            Dim m As Match = Regex.Match(linea, "(\w+)BL\.")

            If m.Success Then
                If lstBLs.Find(Function(p As String) p = m.Groups(1).Value & "BL") = "" Then

                    lstBLs.Add(m.Groups(1).Value & "BL")
                End If
            End If
        Next

        Return lstBLs
    End Function

    Private Function BuscarFuncionesBL(filePath As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstFuncionesBLs As New List(Of String)

        For Each linea As String In lstLineas
            If linea.Contains("Function") Or linea.Contains("Sub") Then
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

    Private Function BuscarMappers(filePath As String) As List(Of String)
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

    Private Function BuscarFuncionesMapper(filePath As String) As List(Of String)
        Dim lstLineas As List(Of String) = File.ReadAllLines(filePath).ToList
        Dim lstFuncionesMappers As New List(Of String)

        For Each linea As String In lstLineas
            If linea.Contains("Function") Or linea.Contains("Sub") Then
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

    Private Function BuscarScripts(filePath As String) As List(Of String)
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
