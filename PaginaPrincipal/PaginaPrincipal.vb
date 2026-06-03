Imports System.IO

Public Class frmPaginaPrincipal
    Private _dtArchivos As DataTable
    Private _dtBls As DataTable
    Private _dtMappers As DataTable
    Private _dtScripts As DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        dgvResultados.DataSource = BuscarTextoEnArchivosVB("C:\Users\Fernando\Dropbox\Fernando.viu\Programacion\RecFaecys2\Sistema\Actas", "BL.")
    End Sub

    Function BuscarTextoEnArchivosVB(carpeta As String, textoBuscado As String) As DataTable
        Dim idArchivos As Integer = 0
        Dim idBls As Integer = 0
        Dim idMappers As Integer = 0
        Dim lineas() As String
        Dim contenido As String

        _dtArchivos = New DataTable
        With _dtArchivos
            .Columns.Add("idArchivo", GetType(Integer))
            .Columns.Add("Nombre", GetType(String))
        End With

        Dim lstArchivos As New List(Of String)

        For Each archivo As String In Directory.GetFiles(carpeta, "*.vb", SearchOption.AllDirectories)

            contenido = File.ReadAllText(archivo)

            If contenido.Contains(textoBuscado) Then
                idArchivos += 1
                _dtArchivos.Rows.Add(idArchivos, archivo)
            End If
        Next

        If _dtArchivos.Rows.Count > 0 Then

            If _dtBls Is Nothing Then
                _dtBls = New DataTable
                With _dtBls
                    .Columns.Add("idBl", GetType(Integer))
                    .Columns.Add("idArchivo", GetType(Integer))
                    .Columns.Add("Nombre", GetType(String))
                End With
            Else
                _dtBls.Rows.Clear()
            End If

            For Each dr As DataRow In _dtArchivos.Rows
                Dim idArchivo As Integer = dr("idArchivo")
                Dim Archivo As String = dr("Nombre")

                lineas = File.ReadAllLines(Archivo)

                For Each linea As String In lineas
                    If linea.Contains(textoBuscado) Then
                        idBls += 1

                        _dtBls.Rows.Add(idBls, idArchivo, linea)
                    End If
                Next
            Next
        End If

        Return _dtBls

    End Function
End Class
