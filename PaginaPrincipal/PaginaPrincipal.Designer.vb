<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPaginaPrincipal
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmbPlataformas = New System.Windows.Forms.ComboBox()
        Me.tvEstructura = New System.Windows.Forms.TreeView()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(219, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 21)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cmbPlataformas
        '
        Me.cmbPlataformas.FormattingEnabled = True
        Me.cmbPlataformas.Location = New System.Drawing.Point(35, 12)
        Me.cmbPlataformas.Name = "cmbPlataformas"
        Me.cmbPlataformas.Size = New System.Drawing.Size(163, 21)
        Me.cmbPlataformas.TabIndex = 2
        '
        'tvEstructura
        '
        Me.tvEstructura.Location = New System.Drawing.Point(35, 39)
        Me.tvEstructura.Name = "tvEstructura"
        Me.tvEstructura.Size = New System.Drawing.Size(584, 535)
        Me.tvEstructura.TabIndex = 3
        '
        'frmPaginaPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(814, 597)
        Me.Controls.Add(Me.tvEstructura)
        Me.Controls.Add(Me.cmbPlataformas)
        Me.Controls.Add(Me.Button1)
        Me.Name = "frmPaginaPrincipal"
        Me.Text = "Hojete Rutero"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents cmbPlataformas As ComboBox
    Friend WithEvents tvEstructura As TreeView
End Class
