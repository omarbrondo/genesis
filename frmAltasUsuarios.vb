Imports System.Data.OleDb

Public Class frmAltasUsuarios

    Private Database1DataSet As DataSet

    Dim Privilegio As String

    Dim opcion As Integer

    Public Property Database1DataSet1 As DataSet
        Get
            Return Database1DataSet
        End Get
        Set(value As DataSet)
            Database1DataSet = value
        End Set
    End Property

    Private Sub ConexionDataBase1()

        Using conexion = ConnectionFactory.CreateConnection()

            conexion.Open()

            Dim UsuariosTableAdapter = New OleDbDataAdapter()

            UsuariosTableAdapter = New OleDbDataAdapter With {
                .SelectCommand = New OleDbCommand("SELECT * FROM Usuarios", conexion)
            }

            Database1DataSet1 = New DataSet

            Database1DataSet1.Tables.Add("Usuarios")

            UsuariosTableAdapter.Fill(Database1DataSet1.Tables("Usuarios"))

            dtgUsuarios.DataSource = Database1DataSet1.Tables("Usuarios")

            If chkContraseña.Checked = True Then

                Me.dtgUsuarios.Columns("Contraseña").Visible = True

            Else

                dtgUsuarios.Columns("Contraseña").Visible = False

            End If

            conexion.Close()

            OleDbConnection.ReleaseObjectPool()

        End Using

    End Sub

    Private Sub btnSalir_Click_1(sender As Object, e As EventArgs) Handles btnSalir.Click

        frmTareaDia.Show()

        Me.Close()

    End Sub

    Private Sub AltaUsuario()

        'DAR DE ALTA A USUARIOS EN LA TABLA USUARIOS

        Dim NuevoUsuario As DataRow

        If Not Privilegio = String.Empty Then

            NuevoUsuario = Database1DataSet1.Tables("Usuarios").NewRow

            NuevoUsuario = NewMethod1(NuevoUsuario)

            Database1DataSet1.Tables("Usuarios").Rows.Add(NuevoUsuario)

            Using conexion = ConnectionFactory.CreateConnection()

                conexion.Open()

                Dim UsuariosTableAdapter = New OleDbDataAdapter()

                UsuariosTableAdapter.InsertCommand = New OleDbCommand() With {
                    .CommandText = "INSERT INTO Usuarios (Nombre, Usuario, Contraseña, Privilegio ) VALUES (@Nombre, @Usuario, @Contraseña, @Privilegio)",
                    .Connection = conexion
                }

                NewMethod(UsuariosTableAdapter)

                conexion.Close()

                OleDbConnection.ReleaseObjectPool()

            End Using

            txtNombreApellido.Text = String.Empty

            txtUsuario.Text = String.Empty

            mtbContrasenia.Text = String.Empty

            txtNombreApellido.Focus()

        Else

            MessageBox.Show("DEBE SELECCIONAR UN PRIVILEGIO", "ERROR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)

        End If

    End Sub

    Private Function NewMethod1(NuevoUsuario As DataRow) As DataRow
        NuevoUsuario("Nombre") = txtNombreApellido.Text

        NuevoUsuario("Usuario") = txtUsuario.Text

        NuevoUsuario("Contraseña") = mtbContrasenia.Text

        NuevoUsuario("Privilegio") = Privilegio
        Return NuevoUsuario
    End Function

    Private Sub NewMethod(UsuariosTableAdapter As OleDbDataAdapter)
        UsuariosTableAdapter.InsertCommand.Parameters.Add("@Nombre", OleDbType.VarChar, 128, "Nombre")

        UsuariosTableAdapter.InsertCommand.Parameters.Add("@Usuario", OleDbType.VarChar, 128, "Usuario")

        UsuariosTableAdapter.InsertCommand.Parameters.Add("@Contraseña", OleDbType.VarChar, 128, "Contraseña")

        UsuariosTableAdapter.InsertCommand.Parameters.Add("@Privilegio", OleDbType.VarChar, 128, "Privilegio")

        UsuariosTableAdapter.Update(Database1DataSet1.Tables("Usuarios"))
    End Sub

    Private Sub ModificacionUsuario()


        If Not Privilegio = String.Empty Then

            Using conexion = ConnectionFactory.CreateConnection()

                Dim cmd As New OleDbCommand With {
                    .CommandType = CommandType.Text,
                    .CommandText = "UPDATE Usuarios SET Nombre= '" + txtNombreApellido.Text + "' WHERE idUsuario=" + txtID.Text,
                    .Connection = conexion
                }

                conexion.Open()

                cmd.ExecuteNonQuery()

            End Using

            Using conexion = ConnectionFactory.CreateConnection()

                Dim cmd As New OleDbCommand With {
                    .CommandType = CommandType.Text,
                    .CommandText = "UPDATE Usuarios SET Usuario= '" + txtUsuario.Text + "' WHERE idUsuario=" + txtID.Text,
                    .Connection = conexion
                }

                conexion.Open()

                cmd.ExecuteNonQuery()

            End Using

            Using conexion = ConnectionFactory.CreateConnection()

                Dim cmd As New OleDbCommand With {
                    .CommandType = CommandType.Text,
                    .CommandText = "UPDATE Usuarios Set Contraseña= '" + mtbContrasenia.Text + "' WHERE idUsuario=" + txtID.Text,
                    .Connection = conexion
                }

                conexion.Open()

                cmd.ExecuteNonQuery()

            End Using

            Using conexion = ConnectionFactory.CreateConnection()

                Dim cmd As New OleDbCommand With {
                    .CommandType = CommandType.Text,
                    .CommandText = "UPDATE Usuarios SET Privilegio= '" + Privilegio + "' WHERE idUsuario=" + txtID.Text,
                    .Connection = conexion
                }

                conexion.Open()

                cmd.ExecuteNonQuery()

            End Using

        Else

            MessageBox.Show("DEBE SELECCIONAR UN PRIVILEGIO", "ERROR",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)

        End If

    End Sub

    Private Sub EliminarUsuario()

        Using conexion = ConnectionFactory.CreateConnection()

            Dim cmd As New OleDbCommand With {
                .CommandType = CommandType.Text,
                .CommandText = "DELETE FROM Usuarios WHERE Nombre= '" + txtNombreApellido.Text + "'",
                .Connection = conexion
            }

            conexion.Open()

            cmd.ExecuteNonQuery()

        End Using

    End Sub
    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click

        Select Case opcion

            Case 1

                AltaUsuario()

                ConexionDataBase1()

            Case 2

                ModificacionUsuario()

                ConexionDataBase1()

            Case 3

                EliminarUsuario()

                ConexionDataBase1()

                txtID.Text = String.Empty

                txtNombreApellido.Text = String.Empty

                txtUsuario.Text = String.Empty

                mtbContrasenia.Text = String.Empty

                ConexionDataBase1()

                pnlOpciones.Enabled = False

            Case Else

                MessageBox.Show("DEBE SELECCIONAR UNA OPCION")

        End Select

    End Sub

    Private Sub frmAltasUsuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ConexionDataBase1()

    End Sub

    Private Sub btnNuevoUsuario_Click(sender As Object, e As EventArgs) Handles btnNuevoUsuario.Click

        opcion = 1

        txtID.Text = String.Empty

        txtNombreApellido.Text = String.Empty

        txtUsuario.Text = String.Empty

        mtbContrasenia.Text = String.Empty

        ConexionDataBase1()

        pnlOpciones.Enabled = True

    End Sub

    Private Sub PermisosUsuarios()

        'ADMINISTRADOR

        If rbAdministrador.Checked = True Then

            Privilegio = "Administrador"

        End If

        'SUPERVISOR

        If rbSupervisor.Checked = True Then

            Privilegio = "Supervisor"

        End If

        'CONTROLADOR

        If rbControlador.Checked = True Then

            Privilegio = "Controlador"

        End If

        'OPERADOR

        If rbOperador.Checked = True Then

            Privilegio = "Operador"

        End If

    End Sub

    Private Sub rbAdministrador_CheckedChanged(sender As Object, e As EventArgs) Handles rbAdministrador.CheckedChanged

        PermisosUsuarios()

    End Sub

    Private Sub rbSupervisor_CheckedChanged(sender As Object, e As EventArgs) Handles rbSupervisor.CheckedChanged

        PermisosUsuarios()

    End Sub

    Private Sub rbControlador_CheckedChanged(sender As Object, e As EventArgs) Handles rbControlador.CheckedChanged

        PermisosUsuarios()

    End Sub

    Private Sub rbOperador_CheckedChanged(sender As Object, e As EventArgs) Handles rbOperador.CheckedChanged

        PermisosUsuarios()

    End Sub

    Private Sub btnModificarEliminar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click

        opcion = 2

        txtID.Text = String.Empty

        txtNombreApellido.Text = String.Empty

        txtUsuario.Text = String.Empty

        mtbContrasenia.Text = String.Empty

        ConexionDataBase1()

        pnlOpciones.Enabled = True

    End Sub
    Private Sub dtgUsuarios_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtgUsuarios.CellClick

        Try

            txtID.Text = dtgUsuarios.Rows(e.RowIndex).Cells(0).Value()

            txtNombreApellido.Text = Me.dtgUsuarios.Rows(e.RowIndex).Cells(1).Value()

            txtUsuario.Text = Me.dtgUsuarios.Rows(e.RowIndex).Cells(2).Value()

            mtbContrasenia.Text = Me.dtgUsuarios.Rows(e.RowIndex).Cells(3).Value()

        Catch ex As Exception

        End Try

    End Sub
    Private Sub txtNombreApellido_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNombreApellido.KeyDown

        If e.KeyCode = Keys.Enter Then

            txtUsuario.Focus()

        End If

    End Sub

    Private Sub txtUsuario_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUsuario.KeyDown

        If e.KeyCode = Keys.Enter Then

            mtbContrasenia.Focus()

        End If

    End Sub

    Private Sub mtbContrasenia_KeyDown(sender As Object, e As KeyEventArgs) Handles mtbContrasenia.KeyDown

        If e.KeyCode = Keys.Enter Then

            rbOperador.Focus()

        End If

    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

        opcion = 3

        txtID.Text = String.Empty

        txtNombreApellido.Text = String.Empty

        txtUsuario.Text = String.Empty

        mtbContrasenia.Text = String.Empty

        ConexionDataBase1()

        pnlOpciones.Enabled = True

    End Sub

End Class