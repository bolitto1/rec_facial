Imports System

Imports System.IO

Imports System.Configuration
Imports System.Data.SqlClient


Imports System.Data.SqlTypes


Public Module variosvb
	Public con As SqlConnection
	Public cmd As SqlCommand
	Public da As SqlDataAdapter

	Public Function probar_con() As Boolean
		Dim cadena As String
		Dim estado As Boolean
		estado = False
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
			estado = True
		Catch ex As Exception
			estado = False

		End Try
		Return estado

	End Function

	Public Function isnumericvb(ByVal valor As Object) As Boolean
		If IsNumeric(valor) Then
			Return True
		Else
			Return False
		End If
	End Function
	Public Function instrvb(ByVal texto As String, ByVal buscado As String) As Integer
		Dim pos As Integer
		pos = InStr(texto, buscado)
		Return pos
	End Function
	Public Function midvb(ByVal texto As String, ByVal ini As Integer, ByVal cuantos As Integer) As String
		Dim cad As String

		cad = Mid(texto, ini, cuantos)
		Return cad
	End Function

	Public Function insertar_empleado(
		 em_nomlar As String,
		   em_fecha_nacimiento As DateTime,
		   ced As String,
		   em_cargo As String,
		   em_estado_civil As String,
		   em_sexo As String,
		   em_direccion As String,
		   em_telefono As String,
		   em_fecha_contratacion As DateTime,
		   em_comentario As String,
		   em_usuario_crea As String,
		   em_sueldo As Double,
		   em_sobretiempo As String,
		   em_cargas As String,
		   em_estado As String,
		   em_sueldo_extra As Double,
		   em_des_seccion As String,
		   em_pago_fdo_reserva As String) As Int32

		Dim dato As String
		Dim cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection
		dato = ""
		con.ConnectionString = cadena
		Try
			con.Open()
			cmd = New SqlCommand("empleado_ins", con)
			cmd.CommandType = CommandType.StoredProcedure
			cmd.Parameters.Add("@em_nomlar", SqlDbType.VarChar)
			cmd.Parameters("@em_nomlar").Value = em_nomlar

			cmd.Parameters.Add("@em_fecha_nacimiento", SqlDbType.DateTime)
			cmd.Parameters("@em_fecha_nacimiento").Value = em_fecha_nacimiento

			cmd.Parameters.Add("@ced", SqlDbType.VarChar)
			cmd.Parameters("@ced").Value = ced

			cmd.Parameters.Add("@em_cargo", SqlDbType.VarChar)
			cmd.Parameters("@em_cargo").Value = em_cargo

			cmd.Parameters.Add("@em_estado_civil", SqlDbType.VarChar)
			cmd.Parameters("@em_estado_civil").Value = em_estado_civil

			cmd.Parameters.Add("@em_sexo", SqlDbType.VarChar)
			cmd.Parameters("@em_sexo").Value = em_sexo

			cmd.Parameters.Add("@em_direccion", SqlDbType.VarChar)
			cmd.Parameters("@em_direccion").Value = em_direccion

			cmd.Parameters.Add("@em_telefono", SqlDbType.VarChar)
			cmd.Parameters("@em_telefono").Value = em_telefono

			cmd.Parameters.Add("@em_fecha_contratacion", SqlDbType.DateTime)
			cmd.Parameters("@em_fecha_contratacion").Value = em_fecha_contratacion

			cmd.Parameters.Add("@em_comentario", SqlDbType.VarChar)
			cmd.Parameters("@em_comentario").Value = em_comentario

			cmd.Parameters.Add("@em_usuario_crea", SqlDbType.VarChar)
			cmd.Parameters("@em_usuario_crea").Value = em_usuario_crea

			cmd.Parameters.Add("@em_sueldo", SqlDbType.Money)
			cmd.Parameters("@em_sueldo").Value = em_sueldo

			cmd.Parameters.Add("@em_sobretiempo", SqlDbType.VarChar)
			cmd.Parameters("@em_sobretiempo").Value = em_sobretiempo

			cmd.Parameters.Add("@em_cargas", SqlDbType.VarChar)
			cmd.Parameters("@em_cargas").Value = em_cargas

			cmd.Parameters.Add("@em_estado", SqlDbType.VarChar)
			cmd.Parameters("@em_estado").Value = em_estado

			cmd.Parameters.Add("@em_sueldo_extra", SqlDbType.Money)
			cmd.Parameters("@em_sueldo_extra").Value = em_sueldo_extra

			cmd.Parameters.Add("@em_des_seccion", SqlDbType.VarChar)
			cmd.Parameters("@em_des_seccion").Value = em_des_seccion

			cmd.Parameters.Add("@em_pago_fdo_reserva", SqlDbType.VarChar)
			cmd.Parameters("@em_pago_fdo_reserva").Value = em_pago_fdo_reserva


			cmd.ExecuteNonQuery()

			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select @@identity"
			Dim id As Int32
			id = Convert.ToInt32(cmd.ExecuteScalar())
			Return id

		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try
		con.Close()

	End Function

	Public Function ins_rec_lote(
								inicio As DateTime,
		   fin As DateTime,
		   nombre_archivo As String,
		   cedula As String,
		fecha As DateTime,
		   comentario As String,
		   usuario As String,
		   metodo As String,
		   distancia As Double,
		   milisegundos As Double, num As String, id_p As Int32,
			mili_recon As Double
		   ) As Int32
		Dim cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Console.Write(ex.Message)
			Return False
		End Try

		Try
			cmd = New SqlCommand("sp_insertar_reconocimiento_lotes", con)
			cmd.CommandType = CommandType.StoredProcedure

			cmd.Parameters.AddWithValue("@inicio", inicio)
			cmd.Parameters.AddWithValue("@fin", fin)
			cmd.Parameters.AddWithValue("@nombre_archivo", nombre_archivo)
			cmd.Parameters.AddWithValue("@cedula", cedula)
			cmd.Parameters.AddWithValue("@comentario", comentario)
			cmd.Parameters.AddWithValue("@distancia", distancia)
			cmd.Parameters.AddWithValue("@milisegundos", milisegundos)
			cmd.Parameters.AddWithValue("@num", num)
			cmd.Parameters.AddWithValue("@id_p", id_p)
			cmd.Parameters.AddWithValue("@mili_recon", mili_recon)


			cmd.ExecuteNonQuery()

			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select @@identity"
			Dim id As Int32
			id = Convert.ToInt32(cmd.ExecuteScalar())

			Return id

		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Console.Write(ex.Message)
			Return -1
		End Try
		Return True
	End Function

	Public Function ins_nueva_rec_lotr(metodo As String, fecha As DateTime) As Int32
		Dim cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Console.Write(ex.Message)
			Return False
		End Try

		Try
			cmd = New SqlCommand("ins_recono", con)
			cmd.CommandType = CommandType.StoredProcedure
			cmd.Parameters.AddWithValue("@fecha", fecha)
			cmd.Parameters.AddWithValue("@metodo", metodo)

			cmd.ExecuteNonQuery()

			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select @@identity"
			Dim id As Int32
			id = Convert.ToInt32(cmd.ExecuteScalar())


			Return id

		Catch ex As Exception
			Console.Write(ex.Message)
			Return -1
		End Try
		Return True
	End Function


	Public Function ins_nueva(cedula As String) As Int32
		Dim cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Console.Write(ex.Message)
			Return False
		End Try

		Try
			cmd = New SqlCommand("sp_inserta_pos", con)
			cmd.CommandType = CommandType.StoredProcedure
			cmd.Parameters.Add("@ced", SqlDbType.VarChar)
			cmd.Parameters("@ced").Value = cedula
			cmd.ExecuteNonQuery()

			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select @@identity"
			Dim id As Int32
			id = Convert.ToInt32(cmd.ExecuteScalar())

			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select pos from tmp_fotos where id=" + id.ToString()


			id = Convert.ToInt32(cmd.ExecuteScalar())

			Return id

		Catch ex As Exception
			Console.Write(ex.Message)
			Return -1
		End Try
		Return True
	End Function

	Public Function reg_asistencia(cedula As String, fotos_en_bd As Integer, fecha As DateTime, hora As DateTime) As String
		Dim fe, fa, cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Console.Write(ex.Message)
		End Try

		fe = CDate(fecha).ToString("yyyy-MM-dd 00:00:00")
		fa = CDate(hora).ToString("yyyy-MM-dd HH:mm:ss.fff")

		Try

			cmd = New SqlCommand("sp_insertar_asistencia", con)
			cmd.CommandType = CommandType.StoredProcedure
			cmd.Parameters.Add("@marcacion", SqlDbType.DateTime)
			cmd.Parameters("@marcacion").Value = CDate(fa)

			cmd.Parameters.Add("@cedula", SqlDbType.VarChar)
			cmd.Parameters("@cedula").Value = cedula

			cmd.Parameters.Add("@fecha", SqlDbType.DateTime)
			cmd.Parameters("@fecha").Value = CDate(fe)

			cmd.Parameters.Add("@fotos_en_bd", SqlDbType.Int)
			cmd.Parameters("@fotos_en_bd").Value = fotos_en_bd

			cmd.Parameters.Add("@comentario", SqlDbType.VarChar)
			cmd.Parameters("@comentario").Value = ""

			cmd.CommandType = CommandType.StoredProcedure

			cmd.ExecuteNonQuery()


			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select @@identity"
			Dim id As Int64
			id = Convert.ToInt64(cmd.ExecuteScalar())
			Return id.ToString()

		Catch ex As Exception
			MessageBox.Show(ex.Message)

		End Try

	End Function

	Public Sub actualizar_pos(pos As Int32, id As Long)
		Dim cad As String
		cad = "UPDATE  tmp_fotos SET  "
		cad = cad + "pos= " + pos.ToString
		cad = cad + "WHERE id =" + id.ToString()


		Dim cadena, fe As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return
		End Try

		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cad
		Try
			cmd.ExecuteNonQuery()
			con.Close()
			Return
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return
		End Try


	End Sub

	Public Sub actualizar_asis(comentario As String, id As String)
		Dim cad As String
		cad = "UPDATE [dbo].[asistencia]  SET  "
		cad = cad + " [comentario] = '" + comentario + "' "
		cad = cad + "WHERE id =" + id
		Dim cadena, fe As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return
		End Try

		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cad
		Try
			cmd.ExecuteNonQuery()
			Return
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return
		End Try


	End Sub

	Public Function get_asis(fecha As String) As DataTable
		Dim cad As String
		Dim dt As DataTable = New DataTable

		cad = "Select  asistencia.id, asistencia.marcacion, asistencia.cedula, "
		cad = cad + " asistencia.fecha, asistencia.fotos_en_bd, asistencia.comentario,"
		cad = cad + " asistencia.mini, empleados.em_nomlar From asistencia INNER Join "
		cad = cad + " empleados On asistencia.cedula = empleados.ced "
		cad = cad + " Where (asistencia.fecha = CONVERT(DATETIME, '" + fecha + "', 102)) Order By asistencia.id"
		Dim cadena, fe As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return dt
		End Try
		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cad
		Dim da As SqlDataAdapter
		da = New SqlDataAdapter(cmd)
		Try
			da.Fill(dt)
			Return dt
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return dt
		End Try
	End Function


	Public Function get_asis50() As DataTable
		Dim cad As String
		Dim dt As DataTable = New DataTable

		cad = "Select top 50 empleados.em_nomlar, asistencia.cedula, asistencia.marcacion, "
		cad = cad + " asistencia.fecha, asistencia.fotos_en_bd, asistencia.comentario,"
		cad = cad + " asistencia.mini,   asistencia.id From asistencia INNER Join "
		cad = cad + " empleados On asistencia.cedula = empleados.ced "
		cad = cad + "  Order By asistencia.id desc"
		Dim cadena, fe As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return dt
		End Try
		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cad
		Dim da As SqlDataAdapter
		da = New SqlDataAdapter(cmd)
		Try
			da.Fill(dt)
			Return dt
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return dt
		End Try
	End Function

	Public Function get_asis_falsos(fecha As String) As DataTable
		Dim cad As String
		Dim dt As DataTable = New DataTable

		cad = "Select  asistencia.id, asistencia.marcacion, asistencia.cedula, "
		cad = cad + " asistencia.fecha, asistencia.fotos_en_bd, asistencia.comentario,"
		cad = cad + " asistencia.mini, empleados.em_nomlar From asistencia INNER Join "
		cad = cad + " empleados On asistencia.cedula = empleados.ced "
		cad = cad + " Where (asistencia.fecha = CONVERT(DATETIME, '" + fecha + "', 102)) "
		cad = cad + " AND (NOT (dbo.asistencia.mini IS NULL)) Order By asistencia.id"
		Dim cadena, fe As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return dt
		End Try
		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cad
		Dim da As SqlDataAdapter
		da = New SqlDataAdapter(cmd)
		Try
			da.Fill(dt)
			Return dt
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return dt
		End Try
	End Function


	Public Function get_asis_falsos_todos() As DataTable
		Dim cad As String
		Dim dt As DataTable = New DataTable

		cad = "Select  asistencia.id, asistencia.marcacion, asistencia.cedula, "
		cad = cad + " asistencia.fecha, asistencia.fotos_en_bd, asistencia.comentario,"
		cad = cad + " asistencia.mini, empleados.em_nomlar From asistencia INNER Join "
		cad = cad + " empleados On asistencia.cedula = empleados.ced "
		cad = cad + " WHERE  (NOT (dbo.asistencia.mini IS NULL)) Order By asistencia.id DESC"
		Dim cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return dt
		End Try
		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cad
		Dim da As SqlDataAdapter
		da = New SqlDataAdapter(cmd)
		Try
			da.Fill(dt)
			Return dt
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return dt
		End Try
	End Function

	Public Function actualizar_fe_sal(ByVal valor As String, ByVal id As String) As Boolean
		Dim cadena, fe As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection

		con.ConnectionString = cadena
		Try
			con.Open()
		Catch ex As Exception
			Return False
		End Try

		If (valor = "NULO") Then
			cadena = "UPDATE [dbo].[empleados] Set  [em_fecha_sal] = null WHERE id=" + id
		Else

			fe = CDate(valor).ToString("yyyy-MM-dd")
			cadena = "UPDATE [dbo].[empleados] "
			cadena = cadena + " Set  [em_fecha_sal] =  CONVERT(DATETIME, '" + fe + " 00:00:00', 102) "
			cadena = cadena + " WHERE id=" + id
		End If
		cmd = New SqlCommand
		cmd.Connection = con
		cmd.CommandType = CommandType.Text
		cmd.CommandText = cadena
		Try
			cmd.ExecuteNonQuery()
			Return True
		Catch ex As Exception
			MessageBox.Show(ex.Message)
			Return False
		End Try



	End Function


	Public Function getId() As String

		Dim ds1 As New DataSet
		Dim dato As String
		Dim cadena As String
		cadena = ConfigurationManager.ConnectionStrings("Detector_facial.Properties.Settings.tesisConnectionString").ConnectionString
		con = New SqlConnection
		dato = ""
		con.ConnectionString = cadena
		Try
			con.Open()
			cmd = New SqlCommand
			cmd.Connection = con
			cmd.CommandType = CommandType.Text
			cmd.CommandText = "Select @@identity"
			Dim id As Int64
			id = Convert.ToInt32(cmd.ExecuteScalar())
			Return id.ToString()

		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


		Return dato
	End Function

	'...
End Module