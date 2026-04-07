using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace Persal._003_Physical_Counting_2
{
    /// <summary>
    /// Repositorio de datos para el Wizard de inventario físico cíclico.
    /// Centraliza todas las operaciones de base de datos de PhysicalCountCyclicForm.
    /// </summary>
    public class PhysicalCountRepository
    {
        private readonly string _connectionString;

        public PhysicalCountRepository(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
                throw new ArgumentException("La cadena de conexión no puede estar vacía.", nameof(connectionString));
            _connectionString = connectionString;
        }

        // =====================================================================
        // PASO 1 — ENCABEZADO DE INVENTARIO
        // =====================================================================
        #region PASO 1 — ENCABEZADO DE INVENTARIO

        /// <summary>
        /// Busca un responsable por número de empleado, login o nombre parcial.
        /// Equivalente al método SearchResponsibleByText del formulario.
        /// </summary>
        /// <param name="valor">Texto a buscar (número de empleado, login o parte del nombre).</param>
        /// <returns>Información del responsable encontrado, o null si no existe.</returns>
        public ResponsibleInfo BuscarResponsablePorTexto(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return null;

            const string sql = @"
                SELECT TOP 1
                    User_ID,
                    Employee_Number,
                    Name,
                    Job_Position
                FROM dbo.Users
                WHERE Active = 'YES'
                  AND (
                        Employee_Number = @p1
                     OR Login           = @p1
                     OR Name LIKE @p2
                  )";

            DataTable dt = EjecutarConsulta(sql, valor, "%" + valor + "%");

            if (dt.Rows.Count == 0)
                return null;

            DataRow fila = dt.Rows[0];
            return new ResponsibleInfo
            {
                UserId         = Convert.ToInt32(fila["User_ID"]),
                EmployeeNumber = fila["Employee_Number"]?.ToString(),
                Name           = fila["Name"]?.ToString(),
                JobPosition    = fila["Job_Position"]?.ToString()
            };
        }

        /// <summary>
        /// Obtiene la información de un responsable a partir de su ID de usuario.
        /// Equivalente al método LoadResponsibleByUserId del formulario.
        /// </summary>
        /// <param name="userId">Identificador único del usuario.</param>
        /// <returns>Información del responsable, o null si no existe o está inactivo.</returns>
        public ResponsibleInfo ObtenerResponsablePorId(int userId)
        {
            const string sql = @"
                SELECT
                    User_ID,
                    Employee_Number,
                    Name,
                    Job_Position
                FROM dbo.Users
                WHERE User_ID = @p1
                  AND Active  = 'YES'";

            DataTable dt = EjecutarConsulta(sql, userId);

            if (dt.Rows.Count == 0)
                return null;

            DataRow fila = dt.Rows[0];
            return new ResponsibleInfo
            {
                UserId         = Convert.ToInt32(fila["User_ID"]),
                EmployeeNumber = fila["Employee_Number"]?.ToString(),
                Name           = fila["Name"]?.ToString(),
                JobPosition    = fila["Job_Position"]?.ToString()
            };
        }

        /// <summary>
        /// Obtiene los tipos de lista de conteo disponibles.
        /// Equivalente al método LoadPhysicalCountListTypes del formulario.
        /// </summary>
        /// <returns>DataTable con los tipos de lista de conteo físico.</returns>
        public DataTable ObtenerTiposListaConteo()
        {
            const string sql = @"
                SELECT
                    Physical_Count_List_Type_ID,
                    Physical_Count_List_Type
                FROM dbo.Physical_Count_List_Types
                WHERE Physical_Count_List_Type_ID = '1'";

            return EjecutarConsulta(sql);
        }

        /// <summary>
        /// Obtiene el ID del almacén predeterminado para conteo físico.
        /// Delega en la función escalar dbo.Return_Default_Warehouse_Location_ID_For_Count().
        /// </summary>
        /// <returns>ID del almacén predeterminado.</returns>
        public int ObtenerAlmacenPredeterminado()
        {
            // Nota: delega en la función escalar dbo.Return_Default_Warehouse_Location_ID_For_Count().
            const string sql = @"
                SELECT dbo.Return_Default_Warehouse_Location_ID_For_Count()";

            return EjecutarEscalarInt(sql);
        }

        #endregion

        // =====================================================================
        // PASO 2 — SELECCIÓN DE UBICACIONES
        // =====================================================================
        #region PASO 2 — SELECCIÓN DE UBICACIONES

        /// <summary>
        /// Obtiene la lista de almacenes disponibles, excluyendo el almacén fijo de referencia.
        /// Equivalente al método LoadAlmacenes del formulario.
        /// </summary>
        /// <param name="almacenFijoId">ID del almacén fijo que se excluye del resultado.</param>
        /// <returns>DataTable con columnas Location_ID, Location_Name, Selected.</returns>
        public DataTable ObtenerAlmacenes(int almacenFijoId)
        {
            const string sql = @"
                SELECT
                    Location_ID,
                    Location_Name,
                    CAST(0 AS BIT) AS Selected
                FROM dbo.Locations
                WHERE Location_Level = 1
                  AND Active         = 'YES'
                  AND Location_ID   <> @p1
                ORDER BY Location_Name";

            return EjecutarConsulta(sql, almacenFijoId);
        }

        /// <summary>
        /// Obtiene las áreas asociadas a un almacén específico.
        /// Equivalente al método LoadAreasForFocusedAlmacen del formulario.
        /// </summary>
        /// <param name="almacenId">ID del almacén padre.</param>
        /// <returns>DataTable con columnas Location_ID, Location_Name, Selected.</returns>
        public DataTable ObtenerAreasPorAlmacen(int almacenId)
        {
            const string sql = @"
                SELECT
                    Location_ID,
                    Location_Name,
                    CAST(0 AS BIT) AS Selected
                FROM dbo.Locations
                WHERE Parent_Location_ID = @p1
                  AND Active             = 'YES'
                ORDER BY Location_Name";

            return EjecutarConsulta(sql, almacenId);
        }

        /// <summary>
        /// Obtiene las sub-áreas asociadas a un área específica.
        /// Equivalente al método LoadSubAreasForFocusedArea del formulario.
        /// </summary>
        /// <param name="areaId">ID del área padre.</param>
        /// <returns>DataTable con columnas Location_ID, Location_Name.</returns>
        public DataTable ObtenerSubAreasPorArea(int areaId)
        {
            const string sql = @"
                SELECT
                    Location_ID,
                    Location_Name
                FROM dbo.Locations
                WHERE Parent_Location_ID = @p1
                  AND Active             = 'YES'
                ORDER BY Location_Name";

            return EjecutarConsulta(sql, areaId);
        }

        /// <summary>
        /// Obtiene las sub-áreas correspondientes a un conjunto de áreas seleccionadas.
        /// Equivalente al método LoadSubAreasForCheckedAreas del formulario.
        /// Construye un IN parametrizado para evitar concatenación de SQL.
        /// </summary>
        /// <param name="areaIds">Colección de IDs de áreas.</param>
        /// <returns>DataTable con las sub-áreas de todas las áreas indicadas.</returns>
        public DataTable ObtenerSubAreasPorAreas(IEnumerable<int> areaIds)
        {
            if (areaIds == null)
                throw new ArgumentNullException(nameof(areaIds));

            var listaIds = new List<int>(areaIds);

            if (listaIds.Count == 0)
                return new DataTable();

            // Construcción parametrizada del IN
            var parametros = new List<SqlParameter>();
            var placeholders = new List<string>();

            for (int i = 0; i < listaIds.Count; i++)
            {
                string parameterName = "@area" + i;
                placeholders.Add(parameterName);
                parametros.Add(new SqlParameter(parameterName, listaIds[i]));
            }

            string sql = $@"
                SELECT
                    Location_ID,
                    Location_Name
                FROM dbo.Locations
                WHERE Parent_Location_ID IN ({string.Join(", ", placeholders)})
                  AND Active = 'YES'
                ORDER BY Location_Name";

            return EjecutarConsultaConParametros(sql, parametros);
        }

        /// <summary>
        /// Cuenta la cantidad de hijos directos que tiene una ubicación.
        /// Equivalente a la query inline de RebuildEffectiveSelection del formulario.
        /// </summary>
        /// <param name="locationId">ID de la ubicación padre.</param>
        /// <returns>Número de ubicaciones hijas activas.</returns>
        public int ContarHijosDeUbicacion(int locationId)
        {
            const string sql = @"
                SELECT COUNT(1)
                FROM dbo.Locations
                WHERE Parent_Location_ID = @p1
                  AND Active             = 'YES'";

            return EjecutarEscalarInt(sql, locationId);
        }

        /// <summary>
        /// Guarda una fila de ubicación seleccionada para el inventario físico.
        /// Equivalente al método SaveSelectedLocationRow del formulario.
        /// Invoca el stored procedure dbo.sp_PC_SaveSelectedLocationRow.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <param name="selectedId">ID de la ubicación seleccionada por el usuario.</param>
        /// <param name="selectedLevel">Nivel jerárquico de la ubicación seleccionada.</param>
        /// <param name="effectiveId">ID de la ubicación efectiva para el conteo.</param>
        /// <param name="effectiveLevel">Nivel jerárquico de la ubicación efectiva.</param>
        /// <param name="parentSelectedId">ID de la ubicación padre seleccionada (null si no aplica).</param>
        /// <param name="userId">ID del usuario que realiza la operación.</param>
        public void GuardarUbicacionSeleccionada(
            int  physicalCountId,
            int  selectedId,
            int  selectedLevel,
            int  effectiveId,
            int  effectiveLevel,
            int? parentSelectedId,
            int  userId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_SaveSelectedLocationRow", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID",       physicalCountId);
                cmd.Parameters.AddWithValue("@Selected_Location_ID",    selectedId);
                cmd.Parameters.AddWithValue("@Selected_Location_Level", selectedLevel);
                cmd.Parameters.AddWithValue("@Effective_Location_ID",   effectiveId);
                cmd.Parameters.AddWithValue("@Effective_Location_Level",effectiveLevel);
                cmd.Parameters.AddWithValue("@Parent_Selected_ID",
                    parentSelectedId.HasValue ? (object)parentSelectedId.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@User_ID", userId);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Elimina todas las ubicaciones seleccionadas de un inventario físico.
        /// Equivalente al método DeleteSelectedLocationsByPhysicalCountId del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        public void EliminarUbicacionesSeleccionadas(int physicalCountId)
        {
            const string sql = @"
                DELETE FROM dbo.Physical_Count_Selected_Locations
                WHERE Physical_Count_ID = @p1";

            EjecutarNoConsulta(sql, physicalCountId);
        }

        /// <summary>
        /// Cuenta cuántas ubicaciones del inventario ya han sido contadas.
        /// Equivalente al método HasCountedLocations del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>Número de ubicaciones con Is_Counted = '1'.</returns>
        public int ContarUbicacionesContadas(int physicalCountId)
        {
            const string sql = @"
                SELECT COUNT(1)
                FROM dbo.Physical_Count_Selected_Locations
                WHERE Physical_Count_ID = @p1
                  AND Is_Counted        = '1'";

            return EjecutarEscalarInt(sql, physicalCountId);
        }

        /// <summary>
        /// Cuenta cuántas ubicaciones han sido guardadas para el inventario.
        /// Equivalente a la query de verificación de SaveSelectedLocations del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>Número de ubicaciones guardadas.</returns>
        public int ContarUbicacionesGuardadas(int physicalCountId)
        {
            const string sql = @"
                SELECT COUNT(1)
                FROM dbo.Physical_Count_Selected_Locations
                WHERE Physical_Count_ID = @p1";

            return EjecutarEscalarInt(sql, physicalCountId);
        }

        #endregion

        // =====================================================================
        // PASO 3 — CONTEO Y SNAPSHOTS
        // =====================================================================
        #region PASO 3 — CONTEO Y SNAPSHOTS

        /// <summary>
        /// Asegura que los snapshots del inventario existan; los genera si aún no se han creado.
        /// Equivalente al método EnsureSnapshotsExist del formulario.
        /// Invoca el stored procedure dbo.sp_PC_EnsureSnapshots.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>SnapshotResult con el resultado de la operación.</returns>
        public SnapshotResult AsegurarSnapshotsExistentes(int physicalCountId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_EnsureSnapshots", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID", physicalCountId);

                // Parámetros de salida del SP
                var pExisting  = new SqlParameter("@Existing_Snapshots",  SqlDbType.Int) { Direction = ParameterDirection.Output };
                var pGenerated = new SqlParameter("@Generated_Snapshots", SqlDbType.Int) { Direction = ParameterDirection.Output };
                var pMessage   = new SqlParameter("@Message", SqlDbType.NVarChar, 500)   { Direction = ParameterDirection.Output };

                cmd.Parameters.Add(pExisting);
                cmd.Parameters.Add(pGenerated);
                cmd.Parameters.Add(pMessage);

                conn.Open();
                cmd.ExecuteNonQuery();

                return new SnapshotResult
                {
                    Success            = true,
                    ExistingSnapshots  = pExisting.Value  != DBNull.Value ? Convert.ToInt32(pExisting.Value)  : 0,
                    GeneratedSnapshots = pGenerated.Value != DBNull.Value ? Convert.ToInt32(pGenerated.Value) : 0,
                    Message            = pMessage.Value   != DBNull.Value ? pMessage.Value.ToString()         : string.Empty
                };
            }
        }

        /// <summary>
        /// Refresca los datos de la grilla del Paso 3 para el inventario indicado.
        /// Equivalente al método RefreshStep3Grid del formulario.
        /// Invoca el stored procedure dbo.sp_PC_RefreshStep3Grid.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>DataTable con las filas de la grilla del Paso 3.</returns>
        public DataTable RefrescarGridPaso3(int physicalCountId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_RefreshStep3Grid", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID", physicalCountId);

                var dt = new DataTable();
                conn.Open();
                using (var reader = cmd.ExecuteReader())
                    dt.Load(reader);

                return dt;
            }
        }

        /// <summary>
        /// Obtiene el estado actual del conteo en la plataforma.
        /// Equivalente al método ValidateCountIsFinished del formulario.
        /// Invoca el stored procedure dbo.sp_PC_GetCountStatus.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>CountStatusResult con el estado del conteo.</returns>
        public CountStatusResult ObtenerEstadoConteo(int physicalCountId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_GetCountStatus", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID", physicalCountId);

                conn.Open();
                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new InvalidOperationException(
                            $"El inventario físico con ID {physicalCountId} no fue encontrado.");

                    int  statusId   = Convert.ToInt32(reader["Status_ID"]);
                    string statusName = reader["Status_Name"]?.ToString();
                    bool isFinished = Convert.ToBoolean(reader["Is_Finished"]);

                    return new CountStatusResult
                    {
                        StatusId   = statusId,
                        StatusName = statusName,
                        IsFinished = isFinished
                    };
                }
            }
        }

        /// <summary>
        /// Obtiene el nombre del inventario físico indicado.
        /// Equivalente a la parte de nombre en RefreshStep3Summary del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>Nombre del inventario, o null si no existe.</returns>
        public string ObtenerNombreConteo(int physicalCountId)
        {
            const string sql = @"
                SELECT Physical_Count_Name
                FROM dbo.Physical_Counts
                WHERE Physical_Count_ID = @p1";

            return EjecutarEscalarString(sql, physicalCountId);
        }

        /// <summary>
        /// Cuenta el número de ubicaciones distintas asociadas al inventario.
        /// Equivalente a la parte de ubicaciones en RefreshStep3Summary del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>Número de ubicaciones distintas seleccionadas.</returns>
        public int ContarUbicacionesDistintas(int physicalCountId)
        {
            const string sql = @"
                SELECT COUNT(DISTINCT Location_ID)
                FROM dbo.Physical_Count_Selected_Locations
                WHERE Physical_Count_ID = @p1";

            return EjecutarEscalarInt(sql, physicalCountId);
        }

        /// <summary>
        /// Elimina los datos del Paso 3: snapshots y ubicaciones seleccionadas.
        /// Equivalente al método DeleteStep3DataByPhysicalCountId del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        public void EliminarDatosPaso3(int physicalCountId)
        {
            // Eliminar primero los snapshots para respetar integridad referencial
            const string sqlSnapshots = @"
                DELETE FROM dbo.Physical_Count_Snapshots
                WHERE Physical_Count_ID = @p1";

            const string sqlLocaciones = @"
                DELETE FROM dbo.Physical_Count_Selected_Locations
                WHERE Physical_Count_ID = @p1";

            EjecutarNoConsulta(sqlSnapshots,  physicalCountId);
            EjecutarNoConsulta(sqlLocaciones, physicalCountId);
        }

        /// <summary>
        /// Obtiene los archivos adjuntos asociados a un inventario físico.
        /// Equivalente al método LoadAttachments del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>DataTable con columnas de la tabla Physical_Count_Attachments.</returns>
        public DataTable ObtenerArchivosAdjuntos(int physicalCountId)
        {
            const string sql = @"
                SELECT
                    Attachment_ID,
                    Physical_Count_ID,
                    File_Name,
                    File_Extension,
                    Content_Type,
                    File_Size,
                    Created_Date,
                    Created_By_User_ID
                FROM dbo.Physical_Count_Attachments
                WHERE Physical_Count_ID = @p1
                ORDER BY Created_Date DESC";

            return EjecutarConsulta(sql, physicalCountId);
        }

        /// <summary>
        /// Guarda un archivo adjunto asociado al inventario físico.
        /// Equivalente a la lógica de AttachAdd_Click del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <param name="fileName">Nombre del archivo sin extensión.</param>
        /// <param name="extension">Extensión del archivo (ej. ".pdf").</param>
        /// <param name="contentType">Tipo MIME del archivo.</param>
        /// <param name="fileSize">Tamaño del archivo en bytes.</param>
        /// <param name="fileData">Contenido binario del archivo.</param>
        /// <param name="userId">ID del usuario que adjunta el archivo.</param>
        public void GuardarAdjunto(
            int    physicalCountId,
            string fileName,
            string extension,
            string contentType,
            long   fileSize,
            byte[] fileData,
            int    userId)
        {
            if (fileData == null || fileData.Length == 0)
                throw new ArgumentException("El contenido del archivo no puede estar vacío.", nameof(fileData));

            const string sql = @"
                INSERT INTO dbo.Physical_Count_Attachments
                    (Physical_Count_ID, File_Name, File_Extension, Content_Type, File_Size, File_Data, Created_Date, Created_By_User_ID)
                VALUES
                    (@p1, @p2, @p3, @p4, @p5, @p6, GETDATE(), @p7)";

            EjecutarNoConsulta(sql, physicalCountId, fileName, extension, contentType, fileSize, fileData, userId);
        }

        /// <summary>
        /// Obtiene el contenido binario y el nombre de un archivo adjunto.
        /// Equivalente a la lógica de AttachOpen_Click del formulario.
        /// </summary>
        /// <param name="attachmentId">ID del adjunto.</param>
        /// <returns>AttachmentData con nombre y datos del archivo.</returns>
        public AttachmentData ObtenerAdjunto(int attachmentId)
        {
            const string sql = @"
                SELECT File_Name, File_Extension, File_Data
                FROM dbo.Physical_Count_Attachments
                WHERE Attachment_ID = @p1";

            DataTable dt = EjecutarConsulta(sql, attachmentId);

            if (dt.Rows.Count == 0)
                throw new InvalidOperationException(
                    $"No se encontró el adjunto con ID {attachmentId}.");

            DataRow fila = dt.Rows[0];
            return new AttachmentData
            {
                FileName = fila["File_Name"]?.ToString() + fila["File_Extension"]?.ToString(),
                FileData = fila["File_Data"] as byte[]
            };
        }

        /// <summary>
        /// Elimina un archivo adjunto del inventario físico.
        /// Equivalente a la lógica de AttachRemove_Click del formulario.
        /// </summary>
        /// <param name="attachmentId">ID del adjunto a eliminar.</param>
        public void EliminarAdjunto(int attachmentId)
        {
            const string sql = @"
                DELETE FROM dbo.Physical_Count_Attachments
                WHERE Attachment_ID = @p1";

            EjecutarNoConsulta(sql, attachmentId);
        }

        #endregion

        // =====================================================================
        // PASO 4 — REVISIÓN Y FINALIZACIÓN
        // =====================================================================
        #region PASO 4 — REVISIÓN Y FINALIZACIÓN

        /// <summary>
        /// Obtiene las diferencias de inventario para una ubicación específica.
        /// Equivalente al método LoadStep4DifferencesByLocation del formulario.
        /// Cuando Found es NULL o 0, la diferencia equivale a la cantidad original.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <param name="locationId">ID de la ubicación a revisar.</param>
        /// <returns>DataTable con las diferencias por artículo.</returns>
        public DataTable ObtenerDiferenciasPorUbicacion(int physicalCountId, int locationId)
        {
            const string sql = @"
                SELECT
                    s.Item_ID,
                    s.Item_Code,
                    s.Item_Name,
                    s.Location_ID,
                    s.Location_Name,
                    s.Quantity,
                    ISNULL(s.Found, 0)                                              AS Found,
                    CASE
                        WHEN s.Found IS NULL OR s.Found = 0
                        THEN s.Quantity
                        ELSE s.Found - s.Quantity
                    END                                                             AS Difference
                FROM dbo.Physical_Count_Snapshots s
                WHERE s.Physical_Count_ID = @p1
                  AND s.Location_ID       = @p2
                ORDER BY s.Item_Code";

            return EjecutarConsulta(sql, physicalCountId, locationId);
        }

        /// <summary>
        /// Obtiene las ubicaciones que cuentan con snapshot para un inventario.
        /// Equivalente al método LoadStep4LocationsCombo del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>DataTable con columnas Location_ID y Location_Name.</returns>
        public DataTable ObtenerUbicacionesSnapshot(int physicalCountId)
        {
            const string sql = @"
                SELECT DISTINCT
                    s.Location_ID,
                    l.Location_Name
                FROM dbo.Physical_Count_Snapshots s
                INNER JOIN dbo.Locations l ON l.Location_ID = s.Location_ID
                WHERE s.Physical_Count_ID = @p1
                ORDER BY l.Location_Name";

            return EjecutarConsulta(sql, physicalCountId);
        }

        /// <summary>
        /// Obtiene el resumen del encabezado del inventario para el Paso 5.
        /// Equivalente al método LoadStep5HeaderInfo del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>CountHeaderInfo con los datos del encabezado.</returns>
        public CountHeaderInfo ObtenerResumenEncabezado(int physicalCountId)
        {
            const string sql = @"
                SELECT
                    pc.Physical_Count_ID,
                    pc.Physical_Count_Name,
                    pc.Start_Date,
                    pc.End_Date,
                    u.Employee_Number,
                    u.Name          AS Responsible_Name,
                    u.Job_Position
                FROM dbo.Physical_Counts pc
                INNER JOIN dbo.Users u ON u.User_ID = pc.Responsible_User_ID
                WHERE pc.Physical_Count_ID = @p1";

            DataTable dt = EjecutarConsulta(sql, physicalCountId);

            if (dt.Rows.Count == 0)
                throw new InvalidOperationException(
                    $"No se encontró el inventario físico con ID {physicalCountId}.");

            DataRow fila = dt.Rows[0];
            return new CountHeaderInfo
            {
                PhysicalCountId  = Convert.ToInt32(fila["Physical_Count_ID"]),
                CountName        = fila["Physical_Count_Name"]?.ToString(),
                StartDate        = Convert.ToDateTime(fila["Start_Date"]),
                EndDate          = Convert.ToDateTime(fila["End_Date"]),
                EmployeeNumber   = fila["Employee_Number"]?.ToString(),
                ResponsibleName  = fila["Responsible_Name"]?.ToString(),
                JobPosition      = fila["Job_Position"]?.ToString()
            };
        }

        /// <summary>
        /// Carga los datos combinados (detalle y ejecutivo) del Paso 5.
        /// Equivalente al método LoadStep5Data del formulario.
        /// Invoca el stored procedure dbo.sp_PC_LoadStep5Data.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>Step5DataResult con las tablas de detalle y resumen ejecutivo.</returns>
        public Step5DataResult CargarDatosPaso5(int physicalCountId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_LoadStep5Data", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID", physicalCountId);

                var result = new Step5DataResult
                {
                    Detail    = new DataTable(),
                    Executive = new DataTable()
                };

                conn.Open();
                using (var reader = cmd.ExecuteReader())
                {
                    result.Detail.Load(reader);

                    if (reader.NextResult())
                        result.Executive.Load(reader);
                }

                return result;
            }
        }

        /// <summary>
        /// Carga el detalle del Paso 5 filtrado por ubicación.
        /// Equivalente al método LoadStep5DetailForLocation del formulario.
        /// Invoca el stored procedure dbo.sp_PC_LoadStep5DetailByLocation.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <param name="locationId">ID de la ubicación.</param>
        /// <returns>DataTable con el detalle de artículos para la ubicación.</returns>
        public DataTable CargarDetallePorUbicacion(int physicalCountId, int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_LoadStep5DetailByLocation", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID", physicalCountId);
                cmd.Parameters.AddWithValue("@Location_ID",       locationId);

                var dt = new DataTable();
                conn.Open();
                using (var reader = cmd.ExecuteReader())
                    dt.Load(reader);

                return dt;
            }
        }

        /// <summary>
        /// Carga el resumen ejecutivo del Paso 5 filtrado por ubicación.
        /// Equivalente al método LoadStep5ExecutiveForLocation del formulario.
        /// Invoca el stored procedure dbo.sp_PC_LoadStep5ExecutiveByLocation.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <param name="locationId">ID de la ubicación.</param>
        /// <returns>DataTable con el resumen ejecutivo para la ubicación.</returns>
        public DataTable CargarResumenEjecutivoPorUbicacion(int physicalCountId, int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_LoadStep5ExecutiveByLocation", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Physical_Count_ID", physicalCountId);
                cmd.Parameters.AddWithValue("@Location_ID",       locationId);

                var dt = new DataTable();
                conn.Open();
                using (var reader = cmd.ExecuteReader())
                    dt.Load(reader);

                return dt;
            }
        }

        /// <summary>
        /// Cuenta las líneas de snapshot que no han sido contadas (Found es NULL o 0).
        /// Equivalente a la validación en FinalizeWizard del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>Número de líneas pendientes de conteo.</returns>
        public int ContarLineasSinConteo(int physicalCountId)
        {
            const string sql = @"
                SELECT COUNT(1)
                FROM dbo.Physical_Count_Snapshots
                WHERE Physical_Count_ID = @p1
                  AND (Found IS NULL OR Found = 0)";

            return EjecutarEscalarInt(sql, physicalCountId);
        }

        /// <summary>
        /// Obtiene las ubicaciones distintas con snapshot, para el resumen del Paso 4.
        /// Equivalente a la parte de ubicaciones de LoadStep4Summary del formulario.
        /// </summary>
        /// <param name="physicalCountId">ID del inventario físico.</param>
        /// <returns>DataTable con la columna Location_Name.</returns>
        public DataTable ObtenerUbicacionesSnapshot_Distintas(int physicalCountId)
        {
            const string sql = @"
                SELECT DISTINCT l.Location_Name
                FROM dbo.Physical_Count_Snapshots s
                INNER JOIN dbo.Locations l ON l.Location_ID = s.Location_ID
                WHERE s.Physical_Count_ID = @p1
                ORDER BY l.Location_Name";

            return EjecutarConsulta(sql, physicalCountId);
        }

        #endregion

        // =====================================================================
        // MÉTODOS AUXILIARES PRIVADOS
        // =====================================================================
        #region MÉTODOS AUXILIARES PRIVADOS

        /// <summary>
        /// Ejecuta una consulta SQL y retorna los resultados en un DataTable.
        /// Equivalente a ExecFillTable del formulario.
        /// Los argumentos se asignan como @p1, @p2, @p3... en el mismo orden.
        /// </summary>
        /// <param name="sql">Sentencia SQL a ejecutar.</param>
        /// <param name="args">Parámetros posicionales de la consulta.</param>
        /// <returns>DataTable con los resultados.</returns>
        private DataTable EjecutarConsulta(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                AgregarParametrosPosicionales(cmd, args);

                var dt = new DataTable();
                conn.Open();
                using (var reader = cmd.ExecuteReader())
                    dt.Load(reader);

                return dt;
            }
        }

        /// <summary>
        /// Ejecuta una consulta SQL con parámetros explícitos (usada para cláusulas IN).
        /// </summary>
        /// <param name="sql">Sentencia SQL a ejecutar.</param>
        /// <param name="parametros">Lista de SqlParameter con nombre y valor.</param>
        /// <returns>DataTable con los resultados.</returns>
        private DataTable EjecutarConsultaConParametros(string sql, IEnumerable<SqlParameter> parametros)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                foreach (var p in parametros)
                    cmd.Parameters.Add(p);

                var dt = new DataTable();
                conn.Open();
                using (var reader = cmd.ExecuteReader())
                    dt.Load(reader);

                return dt;
            }
        }

        /// <summary>
        /// Ejecuta una consulta escalar y retorna el resultado como entero.
        /// Equivalente a ExecScalarInt del formulario.
        /// </summary>
        /// <param name="sql">Sentencia SQL a ejecutar.</param>
        /// <param name="args">Parámetros posicionales de la consulta.</param>
        /// <returns>Valor entero retornado por la consulta escalar.</returns>
        private int EjecutarEscalarInt(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                AgregarParametrosPosicionales(cmd, args);

                conn.Open();
                object resultado = cmd.ExecuteScalar();

                return resultado != null && resultado != DBNull.Value
                    ? Convert.ToInt32(resultado)
                    : 0;
            }
        }

        /// <summary>
        /// Ejecuta una consulta escalar y retorna el resultado como cadena.
        /// Equivalente a ExecScalarString del formulario.
        /// </summary>
        /// <param name="sql">Sentencia SQL a ejecutar.</param>
        /// <param name="args">Parámetros posicionales de la consulta.</param>
        /// <returns>Valor string retornado por la consulta escalar, o null.</returns>
        private string EjecutarEscalarString(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                AgregarParametrosPosicionales(cmd, args);

                conn.Open();
                object resultado = cmd.ExecuteScalar();

                return resultado != null && resultado != DBNull.Value
                    ? resultado.ToString()
                    : null;
            }
        }

        /// <summary>
        /// Ejecuta una sentencia SQL que no retorna filas (INSERT, UPDATE, DELETE).
        /// Equivalente a ExecNonQuery del formulario.
        /// </summary>
        /// <param name="sql">Sentencia SQL a ejecutar.</param>
        /// <param name="args">Parámetros posicionales de la sentencia.</param>
        private void EjecutarNoConsulta(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                AgregarParametrosPosicionales(cmd, args);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Agrega parámetros posicionales a un SqlCommand usando la convención @p1, @p2, @p3...
        /// Los valores null se convierten automáticamente en DBNull.Value.
        /// </summary>
        /// <param name="cmd">Comando al que se agregarán los parámetros.</param>
        /// <param name="args">Valores de los parámetros en orden posicional.</param>
        private static void AgregarParametrosPosicionales(SqlCommand cmd, object[] args)
        {
            if (args == null)
                return;

            for (int i = 0; i < args.Length; i++)
            {
                object valor = args[i] ?? DBNull.Value;
                cmd.Parameters.AddWithValue("@p" + (i + 1), valor);
            }
        }

        #endregion

        // =========================================================================
        // MÉTODOS NO MIGRADOS (permanecen en PhysicalCountCyclicForm)
        // =========================================================================
        // Los siguientes métodos contienen lógica de UI que no puede separarse
        // sin un mayor refactor de la arquitectura:
        //   - SaveInventoryHeader      → usa Persal.System_Functions directamente
        //   - FinalizeWizard           → coordina UI + Set_Physical_Count_Status
        //   - LoadStep4Summary         → mezcla cálculo con actualización de Label
        //   - RefreshStep3Summary      → actualiza Labels del formulario
        // Estos son candidatos para una segunda fase de refactorización.
        // =========================================================================
    }

    // =========================================================================
    // DTOs — Clases de transferencia de datos
    // =========================================================================

    /// <summary>Información del responsable del inventario.</summary>
    public class ResponsibleInfo
    {
        public int    UserId         { get; set; }
        public string EmployeeNumber { get; set; }
        public string Name           { get; set; }
        public string JobPosition    { get; set; }
    }

    /// <summary>Resultado de la verificación/generación de snapshots.</summary>
    public class SnapshotResult
    {
        public bool   Success            { get; set; }
        public int    ExistingSnapshots  { get; set; }
        public int    GeneratedSnapshots { get; set; }
        public string Message            { get; set; }
    }

    /// <summary>Estado actual del conteo en la plataforma web.</summary>
    public class CountStatusResult
    {
        public int    StatusId   { get; set; }
        public string StatusName { get; set; }
        public bool   IsFinished { get; set; }
    }

    /// <summary>Datos de un archivo adjunto al inventario.</summary>
    public class AttachmentData
    {
        public string FileName { get; set; }
        public byte[] FileData { get; set; }
    }

    /// <summary>Información del encabezado del inventario para el Paso 5.</summary>
    public class CountHeaderInfo
    {
        public int      PhysicalCountId  { get; set; }
        public string   CountName        { get; set; }
        public DateTime StartDate        { get; set; }
        public DateTime EndDate          { get; set; }
        public string   EmployeeNumber   { get; set; }
        public string   ResponsibleName  { get; set; }
        public string   JobPosition      { get; set; }
    }

    /// <summary>Resultado combinado Detail + Executive para el Paso 5.</summary>
    public class Step5DataResult
    {
        public DataTable Detail    { get; set; }
        public DataTable Executive { get; set; }
    }
}
