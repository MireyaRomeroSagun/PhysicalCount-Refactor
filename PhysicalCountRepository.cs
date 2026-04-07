using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace Persal._003_Physical_Counting_2
{
    // =========================================================================
    // PhysicalCountRepository — Acceso a datos del módulo de inventario cíclico
    // =========================================================================
    // Centraliza TODA la comunicación con SQL Server.
    // PhysicalCountCyclicForm solo llama a métodos de esta clase;
    // no contiene ni cadenas SQL ni objetos SqlConnection propios.
    // =========================================================================
    internal sealed class PhysicalCountRepository
    {
        private readonly string _connectionString;

        /// <summary>
        /// Constructor. Recibe la cadena de conexión activa.
        /// </summary>
        public PhysicalCountRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // =====================================================================
        #region PASO 1 — CREAR INVENTARIO
        // =====================================================================

        /// <summary>
        /// Carga los tipos de lista de conteo disponibles.
        /// </summary>
        public DataTable LoadListTypes()
        {
            return ExecFillTable(
                "SELECT Physical_Count_List_Type_ID, Physical_Count_List_Type " +
                "FROM Physical_Count_List_Types " +
                "WHERE Physical_Count_List_Type_ID = '1'");
        }

        /// <summary>
        /// Carga los datos de un responsable a partir de su User_ID.
        /// Devuelve true si se encontró; llena los parámetros de salida.
        /// </summary>
        public bool LoadResponsibleById(
            int userId,
            out int responsableUserId,
            out string empNo,
            out string nombre,
            out string puesto)
        {
            responsableUserId = 0;
            empNo             = string.Empty;
            nombre            = string.Empty;
            puesto            = string.Empty;

            if (userId <= 0) return false;

            const string sql =
                "SELECT TOP 1 User_ID, Name, Employee_Number, Job_Position " +
                "FROM Users " +
                "WHERE User_ID = @p1 AND Active = 'YES';";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@p1", userId);
                conn.Open();

                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return false;

                    responsableUserId = Convert.ToInt32(r["User_ID"]);
                    empNo             = Convert.ToString(r["Employee_Number"]);
                    nombre            = Convert.ToString(r["Name"]);
                    puesto            = Convert.ToString(r["Job_Position"]);
                    return true;
                }
            }
        }

        /// <summary>
        /// Busca un responsable activo por número de empleado, login o nombre.
        /// Devuelve true si se encontró; llena los parámetros de salida.
        /// </summary>
        public bool SearchResponsibleByText(
            string texto,
            out int responsableUserId,
            out string empNo,
            out string nombre,
            out string puesto)
        {
            responsableUserId = 0;
            empNo             = string.Empty;
            nombre            = string.Empty;
            puesto            = string.Empty;

            if (string.IsNullOrWhiteSpace(texto)) return false;

            texto = texto.Trim();

            const string sql =
                "SELECT TOP 1 User_ID, Name, Employee_Number, Job_Position " +
                "FROM Users " +
                "WHERE Active = 'YES' " +
                "  AND (Employee_Number = @p1 OR Login = @p1 OR Name LIKE @p2) " +
                "ORDER BY Name;";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@p1", texto);
                cmd.Parameters.AddWithValue("@p2", "%" + texto + "%");
                conn.Open();

                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return false;

                    responsableUserId = Convert.ToInt32(r["User_ID"]);
                    empNo             = Convert.ToString(r["Employee_Number"]);
                    nombre            = Convert.ToString(r["Name"]);
                    puesto            = Convert.ToString(r["Job_Position"]);
                    return true;
                }
            }
        }

        /// <summary>
        /// Obtiene el ID del último conteo creado por el usuario indicado.
        /// </summary>
        public int GetLastCountIdByUser(int userId)
        {
            return ExecScalarInt(
                "SELECT TOP 1 Physical_Count_ID FROM Physical_Counts " +
                "WHERE Created_By_User_ID = @p1 ORDER BY Physical_Count_ID DESC",
                userId);
        }

        #endregion

        // =====================================================================
        #region PASO 2 — SELECCIONAR UBICACIONES
        // =====================================================================

        /// <summary>
        /// Carga el almacén fijo con columna Selected = false.
        /// </summary>
        public DataTable LoadWarehouses(int warehouseId)
        {
            return ExecFillTable(
                "SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected " +
                "FROM dbo.Locations WHERE Location_ID = @p1 ORDER BY Location_Name;",
                warehouseId);
        }

        /// <summary>
        /// Carga las áreas hijas de un almacén padre con columna Selected = false.
        /// </summary>
        public DataTable LoadAreas(int parentId)
        {
            return ExecFillTable(
                "SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected " +
                "FROM dbo.Locations WHERE Parent_Location_ID = @p1 ORDER BY Location_Name;",
                parentId);
        }

        /// <summary>
        /// Carga las sub-locaciones hijas de las áreas indicadas.
        /// Devuelve un DataTable vacío si la lista está vacía.
        /// </summary>
        public DataTable LoadSubLocations(IList<int> areaIds)
        {
            if (areaIds == null || areaIds.Count == 0) return new DataTable();

            var sb = new StringBuilder();
            sb.Append("SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected ");
            sb.Append("FROM dbo.Locations WHERE Parent_Location_ID IN (");

            for (int i = 0; i < areaIds.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append("@p").Append(i + 1);
            }
            sb.Append(") ORDER BY Location_Name;");

            var args = new object[areaIds.Count];
            for (int i = 0; i < areaIds.Count; i++) args[i] = areaIds[i];

            return ExecFillTable(sb.ToString(), args);
        }

        /// <summary>
        /// Cuenta cuántas locaciones hijas tiene la locación indicada.
        /// </summary>
        public int CountChildLocations(int locationId)
        {
            return ExecScalarInt(
                "SELECT COUNT(1) FROM dbo.Locations WHERE Parent_Location_ID = @p1;",
                locationId);
        }

        /// <summary>
        /// Obtiene el Parent_Location_ID de una locación. Devuelve null si no tiene padre.
        /// </summary>
        public int? GetLocationParentId(int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(
                "SELECT TOP 1 Parent_Location_ID FROM dbo.Locations WHERE Location_ID = @p1;",
                conn))
            {
                cmd.Parameters.AddWithValue("@p1", locationId);
                conn.Open();
                object o = cmd.ExecuteScalar();
                if (o == null || o == DBNull.Value) return null;
                return Convert.ToInt32(o);
            }
        }

        /// <summary>
        /// Elimina las locaciones seleccionadas del conteo (antes de re-guardar).
        /// </summary>
        public void DeleteSelectedLocations(int countId)
        {
            ExecNonQuery(
                "DELETE FROM dbo.Physical_Count_Selected_Locations WHERE Physical_Count_ID = @p1;",
                countId);
        }

        /// <summary>
        /// Guarda una fila en Physical_Count_Selected_Locations usando el SP correspondiente.
        /// Lanza excepción si el SP reporta fallo.
        /// </summary>
        public void SaveSelectedLocationRow(
            int    countId,
            int    selectedId,
            int    selectedLevel,
            int    effectiveId,
            int    effectiveLevel,
            int?   parentId,
            int    userId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_SaveSelectedLocationRow", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId",         countId);
                cmd.Parameters.AddWithValue("@SelectedLocationId",      selectedId);
                cmd.Parameters.AddWithValue("@SelectedLevel",           selectedLevel);
                cmd.Parameters.AddWithValue("@EffectiveLocationId",     effectiveId);
                cmd.Parameters.AddWithValue("@EffectiveLevel",          effectiveLevel);
                cmd.Parameters.AddWithValue("@ParentSelectedLocationId",
                    (object)parentId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@UserId", userId);
                conn.Open();

                using (var r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        bool ok = r["Success"] != DBNull.Value
                               && Convert.ToBoolean(r["Success"]);
                        if (!ok)
                            throw new Exception("sp_PC_SaveSelectedLocationRow reportó fallo.");
                    }
                }
            }
        }

        /// <summary>
        /// Cuenta las locaciones seleccionadas guardadas en BD para el conteo.
        /// </summary>
        public int CountSavedLocations(int countId)
        {
            return ExecScalarInt(
                "SELECT COUNT(1) FROM dbo.Physical_Count_Selected_Locations " +
                "WHERE Physical_Count_ID = @p1",
                countId);
        }

        /// <summary>
        /// Verifica si alguna locación del conteo ya fue contada (impide retroceder).
        /// </summary>
        public bool HasCountedLocations(int countId)
        {
            return ExecScalarInt(
                "SELECT COUNT(*) FROM Physical_Count_Selected_Locations " +
                "WHERE Physical_Count_Id = @p1 AND Is_Counted = '1'",
                countId) > 0;
        }

        /// <summary>
        /// Elimina los datos del Paso 3: snapshots y locaciones seleccionadas.
        /// </summary>
        public void DeleteStep3Data(int countId)
        {
            ExecNonQuery(
                "DELETE FROM dbo.Physical_Count_Snapshots " +
                "  WHERE Physical_Count_ID = @p1; " +
                "DELETE FROM dbo.Physical_Count_Selected_Locations " +
                "  WHERE Physical_Count_ID = @p1;",
                countId);
        }

        #endregion

        // =====================================================================
        #region PASO 3 — CONTEO EN PLATAFORMA WEB
        // =====================================================================

        /// <summary>
        /// Asegura que existen snapshots para el conteo (SP sp_PC_EnsureSnapshots).
        /// Devuelve true si los snapshots ya existían o se generaron con éxito.
        /// Los parámetros de salida permiten que el Form muestre mensajes apropiados.
        /// </summary>
        public bool EnsureSnapshots(
            int     countId,
            out bool existentes,
            out int  generados,
            out string mensaje)
        {
            existentes = false;
            generados  = 0;
            mensaje    = null;

            if (countId <= 0) return false;

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_EnsureSnapshots", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", countId);
                conn.Open();

                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return false;

                    bool ok = r["Success"] != DBNull.Value
                           && Convert.ToBoolean(r["Success"]);
                    int snapExistentes = r["ExistingSnapshots"]  == DBNull.Value ? 0
                                        : Convert.ToInt32(r["ExistingSnapshots"]);
                    int snapGenerados  = r["GeneratedSnapshots"] == DBNull.Value ? 0
                                        : Convert.ToInt32(r["GeneratedSnapshots"]);
                    string msg = r["Message"] == DBNull.Value ? null
                                 : Convert.ToString(r["Message"]);

                    existentes = snapExistentes > 0;
                    generados  = snapGenerados;
                    mensaje    = msg;

                    return ok || existentes;
                }
            }
        }

        /// <summary>
        /// Consulta el estado del conteo (SP sp_PC_GetCountStatus).
        /// Devuelve true si se pudo leer el estado.
        /// </summary>
        public bool GetCountStatus(
            int     countId,
            out int    statusId,
            out string statusName,
            out bool   isFinished)
        {
            statusId   = 0;
            statusName = string.Empty;
            isFinished = false;

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_GetCountStatus", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", countId);
                conn.Open();

                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return false;

                    if (r["Physical_Count_Status_ID"] != DBNull.Value)
                        statusId = Convert.ToInt32(r["Physical_Count_Status_ID"]);
                    if (r["Physical_Count_Status"] != DBNull.Value)
                        statusName = Convert.ToString(r["Physical_Count_Status"]);
                    if (r["IsFinished"] != DBNull.Value)
                        isFinished = Convert.ToBoolean(r["IsFinished"]);

                    return true;
                }
            }
        }

        /// <summary>
        /// Carga la grilla del Paso 3 (SP sp_PC_RefreshStep3Grid).
        /// </summary>
        public DataTable LoadStep3Grid(int countId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_RefreshStep3Grid", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", countId);

                using (var da = new SqlDataAdapter(cmd))
                {
                    var dt = new DataTable();
                    da.Fill(dt);
                    return dt;
                }
            }
        }

        /// <summary>
        /// Obtiene el nombre del conteo desde la tabla Physical_Counts.
        /// </summary>
        public string GetCountName(int countId)
        {
            return ExecScalarString(
                "SELECT Physical_Count_Name FROM Physical_Counts " +
                "WHERE Physical_Count_ID = @p1",
                countId);
        }

        /// <summary>
        /// Cuenta las locaciones distintas seleccionadas para el conteo.
        /// </summary>
        public int CountSelectedLocations(int countId)
        {
            return ExecScalarInt(
                "SELECT COUNT(DISTINCT Location_ID) " +
                "FROM Physical_Count_Selected_Locations " +
                "WHERE Physical_Count_ID = @p1",
                countId);
        }

        // ── Adjuntos ─────────────────────────────────────────────────────────

        /// <summary>
        /// Carga los adjuntos del conteo en un DataTable.
        /// </summary>
        public DataTable LoadAttachments(int countId)
        {
            return ExecFillTable(
                "SELECT Physical_Count_Attachment_ID, File_Name, File_Extension, " +
                "File_Size, Created_Date " +
                "FROM dbo.Physical_Count_Attachments " +
                "WHERE Physical_Count_ID = @p1 ORDER BY File_Name;",
                countId);
        }

        /// <summary>
        /// Guarda un adjunto en la base de datos.
        /// </summary>
        public void SaveAttachment(
            int    countId,
            string fileName,
            string ext,
            string contentType,
            long   size,
            byte[] data,
            int    userId)
        {
            const string sql =
                "INSERT INTO dbo.Physical_Count_Attachments " +
                "(Physical_Count_ID, File_Name, File_Extension, Content_Type, " +
                " File_Size, File_Data, Created_By_User_ID) " +
                "VALUES (@p1, @p2, @p3, @p4, @p5, @p6, @p7);";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@p1", countId);
                cmd.Parameters.AddWithValue("@p2", fileName);
                cmd.Parameters.AddWithValue("@p3", ext);
                cmd.Parameters.AddWithValue("@p4", contentType);
                cmd.Parameters.AddWithValue("@p5", size);
                cmd.Parameters.Add("@p6", SqlDbType.VarBinary, data.Length).Value = data;
                cmd.Parameters.AddWithValue("@p7", userId);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Obtiene los datos binarios de un adjunto.
        /// Devuelve true si se encontró; llena los parámetros de salida.
        /// </summary>
        public bool GetAttachmentData(
            int     attachmentId,
            out string fileName,
            out byte[] data)
        {
            fileName = null;
            data     = null;

            const string sql =
                "SELECT TOP 1 File_Name, File_Data " +
                "FROM dbo.Physical_Count_Attachments " +
                "WHERE Physical_Count_Attachment_ID = @p1;";

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@p1", attachmentId);
                conn.Open();

                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return false;
                    fileName = Convert.ToString(r["File_Name"]);
                    data     = (byte[])r["File_Data"];
                    return true;
                }
            }
        }

        /// <summary>
        /// Elimina un adjunto por su ID.
        /// </summary>
        public void DeleteAttachment(int attachmentId)
        {
            ExecNonQuery(
                "DELETE FROM dbo.Physical_Count_Attachments " +
                "WHERE Physical_Count_Attachment_ID = @p1;",
                attachmentId);
        }

        #endregion

        // =====================================================================
        #region PASO 4 — REVISIÓN Y CIERRE
        // =====================================================================

        /// <summary>
        /// Carga los datos de detalle y resumen ejecutivo del Paso 4
        /// (SP sp_PC_LoadStep5Data — primer y segundo conjunto de resultados).
        /// </summary>
        public void LoadReviewData(
            int         countId,
            out DataTable detail,
            out DataTable executive)
        {
            detail    = new DataTable();
            executive = new DataTable();

            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_LoadStep5Data", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", countId);

                using (var da = new SqlDataAdapter(cmd))
                {
                    var ds = new DataSet();
                    da.Fill(ds);
                    if (ds.Tables.Count > 0) detail    = ds.Tables[0];
                    if (ds.Tables.Count > 1) executive = ds.Tables[1];
                }
            }
        }

        /// <summary>
        /// Carga los datos del encabezado de revisión (ID, nombre, fechas, responsable).
        /// </summary>
        public DataTable LoadReviewHeader(int countId)
        {
            const string sql =
                "SELECT PC.Physical_Count_ID, PC.Physical_Count_Name, " +
                "PC.Start_Date, PC.End_Date, " +
                "U.Employee_Number, U.Name, U.Job_Position " +
                "FROM Physical_Counts PC " +
                "LEFT JOIN Users U " +
                "  ON U.User_ID = PC.Default_Count_Responsible_ID " +
                "WHERE PC.Physical_Count_ID = @p1;";

            return ExecFillTable(sql, countId);
        }

        /// <summary>
        /// Carga el detalle filtrado por locación (SP sp_PC_LoadStep5DetailByLocation).
        /// Pasar locationId = -1 para obtener todas las locaciones.
        /// </summary>
        public DataTable LoadDetailByLocation(int countId, int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_LoadStep5DetailByLocation", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", countId);
                cmd.Parameters.AddWithValue("@LocationId",      locationId);

                using (var da = new SqlDataAdapter(cmd))
                {
                    var dt = new DataTable();
                    da.Fill(dt);
                    return dt;
                }
            }
        }

        /// <summary>
        /// Carga el resumen ejecutivo filtrado por locación
        /// (SP sp_PC_LoadStep5ExecutiveByLocation).
        /// Pasar locationId = -1 para obtener todas las locaciones.
        /// </summary>
        public DataTable LoadExecutiveByLocation(int countId, int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = new SqlCommand("dbo.sp_PC_LoadStep5ExecutiveByLocation", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", countId);
                cmd.Parameters.AddWithValue("@LocationId",      locationId);

                using (var da = new SqlDataAdapter(cmd))
                {
                    var dt = new DataTable();
                    da.Fill(dt);
                    return dt;
                }
            }
        }

        /// <summary>
        /// Cuenta los ítems del conteo que aún no han sido contados
        /// (Found IS NULL o Found = 0).
        /// </summary>
        public int CountUncountedItems(int countId)
        {
            return ExecScalarInt(
                "SELECT COUNT(1) FROM dbo.Physical_Count_Snapshots " +
                "WHERE Physical_Count_ID = @p1 AND (Found IS NULL OR Found = 0)",
                countId);
        }

        #endregion

        // =====================================================================
        #region HELPERS PRIVADOS ADO.NET
        // =====================================================================

        /// <summary>
        /// Ejecuta un SELECT y devuelve los resultados en un DataTable.
        /// Los parámetros se nombran @p1, @p2, @p3…
        /// </summary>
        private DataTable ExecFillTable(string sql, params object[] args)
        {
            var dt = new DataTable();
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = conn.CreateCommand())
            using (var da   = new SqlDataAdapter(cmd))
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                da.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// Ejecuta un SELECT escalar y devuelve el resultado como int (0 si nulo).
        /// Los parámetros se nombran @p1, @p2, @p3…
        /// </summary>
        private int ExecScalarInt(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                conn.Open();
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o);
            }
        }

        /// <summary>
        /// Ejecuta un SELECT escalar y devuelve el resultado como string (null si nulo).
        /// Los parámetros se nombran @p1, @p2, @p3…
        /// </summary>
        private string ExecScalarString(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                conn.Open();
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? null : Convert.ToString(o);
            }
        }

        /// <summary>
        /// Ejecuta una instrucción DML (INSERT / UPDATE / DELETE).
        /// Los parámetros se nombran @p1, @p2, @p3…
        /// </summary>
        private void ExecNonQuery(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd  = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        #endregion
    }
}
