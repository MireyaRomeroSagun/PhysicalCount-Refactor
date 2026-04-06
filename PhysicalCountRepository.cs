using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Persal._003_Physical_Counting_2
{
    /// Resultado estándar para operaciones del repositorio.
    /// Evita que el repositorio lance excepciones directamente a la UI.
    public class PhysicalCountResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public int GeneratedId { get; set; }
        public int RowsAffected { get; set; }
        public object Data { get; set; }
    }

    /// Repositorio de datos para el conteo físico cíclico.
    /// Centraliza todas las operaciones de base de datos, exportación e impresión.
    public class PhysicalCountRepository
    {
        private readonly string _connectionString;

        public PhysicalCountRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Step 1 — Encabezado del conteo

        public PhysicalCountResult SaveInventoryHeader(int physicalCountId, int countTypeId, int listTypeId, int responsibleUserId, int userId, string description, DateTime startDate, DateTime finishDate, int warehouseId)
        {
            try
            {
                var sf = new Persal.System_Functions();
                bool saved = sf.Save_Physical_Count(physicalCountId, countTypeId, listTypeId, responsibleUserId, userId, 1, description, startDate, finishDate, 100.0M, warehouseId);
                if (!saved) return new PhysicalCountResult { Success = false, ErrorMessage = sf.G.Error_Message };
                int newId = physicalCountId;
                if (physicalCountId == 0)
                    newId = ExecScalarInt("SELECT TOP 1 Physical_Count_ID FROM Physical_Counts WHERE Created_By_User_ID = @p1 ORDER BY Physical_Count_ID DESC", userId);
                return new PhysicalCountResult { Success = true, GeneratedId = newId };
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public DataTable LoadPhysicalCountListTypes()
        {
            return ExecFillTable("SELECT Physical_Count_List_Type_ID, Physical_Count_List_Type FROM Physical_Count_List_Types WHERE Physical_Count_List_Type_ID = '1'");
        }

        public PhysicalCountResult LoadResponsibleByUserId(int userId)
        {
            if (userId <= 0) return new PhysicalCountResult { Success = false, ErrorMessage = "Invalid user ID." };
            var dt = ExecFillTable("SELECT TOP 1 User_ID, Name, Employee_Number, Job_Position FROM Users WHERE User_ID = @p1 AND Active = 'YES';", userId);
            if (dt.Rows.Count == 0) return new PhysicalCountResult { Success = false, ErrorMessage = "Selected user was not found." };
            return new PhysicalCountResult { Success = true, Data = dt.Rows[0] };
        }

        public PhysicalCountResult SearchResponsibleByText(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return new PhysicalCountResult { Success = false, ErrorMessage = "Search value is empty." };
            value = value.Trim();
            const string sql = "SELECT TOP 1 User_ID, Name, Employee_Number, Login, Job_Position FROM Users WHERE Active = 'YES' AND (Employee_Number = @p1 OR Login = @p1 OR Name LIKE @p2) ORDER BY Name;";
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@p1", value);
                    cmd.Parameters.AddWithValue("@p2", "%" + value + "%");
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        if (!r.Read()) return new PhysicalCountResult { Success = false, ErrorMessage = "Active user not found. Try Employee Number, Login, Name, or use F3." };
                        int uid = Convert.ToInt32(r["User_ID"]);
                        return new PhysicalCountResult { Success = true, GeneratedId = uid, Data = new { UserId = uid, EmployeeNumber = Convert.ToString(r["Employee_Number"]), Name = Convert.ToString(r["Name"]), JobPosition = Convert.ToString(r["Job_Position"]) } };
                    }
                }
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        #endregion

        #region Step 2 — Ubicaciones

        public DataTable LoadAlmacenes(int fixedAlmacenId)
        {
            return ExecFillTable("SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected FROM dbo.Locations WHERE Location_ID = @p1 ORDER BY Location_Name;", fixedAlmacenId);
        }

        public DataTable LoadAreasForAlmacen(int almacenId)
        {
            return ExecFillTable("SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected FROM dbo.Locations WHERE Parent_Location_ID = @p1 ORDER BY Location_Name;", almacenId);
        }

        public DataTable LoadSubAreasForAreas(List<int> areaIds)
        {
            if (areaIds == null || areaIds.Count == 0) return new DataTable();
            var sb = new StringBuilder("SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected FROM dbo.Locations WHERE Parent_Location_ID IN (");
            for (int i = 0; i < areaIds.Count; i++) { if (i > 0) sb.Append(", "); sb.Append("@p" + (i + 1)); }
            sb.Append(") ORDER BY Location_Name;");
            return ExecFillTable(sb.ToString(), areaIds.Cast<object>().ToArray());
        }

        public PhysicalCountResult DeleteSelectedLocations(int physicalCountId)
        {
            try { ExecNonQuery("DELETE FROM dbo.Physical_Count_Selected_Locations WHERE Physical_Count_ID = @p1;", physicalCountId); return new PhysicalCountResult { Success = true }; }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public PhysicalCountResult DeleteStep3Data(int physicalCountId)
        {
            try
            {
                if (physicalCountId <= 0) return new PhysicalCountResult { Success = false, ErrorMessage = "Invalid physical count ID." };
                ExecNonQuery("DELETE FROM dbo.Physical_Count_Snapshots WHERE Physical_Count_ID = @p1; DELETE FROM dbo.Physical_Count_Selected_Locations WHERE Physical_Count_ID = @p1;", physicalCountId);
                return new PhysicalCountResult { Success = true };
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public bool HasCountedLocations(int physicalCountId)
        {
            return ExecScalarInt("SELECT COUNT(*) FROM Physical_Count_Selected_Locations WHERE Physical_Count_Id = @p1 AND Is_Counted = '1'", physicalCountId) > 0;
        }

        public PhysicalCountResult SaveSelectedLocationRow(int physicalCountId, int selectedId, int selectedLevel, int effectiveId, int effectiveLevel, int? parentSelectedId, int userId)
        {
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                using (var cmd = new SqlCommand("dbo.sp_PC_SaveSelectedLocationRow", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                    cmd.Parameters.AddWithValue("@SelectedLocationId", selectedId);
                    cmd.Parameters.AddWithValue("@SelectedLevel", selectedLevel);
                    cmd.Parameters.AddWithValue("@EffectiveLocationId", effectiveId);
                    cmd.Parameters.AddWithValue("@EffectiveLevel", effectiveLevel);
                    cmd.Parameters.AddWithValue("@ParentSelectedLocationId", (object)parentSelectedId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@UserId", userId);
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            bool ok = r["Success"] != DBNull.Value && Convert.ToBoolean(r["Success"]);
                            if (!ok) return new PhysicalCountResult { Success = false, ErrorMessage = "sp_PC_SaveSelectedLocationRow returned failure." };
                        }
                    }
                }
                return new PhysicalCountResult { Success = true };
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        #endregion

        #region Step 3 — Snapshots y conteo web

        public PhysicalCountResult EnsureSnapshotsExist(int physicalCountId)
        {
            if (physicalCountId <= 0) return new PhysicalCountResult { Success = false, ErrorMessage = "Invalid physical count ID." };
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                using (var cmd = new SqlCommand("dbo.sp_PC_EnsureSnapshots", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        if (!r.Read()) return new PhysicalCountResult { Success = false, ErrorMessage = "No response from sp_PC_EnsureSnapshots." };
                        bool success = r["Success"] != DBNull.Value && Convert.ToBoolean(r["Success"]);
                        int existing = r["ExistingSnapshots"] == DBNull.Value ? 0 : Convert.ToInt32(r["ExistingSnapshots"]);
                        int generated = r["GeneratedSnapshots"] == DBNull.Value ? 0 : Convert.ToInt32(r["GeneratedSnapshots"]);
                        string message = r["Message"] == DBNull.Value ? null : Convert.ToString(r["Message"]);
                        if (existing > 0) return new PhysicalCountResult { Success = true, RowsAffected = existing };
                        if (!success || generated <= 0) return new PhysicalCountResult { Success = false, ErrorMessage = message ?? "No snapshots generated." };
                        return new PhysicalCountResult { Success = true, RowsAffected = generated, Data = message };
                    }
                }
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public DataTable RefreshStep3Grid(int physicalCountId)
        {
            if (physicalCountId <= 0) return null;
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand("dbo.sp_PC_RefreshStep3Grid", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                using (var da = new SqlDataAdapter(cmd)) { var dt = new DataTable(); da.Fill(dt); return dt; }
            }
        }

        public PhysicalCountResult ValidateCountIsFinished(int physicalCountId)
        {
            if (physicalCountId <= 0) return new PhysicalCountResult { Success = false, ErrorMessage = "No valid inventory ID." };
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                using (var cmd = new SqlCommand("dbo.sp_PC_GetCountStatus", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        if (!r.Read()) return new PhysicalCountResult { Success = false, ErrorMessage = "Could not retrieve count status." };
                        int statusId = r["Physical_Count_Status_ID"] != DBNull.Value ? Convert.ToInt32(r["Physical_Count_Status_ID"]) : 0;
                        string statusName = r["Physical_Count_Status"] != DBNull.Value ? Convert.ToString(r["Physical_Count_Status"]) : "";
                        bool isFinished = r["IsFinished"] != DBNull.Value && Convert.ToBoolean(r["IsFinished"]);
                        if (!isFinished) return new PhysicalCountResult { Success = false, ErrorMessage = $"The count has not been completed on the web platform yet.\n\nCurrent status: {statusName} (ID: {statusId})" };
                        return new PhysicalCountResult { Success = true };
                    }
                }
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public PhysicalCountResult GetCountSummary(int physicalCountId)
        {
            try
            {
                string countName = ExecScalarString("SELECT Physical_Count_Name FROM Physical_Counts WHERE Physical_Count_ID = @p1", physicalCountId);
                int locationCount = ExecScalarInt("SELECT COUNT(DISTINCT Location_ID) FROM Physical_Count_Selected_Locations WHERE Physical_Count_ID = @p1", physicalCountId);
                return new PhysicalCountResult { Success = true, Data = new { CountName = countName ?? "", LocationCount = locationCount } };
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public DataTable LoadAttachments(int physicalCountId)
        {
            return ExecFillTable("SELECT Physical_Count_Attachment_ID, File_Name, File_Extension, File_Size, Created_Date FROM dbo.Physical_Count_Attachments WHERE Physical_Count_ID = @p1 ORDER BY File_Name;", physicalCountId);
        }

        public PhysicalCountResult SaveAttachment(int physicalCountId, string fileName, string ext, string contentType, long size, byte[] data, int userId)
        {
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                using (var cmd = new SqlCommand("INSERT INTO dbo.Physical_Count_Attachments (Physical_Count_ID, File_Name, File_Extension, Content_Type, File_Size, File_Data, Created_By_User_ID) VALUES (@p1,@p2,@p3,@p4,@p5,@p6,@p7);", conn))
                {
                    cmd.Parameters.AddWithValue("@p1", physicalCountId);
                    cmd.Parameters.AddWithValue("@p2", fileName);
                    cmd.Parameters.AddWithValue("@p3", ext);
                    cmd.Parameters.AddWithValue("@p4", contentType);
                    cmd.Parameters.AddWithValue("@p5", size);
                    cmd.Parameters.Add("@p6", SqlDbType.VarBinary, data.Length).Value = data;
                    cmd.Parameters.AddWithValue("@p7", userId);
                    conn.Open(); cmd.ExecuteNonQuery();
                }
                return new PhysicalCountResult { Success = true };
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public PhysicalCountResult OpenAttachment(int attachmentId)
        {
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                using (var cmd = new SqlCommand("SELECT TOP 1 File_Name, File_Data FROM dbo.Physical_Count_Attachments WHERE Physical_Count_Attachment_ID = @p1;", conn))
                {
                    cmd.Parameters.AddWithValue("@p1", attachmentId);
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        if (!r.Read()) return new PhysicalCountResult { Success = false, ErrorMessage = "Attachment not found." };
                        return new PhysicalCountResult { Success = true, Data = new { FileName = Convert.ToString(r["File_Name"]), FileData = (byte[])r["File_Data"] } };
                    }
                }
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public PhysicalCountResult DeleteAttachment(int attachmentId)
        {
            try { ExecNonQuery("DELETE FROM dbo.Physical_Count_Attachments WHERE Physical_Count_Attachment_ID = @p1;", attachmentId); return new PhysicalCountResult { Success = true }; }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        #endregion

        #region Step 4 — Diferencias

        public DataTable LoadStep4Locations(int physicalCountId)
        {
            return ExecFillTable("SELECT DISTINCT Location_ID, ISNULL(Location_Name,'EMPTY') AS [Location Name] FROM dbo.Physical_Count_Snapshots WHERE Physical_Count_ID = @p1 ORDER BY [Location Name]", physicalCountId);
        }

        public DataTable LoadStep4DifferencesByLocation(int physicalCountId, int locationId)
        {
            const string sql = @"SELECT PCS.Physical_Count_Snapshot_ID, PCS.Location_ID, PCS.Part_ID, PCS.Dimension_ID, PCS.Part_Number AS [Part Number], PCS.Description AS [Description], ISNULL(PCS.Dimension,'') AS [Dimension], CAST(ISNULL(PCS.Quantity,0) AS DECIMAL(18,4)) AS [System Stock], CAST(CASE WHEN PCS.Found IS NULL OR PCS.Found=0 THEN 0 ELSE PCS.Found END AS DECIMAL(18,4)) AS [Counted Qty], CAST(CASE WHEN PCS.Found IS NULL OR PCS.Found=0 THEN ISNULL(PCS.Quantity,0) ELSE ABS(ISNULL(PCS.Quantity,0)-PCS.Found) END AS DECIMAL(18,4)) AS [Difference], CAST(CASE WHEN PCS.Found IS NULL OR PCS.Found=0 THEN 100 WHEN ABS(ISNULL(PCS.Quantity,0)-PCS.Found)=0 THEN 0 WHEN ISNULL(PCS.Quantity,0)=0 THEN 100 ELSE CASE WHEN (ABS(ISNULL(PCS.Quantity,0)-PCS.Found)*100.0/PCS.Quantity)>100 THEN 100 ELSE ABS(ISNULL(PCS.Quantity,0)-PCS.Found)*100.0/PCS.Quantity END END AS DECIMAL(18,2)) AS [Diff %] FROM dbo.Physical_Count_Snapshots PCS WHERE PCS.Physical_Count_ID=@p1 AND PCS.Location_ID=@p2 ORDER BY PCS.Part_Number;";
            return ExecFillTable(sql, physicalCountId, locationId);
        }

        public DataTable LoadStep4PrintData(int physicalCountId, int locationId)
        {
            const string sql = @"SELECT PCS.Part_Number AS [Part Number], PCS.Description AS [Description], ISNULL(PCS.Dimension,'') AS [Dimension], CAST(ISNULL(PCS.Quantity,0) AS DECIMAL(18,4)) AS [System Stock], CAST(ISNULL(PCS.Found,0) AS DECIMAL(18,4)) AS [Counted Qty], CAST(CASE WHEN PCS.Found IS NULL OR PCS.Found=0 THEN ISNULL(PCS.Quantity,0) ELSE ABS(ISNULL(PCS.Quantity,0)-PCS.Found) END AS DECIMAL(18,4)) AS [Difference] FROM dbo.Physical_Count_Snapshots PCS WHERE PCS.Physical_Count_ID=@p1 AND PCS.Location_ID=@p2 ORDER BY PCS.Part_Number;";
            return ExecFillTable(sql, physicalCountId, locationId);
        }

        #endregion

        #region Step 5 — Revisión y finalización

        public DataSet LoadStep5Data(int physicalCountId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand("dbo.sp_PC_LoadStep5Data", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                using (var da = new SqlDataAdapter(cmd)) { var ds = new DataSet(); da.Fill(ds); return ds; }
            }
        }

        public DataTable LoadStep5DetailForLocation(int physicalCountId, int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand("dbo.sp_PC_LoadStep5DetailByLocation", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                cmd.Parameters.AddWithValue("@LocationId", locationId);
                using (var da = new SqlDataAdapter(cmd)) { var dt = new DataTable(); da.Fill(dt); return dt; }
            }
        }

        public DataTable LoadStep5ExecutiveForLocation(int physicalCountId, int locationId)
        {
            using (var conn = new SqlConnection(_connectionString))
            using (var cmd = new SqlCommand("dbo.sp_PC_LoadStep5ExecutiveByLocation", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PhysicalCountId", physicalCountId);
                cmd.Parameters.AddWithValue("@LocationId", locationId);
                using (var da = new SqlDataAdapter(cmd)) { var dt = new DataTable(); da.Fill(dt); return dt; }
            }
        }

        public DataTable LoadStep5HeaderInfo(int physicalCountId)
        {
            const string sql = @"SELECT PC.Physical_Count_ID, PC.Physical_Count_Name, PC.Start_Date, PC.End_Date, PC.Default_Count_Responsible_ID AS User_ID, U.Employee_Number, U.Name, U.Job_Position FROM Physical_Counts PC LEFT JOIN Users U ON U.User_ID = PC.Default_Count_Responsible_ID WHERE PC.Physical_Count_ID = @p1;";
            return ExecFillTable(sql, physicalCountId);
        }

        public PhysicalCountResult FinalizeInventory(int physicalCountId, int statusFinished)
        {
            try
            {
                var sf = new Persal.System_Functions();
                bool ok = sf.Set_Physical_Count_Status(statusFinished, physicalCountId);
                if (!ok) return new PhysicalCountResult { Success = false, ErrorMessage = sf.G.Error_Message };
                return new PhysicalCountResult { Success = true };
            }
            catch (Exception ex) { return new PhysicalCountResult { Success = false, ErrorMessage = ex.Message }; }
        }

        public int GetUncountedSnapshots(int physicalCountId)
        {
            return ExecScalarInt("SELECT COUNT(1) FROM dbo.Physical_Count_Snapshots WHERE Physical_Count_ID = @p1 AND (Found IS NULL OR Found = 0)", physicalCountId);
        }

        #endregion

        #region Exportación

        public void ExportToCsv(DataTable dt, string filePath)
        {
            const string sep = ";";
            using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                for (int i = 0; i < dt.Columns.Count; i++) { if (i > 0) sw.Write(sep); sw.Write(EscapeCsvValue(dt.Columns[i].ColumnName)); }
                sw.WriteLine();
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++) { if (c > 0) sw.Write(sep); sw.Write(EscapeCsvValue(dt.Rows[r][c] == DBNull.Value ? "" : dt.Rows[r][c].ToString())); }
                    sw.WriteLine();
                }
            }
        }

        public void ExportDataTableToCsv(DataTable dt, string defaultFileName, Form owner)
        {
            if (dt == null || dt.Columns.Count == 0) return;
            using (var sfd = new SaveFileDialog { Filter = "CSV files (*.csv)|*.csv", FileName = defaultFileName, Title = "Export CSV" })
            {
                if (sfd.ShowDialog(owner) != DialogResult.OK) return;
                var sb = new StringBuilder();
                for (int c = 0; c < dt.Columns.Count; c++) { if (c > 0) sb.Append(","); sb.Append(EscapeCsvValue(dt.Columns[c].ColumnName)); }
                sb.AppendLine();
                foreach (DataRow row in dt.Rows)
                {
                    for (int c = 0; c < dt.Columns.Count; c++) { if (c > 0) sb.Append(","); sb.Append(EscapeCsvValue(Convert.ToString(row[c] ?? ""))); }
                    sb.AppendLine();
                }
                File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                MessageBox.Show("CSV generated: " + sfd.FileName, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void ExportDataTableToExcel(DataTable dt, string sheetName, string defaultFileName, Form owner)
        {
            if (dt == null || dt.Columns.Count == 0) return;
            using (var sfd = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", FileName = defaultFileName, Title = "Export Excel" })
            {
                if (sfd.ShowDialog(owner) != DialogResult.OK) return;
                Excel.Application xlApp = null; Excel.Workbook xlBook = null; Excel.Worksheet xlSheet = null; Excel.Range hdr = null; Excel.Range used = null;
                try
                {
                    xlApp = new Excel.Application(); xlApp.DisplayAlerts = false; xlApp.Visible = false;
                    xlBook = xlApp.Workbooks.Add(); xlSheet = (Excel.Worksheet)xlBook.Worksheets[1]; xlSheet.Name = sheetName;
                    for (int c = 0; c < dt.Columns.Count; c++) xlSheet.Cells[1, c + 1] = dt.Columns[c].ColumnName;
                    hdr = xlSheet.Range[xlSheet.Cells[1, 1], xlSheet.Cells[1, dt.Columns.Count]];
                    hdr.Font.Bold = true; hdr.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 238, 250)); hdr.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    for (int r = 0; r < dt.Rows.Count; r++) for (int c = 0; c < dt.Columns.Count; c++) xlSheet.Cells[r + 2, c + 1] = dt.Rows[r][c] == DBNull.Value ? "" : Convert.ToString(dt.Rows[r][c]);
                    used = xlSheet.Range[xlSheet.Cells[1, 1], xlSheet.Cells[Math.Max(1, dt.Rows.Count + 1), dt.Columns.Count]];
                    used.Borders.LineStyle = Excel.XlLineStyle.xlContinuous; used.Columns.AutoFit();
                    xlApp.ActiveWindow.SplitRow = 1; xlApp.ActiveWindow.FreezePanes = true;
                    xlBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    MessageBox.Show("Excel generated: " + sfd.FileName, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Error generating Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally
                {
                    try { if (xlBook != null) xlBook.Close(false); } catch { }
                    try { if (xlApp != null) xlApp.Quit(); } catch { }
                    ReleaseComObjectSafe(used); ReleaseComObjectSafe(hdr); ReleaseComObjectSafe(xlSheet); ReleaseComObjectSafe(xlBook); ReleaseComObjectSafe(xlApp);
                    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
                }
            }
        }

        public string EscapeCsvValue(string s)
        {
            if (s == null) return "";
            bool mustQuote = s.Contains(";") || s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r");
            if (mustQuote) s ="\" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        public string EscapeCsvValue(object value)
        {
            if (value == null || value == DBNull.Value) return "";
            return EscapeCsvValue(Convert.ToString(value) ?? "");
        }

        private void ReleaseComObjectSafe(object obj)
        {
            try { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); } catch { }
        }

        #endregion
    }
}