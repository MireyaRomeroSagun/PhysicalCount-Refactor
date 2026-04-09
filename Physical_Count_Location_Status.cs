// Requires NuGet: ClosedXML
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace Persal._003_Physical_Counting_2
{
    // =========================================================================
    // Physical_Count_Location_Status — Read-only report: location count status
    // =========================================================================
    public class Physical_Count_Location_Status : Form
    {
        // =====================================================================
        // #region CONSTANTS
        // =====================================================================
        #region CONSTANTS

        private const int FormCornerRadius = 10;

        // Warehouse root location ID (fixed — same as in PhysicalCountCyclicForm Step2FixedAlmacenId)
        private const int WarehouseParentLocationId = 3;

        #endregion

        // =====================================================================
        // #region FIELDS
        // =====================================================================
        #region FIELDS

        private string _connectionString = "";
        private DataTable _dtFull;
        private DataTable _dtFiltered;

        // UI controls
        private DataGridView _dgvMain;
        private ComboBox _cboShow;
        private ComboBox _cboArea;
        private ComboBox _cboPeriod;
        private Label _lblTotal;
        private Label _lblOnTime;
        private Label _lblOverdue;
        private Label _lblNever;
        private Panel _pnlHeader;
        private Button _btnCerrar;
        private Button _btnMaximizar;
        private Button _btnMinimizar;

        // Print fields
        private PrintDocument _printDoc;
        private int _printRowIndex;
        private int _printCurrentPage;
        private int _printTotalPages;
        private float _printY;

        // Column widths for printing
        private readonly int[] _printColWidths = { 160, 120, 100, 90, 90, 80, 100 };
        private readonly string[] _printColHeaders = { "Location", "Area", "Period", "Last Count", "Next Due", "Days Overdue", "Status" };
        private readonly string[] _printColFields = { "Location_Name", "Area_Name", "Period_Name", "Last_Count_Date", "Next_Due_Date", "Days_Overdue", "Status" };

        #endregion

        // =====================================================================
        // #region CONSTRUCTORS
        // =====================================================================
        #region CONSTRUCTORS

        public Physical_Count_Location_Status(string connectionString)
        {
            _connectionString = connectionString ?? "";
            SetupUI();
            this.Load += OnFormLoad;
        }

        public Physical_Count_Location_Status() : this("") { }

        private string GetCs()
        {
            if (_connectionString != null && _connectionString.Trim().Length > 0)
                return _connectionString;
            try { return new con().ConnectionString; }
            catch { return ""; }
        }

        #endregion

        // =====================================================================
        // #region SETUP UI
        // =====================================================================
        #region SETUP UI

        private void SetupUI()
        {
            // Form properties
            this.Text = "Location Count Status";
            this.Font = new Font("Segoe UI", 9F);
            this.BackColor = Color.FromArgb(243, 244, 246);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1100, 700);
            this.Size = new Size(1300, 800);
            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;

            ApplyRoundedRegion(this, FormCornerRadius);

            this.Resize += (s, e) =>
            {
                if (this.WindowState == FormWindowState.Maximized)
                    this.Region = null;
                else
                    ApplyRoundedRegion(this, FormCornerRadius);
            };

            // --- Header ---
            BuildHeader();

            // --- Footer ---
            BuildFooter();

            // --- Filter bar ---
            BuildFilterBar();

            // --- Main card + grid ---
            BuildGridArea();
        }

        private void BuildHeader()
        {
            _pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = Color.FromArgb(31, 41, 55),
                Padding = new Padding(16, 0, 0, 0)
            };

            // Title label
            var lblTitle = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = "Location Count Status Report",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Chrome buttons (right side)
            var headerRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 132,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.FromArgb(31, 41, 55),
                Padding = new Padding(0)
            };

            _btnCerrar = CreateChromeButton("✕");
            _btnMaximizar = CreateChromeButton("□");
            _btnMinimizar = CreateChromeButton("—");

            _btnCerrar.Click += (s, e) => this.Close();
            _btnMaximizar.Click += (s, e) => ToggleMaximize();
            _btnMinimizar.Click += (s, e) => this.WindowState = FormWindowState.Minimized;

            headerRight.Controls.Add(_btnCerrar);
            headerRight.Controls.Add(_btnMaximizar);
            headerRight.Controls.Add(_btnMinimizar);

            // Drag support
            MouseEventHandler dragHandler = (s, e) =>
            {
                if (this.WindowState == FormWindowState.Maximized) return;
                if (e.Button != MouseButtons.Left) return;
                ReleaseCapture();
                SendMessage(this.Handle, 0xA1, 0x2, 0);
            };

            _pnlHeader.MouseDown += dragHandler;
            lblTitle.MouseDown += dragHandler;

            _pnlHeader.Controls.Add(lblTitle);
            _pnlHeader.Controls.Add(headerRight);

            this.Controls.Add(_pnlHeader);
        }

        private void BuildFooter()
        {
            var pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = Color.White,
                Padding = new Padding(12, 0, 12, 0)
            };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.White,
                Padding = new Padding(0, 8, 0, 0)
            };

            _lblTotal = new Label
            {
                AutoSize = true,
                Text = "Total: 0",
                ForeColor = Color.FromArgb(31, 41, 55),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Margin = new Padding(0, 0, 20, 0)
            };
            _lblOnTime = new Label
            {
                AutoSize = true,
                Text = "✅ On Time: 0",
                ForeColor = Color.FromArgb(6, 95, 70),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Margin = new Padding(0, 0, 20, 0)
            };
            _lblOverdue = new Label
            {
                AutoSize = true,
                Text = "⚠ Overdue: 0",
                ForeColor = Color.FromArgb(120, 53, 15),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Margin = new Padding(0, 0, 20, 0)
            };
            _lblNever = new Label
            {
                AutoSize = true,
                Text = "❌ Never Counted: 0",
                ForeColor = Color.FromArgb(153, 27, 27),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 0)
            };

            flow.Controls.Add(_lblTotal);
            flow.Controls.Add(_lblOnTime);
            flow.Controls.Add(_lblOverdue);
            flow.Controls.Add(_lblNever);

            pnlFooter.Controls.Add(flow);
            this.Controls.Add(pnlFooter);
        }

        private void BuildFilterBar()
        {
            var pnlFilter = new Panel
            {
                Dock = DockStyle.Top,
                Height = 56,
                BackColor = Color.White,
                Padding = new Padding(12, 8, 12, 8)
            };

            // Left side — combos + search button
            var flowLeft = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.White,
                Padding = new Padding(0)
            };

            // Show:
            var lblShow = new Label { Text = "Show:", AutoSize = true, Margin = new Padding(0, 8, 4, 0), ForeColor = Color.FromArgb(31, 41, 55) };
            _cboShow = new ComboBox
            {
                Width = 130,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 4, 12, 0),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };

            // Area:
            var lblArea = new Label { Text = "Area:", AutoSize = true, Margin = new Padding(0, 8, 4, 0), ForeColor = Color.FromArgb(31, 41, 55) };
            _cboArea = new ComboBox
            {
                Width = 160,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 4, 12, 0),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };

            // Period:
            var lblPeriod = new Label { Text = "Period:", AutoSize = true, Margin = new Padding(0, 8, 4, 0), ForeColor = Color.FromArgb(31, 41, 55) };
            _cboPeriod = new ComboBox
            {
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 4, 12, 0),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };

            // Search button
            var btnSearch = new Button { Text = "🔍 Search", Width = 100, Height = 32, Margin = new Padding(0, 2, 0, 0) };
            StyleButtonPrimary(btnSearch);
            btnSearch.Height = 32;
            btnSearch.Click += (s, e) => LoadData();

            flowLeft.Controls.Add(lblShow);
            flowLeft.Controls.Add(_cboShow);
            flowLeft.Controls.Add(lblArea);
            flowLeft.Controls.Add(_cboArea);
            flowLeft.Controls.Add(lblPeriod);
            flowLeft.Controls.Add(_cboPeriod);
            flowLeft.Controls.Add(btnSearch);

            // Right side — Print + Export
            var flowRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 230,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.White,
                Padding = new Padding(0)
            };

            var btnExport = new Button { Text = "📊 Export", Width = 110, Height = 32, Margin = new Padding(4, 2, 0, 0) };
            StyleButtonGhost(btnExport);
            btnExport.Height = 32;
            btnExport.Click += (s, e) => ExportToExcel();

            var btnPrint = new Button { Text = "🖨 Print", Width = 100, Height = 32, Margin = new Padding(0, 2, 0, 0) };
            StyleButtonGhost(btnPrint);
            btnPrint.Height = 32;
            btnPrint.Click += (s, e) => PrintReport();

            flowRight.Controls.Add(btnExport);
            flowRight.Controls.Add(btnPrint);

            pnlFilter.Controls.Add(flowLeft);
            pnlFilter.Controls.Add(flowRight);

            this.Controls.Add(pnlFilter);
        }

        private void BuildGridArea()
        {
            var card = new ModernCardPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(1)
            };

            _dgvMain = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ScrollBars = ScrollBars.Both
            };

            StyleGrid(_dgvMain);
            BuildGridColumns();

            _dgvMain.CellFormatting += DgvMain_CellFormatting;
            _dgvMain.CellPainting += DgvMain_CellPainting;

            card.Controls.Add(_dgvMain);
            this.Controls.Add(card);
        }

        private void BuildGridColumns()
        {
            _dgvMain.Columns.Clear();

            // Location_ID (hidden)
            var colId = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Location_ID",
                HeaderText = "ID",
                Width = 0,
                Visible = false
            };
            _dgvMain.Columns.Add(colId);

            // Location_Name (fill)
            var colName = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Location_Name",
                HeaderText = "Location",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            _dgvMain.Columns.Add(colName);

            // Area_Name
            var colArea = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Area_Name",
                HeaderText = "Area",
                Width = 140
            };
            _dgvMain.Columns.Add(colArea);

            // Period_Name
            var colPeriod = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Period_Name",
                HeaderText = "Period",
                Width = 120
            };
            _dgvMain.Columns.Add(colPeriod);

            // Last_Count_Date
            var colLastDate = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Last_Count_Date",
                HeaderText = "Last Count",
                Width = 110
            };
            _dgvMain.Columns.Add(colLastDate);

            // Last_Count_Name
            var colLastName = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Last_Count_Name",
                HeaderText = "Last Inventory",
                Width = 180
            };
            _dgvMain.Columns.Add(colLastName);

            // Next_Due_Date
            var colNextDate = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Next_Due_Date",
                HeaderText = "Next Due",
                Width = 110
            };
            _dgvMain.Columns.Add(colNextDate);

            // Days_Overdue
            var colDays = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Days_Overdue",
                HeaderText = "Days Overdue",
                Width = 100
            };
            _dgvMain.Columns.Add(colDays);

            // Status (custom painted)
            var colStatus = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Status",
                HeaderText = "Status",
                Width = 130
            };
            _dgvMain.Columns.Add(colStatus);
        }

        #endregion

        // =====================================================================
        // #region FORM LOAD
        // =====================================================================
        #region FORM LOAD

        private void OnFormLoad(object sender, EventArgs e)
        {
            LoadShowCombo();
            LoadAreaCombo();
            LoadPeriodCombo();
            LoadData();
        }

        private void LoadShowCombo()
        {
            var dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add(0, "All");
            dt.Rows.Add(1, "Overdue");
            dt.Rows.Add(2, "Never Counted");

            _cboShow.DataSource = dt;
            _cboShow.DisplayMember = "Name";
            _cboShow.ValueMember = "ID";
            _cboShow.SelectedIndex = 0;
        }

        private void LoadAreaCombo()
        {
            var dt = new DataTable();
            dt.Columns.Add("Location_ID", typeof(int));
            dt.Columns.Add("Location_Name", typeof(string));
            dt.Rows.Add(0, "— All Areas —");

            try
            {
                string sql = string.Format(@"
SELECT Location_ID, Location_Name
FROM dbo.Locations
WHERE Parent_Location_ID = {0}
ORDER BY Location_Name", WarehouseParentLocationId);

                using (var conn = new SqlConnection(GetCs()))
                using (var cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    using (var da = new SqlDataAdapter(cmd))
                    {
                        var dtDb = new DataTable();
                        da.Fill(dtDb);
                        foreach (DataRow r in dtDb.Rows)
                            dt.Rows.Add(r["Location_ID"], r["Location_Name"]);
                    }
                }
            }
            catch { }

            _cboArea.DataSource = dt;
            _cboArea.DisplayMember = "Location_Name";
            _cboArea.ValueMember = "Location_ID";
            _cboArea.SelectedIndex = 0;
        }

        private void LoadPeriodCombo()
        {
            var dt = new DataTable();
            dt.Columns.Add("Period_ID", typeof(int));
            dt.Columns.Add("Period_Name", typeof(string));
            dt.Rows.Add(0, "— All Periods —");

            try
            {
                const string sql = @"
SELECT Period_ID, Period_Name
FROM dbo.Physical_Count_Period
WHERE Activo = 1
ORDER BY Period_Days ASC";

                using (var conn = new SqlConnection(GetCs()))
                using (var cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    using (var da = new SqlDataAdapter(cmd))
                    {
                        var dtDb = new DataTable();
                        da.Fill(dtDb);
                        foreach (DataRow r in dtDb.Rows)
                            dt.Rows.Add(r["Period_ID"], r["Period_Name"]);
                    }
                }
            }
            catch { }

            _cboPeriod.DataSource = dt;
            _cboPeriod.DisplayMember = "Period_Name";
            _cboPeriod.ValueMember = "Period_ID";
            _cboPeriod.SelectedIndex = 0;
        }

        #endregion

        // =====================================================================
        // #region DATA LOADING
        // =====================================================================
        #region DATA LOADING

        private void LoadData()
        {
            int areaId = 0;
            int periodId = 0;

            try
            {
                if (_cboArea.SelectedValue != null)
                    int.TryParse(_cboArea.SelectedValue.ToString(), out areaId);
            }
            catch { }

            try
            {
                if (_cboPeriod.SelectedValue != null)
                    int.TryParse(_cboPeriod.SelectedValue.ToString(), out periodId);
            }
            catch { }

            var sbSql = new System.Text.StringBuilder();
            sbSql.Append(string.Format(@"
SELECT
    L.Location_ID,
    L.Location_Name,
    Parent.Location_Name        AS Area_Name,
    P.Period_Name,
    P.Period_Days,
    MAX(PCSL.Counted_Date)      AS Last_Count_Date,
    MAX(PC.Physical_Count_Name) AS Last_Count_Name,
    CASE
        WHEN MAX(PCSL.Counted_Date) IS NULL
            THEN 'Never Counted'
        WHEN DATEDIFF(day, MAX(PCSL.Counted_Date), GETDATE()) > P.Period_Days
            THEN 'Overdue'
        ELSE 'On Time'
    END AS Status,
    CASE
        WHEN MAX(PCSL.Counted_Date) IS NULL THEN NULL
        ELSE DATEADD(day, P.Period_Days, MAX(PCSL.Counted_Date))
    END AS Next_Due_Date,
    CASE
        WHEN MAX(PCSL.Counted_Date) IS NULL THEN NULL
        ELSE DATEDIFF(day,
                DATEADD(day, P.Period_Days, MAX(PCSL.Counted_Date)),
                GETDATE())
    END AS Days_Overdue
FROM dbo.Locations L
LEFT JOIN dbo.Locations Parent
    ON Parent.Location_ID = L.Parent_Location_ID
LEFT JOIN dbo.Physical_Count_Period P
    ON P.Period_ID = L.Period_ID AND P.Activo = 1
LEFT JOIN dbo.Physical_Count_Selected_Locations PCSL
    ON PCSL.Effective_Location_ID = L.Location_ID
    AND PCSL.Is_Counted = 1
LEFT JOIN dbo.Physical_Counts PC
    ON PC.Physical_Count_ID = PCSL.Physical_Count_ID
    AND PC.Physical_Count_Status_ID = 3
WHERE L.Parent_Location_ID IN (
    SELECT Location_ID FROM dbo.Locations
    WHERE Parent_Location_ID = {0}
)", WarehouseParentLocationId));

            if (areaId > 0)
                sbSql.Append(" AND Parent.Location_ID = @areaId");

            if (periodId > 0)
                sbSql.Append(" AND P.Period_ID = @periodId");

            sbSql.Append(@"
GROUP BY
    L.Location_ID, L.Location_Name,
    Parent.Location_Name,
    P.Period_Name, P.Period_Days
ORDER BY Status DESC, L.Location_Name ASC");

            _dtFull = new DataTable();

            try
            {
                using (var conn = new SqlConnection(GetCs()))
                using (var cmd = new SqlCommand(sbSql.ToString(), conn))
                {
                    if (areaId > 0)
                        cmd.Parameters.AddWithValue("@areaId", areaId);
                    if (periodId > 0)
                        cmd.Parameters.AddWithValue("@periodId", periodId);

                    conn.Open();
                    using (var da = new SqlDataAdapter(cmd))
                        da.Fill(_dtFull);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error loading data: {0}", ex.Message), "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _dtFull = new DataTable();
            }

            ApplyShowFilter();
        }

        private void ApplyShowFilter()
        {
            int showVal = 0;
            try
            {
                if (_cboShow.SelectedValue != null)
                    int.TryParse(_cboShow.SelectedValue.ToString(), out showVal);
            }
            catch { }

            if (_dtFull == null)
                _dtFull = new DataTable();

            if (showVal == 0)
            {
                // All
                _dtFiltered = _dtFull.Copy();
            }
            else
            {
                string filterStatus = showVal == 1 ? "Overdue" : "Never Counted";
                _dtFiltered = _dtFull.Clone(); // same schema, no rows
                foreach (DataRow row in _dtFull.Rows)
                {
                    string st = Convert.ToString(row["Status"]);
                    if (st == filterStatus)
                        _dtFiltered.ImportRow(row);
                }
            }

            _dgvMain.DataSource = null;
            _dgvMain.DataSource = _dtFiltered;

            // Fix date columns display
            foreach (DataGridViewRow row in _dgvMain.Rows)
            {
                FormatDateCell(row, "Last_Count_Date");
                FormatDateCell(row, "Next_Due_Date");
            }

            UpdateFooterSummary();
        }

        private void FormatDateCell(DataGridViewRow row, string dataPropertyName)
        {
            foreach (DataGridViewColumn col in _dgvMain.Columns)
            {
                if (col.DataPropertyName == dataPropertyName)
                {
                    var cell = row.Cells[col.Index];
                    if (cell.Value == null || cell.Value == DBNull.Value)
                        cell.Value = "";
                    else
                    {
                        try
                        {
                            DateTime dt = Convert.ToDateTime(cell.Value);
                            cell.Value = dt.ToString("MM/dd/yyyy");
                        }
                        catch { }
                    }
                    break;
                }
            }
        }

        private void UpdateFooterSummary()
        {
            if (_dtFiltered == null)
            {
                _lblTotal.Text = "Total: 0";
                _lblOnTime.Text = "✅ On Time: 0";
                _lblOverdue.Text = "⚠ Overdue: 0";
                _lblNever.Text = "❌ Never Counted: 0";
                return;
            }

            int total = _dtFiltered.Rows.Count;
            int onTime = 0, overdue = 0, never = 0;

            foreach (DataRow r in _dtFiltered.Rows)
            {
                string s = Convert.ToString(r["Status"]);
                if (s == "On Time") onTime++;
                else if (s == "Overdue") overdue++;
                else if (s == "Never Counted") never++;
            }

            _lblTotal.Text = string.Format("Total: {0}", total);
            _lblOnTime.Text = string.Format("✅ On Time: {0}", onTime);
            _lblOverdue.Text = string.Format("⚠ Overdue: {0}", overdue);
            _lblNever.Text = string.Format("❌ Never Counted: {0}", never);
        }

        #endregion

        // =====================================================================
        // #region GRID EVENTS
        // =====================================================================
        #region GRID EVENTS

        private void DgvMain_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || _dgvMain.Rows[e.RowIndex].IsNewRow) return;

            // Determine row status for background coloring
            int statusColIndex = -1;
            foreach (DataGridViewColumn col in _dgvMain.Columns)
            {
                if (col.DataPropertyName == "Status") { statusColIndex = col.Index; break; }
            }
            if (statusColIndex < 0) return;

            string status = Convert.ToString(_dgvMain.Rows[e.RowIndex].Cells[statusColIndex].Value);

            Color rowColor;
            switch (status)
            {
                case "Overdue":
                    rowColor = Color.FromArgb(254, 243, 199);
                    break;
                case "Never Counted":
                    rowColor = Color.FromArgb(254, 226, 226);
                    break;
                default:
                    rowColor = Color.White;
                    break;
            }

            // Apply to all cells except the Status column (handled by CellPainting)
            int statusPaintColIndex = -1;
            foreach (DataGridViewColumn col in _dgvMain.Columns)
            {
                if (col.DataPropertyName == "Status") { statusPaintColIndex = col.Index; break; }
            }

            if (e.ColumnIndex != statusPaintColIndex)
            {
                _dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = rowColor;
                _dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.SelectionBackColor =
                    ControlPaint.Dark(rowColor, 0.05f);
            }

            // Format dates
            string dpn = _dgvMain.Columns[e.ColumnIndex].DataPropertyName;
            if (dpn == "Last_Count_Date" || dpn == "Next_Due_Date")
            {
                if (e.Value == null || e.Value == DBNull.Value || Convert.ToString(e.Value).Trim().Length == 0)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }
                else
                {
                    DateTime d;
                    if (DateTime.TryParse(Convert.ToString(e.Value), out d))
                    {
                        e.Value = d.ToString("MM/dd/yyyy");
                        e.FormattingApplied = true;
                    }
                }
            }

            // Days_Overdue — show empty if null
            if (dpn == "Days_Overdue")
            {
                if (e.Value == null || e.Value == DBNull.Value)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }
            }
        }

        private void DgvMain_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Find the Status column index
            int statusColIndex = -1;
            foreach (DataGridViewColumn col in _dgvMain.Columns)
            {
                if (col.DataPropertyName == "Status") { statusColIndex = col.Index; break; }
            }
            if (e.ColumnIndex != statusColIndex) return;

            string status = Convert.ToString(e.Value);

            Color fillColor;
            Color textColor;

            switch (status)
            {
                case "On Time":
                    fillColor = Color.FromArgb(209, 250, 229);
                    textColor = Color.FromArgb(6, 95, 70);
                    break;
                case "Overdue":
                    fillColor = Color.FromArgb(254, 243, 199);
                    textColor = Color.FromArgb(120, 53, 15);
                    break;
                case "Never Counted":
                    fillColor = Color.FromArgb(254, 226, 226);
                    textColor = Color.FromArgb(153, 27, 27);
                    break;
                default:
                    return; // Let default painting handle unknown
            }

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Row background
            string rowStatus = status;
            Color rowBg;
            switch (rowStatus)
            {
                case "Overdue": rowBg = Color.FromArgb(254, 243, 199); break;
                case "Never Counted": rowBg = Color.FromArgb(254, 226, 226); break;
                default: rowBg = Color.White; break;
            }

            using (var bgBrush = new SolidBrush(rowBg))
                e.Graphics.FillRectangle(bgBrush, e.CellBounds);

            // Pill
            int pillH = 22;
            int pillW = e.CellBounds.Width - 16;
            int pillX = e.CellBounds.X + 8;
            int pillY = e.CellBounds.Y + (e.CellBounds.Height - pillH) / 2;

            var pillRect = new Rectangle(pillX, pillY, pillW, pillH);
            int radius = pillH / 2;

            using (var path = new GraphicsPath())
            {
                int d = radius * 2;
                path.AddArc(pillRect.X, pillRect.Y, d, d, 180, 90);
                path.AddArc(pillRect.Right - d, pillRect.Y, d, d, 270, 90);
                path.AddArc(pillRect.Right - d, pillRect.Bottom - d, d, d, 0, 90);
                path.AddArc(pillRect.X, pillRect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();

                using (var fillBrush = new SolidBrush(fillColor))
                    e.Graphics.FillPath(fillBrush, path);
            }

            // Text
            using (var txtBrush = new SolidBrush(textColor))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                e.Graphics.DrawString(status, new Font("Segoe UI", 8F, FontStyle.Bold), txtBrush, pillRect, sf);
            }

            // Grid line (bottom)
            using (var gridPen = new Pen(Color.FromArgb(229, 231, 235)))
                e.Graphics.DrawLine(gridPen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);

            e.Handled = true;
        }

        #endregion

        // =====================================================================
        // #region PRINT
        // =====================================================================
        #region PRINT

        private void PrintReport()
        {
            if (_dtFiltered == null || _dtFiltered.Rows.Count == 0)
            {
                MessageBox.Show("No data to print.", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Pre-calculate total pages
            _printTotalPages = CalculatePrintTotalPages();

            _printDoc = new PrintDocument();
            _printDoc.PrintPage += PrintDoc_PrintPage;

            _printRowIndex = 0;
            _printCurrentPage = 0;

            using (var ppd = new PrintPreviewDialog())
            {
                ppd.Document = _printDoc;
                ppd.WindowState = FormWindowState.Maximized;
                ppd.ShowDialog(this);
            }
        }

        private int CalculatePrintTotalPages()
        {
            if (_dtFiltered == null) return 1;

            // Use a dummy measurement
            var settings = new System.Drawing.Printing.PageSettings();
            float pageH = 1100f; // approximate A4 landscape in pixels at 100dpi
            float margin = 60f;
            float usableH = pageH - margin * 2;
            float headerH = 80f; // header block height estimate
            float rowH = 20f;
            float footerH = 20f;

            float availableForRows = usableH - headerH - footerH;
            int rowsPerPage = Math.Max(1, (int)(availableForRows / rowH));
            int total = (int)Math.Ceiling((double)_dtFiltered.Rows.Count / rowsPerPage);
            return Math.Max(1, total);
        }

        private void PrintDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            _printCurrentPage++;
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            float margin = 60f;
            float pageWidth = e.PageBounds.Width;
            float pageHeight = e.PageBounds.Height;
            float contentWidth = pageWidth - margin * 2;
            float y = margin;

            // ---- Header ----
            // Company name
            using (var fBig = new Font("Segoe UI", 16F, FontStyle.Bold))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center })
            {
                g.DrawString("Persal", fBig, Brushes.Black,
                    new RectangleF(margin, y, contentWidth, 24), sf);
                y += 26;
            }

            // Report title
            using (var fTitle = new Font("Segoe UI", 12F, FontStyle.Bold))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center })
            {
                g.DrawString("Location Count Status Report", fTitle, Brushes.Black,
                    new RectangleF(margin, y, contentWidth, 20), sf);
                y += 22;
            }

            // Printed on — right aligned
            using (var fSmall = new Font("Segoe UI", 9F))
            using (var sf = new StringFormat { Alignment = StringAlignment.Far })
            {
                string printedOn = string.Format("Printed on: {0}", DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                g.DrawString(printedOn, fSmall, Brushes.Gray,
                    new RectangleF(margin, y, contentWidth, 16), sf);
                y += 18;
            }

            // Horizontal rule
            g.DrawLine(Pens.LightGray, margin, y, pageWidth - margin, y);
            y += 8;

            // Column headers
            float[] colWidths = GetPrintColWidths(contentWidth);
            using (var fHeader = new Font("Segoe UI", 9F, FontStyle.Bold))
            {
                float x = margin;
                using (var headerBrush = new SolidBrush(Color.FromArgb(229, 231, 235)))
                    g.FillRectangle(headerBrush, margin, y, contentWidth, 18);

                for (int c = 0; c < _printColHeaders.Length; c++)
                {
                    g.DrawString(_printColHeaders[c], fHeader, Brushes.Black,
                        new RectangleF(x + 2, y, colWidths[c] - 4, 18));
                    x += colWidths[c];
                }
                y += 20;
            }

            // ---- Data rows ----
            float bottomLimit = pageHeight - margin - 20; // leave room for footer
            using (var fRow = new Font("Segoe UI", 8.5F))
            {
                while (_printRowIndex < _dtFiltered.Rows.Count)
                {
                    if (y + 18 > bottomLimit) break;

                    DataRow dr = _dtFiltered.Rows[_printRowIndex];
                    float x = margin;

                    // Alternate row background
                    Color rowBg = (_printRowIndex % 2 == 0) ? Color.White : Color.FromArgb(249, 250, 251);
                    using (var rowBrush = new SolidBrush(rowBg))
                        g.FillRectangle(rowBrush, margin, y, contentWidth, 18);

                    for (int c = 0; c < _printColFields.Length; c++)
                    {
                        string fieldName = _printColFields[c];
                        string val = "";

                        if (_dtFiltered.Columns.Contains(fieldName))
                        {
                            object v = dr[fieldName];
                            if (v == null || v == DBNull.Value)
                                val = "";
                            else if (fieldName == "Last_Count_Date" || fieldName == "Next_Due_Date")
                            {
                                DateTime dt;
                                if (DateTime.TryParse(v.ToString(), out dt))
                                    val = dt.ToString("MM/dd/yyyy");
                            }
                            else
                                val = Convert.ToString(v);
                        }

                        g.DrawString(val, fRow, Brushes.Black,
                            new RectangleF(x + 2, y + 1, colWidths[c] - 4, 17),
                            new StringFormat { FormatFlags = StringFormatFlags.NoWrap });
                        x += colWidths[c];
                    }

                    y += 18;
                    _printRowIndex++;
                }
            }

            // ---- Footer ----
            string pageLabel = string.Format("Page {0} of {1}", _printCurrentPage, _printTotalPages);
            using (var fSmall = new Font("Segoe UI", 8F))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center })
            {
                g.DrawString(pageLabel, fSmall, Brushes.Gray,
                    new RectangleF(margin, pageHeight - margin, contentWidth, 16), sf);
            }

            e.HasMorePages = (_printRowIndex < _dtFiltered.Rows.Count);
        }

        private float[] GetPrintColWidths(float totalWidth)
        {
            int total = 0;
            foreach (int w in _printColWidths) total += w;

            float[] result = new float[_printColWidths.Length];
            for (int i = 0; i < _printColWidths.Length; i++)
                result[i] = (float)_printColWidths[i] / total * totalWidth;

            return result;
        }

        #endregion

        // =====================================================================
        // #region EXCEL EXPORT
        // =====================================================================
        #region EXCEL EXPORT

        private void ExportToExcel()
        {
            if (_dtFiltered == null || _dtFiltered.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Excel File";
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                sfd.FileName = string.Format("LocationStatus_{0}.xlsx", DateTime.Now.ToString("yyyyMMdd_HHmmss"));

                if (sfd.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Location Status");

                        string[] exportCols = { "Location_Name", "Area_Name", "Period_Name", "Last_Count_Date", "Next_Due_Date", "Days_Overdue", "Status" };
                        string[] exportHeaders = { "Location", "Area", "Period", "Last Count", "Next Due", "Days Overdue", "Status" };
                        int lastCol = exportHeaders.Length;
                        string lastColLetter = GetExcelColumnLetter(lastCol);

                        // Row 1: Company name
                        ws.Range(string.Format("A1:{0}1", lastColLetter)).Merge();
                        var r1 = ws.Cell("A1");
                        r1.Value = "Persal";
                        r1.Style.Font.Bold = true;
                        r1.Style.Font.FontSize = 14;
                        r1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        r1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        ws.Row(1).Height = 24;

                        // Row 2: Report title
                        ws.Range(string.Format("A2:{0}2", lastColLetter)).Merge();
                        var r2 = ws.Cell("A2");
                        r2.Value = "Location Count Status Report";
                        r2.Style.Font.Bold = true;
                        r2.Style.Font.FontSize = 12;
                        r2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        r2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        ws.Row(2).Height = 20;

                        // Row 3: Exported on
                        ws.Range(string.Format("A3:{0}3", lastColLetter)).Merge();
                        var r3 = ws.Cell("A3");
                        r3.Value = string.Format("Exported on: {0}", DateTime.Now.ToString("MM/dd/yyyy hh:mm tt"));
                        r3.Style.Font.Italic = true;
                        r3.Style.Font.FontSize = 10;
                        r3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        r3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        ws.Row(3).Height = 16;

                        // Row 4: empty
                        ws.Row(4).Height = 8;

                        // Row 5: Column headers
                        ws.Row(5).Height = 18;
                        for (int c = 0; c < exportHeaders.Length; c++)
                        {
                            var hCell = ws.Cell(5, c + 1);
                            hCell.Value = exportHeaders[c];
                            hCell.Style.Font.Bold = true;
                            hCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1F2937");
                            hCell.Style.Font.FontColor = XLColor.White;
                            hCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }

                        // Data rows starting at row 6
                        for (int r = 0; r < _dtFiltered.Rows.Count; r++)
                        {
                            DataRow dr = _dtFiltered.Rows[r];
                            int xlRow = r + 6;
                            XLColor rowBg = (r % 2 == 0) ? XLColor.White : XLColor.FromHtml("#F9FAFB");

                            for (int c = 0; c < exportCols.Length; c++)
                            {
                                string fieldName = exportCols[c];
                                var cell = ws.Cell(xlRow, c + 1);

                                if (_dtFiltered.Columns.Contains(fieldName))
                                {
                                    object v = dr[fieldName];
                                    if (v == null || v == DBNull.Value)
                                        cell.Value = "";
                                    else if (fieldName == "Last_Count_Date" || fieldName == "Next_Due_Date")
                                    {
                                        DateTime dt;
                                        if (DateTime.TryParse(v.ToString(), out dt))
                                        {
                                            cell.Value = dt;
                                            cell.Style.DateFormat.Format = "MM/dd/yyyy";
                                        }
                                        else
                                            cell.Value = "";
                                    }
                                    else if (fieldName == "Days_Overdue")
                                    {
                                        int daysVal;
                                        if (int.TryParse(v.ToString(), out daysVal))
                                            cell.Value = daysVal;
                                        else
                                            cell.Value = "";
                                    }
                                    else
                                        cell.Value = Convert.ToString(v);
                                }

                                // Status cell background
                                if (fieldName == "Status")
                                {
                                    string st = Convert.ToString(dr[fieldName]);
                                    if (st == "On Time")
                                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D1FAE5");
                                    else if (st == "Overdue")
                                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FEF3C7");
                                    else if (st == "Never Counted")
                                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FEE2E2");
                                    else
                                        cell.Style.Fill.BackgroundColor = rowBg;
                                }
                                else
                                {
                                    cell.Style.Fill.BackgroundColor = rowBg;
                                }
                            }
                        }

                        // Auto-fit columns
                        ws.Columns().AdjustToContents();

                        wb.SaveAs(sfd.FileName);
                    }

                    MessageBox.Show("Export completed successfully.", "Export",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Error exporting: {0}", ex.Message), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string GetExcelColumnLetter(int colNumber)
        {
            string result = "";
            while (colNumber > 0)
            {
                int mod = (colNumber - 1) % 26;
                result = Convert.ToChar('A' + mod) + result;
                colNumber = (colNumber - mod) / 26;
            }
            return result;
        }

        #endregion

        // =====================================================================
        // #region UI STYLE HELPERS
        // =====================================================================
        #region UI STYLE HELPERS

        private Button CreateChromeButton(string text)
        {
            var b = new Button
            {
                Text = text,
                Width = 44,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(31, 41, 55),
                Cursor = Cursors.Hand,
                Margin = new Padding(6, 0, 0, 0),
                TabStop = false
            };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(55, 65, 81);
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(17, 24, 39);
            return b;
        }

        private void StyleButtonPrimary(Button b)
        {
            if (b == null) return;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = Color.FromArgb(37, 99, 235);
            b.ForeColor = Color.White;
            b.Height = 40;
            b.Width = Math.Max(140, b.Width);
            b.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            b.Cursor = Cursors.Hand;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(29, 78, 216);
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(30, 64, 175);
        }

        private void StyleButtonGhost(Button b)
        {
            if (b == null) return;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = Color.White;
            b.ForeColor = Color.FromArgb(107, 114, 128);
            b.Height = 40;
            b.Width = Math.Max(120, b.Width);
            b.Font = new Font("Segoe UI", 9F);
            b.Cursor = Cursors.Hand;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(249, 250, 251);
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(243, 244, 246);
        }

        private void StyleGrid(DataGridView dgv)
        {
            if (dgv == null) return;

            dgv.BorderStyle = BorderStyle.None;
            dgv.BackgroundColor = Color.White;
            dgv.GridColor = Color.FromArgb(229, 231, 235);
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(243, 244, 246);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 41, 55);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 38;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(17, 24, 39);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgv.DefaultCellStyle.SelectionForeColor = Color.FromArgb(17, 24, 39);
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            dgv.RowTemplate.Height = 32;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 250, 251);
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.RowHeadersVisible = false;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgv.ScrollBars = ScrollBars.Both;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        #endregion

        // =====================================================================
        // #region ROUNDED REGION & WINDOW CHROME
        // =====================================================================
        #region ROUNDED REGION AND WINDOW CHROME

        private void ToggleMaximize()
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
                ApplyRoundedRegion(this, FormCornerRadius);
                if (_btnMaximizar != null) _btnMaximizar.Text = "□";
            }
            else
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Maximized;
                this.Region = null;
                if (_btnMaximizar != null) _btnMaximizar.Text = "❐";
            }
        }

        private void ApplyRoundedRegion(Control c, int radius)
        {
            try
            {
                if (c.Width <= 0 || c.Height <= 0) return;
                using (var path = BuildRoundRectPath(new Rectangle(0, 0, c.Width, c.Height), radius))
                    c.Region = new Region(path);
            }
            catch { }
        }

        private GraphicsPath BuildRoundRectPath(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;

            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        #endregion

        // =====================================================================
        // #region MODERN CARD PANEL (inner class)
        // =====================================================================
        #region MODERN CARD PANEL

        private class ModernCardPanel : Panel
        {
            public int Radius { get; set; } = 10;
            public Color BorderColor { get; set; } = Color.FromArgb(229, 231, 235);
            public int BorderThickness { get; set; } = 1;

            public ModernCardPanel()
            {
                this.DoubleBuffered = true;
                this.BackColor = Color.White;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                var rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
                int d = Radius * 2;

                using (var path = new GraphicsPath())
                {
                    path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                    path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                    path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                    path.CloseFigure();

                    using (var pen = new Pen(BorderColor, BorderThickness))
                        e.Graphics.DrawPath(pen, path);
                }
            }
        }

        #endregion

    } // end class Physical_Count_Location_Status
} // end namespace
