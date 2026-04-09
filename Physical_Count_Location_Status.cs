// ============================================================================
// Physical_Count_Location_Status.cs
// Reporte: Estatus de Locaciones para Conteo Cíclico
// Compatible con .NET Framework 4 (sin interpolación, sin ?. en tipos valor)
// ============================================================================

using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ClosedXML.Excel;
using DevExpress.Utils;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;

namespace Persal._003_Physical_Counting_2
{
    // =========================================================================
    // Physical_Count_Location_Status — Formulario de estatus de locaciones
    // =========================================================================
    public class Physical_Count_Location_Status : Form
    {
        // =====================================================================
        // #region CONSTANTES
        // =====================================================================
        #region CONSTANTES

        private const bool UseCustomChrome  = true;
        private const int  FormCornerRadius = 10;

        #endregion

        // =====================================================================
        // #region CAMPOS — DATOS
        // =====================================================================
        #region CAMPOS — DATOS

        private string    _cs         = "";
        private DataTable _dtFull     = null;
        private DataTable _dtFiltered = null;

        #endregion

        // =====================================================================
        // #region CAMPOS — CONTROLES DE FILTRO
        // =====================================================================
        #region CAMPOS — CONTROLES DE FILTRO

        private ComboBox _cboShow;
        private ComboBox _cboArea;
        private ComboBox _cboPeriod;
        private Button   _btnSearch;
        private Button   _btnPrint;
        private Button   _btnExport;

        #endregion

        // =====================================================================
        // #region CAMPOS — GRID DEVEXPRESS
        // =====================================================================
        #region CAMPOS — GRID DEVEXPRESS

        private GridControl _gcMain;
        private GridView    _gvMain;

        #endregion

        // =====================================================================
        // #region CAMPOS — FOOTER
        // =====================================================================
        #region CAMPOS — FOOTER

        private Label _lblTotal;
        private Label _lblOnTime;
        private Label _lblOverdue;
        private Label _lblNever;

        #endregion

        // =====================================================================
        // #region CAMPOS — IMPRESIÓN
        // =====================================================================
        #region CAMPOS — IMPRESIÓN

        private PrintDocument    _pd;
        private PrintPreviewDialog _ppd;
        private int _printRowIndex;
        private int _printCurrentPage;
        private int _printTotalPages;

        private readonly string[] _printColHeaders = {
            "Locación", "Área", "Período",
            "Últ. Conteo", "Próx. Venc.", "Días Venc.", "Estatus"
        };

        private readonly string[] _printColFields = {
            "Location_Name", "Area_Name", "Period_Name",
            "Last_Count_Date", "Next_Due_Date", "Days_Overdue", "Status"
        };

        #endregion

        // =====================================================================
        // #region CONSTRUCTORES
        // =====================================================================
        #region CONSTRUCTORES

        public Physical_Count_Location_Status(string connectionString)
        {
            _cs = (connectionString != null) ? connectionString : "";
            InitForm();
        }

        public Physical_Count_Location_Status()
        {
            _cs = "";
            InitForm();
        }

        #endregion

        // =====================================================================
        // #region INICIALIZACIÓN
        // =====================================================================
        #region INICIALIZACIÓN

        private void InitForm()
        {
            this.SuspendLayout();

            // Propiedades del formulario
            this.Text             = "Estatus de Locaciones";
            this.Font             = new Font("Segoe UI", 9F);
            this.BackColor        = Color.FromArgb(243, 244, 246);
            this.MinimumSize      = new Size(1100, 680);
            this.Size             = new Size(1200, 760);
            this.StartPosition    = FormStartPosition.CenterScreen;

            if (UseCustomChrome)
            {
                this.FormBorderStyle = FormBorderStyle.None;
                this.MaximizedBounds = Screen.PrimaryScreen.WorkingArea;
            }

            // Construir interfaz
            BuildUi();

            // Conectar eventos de botones y combos
            WireEvents();

            // Eventos del formulario
            this.Load   += Form_Load;
            this.Resize += Form_Resize;

            this.ResumeLayout(false);
        }

        private void Form_Load(object sender, EventArgs e)
        {
            if (UseCustomChrome)
                ApplyRoundedRegion(this, FormCornerRadius);

            LoadCombos();
            LoadData();
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            if (UseCustomChrome)
                ApplyRoundedRegion(this, FormCornerRadius);
        }

        #endregion

        // =====================================================================
        // #region CONSTRUCCIÓN DE LA UI
        // =====================================================================
        #region CONSTRUCCIÓN DE LA UI

        private void BuildUi()
        {
            // -----------------------------------------------------------------
            // Barra de título (header chrome)
            // -----------------------------------------------------------------
            var header = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 48,
                BackColor = Color.FromArgb(31, 41, 55)
            };

            var lblTitle = new Label
            {
                Text      = "Estatus de Locaciones",
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 12F, FontStyle.Bold),
                AutoSize  = true,
                Location  = new Point(16, 13)
            };

            var btnClose = CreateChromeButton("✕");
            var btnMax   = CreateChromeButton("□");
            var btnMin   = CreateChromeButton("─");

            btnClose.Click += (s, e) => this.Close();
            btnMax.Click   += (s, e) =>
            {
                this.WindowState = (this.WindowState == FormWindowState.Maximized)
                    ? FormWindowState.Normal
                    : FormWindowState.Maximized;
            };
            btnMin.Click   += (s, e) => this.WindowState = FormWindowState.Minimized;

            // Anclar botones al lado derecho
            btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnMax.Anchor   = AnchorStyles.Top | AnchorStyles.Right;
            btnMin.Anchor   = AnchorStyles.Top | AnchorStyles.Right;

            btnClose.Location = new Point(this.Width - 50,  10);
            btnMax.Location   = new Point(this.Width - 96,  10);
            btnMin.Location   = new Point(this.Width - 142, 10);

            header.Controls.Add(lblTitle);
            header.Controls.Add(btnClose);
            header.Controls.Add(btnMax);
            header.Controls.Add(btnMin);

            header.MouseDown  += Header_MouseDown;
            lblTitle.MouseDown += Header_MouseDown;

            // -----------------------------------------------------------------
            // Panel de filtros
            // -----------------------------------------------------------------
            var pnlFilter = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 60,
                BackColor = Color.White,
                Padding   = new Padding(12, 0, 12, 0)
            };

            pnlFilter.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(229, 231, 235)))
                    e.Graphics.DrawLine(pen, 0, pnlFilter.Height - 1,
                                             pnlFilter.Width, pnlFilter.Height - 1);
            };

            _cboShow   = MakeFilterCombo(120);
            _cboArea   = MakeFilterCombo(150);
            _cboPeriod = MakeFilterCombo(150);
            _btnSearch = MakeButton("🔍 Buscar",   true);
            _btnPrint  = MakeButton("🖨 Imprimir",  false);
            _btnExport = MakeButton("📊 Exportar",  false);

            var flow = new FlowLayoutPanel
            {
                Dock          = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents  = false,
                AutoScroll    = false,
                BackColor     = Color.White,
                Padding       = new Padding(0, 10, 0, 0)
            };

            // Etiquetas y combos de filtro
            var lblShow = MakeFilterLabel("Mostrar:");
            var lblArea = MakeFilterLabel("Área:");
            var lblPer  = MakeFilterLabel("Período:");

            Control[] filterControls = {
                lblShow, _cboShow,
                lblArea, _cboArea,
                lblPer,  _cboPeriod,
                _btnSearch, _btnPrint, _btnExport
            };

            foreach (Control c in filterControls)
            {
                c.Margin = new Padding(4, 4, 4, 0);
                flow.Controls.Add(c);
            }

            pnlFilter.Controls.Add(flow);

            // -----------------------------------------------------------------
            // Footer de resumen
            // -----------------------------------------------------------------
            var pnlFooter = new Panel
            {
                Dock      = DockStyle.Bottom,
                Height    = 40,
                BackColor = Color.White
            };

            pnlFooter.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(229, 231, 235)))
                    e.Graphics.DrawLine(pen, 0, 0, pnlFooter.Width, 0);
            };

            _lblTotal   = MakeFooterLabel();
            _lblOnTime  = MakeFooterLabel();
            _lblOverdue = MakeFooterLabel();
            _lblNever   = MakeFooterLabel();

            var footerFlow = new FlowLayoutPanel
            {
                Dock          = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents  = false,
                BackColor     = Color.White,
                Padding       = new Padding(12, 8, 12, 0)
            };

            Label[] footerLabels = { _lblTotal, _lblOnTime, _lblOverdue, _lblNever };
            foreach (Label lbl in footerLabels)
            {
                lbl.Margin = new Padding(0, 0, 24, 0);
                footerFlow.Controls.Add(lbl);
            }

            pnlFooter.Controls.Add(footerFlow);

            // -----------------------------------------------------------------
            // Grid DevExpress XtraGrid
            // -----------------------------------------------------------------
            _gcMain            = new GridControl { Dock = DockStyle.Fill };
            _gvMain            = new GridView();
            _gcMain.MainView   = _gvMain;
            _gcMain.ViewCollection.Add(_gvMain);

            _gvMain.OptionsView.ShowGroupPanel          = false;
            _gvMain.OptionsView.ColumnAutoWidth         = false;
            _gvMain.OptionsView.ShowIndicator           = false;
            _gvMain.OptionsBehavior.Editable            = false;
            _gvMain.OptionsSelection.MultiSelect        = false;
            _gvMain.OptionsView.EnableAppearanceOddRow  = true;
            _gvMain.OptionsView.EnableAppearanceEvenRow = true;

            _gvMain.RowStyle       += GvMain_RowStyle;
            _gvMain.CustomDrawCell += GvMain_CustomDrawCell;

            BuildGridColumns();

            // Panel contenedor con sangría/padding
            var pnlContent = new Panel
            {
                Dock      = DockStyle.Fill,
                Padding   = new Padding(12),
                BackColor = Color.FromArgb(243, 244, 246)
            };
            pnlContent.Controls.Add(_gcMain);

            // -----------------------------------------------------------------
            // Agregar controles al formulario (orden inverso para Dock)
            // -----------------------------------------------------------------
            this.Controls.Add(pnlContent);
            this.Controls.Add(pnlFooter);
            this.Controls.Add(pnlFilter);
            this.Controls.Add(header);
        }

        #endregion

        // =====================================================================
        // #region HELPERS DE CONSTRUCCIÓN DE UI
        // =====================================================================
        #region HELPERS DE CONSTRUCCIÓN DE UI

        private Label MakeFilterLabel(string text)
        {
            return new Label
            {
                Text      = text,
                AutoSize  = true,
                ForeColor = Color.FromArgb(55, 65, 81),
                Font      = new Font("Segoe UI", 9F),
                Margin    = new Padding(0, 8, 4, 0)
            };
        }

        private ComboBox MakeFilterCombo(int width)
        {
            return new ComboBox
            {
                Width         = width,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font          = new Font("Segoe UI", 9F),
                FlatStyle     = FlatStyle.Flat
            };
        }

        private Button MakeButton(string text, bool primary)
        {
            var btn = new Button
            {
                Text      = text,
                Height    = 30,
                Width     = 115,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor    = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;

            if (primary)
            {
                btn.BackColor                         = Color.FromArgb(37, 99, 235);
                btn.ForeColor                         = Color.White;
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(29, 78, 216);
                btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(30, 64, 175);
            }
            else
            {
                btn.BackColor                         = Color.White;
                btn.ForeColor                         = Color.FromArgb(55, 65, 81);
                btn.FlatAppearance.BorderSize         = 1;
                btn.FlatAppearance.BorderColor        = Color.FromArgb(209, 213, 219);
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(243, 244, 246);
                btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(229, 231, 235);
            }

            return btn;
        }

        private Label MakeFooterLabel()
        {
            return new Label
            {
                AutoSize  = true,
                ForeColor = Color.FromArgb(55, 65, 81),
                Font      = new Font("Segoe UI", 9F),
                Text      = ""
            };
        }

        private Button CreateChromeButton(string text)
        {
            var b = new Button
            {
                Text      = text,
                Width     = 44,
                Height    = 28,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(31, 41, 55),
                Cursor    = Cursors.Hand,
                TabStop   = false
            };
            b.FlatAppearance.BorderSize             = 0;
            b.FlatAppearance.MouseOverBackColor     = Color.FromArgb(55, 65, 81);
            b.FlatAppearance.MouseDownBackColor     = Color.FromArgb(17, 24, 39);
            return b;
        }

        #endregion

        // =====================================================================
        // #region COLUMNAS DEL GRID
        // =====================================================================
        #region COLUMNAS DEL GRID

        private void BuildGridColumns()
        {
            _gvMain.Columns.Clear();

            // Columna oculta — ID interno
            var colId = new GridColumn
            {
                FieldName = "Location_ID",
                Caption   = "ID",
                Visible   = false
            };
            _gvMain.Columns.Add(colId);

            // Locación
            var colName = new GridColumn
            {
                FieldName    = "Location_Name",
                Caption      = "Locación",
                Width        = 200,
                VisibleIndex = 0
            };
            colName.Options.AutoFilterCondition = AutoFilterCondition.Contains;
            _gvMain.Columns.Add(colName);

            // Área
            var colArea = new GridColumn
            {
                FieldName    = "Area_Name",
                Caption      = "Área",
                Width        = 150,
                VisibleIndex = 1
            };

            // Período
            var colPer = new GridColumn
            {
                FieldName    = "Period_Name",
                Caption      = "Período",
                Width        = 120,
                VisibleIndex = 2
            };

            // Último Conteo
            var colLD = new GridColumn
            {
                FieldName    = "Last_Count_Date",
                Caption      = "Último Conteo",
                Width        = 110,
                VisibleIndex = 3
            };
            colLD.DisplayFormat.FormatType   = FormatType.DateTime;
            colLD.DisplayFormat.FormatString = "MM/dd/yyyy";

            // Inventario (nombre del último conteo)
            var colLN = new GridColumn
            {
                FieldName    = "Last_Count_Name",
                Caption      = "Inventario",
                Width        = 180,
                VisibleIndex = 4
            };

            // Próxima Vencimiento
            var colND = new GridColumn
            {
                FieldName    = "Next_Due_Date",
                Caption      = "Próx. Venc.",
                Width        = 110,
                VisibleIndex = 5
            };
            colND.DisplayFormat.FormatType   = FormatType.DateTime;
            colND.DisplayFormat.FormatString = "MM/dd/yyyy";

            // Días Vencidos
            var colDays = new GridColumn
            {
                FieldName    = "Days_Overdue",
                Caption      = "Días Vencidos",
                Width        = 100,
                VisibleIndex = 6
            };

            // Estatus
            var colSt = new GridColumn
            {
                FieldName    = "Status",
                Caption      = "Estatus",
                Width        = 130,
                VisibleIndex = 7
            };

            _gvMain.Columns.AddRange(new GridColumn[]
            {
                colArea, colPer, colLD, colLN, colND, colDays, colSt
            });
        }

        #endregion

        // =====================================================================
        // #region EVENTOS Y CONEXIÓN
        // =====================================================================
        #region EVENTOS Y CONEXIÓN

        private void WireEvents()
        {
            _btnSearch.Click += (s, e) => LoadData();
            _btnPrint.Click  += BtnPrint_Click;
            _btnExport.Click += BtnExport_Click;
            _cboShow.SelectedIndexChanged += (s, e) => ApplyShowFilter();
        }

        #endregion

        // =====================================================================
        // #region CARGA DE COMBOS
        // =====================================================================
        #region CARGA DE COMBOS

        private void LoadCombos()
        {
            // --- Combo "Mostrar" ---
            var dtShow = new DataTable();
            dtShow.Columns.Add("Val", typeof(int));
            dtShow.Columns.Add("Txt", typeof(string));
            dtShow.Rows.Add(0, "Todos");
            dtShow.Rows.Add(1, "Vencido");
            dtShow.Rows.Add(2, "Nunca Contado");
            dtShow.Rows.Add(3, "Al Día");

            _cboShow.DataSource    = dtShow;
            _cboShow.ValueMember   = "Val";
            _cboShow.DisplayMember = "Txt";
            _cboShow.SelectedIndex = 0;

            // --- Combo "Área" ---
            try
            {
                var dtArea = ExecFillTable(
                    "SELECT 0 AS Area_ID, 'Todas' AS Area_Name " +
                    "UNION ALL " +
                    "SELECT Location_ID, Location_Name " +
                    "  FROM dbo.Locations " +
                    " WHERE Location_Level = 2 AND Is_Active = 1 " +
                    " ORDER BY Area_Name;");

                _cboArea.DataSource    = dtArea;
                _cboArea.ValueMember   = "Area_ID";
                _cboArea.DisplayMember = "Area_Name";
                _cboArea.SelectedIndex = 0;
            }
            catch { /* Dejar combo vacío si falla */ }

            // --- Combo "Período" ---
            try
            {
                var dtPer = ExecFillTable(
                    "SELECT 0 AS Period_ID, 'Todos' AS Period_Name " +
                    "UNION ALL " +
                    "SELECT Cyclic_Period_ID, Period_Name " +
                    "  FROM dbo.Cyclic_Count_Periods " +
                    " WHERE Is_Active = 1 " +
                    " ORDER BY Period_Name;");

                _cboPeriod.DataSource    = dtPer;
                _cboPeriod.ValueMember   = "Period_ID";
                _cboPeriod.DisplayMember = "Period_Name";
                _cboPeriod.SelectedIndex = 0;
            }
            catch { /* Dejar combo vacío si falla */ }
        }

        #endregion

        // =====================================================================
        // #region CARGA Y FILTRADO DE DATOS
        // =====================================================================
        #region CARGA Y FILTRADO DE DATOS

        private void LoadData()
        {
            try
            {
                int areaId   = 0;
                int periodId = 0;

                if (_cboArea.SelectedValue != null)
                    int.TryParse(_cboArea.SelectedValue.ToString(), out areaId);

                if (_cboPeriod.SelectedValue != null)
                    int.TryParse(_cboPeriod.SelectedValue.ToString(), out periodId);

                string sql = BuildSql();
                _dtFull = ExecFillTable(sql, areaId, periodId);
            }
            catch (Exception ex)
            {
                _dtFull = new DataTable();
                MessageBox.Show(
                    string.Format("Error al cargar datos: {0}", ex.Message),
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ApplyShowFilter();
        }

        private string BuildSql()
        {
            return @"
WITH LastCount AS (
    SELECT
        PCS.Location_ID,
        MAX(PC.End_Date) AS Last_Count_Date
    FROM dbo.Physical_Count_Snapshots PCS
    INNER JOIN dbo.Physical_Counts PC
        ON PC.Physical_Count_ID = PCS.Physical_Count_ID
    WHERE PC.Status_ID = 3
    GROUP BY PCS.Location_ID
),
LastCountWithName AS (
    SELECT
        LC.Location_ID,
        LC.Last_Count_Date,
        PC2.Physical_Count_Name AS Last_Count_Name
    FROM LastCount LC
    INNER JOIN dbo.Physical_Count_Snapshots PCS2
        ON PCS2.Location_ID = LC.Location_ID
    INNER JOIN dbo.Physical_Counts PC2
        ON PC2.Physical_Count_ID = PCS2.Physical_Count_ID
       AND PC2.End_Date          = LC.Last_Count_Date
       AND PC2.Status_ID         = 3
)
SELECT
    L.Location_ID,
    L.Location_Name,
    ISNULL(A.Location_Name, '')     AS Area_Name,
    ISNULL(CCP.Period_Name, '')     AS Period_Name,
    LCN.Last_Count_Date,
    ISNULL(LCN.Last_Count_Name, '') AS Last_Count_Name,
    CASE
        WHEN LCN.Last_Count_Date IS NULL THEN NULL
        ELSE DATEADD(DAY, ISNULL(CCP.Days_Between_Counts, 365), LCN.Last_Count_Date)
    END AS Next_Due_Date,
    CASE
        WHEN LCN.Last_Count_Date IS NULL THEN NULL
        WHEN DATEDIFF(DAY,
             DATEADD(DAY, ISNULL(CCP.Days_Between_Counts, 365), LCN.Last_Count_Date),
             GETDATE()) > 0
        THEN DATEDIFF(DAY,
             DATEADD(DAY, ISNULL(CCP.Days_Between_Counts, 365), LCN.Last_Count_Date),
             GETDATE())
        ELSE 0
    END AS Days_Overdue,
    CASE
        WHEN LCN.Last_Count_Date IS NULL                    THEN 'Nunca Contado'
        WHEN DATEDIFF(DAY,
             DATEADD(DAY, ISNULL(CCP.Days_Between_Counts, 365), LCN.Last_Count_Date),
             GETDATE()) > 0                                 THEN 'Vencido'
        ELSE                                                     'Al Día'
    END AS Status
FROM dbo.Locations L
LEFT JOIN dbo.Locations A
    ON A.Location_ID = L.Parent_Location_ID
LEFT JOIN dbo.Cyclic_Count_Periods CCP
    ON CCP.Cyclic_Period_ID = L.Cyclic_Period_ID
LEFT JOIN LastCountWithName LCN
    ON LCN.Location_ID = L.Location_ID
WHERE L.Is_Active       = 1
  AND L.Location_Level  = 3
  AND (@p1 = 0 OR L.Parent_Location_ID = @p1)
  AND (@p2 = 0 OR L.Cyclic_Period_ID   = @p2)
ORDER BY
    CASE Status
        WHEN 'Vencido'       THEN 1
        WHEN 'Nunca Contado' THEN 2
        ELSE                      3
    END,
    L.Location_Name ASC;";
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

            if (_dtFull == null) _dtFull = new DataTable();

            if (showVal == 0)
            {
                _dtFiltered = _dtFull.Copy();
            }
            else
            {
                string filterStatus;
                switch (showVal)
                {
                    case 1:  filterStatus = "Vencido";       break;
                    case 2:  filterStatus = "Nunca Contado"; break;
                    case 3:  filterStatus = "Al Día";        break;
                    default: filterStatus = "";              break;
                }

                _dtFiltered = _dtFull.Clone();
                foreach (DataRow row in _dtFull.Rows)
                {
                    if (Convert.ToString(row["Status"]) == filterStatus)
                        _dtFiltered.ImportRow(row);
                }
            }

            _gcMain.DataSource = _dtFiltered;
            UpdateFooterSummary();
        }

        #endregion

        // =====================================================================
        // #region FOOTER — RESUMEN
        // =====================================================================
        #region FOOTER — RESUMEN

        private void UpdateFooterSummary()
        {
            int total = 0, onTime = 0, overdue = 0, never = 0;

            if (_dtFiltered != null)
            {
                total = _dtFiltered.Rows.Count;
                foreach (DataRow row in _dtFiltered.Rows)
                {
                    string s = Convert.ToString(row["Status"]);
                    if (s == "Al Día")           onTime++;
                    else if (s == "Vencido")     overdue++;
                    else if (s == "Nunca Contado") never++;
                }
            }

            _lblTotal.Text   = string.Format("Total: {0}", total);
            _lblOnTime.Text  = string.Format("✅ Al Día: {0}", onTime);
            _lblOverdue.Text = string.Format("⚠ Vencido: {0}", overdue);
            _lblNever.Text   = string.Format("❌ Nunca Contado: {0}", never);
        }

        #endregion

        // =====================================================================
        // #region COLORES DE FILA — DEVEXPRESS ROWSTYLE
        // =====================================================================
        #region COLORES DE FILA — DEVEXPRESS ROWSTYLE

        private void GvMain_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle < 0) return;

            var view   = (GridView)sender;
            string status = Convert.ToString(view.GetRowCellValue(e.RowHandle, "Status"));

            switch (status)
            {
                case "Vencido":
                    e.Appearance.BackColor  = Color.FromArgb(254, 243, 199);
                    e.Appearance.BackColor2 = Color.FromArgb(254, 243, 199);
                    e.HighPriority = true;
                    break;

                case "Nunca Contado":
                    e.Appearance.BackColor  = Color.FromArgb(254, 226, 226);
                    e.Appearance.BackColor2 = Color.FromArgb(254, 226, 226);
                    e.HighPriority = true;
                    break;

                case "Al Día":
                    e.Appearance.BackColor  = Color.FromArgb(220, 252, 231);
                    e.Appearance.BackColor2 = Color.FromArgb(220, 252, 231);
                    e.HighPriority = true;
                    break;
            }
        }

        #endregion

        // =====================================================================
        // #region PILL/BADGE EN COLUMNA ESTATUS — CUSTOMDRAWCELL
        // =====================================================================
        #region PILL/BADGE EN COLUMNA ESTATUS — CUSTOMDRAWCELL

        private void GvMain_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName != "Status") return;

            string status = Convert.ToString(e.CellValue);
            Color  fillColor;
            Color  textColor;

            switch (status)
            {
                case "Al Día":
                    fillColor = Color.FromArgb(209, 250, 229);
                    textColor = Color.FromArgb(6,   95,  70);
                    break;

                case "Vencido":
                    fillColor = Color.FromArgb(254, 243, 199);
                    textColor = Color.FromArgb(120,  53,  15);
                    break;

                case "Nunca Contado":
                    fillColor = Color.FromArgb(254, 226, 226);
                    textColor = Color.FromArgb(153,  27,  27);
                    break;

                default:
                    return;
            }

            e.Handled = true;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Fondo de la celda (usa el color de la fila para consistencia)
            Color bg = (e.Appearance.BackColor == Color.Empty)
                ? Color.White
                : e.Appearance.BackColor;

            using (var bgBrush = new SolidBrush(bg))
                e.Graphics.FillRectangle(bgBrush, e.Bounds);

            // Pill redondeado
            int pH = 20;
            int pW = Math.Max(0, e.Bounds.Width - 16);
            int pX = e.Bounds.X + 8;
            int pY = e.Bounds.Y + (e.Bounds.Height - pH) / 2;

            var pillRect = new Rectangle(pX, pY, pW, pH);
            int rad = pH / 2;
            int d   = rad * 2;

            using (var path = new GraphicsPath())
            {
                path.AddArc(pillRect.X,             pillRect.Y,             d, d, 180, 90);
                path.AddArc(pillRect.Right - d,     pillRect.Y,             d, d, 270, 90);
                path.AddArc(pillRect.Right - d,     pillRect.Bottom - d,    d, d,   0, 90);
                path.AddArc(pillRect.X,             pillRect.Bottom - d,    d, d,  90, 90);
                path.CloseFigure();

                using (var fb = new SolidBrush(fillColor))
                    e.Graphics.FillPath(fb, path);
            }

            // Texto del pill
            using (var tb = new SolidBrush(textColor))
            using (var sf = new StringFormat
            {
                Alignment     = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            })
            {
                e.Graphics.DrawString(
                    status,
                    new Font("Segoe UI", 8F, FontStyle.Bold),
                    tb,
                    pillRect,
                    sf);
            }
        }

        #endregion

        // =====================================================================
        // #region IMPRESIÓN
        // =====================================================================
        #region IMPRESIÓN

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (_dtFiltered == null || _dtFiltered.Rows.Count == 0)
            {
                MessageBox.Show(
                    "No hay datos para imprimir.",
                    "Información",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Precalcular total de páginas (aprox. 32 filas por página)
            int rowsPerPage = 32;
            _printTotalPages = (_dtFiltered.Rows.Count + rowsPerPage - 1) / rowsPerPage;
            if (_printTotalPages < 1) _printTotalPages = 1;

            _printRowIndex    = 0;
            _printCurrentPage = 0;

            _pd           = new PrintDocument();
            _pd.PrintPage += PrintPage_Handler;

            _ppd = new PrintPreviewDialog
            {
                Document    = _pd,
                WindowState = FormWindowState.Maximized
            };

            // Reiniciar índices antes de mostrar
            _printRowIndex    = 0;
            _printCurrentPage = 0;

            _ppd.ShowDialog(this);
        }

        private void PrintPage_Handler(object sender, PrintPageEventArgs e)
        {
            _printCurrentPage++;

            int left  = e.MarginBounds.Left;
            int top   = e.MarginBounds.Top;
            int right = e.MarginBounds.Right;
            int bottom = e.MarginBounds.Bottom;
            int width  = e.MarginBounds.Width;

            using (var fBrand  = new Font("Segoe UI", 16F, FontStyle.Bold))
            using (var fTitle  = new Font("Segoe UI", 11F, FontStyle.Bold))
            using (var fDate   = new Font("Segoe UI",  8F))
            using (var fHeader = new Font("Segoe UI",  8F, FontStyle.Bold))
            using (var fText   = new Font("Segoe UI",  8F))
            using (var fPage   = new Font("Segoe UI",  8F))
            using (var pen     = new Pen(Color.Black, 0.5f))
            {
                int y = top;

                var sfCenter = new StringFormat { Alignment = StringAlignment.Center };

                // Nombre de la empresa (grande, centrado)
                e.Graphics.DrawString(
                    "Persal",
                    fBrand, Brushes.Black,
                    new RectangleF(left, y, width, 26), sfCenter);
                y += 28;

                // Título del reporte
                e.Graphics.DrawString(
                    "Reporte de Estatus de Locaciones",
                    fTitle, Brushes.Black,
                    new RectangleF(left, y, width, 20), sfCenter);
                y += 22;

                // Fecha de impresión
                e.Graphics.DrawString(
                    string.Format("Impreso el: {0}",
                        DateTime.Now.ToString("dd/MM/yyyy hh:mm tt")),
                    fDate, Brushes.Black,
                    new RectangleF(left, y, width, 14), sfCenter);
                y += 18;

                // Línea separadora de encabezado
                e.Graphics.DrawLine(pen, left, y, right, y);
                y += 5;

                // Anchos de columnas (proporcional al ancho de página)
                int wName = (int)(width * 0.20f);
                int wArea = (int)(width * 0.13f);
                int wPer  = (int)(width * 0.10f);
                int wLD   = (int)(width * 0.10f);
                int wND   = (int)(width * 0.10f);
                int wDays = (int)(width * 0.08f);
                int wSt   = width - wName - wArea - wPer - wLD - wND - wDays;

                int[] colW = { wName, wArea, wPer, wLD, wND, wDays, wSt };
                int rowH   = 16;

                // Fila de cabeceras
                var headerRect = new Rectangle(left, y, width, rowH + 4);
                e.Graphics.FillRectangle(Brushes.Gainsboro, headerRect);
                e.Graphics.DrawRectangle(pen, headerRect);

                int x = left;
                for (int i = 0; i < _printColHeaders.Length; i++)
                {
                    PrintCell(e.Graphics, _printColHeaders[i], fHeader, pen,
                              x, y, colW[i], rowH + 4, false);
                    x += colW[i];
                }
                y += rowH + 6;

                // Filas de datos
                while (_printRowIndex < _dtFiltered.Rows.Count)
                {
                    // Verificar si hay espacio para otra fila (reservar 20px para número de página)
                    if (y + rowH > bottom - 20)
                    {
                        // Número de página
                        e.Graphics.DrawString(
                            string.Format("Página {0} de {1}",
                                _printCurrentPage, _printTotalPages),
                            fPage, Brushes.Gray,
                            new RectangleF(left, bottom - 14, width, 14), sfCenter);

                        e.HasMorePages = true;
                        return;
                    }

                    var dr = _dtFiltered.Rows[_printRowIndex];
                    x = left;

                    for (int i = 0; i < _printColFields.Length; i++)
                    {
                        object val = dr[_printColFields[i]];
                        string txt;

                        if (val == null || val == DBNull.Value)
                            txt = "";
                        else if (val is DateTime)
                            txt = ((DateTime)val).ToString("MM/dd/yyyy");
                        else
                            txt = Convert.ToString(val);

                        PrintCell(e.Graphics, txt, fText, pen,
                                  x, y, colW[i], rowH, false);
                        x += colW[i];
                    }

                    y += rowH;
                    _printRowIndex++;
                }

                e.HasMorePages = false;

                // Número de página (última página)
                e.Graphics.DrawString(
                    string.Format("Página {0} de {1}",
                        _printCurrentPage, _printTotalPages),
                    fPage, Brushes.Gray,
                    new RectangleF(left, bottom - 14, width, 14), sfCenter);
            }
        }

        private static void PrintCell(
            Graphics g, string text, Font font, Pen pen,
            int x, int y, int w, int h, bool rightAlign)
        {
            var rect = new Rectangle(x, y, w, h);
            g.DrawRectangle(pen, rect);

            var sf = new StringFormat
            {
                LineAlignment = StringAlignment.Center,
                Trimming      = StringTrimming.EllipsisCharacter,
                FormatFlags   = StringFormatFlags.NoWrap,
                Alignment     = rightAlign
                    ? StringAlignment.Far
                    : StringAlignment.Near
            };

            var inner = new RectangleF(rect.X + 3, rect.Y + 1, rect.Width - 6, rect.Height - 2);
            g.DrawString(text ?? "", font, Brushes.Black, inner, sf);
        }

        #endregion

        // =====================================================================
        // #region EXPORTACIÓN A EXCEL (ClosedXML)
        // =====================================================================
        #region EXPORTACIÓN A EXCEL (ClosedXML)

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_dtFiltered == null || _dtFiltered.Rows.Count == 0)
            {
                MessageBox.Show(
                    "No hay datos para exportar.",
                    "Información",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter   = "Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
                sfd.FileName = string.Format("Estatus_Locaciones_{0}.xlsx",
                    DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                sfd.Title    = "Exportar a Excel";

                if (sfd.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Estatus Locaciones");

                        // Fila 1 — Marca
                        ws.Cell(1, 1).Value                 = "Persal";
                        ws.Cell(1, 1).Style.Font.Bold       = true;
                        ws.Cell(1, 1).Style.Font.FontSize   = 16;

                        // Fila 2 — Título
                        ws.Cell(2, 1).Value           = "Reporte de Estatus de Locaciones";
                        ws.Cell(2, 1).Style.Font.Bold = true;

                        // Fila 3 — Fecha de exportación
                        ws.Cell(3, 1).Value = string.Format("Exportado el: {0}",
                            DateTime.Now.ToString("dd/MM/yyyy hh:mm tt"));

                        // Fila 5 — Encabezados de columnas
                        string[] exHeaders = {
                            "Locación", "Área", "Período",
                            "Últ. Conteo", "Próx. Venc.", "Días Vencidos", "Estatus"
                        };
                        string[] exFields = {
                            "Location_Name", "Area_Name", "Period_Name",
                            "Last_Count_Date", "Next_Due_Date", "Days_Overdue", "Status"
                        };

                        for (int i = 0; i < exHeaders.Length; i++)
                        {
                            var hCell = ws.Cell(5, i + 1);
                            hCell.Value                      = exHeaders[i];
                            hCell.Style.Font.Bold            = true;
                            hCell.Style.Fill.BackgroundColor =
                                XLColor.FromArgb(243, 244, 246);
                        }

                        // Datos a partir de la fila 6
                        for (int r = 0; r < _dtFiltered.Rows.Count; r++)
                        {
                            var    dr    = _dtFiltered.Rows[r];
                            int    exRow = r + 6;
                            string st    = Convert.ToString(dr["Status"]);

                            for (int c = 0; c < exFields.Length; c++)
                            {
                                var    cell = ws.Cell(exRow, c + 1);
                                object val  = dr[exFields[c]];

                                if (val == null || val == DBNull.Value)
                                    cell.Value = "";
                                else if (val is DateTime)
                                    cell.Value = ((DateTime)val).ToString("MM/dd/yyyy");
                                else
                                    cell.Value = Convert.ToString(val);

                                // Colorear columna Estatus
                                if (exFields[c] == "Status")
                                {
                                    XLColor bg;
                                    switch (st)
                                    {
                                        case "Al Día":
                                            bg = XLColor.FromHtml("#D1FAE5"); break;
                                        case "Vencido":
                                            bg = XLColor.FromHtml("#FEF3C7"); break;
                                        case "Nunca Contado":
                                            bg = XLColor.FromHtml("#FEE2E2"); break;
                                        default:
                                            bg = XLColor.NoColor;            break;
                                    }
                                    cell.Style.Fill.BackgroundColor = bg;
                                }
                            }
                        }

                        // Ajustar anchos de columna al contenido
                        ws.Columns().AdjustToContents();

                        wb.SaveAs(sfd.FileName);
                    }

                    MessageBox.Show(
                        string.Format("Archivo guardado: {0}", sfd.FileName),
                        "Exportación Exitosa",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // Abrir el archivo exportado
                    try { System.Diagnostics.Process.Start(sfd.FileName); }
                    catch { /* No crítico si falla al abrir */ }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        string.Format("Error al exportar: {0}", ex.Message),
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        #endregion

        // =====================================================================
        // #region HELPERS DE BASE DE DATOS
        // =====================================================================
        #region HELPERS DE BASE DE DATOS

        private string GetCs()
        {
            if (_cs != null && _cs.Trim().Length > 0) return _cs;
            try { return new con().ConnectionString; }
            catch { return ""; }
        }

        private DataTable ExecFillTable(string sql, params object[] args)
        {
            var dt = new DataTable();

            using (var conn = new SqlConnection(GetCs()))
            using (var cmd  = conn.CreateCommand())
            using (var da   = new SqlDataAdapter(cmd))
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue(
                        "@p" + (i + 1),
                        args[i] ?? DBNull.Value);

                da.Fill(dt);
            }

            return dt;
        }

        #endregion

        // =====================================================================
        // #region CHROME HELPERS (borde, arrastrar, esquinas)
        // =====================================================================
        #region CHROME HELPERS

        private void ApplyRoundedRegion(Control c, int radius)
        {
            try
            {
                if (c.Width <= 0 || c.Height <= 0) return;
                using (var path = BuildRoundRectPath(
                    new Rectangle(0, 0, c.Width, c.Height), radius))
                {
                    c.Region = new Region(path);
                }
            }
            catch { }
        }

        private GraphicsPath BuildRoundRectPath(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            int d    = radius * 2;

            path.AddArc(r.X,         r.Y,          d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y,          d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d,   0, 90);
            path.AddArc(r.X,         r.Bottom - d, d, d,  90, 90);
            path.CloseFigure();

            return path;
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            if (!UseCustomChrome) return;
            if (this.WindowState == FormWindowState.Maximized) return;
            if (e.Button != MouseButtons.Left) return;

            ReleaseCapture();
            SendMessage(this.Handle, 0xA1, 0x2, 0);
        }

        #endregion
    }
}
