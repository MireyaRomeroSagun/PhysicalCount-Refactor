using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Persal._003_Physical_Counting_2
{
    // =========================================================================
    // PhysicalCountCyclicForm — 5-step wizard for cyclic physical inventory
    // =========================================================================
    public partial class PhysicalCountCyclicForm : Form
    {
        // =====================================================================
        // #region CONSTANTES
        // =====================================================================
        #region CONSTANTES

        // --- UI Chrome ---
        private const bool UseCustomChrome = true;
        private const int FormCornerRadius = 10;

        // --- Inventory Status IDs ---
        private const int STATUS_STARTED = 1;
        private const int STATUS_FINISHED = 3;

        #endregion

        // =====================================================================
        // #region CAMPOS
        // =====================================================================
        #region CAMPOS

        // --- Wizard state ---
        private int _step = 1;
        private int _physicalCountId = 0;
        private string _connectionString = "";
        private int _responsibleUserId = 0;

        // --- Step 2 ---
        private DataTable _dtLocations;

        // --- Repositorio de datos ---
        private PhysicalCountRepository _repo;

        // --- Paso 3 ---
        private bool _step3ReportReady = false;
        private DataTable _dtStep3Preview;
        private System.Windows.Forms.Timer _step3RefreshTimer = null;
        private const int Step3RefreshMs = 300000;   // 5 minutos
        private bool _step3Loading = false;
        private List<Label> _step3ExtraLabels = null;

        // --- Step 3 Tabs ---
        private TabControl _tabStep3;
        private TabPage _tabStep3Preview;
        private TabPage _tabStep3Attachments;

        // --- Step 3 Attachments ---
        private ListView _lvAttachments;
        private Button _btnAttachAdd;
        private Button _btnAttachOpen;
        private Button _btnAttachRemove;
        private Label _lblAttachInfo;


        // --- Paso 4 (pnlStep4_Capture permanece oculto; solo se inicializa la grilla) ---

        // --- Step 5 ---              
        private DataGridView _dgvStep5Detail;
        private DataGridView _dgvStep5Executive;
        private Button _btnStep5PreviewDetail;
        private Button _btnStep5PreviewExecutive;
        private Button _btnStep5ExportPdfDetail;
        private Button _btnStep5ExportPdfExecutive;
        private DataTable _dtStep5Detail;
        private DataTable _dtStep5Executive;
        private Label _lblStep5Kpis;
        private List<Label> _step5ExtraLabels = null;


        //Reportes y graficas Step 5
        private TabControl _tabStep5;
        private TabPage _tabStep5Detail;
        private TabPage _tabStep5Charts;
        private TabPage _tabStep5Executive;

        private Panel _pnlStep5ChartsHost;
        private Button _btnStep5PreviewCharts;
        private Button _btnStep5ExportPdfCharts;

        private Panel _pnlStep5ChartItemsByLocation;
        private Panel _pnlStep5ChartAccuracyByLocation;
        private Panel _pnlStep5ChartTopDiscrepancies;
        private Panel _pnlStep5ChartStockVsCount;
        private Panel _pnlStep5ChartStockDistribution;
        private Chart _chtStep5ItemsByLocation;
        private Chart _chtStep5AccuracyByLocation;
        private Chart _chtStep5TopDiscrepancies;
        private Chart _chtStep5StockVsCount;

        private Panel _pnlStep5TreemapSurface;
        private List<Step5TreemapItem> _step5TreemapItems;

        private Panel _pnlStep5HoverTopDiff;
        private Label _lblStep5HoverTopDiffTitle;
        private Label _lblStep5HoverTopDiffValue;
        private int _step5TopDiffHoverIndex = -1;

        private Button _btnStep5ExportExcelDetail;
        private Button _btnStep5ExportCsvDetail;
        private Button _btnStep5ExportExcelExecutive;
        private Button _btnStep5ExportCsvExecutive;

        private Button _btnStep5ExportDetail;
        private Button _btnStep5ExportExecutive;

        private ContextMenuStrip _cmsStep5ExportDetail;
        private ContextMenuStrip _cmsStep5ExportExecutive;

        private ToolStripMenuItem _miStep5ExportPdfDetail;
        private ToolStripMenuItem _miStep5ExportExcelDetail;
        private ToolStripMenuItem _miStep5ExportCsvDetail;

        private ToolStripMenuItem _miStep5ExportPdfExecutive;
        private ToolStripMenuItem _miStep5ExportExcelExecutive;
        private ToolStripMenuItem _miStep5ExportCsvExecutive;

        private Bitmap _step5ChartsPrintBitmap;
        private System.Drawing.Printing.PrintDocument _pdStep5Charts;
        private PrintPreviewDialog _ppdStep5Charts;

        // --- Shell layout ---
        private bool _modernUiReady = false;
        private Panel _pnlShellRoot;
        private Panel _pnlHeaderBar;
        private Panel _pnlStepperBar;
        private Panel _pnlContent;
        private Panel _pnlFooter;
        private Panel _pnlStepHost;
        private ModernCardPanel _cardContent;
        private Label _lblHeaderTitle;
        private Label _lblHeaderSubtitle;
        private StepIndicatorBar _stepper;

        // --- Chrome buttons ---
        private Button _btnClose;
        private Button _btnMin;
        private Button _btnMax;

        // --- Misc ---
        private int _defaultCountResponsibleId = 0;

        #endregion

        // =====================================================================
        // #region CONSTRUCTOR & LOAD
        // =====================================================================
        #region CONSTRUCTOR Y CARGA

        /// <summary>
        /// Main constructor. Requires a valid SQL connection string.
        /// </summary>
        public PhysicalCountCyclicForm(string connectionString)
        {
            InitializeComponent();
            _connectionString = connectionString;

            // Wire form-level events
            this.Load += PhysicalCountCyclicForm_Load;
            this.FormClosing += (s, e) => StopStep3RefreshTimer();

            // Wire wizard navigation buttons
            btnNext.Click -= btnNext_Click;
            btnNext.Click += btnNext_Click;
            btnBack.Click -= btnBack_Click;
            btnBack.Click += btnBack_Click;

            // Wire Step 3 export button
            btnStep3ExportExcel.Click += btnStep3ExportExcel_Click;
        }

        /// <summary>
        /// Parameterless constructor for designer support.
        /// </summary>
        public PhysicalCountCyclicForm() : this("") { }

        /// <summary>
        /// Form Load — resolves connection string, fills combos, configures grids,
        /// applies modern UI shell and shows Step 1.
        /// </summary>
        private void PhysicalCountCyclicForm_Load(object sender, EventArgs e)
        {
            // Resolver cadena de conexión
            if (string.IsNullOrWhiteSpace(_connectionString))
                _connectionString = GetConnectionString();

            if (string.IsNullOrWhiteSpace(_connectionString))
            {
                MessageBox.Show(
                    "No connection string provided. Open this form passing a valid connection string.",
                    "Configuration Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.Close();
                return;
            }

            // Inicializar el repositorio de datos
            _repo = new PhysicalCountRepository(_connectionString);

            // Llenar combo de tipo de conteo
            try
            {
                var sf = new Persal.System_Functions();
                cboCountType.DropDownStyle = ComboBoxStyle.DropDownList;

                bool ok = sf.Fill_Combo_with_Physical_Count_Types(cboCountType);
                if (!ok) sf.Show_Error_Message(sf.G.Error_Message);

                if (cboCountType.Items.Count > 0 && cboCountType.SelectedIndex < 0)
                    cboCountType.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading count types: " + ex.Message);
            }

            this.MaximumSize = Size.Empty;
            // Llenar combo de tipo de lista
            LoadPhysicalCountListTypes();

            // Configurar grillas
            ConfigureStep3PreviewGrid();
            ConfigureStep4Grid();

            // Aplicar shell moderno y mostrar el primer paso
            ApplyModernUiShell();
            ConfigureFormResizeBehavior();
            ShowStep(_startStep);
        }

        #endregion

        // =====================================================================
        // #region NAVEGACIÓN DEL WIZARD
        // =====================================================================
        #region NAVEGACIÓN DEL WIZARD

        /// <summary>
        /// Shows the requested wizard step, hiding all others.
        /// Handles timer, UI state, and step-specific initialization.
        /// </summary>
        private void ShowStep(int step)
        {
            if (step < 1) step = 1;
            if (step > 4) step = 4;

            if (_step == 3 && step != 3)
            {
                StopStep3RefreshTimer();
                _dtStep3Preview = null;
            }

            _step = step;

            pnlStep1_Create.Visible = (step == 1);
            pnlStep2_Locations.Visible = (step == 2);
            pnlStep3_Confirm.Visible = (step == 3);
            pnlStep4_Capture.Visible = false;        // nunca se usa
            pnlStep5_Review.Visible = (step == 4);  // step 4 real = Review & Finalize

            if (prgSteps != null)
            {
                prgSteps.Minimum = 1;
                prgSteps.Maximum = 4;
                prgSteps.Value = step;
            }

            if (btnBack != null)
            {
                bool backEnabled = (step == 2 || step == 3);
                btnBack.Enabled = backEnabled;
                UpdateButtonVisualStyle(btnBack, backEnabled);
            }

            if (lblStepTitle != null)
            {
                switch (step)
                {
                    case 1: lblStepTitle.Text = "Step 1 of 4 — Create cyclic inventory"; break;
                    case 2: lblStepTitle.Text = "Step 2 of 4 — Select locations"; break;
                    case 3: lblStepTitle.Text = "Step 3 of 4 — Waiting for web count"; break;
                    case 4: lblStepTitle.Text = "Step 4 of 4 — Review & Finalize"; break;
                }
                lblStepTitle.BringToFront();
            }

            if (btnNext != null)
            {
                if (step == 3) btnNext.Text = "Verify Count";
                else if (step == 4) btnNext.Text = "Close";
                else btnNext.Text = "Next";
            }

            UpdateStepper(step);
            UpdateHeaderSubtitle(step);

            switch (step)
            {
                case 1:
                    LayoutStep1();
                    if (txtDescription != null && txtDescription.CanFocus)
                        txtDescription.Focus();
                    break;

                case 2:
                    EnsureStep2HierarchyUI();
                    LoadAlmacenes();
                    break;

                case 3:
                    EnsureStep3TabsUI();
                    RefreshStep3Summary();
                    if (_lvAttachments != null) LoadAttachments();
                    if (!_step3ReportReady)
                    {
                        InitStep3Preview();
                        SetStep3GridReadOnly();
                        StartStep3RefreshTimer();
                    }
                    break;

                case 4:
                    EnsureStep5UI();
                    LoadStep5Data();
                    LoadStep5HeaderInfo();
                    break;
            }
        }

        /// <summary>
        /// Next button — validates current step then advances.
        /// </summary>


        private void btnNext_Click(object sender, EventArgs e)
        {
            if (!ValidateStep(_step)) return;

            switch (_step)
            {
                case 1:
                    if (!SaveInventoryHeader()) return;
                    ShowStep(2);
                    break;

                case 2:
                    if (!SaveSelectedLocations()) return;
                    ResetDownstreamSteps();
                    ShowStep(3);
                    break;

                case 3:
                    if (!ValidateCountIsFinished()) return;
                    ShowStep(4);
                    break;

                case 4:
                    FinalizeWizard();

                    break;

            }
        }

        /// <summary>
        /// Back button — returns to previous step with confirmation where needed.
        /// </summary>
        private void btnBack_Click(object sender, EventArgs e)
        {
            if (_step <= 1 || !btnBack.Enabled) return;
            if (_step == 3)
            {
                // ✅ Block going back if any location has already been counted
                if (HasCountedLocations(_physicalCountId))
                {
                    MessageBox.Show(
                        "You cannot go back because one or more locations have already been counted.",
                        "Action Not Allowed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                var result = MessageBox.Show(
                    "Do you want to go back to select locations?\n\nData generated in this step will be lost.",
                    "Confirm",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes) return;

                if (_physicalCountId > 0)
                {
                    if (!DeleteStep3DataByPhysicalCountId(_physicalCountId))
                        return;
                }

                ResetDownstreamSteps();
                ShowStep(2);
                return;
            }

            // Steps 4+ : back not allowed
        }

        private bool HasCountedLocations(int physicalCountId)
        {
            return _repo.HasCountedLocations(physicalCountId);
        }

        private bool DeleteStep3DataByPhysicalCountId(int physicalCountId)
        {
            try
            {
                if (physicalCountId <= 0) return false;
                _repo.DeleteStep3Data(physicalCountId);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error al eliminar datos del Paso 3: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }



        /// <summary>
        /// Validates required fields for the given step before advancing.
        /// </summary>
        private bool ValidateStep(int step)
        {
            switch (step)
            {
                case 1:
                    if (string.IsNullOrWhiteSpace(txtDescription?.Text))
                    {
                        MessageBox.Show("Please enter a description.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtDescription?.Focus();
                        return false;
                    }
                    if (cboCountType == null || cboCountType.SelectedIndex < 0)
                    {
                        MessageBox.Show("Please select a count type.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cboCountType?.Focus();
                        return false;
                    }
                    if (dtpStartDate != null && dtpEndDate != null &&
                        dtpStartDate.Value.Date > dtpEndDate.Value.Date)
                    {
                        MessageBox.Show("Start date cannot be later than end date.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        dtpStartDate.Focus();
                        return false;
                    }
                    if (_responsibleUserId <= 0)
                    {
                        MessageBox.Show("Please select a valid responsible person.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtResponsibleEmpNo?.Focus();
                        return false;
                    }
                    break;

                case 2:
                    if (_physicalCountId <= 0)
                    {
                        MessageBox.Show("Please save the inventory header first (Step 1).", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                    if (_effectiveSelectedLocationIds.Count == 0)
                    {
                        MessageBox.Show("Please select at least one location.", "Validation",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                    break;

                case 3:
                case 4:
                case 5:
                    if (_physicalCountId <= 0)
                    {
                        MessageBox.Show("No valid inventory ID found.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                    break;
            }

            return true;
        }

        /// <summary>
        /// Resets Steps 3–5 data when locations change.
        /// </summary>
        private void ResetDownstreamSteps()
        {
            _step3ReportReady = false;
            _dtStep3Preview = null;
            _dtStep5Detail = null;
            _dtStep5Executive = null;

            if (dgvStep3Preview != null) dgvStep3Preview.DataSource = null;
            if (dgvStep4 != null) dgvStep4.DataSource = null;
            if (_dgvStep5Detail != null) _dgvStep5Detail.DataSource = null;
            if (_dgvStep5Executive != null) _dgvStep5Executive.DataSource = null;
        }

        #endregion

        // =====================================================================z
        // #region PASO 1 — CREAR INVENTARIO
        // =====================================================================
        #region PASO 1 — CREAR INVENTARIO
        private const string ResponsiblePlaceholder = "Search Emp";

        /// <summary>
        /// Builds the layout for Step 1 (general info, dates, responsible).
        /// </summary>
        private void LayoutStep1()
        {
            if (pnlStep1_Create == null) return;

            try
            {
                pnlStep1_Create.SuspendLayout();

                pnlStep1_Create.Controls.Clear();
                pnlStep1_Create.Dock = DockStyle.Fill;
                pnlStep1_Create.Padding = new Padding(0);
                pnlStep1_Create.BackColor = Color.FromArgb(243, 244, 246);

                var scroll = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = pnlStep1_Create.BackColor,
                    Padding = new Padding(24)
                };

                var border = new Panel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    BackColor = Color.FromArgb(209, 213, 219),
                    Padding = new Padding(1),
                    Margin = new Padding(0)
                };

                var card = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    BackColor = Color.White,
                    Padding = new Padding(24, 20, 24, 20),
                    Margin = new Padding(0)
                };

                var content = new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    ColumnCount = 1,
                    RowCount = 5,
                    BackColor = Color.White,
                    Padding = new Padding(0),
                    Margin = new Padding(0)
                };
                content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                for (int i = 0; i < 5; i++)
                    content.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var s1 = BuildSection_GeneralInfo();
                var s2 = BuildSection_Dates();
                var s3 = BuildSection_Responsible();

                if (s1 != null) content.Controls.Add(s1, 0, 0);
                content.Controls.Add(CreateStep1Separator(), 0, 1);

                if (s2 != null) content.Controls.Add(s2, 0, 2);
                content.Controls.Add(CreateStep1Separator(), 0, 3);

                if (s3 != null) content.Controls.Add(s3, 0, 4);

                card.Controls.Add(content);
                border.Controls.Add(card);
                scroll.Controls.Add(border);
                pnlStep1_Create.Controls.Add(scroll);
                pnlStep1_Create.Controls.Clear();
                pnlStep1_Create.Controls.Add(scroll);

                WireStep1ResponsibleEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error building Step 1 layout: " + ex.Message);
            }
            finally
            {
                pnlStep1_Create.ResumeLayout(true);
            }
        }

        private Control CreateStep1Separator()
        {
            return new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = Color.FromArgb(229, 231, 235),
                Margin = new Padding(0, 8, 0, 18)
            };
        }

        /// <summary>Builds the "General Information" section.</summary>
        private Panel BuildSection_GeneralInfo()
        {
            var section = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var container = CreateTableContainer(3, 1);
            AddColumnStyle(container, SizeType.Percent, 100);
            AddRowStyles(container, 3);

            container.Controls.Add(CreateSectionTitle("GENERAL INFORMATION"), 0, 0);

            if (label1 != null && txtDescription != null)
            {
                txtDescription.Parent = null;
                txtDescription.MinimumSize = new Size(320, 36);
                container.Controls.Add(CreateFieldRow(label1.Text, txtDescription), 0, 1);
            }

            var row2 = CreateTwoColumnRow();

            if (label2 != null && cboCountType != null)
            {
                cboCountType.Parent = null;
                container.SuspendLayout();
                row2.Controls.Add(CreateFieldRow(label2.Text, cboCountType), 0, 0);
                container.ResumeLayout();
            }

            if (label16 != null && Combo_Physical_Count_List_Type != null)
            {
                Combo_Physical_Count_List_Type.Parent = null;
                row2.Controls.Add(CreateFieldRow(label16.Text, Combo_Physical_Count_List_Type), 1, 0);
            }

            container.Controls.Add(row2, 0, 2);

            section.Controls.Add(container);
            return section;
        }

        /// <summary>Builds the "Rango de Fechas" section.</summary>
        private Panel BuildSection_Dates()
        {
            if (dtpStartDate == null || dtpEndDate == null) return null;

            var section = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var container = CreateTableContainer(2, 1);
            AddColumnStyle(container, SizeType.Percent, 100);
            AddRowStyles(container, 2);

            container.Controls.Add(CreateSectionTitle("DATE"), 0, 0);

            var dates = CreateTwoColumnRow();
            dates.Controls.Add(CreateFieldRow(lbl_Fecha_Inicio?.Text ?? "Start Date", dtpStartDate), 0, 0);
            dates.Controls.Add(CreateFieldRow(lbl_Fecha_Final?.Text ?? "End Date", dtpEndDate), 1, 0);

            container.Controls.Add(dates, 0, 1);

            section.Controls.Add(container);
            return section;
        }

        /// <summary>Builds the "Responsible" section (search + info card).</summary>
        private Panel BuildSection_Responsible()
        {
            var section = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var container = CreateTableContainer(3, 1);
            AddColumnStyle(container, SizeType.Percent, 100);
            AddRowStyles(container, 3);

            container.Controls.Add(CreateSectionTitle("INFORMATION OF THE RESPONSIBLE"), 0, 0);

            var infoCard = BuildResponsibleInfoCard();
            if (infoCard != null)
                container.Controls.Add(infoCard, 0, 1);

            var searchRow = BuildResponsibleSearchRow();
            if (searchRow != null)
                container.Controls.Add(searchRow, 0, 2);

            section.Controls.Add(container);
            return section;
        }

        /// <summary>Builds the blue info card.</summary>
        private Panel BuildResponsibleInfoCard()
        {
            if (lblResponsibleName == null || lblResponsibleJob == null || lblempleadonum == null)
                return null;

            var border = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(209, 213, 219),
                Padding = new Padding(1),
                Margin = new Padding(0, 0, 0, 12)
            };

            var card = new Panel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.FromArgb(243, 246, 252),
                Padding = new Padding(16, 12, 16, 10)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = card.BackColor,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++)
                grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            AddInfoRow(grid, 0, lbl_num?.Text ?? "No. Empleado", lblempleadonum, false);
            AddInfoRow(grid, 1, lblname?.Text ?? "Nombre", lblResponsibleName, false);
            AddInfoRow(grid, 2, lbl_Puesto?.Text ?? "Puesto", lblResponsibleJob, false);

            card.Controls.Add(grid);
            border.Controls.Add(card);

            return border;
        }

        /// <summary>Builds the search row under the card.</summary>
        private Panel BuildResponsibleSearchRow()
        {
            if (txtResponsibleEmpNo == null || button_Select_Physical_Count_Responsible_ID == null)
                return null;

            var wrapper = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0, 10, 0, 0),
                Margin = new Padding(0)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lbl = new Label
            {
                Text = "Find a Responsible Person",
                Font = new Font("Segoe UI", 10F),
                ForeColor = Color.FromArgb(55, 65, 81),
                AutoSize = true,
                Padding = new Padding(0, 8, 16, 0),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0)
            };

            txtResponsibleEmpNo.Parent = null;
            txtResponsibleEmpNo.Height = 38;
            txtResponsibleEmpNo.Dock = DockStyle.Fill;
            txtResponsibleEmpNo.Margin = new Padding(0, 0, 12, 0);
            StyleControl(txtResponsibleEmpNo);

            if (string.IsNullOrWhiteSpace(txtResponsibleEmpNo.Text))
            {
                txtResponsibleEmpNo.Text = ResponsiblePlaceholder;
                txtResponsibleEmpNo.ForeColor = Color.Gray;
            }

            button_Select_Physical_Count_Responsible_ID.Parent = null;
            button_Select_Physical_Count_Responsible_ID.Text = "Search F3";
            button_Select_Physical_Count_Responsible_ID.Height = 38;
            button_Select_Physical_Count_Responsible_ID.Dock = DockStyle.Fill;
            button_Select_Physical_Count_Responsible_ID.Margin = new Padding(0);
            StyleButtonSecondary(button_Select_Physical_Count_Responsible_ID);

            grid.Controls.Add(lbl, 0, 0);
            grid.Controls.Add(txtResponsibleEmpNo, 1, 0);
            grid.Controls.Add(button_Select_Physical_Count_Responsible_ID, 2, 0);

            wrapper.Controls.Add(grid);
            return wrapper;
        }
        private void txtResponsibleEmpNo_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtResponsibleEmpNo.Text))
            {
                txtResponsibleEmpNo.Text = ResponsiblePlaceholder;
                txtResponsibleEmpNo.ForeColor = Color.Gray;
            }
        }

        private void lblResponsibleName_Click(object sender, EventArgs e)
        {
            // En este diseño el buscador siempre debe permanecer visible.
            if (txtResponsibleEmpNo == null) return;

            txtResponsibleEmpNo.Visible = true;
            txtResponsibleEmpNo.Enabled = true;
        }

        /// <summary>Agrega una fila etiqueta + valor a la tarjeta de información del responsable.</summary>
        private void AddInfoRow(TableLayoutPanel grid, int row, string caption, Label valueLabel, bool bold)
        {
            var lbl = new Label
            {
                Text = caption,
                Font = new Font("Segoe UI", 10F),
                ForeColor = Color.FromArgb(107, 114, 128),
                AutoSize = true,
                Margin = new Padding(0, 0, 12, 8),
                Padding = new Padding(0, 2, 0, 0)
            };

            valueLabel.Parent = null;
            valueLabel.Text = string.IsNullOrWhiteSpace(valueLabel.Text) ? "—" : valueLabel.Text;
            valueLabel.Font = new Font("Segoe UI", 10F, bold ? FontStyle.Bold : FontStyle.Regular);
            valueLabel.ForeColor = Color.FromArgb(17, 24, 39);
            valueLabel.AutoSize = true;
            valueLabel.Margin = new Padding(0, 0, 0, 8);

            grid.Controls.Add(lbl, 0, row);
            grid.Controls.Add(valueLabel, 1, row);
        }

        private void txtResponsibleEmpNo_Enter(object sender, EventArgs e)
        {
            if (txtResponsibleEmpNo.Text == ResponsiblePlaceholder)
            {
                txtResponsibleEmpNo.Text = "";
                txtResponsibleEmpNo.ForeColor = Color.Black;
            }
        }


        /// <summary>
        /// Saves the inventory header (Step 1) to the database.
        /// Returns true on success, false on validation or DB error.
        /// </summary>
        private bool SaveInventoryHeader()
        {
            try
            {
                var sf = new Persal.System_Functions();

                // Validate list type
                int listTypeId = 0;
                if (Combo_Physical_Count_List_Type.SelectedValue != null)
                    listTypeId = Convert.ToInt32(Combo_Physical_Count_List_Type.SelectedValue);

                if (listTypeId <= 0)
                {
                    MessageBox.Show("Please select a count list type.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Combo_Physical_Count_List_Type.Focus();
                    return false;
                }

                // Count type
                int countTypeId = GetComboSelectedValueAsInt(cboCountType);
                if (countTypeId <= 0)
                {
                    MessageBox.Show("Please select a count type.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cboCountType.Focus();
                    return false;
                }

                // Current user
                int userId = GetCurrentUserId();
                if (userId <= 0)
                {
                    MessageBox.Show("Could not retrieve current user ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Responsible
                if (_responsibleUserId <= 0)
                {
                    MessageBox.Show("Please select a valid responsible person.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtResponsibleEmpNo?.Focus();
                    return false;
                }

                // Dates
                DateTime startDate = dtpStartDate.Value.Date;
                DateTime finishDate = dtpEndDate.Value.Date.AddDays(1).AddTicks(-1);

                if (finishDate < startDate)
                {
                    MessageBox.Show("End date cannot be before start date.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dtpEndDate.Focus();
                    return false;
                }

                // Description
                string description = (txtDescription.Text ?? "").Trim();
                if (description.Length == 0)
                {
                    MessageBox.Show("Please enter a description.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtDescription.Focus();
                    return false;
                }

                // Default warehouse
                int warehouseId = sf.Return_Default_Warehouse_Location_ID_For_Count();
                if (warehouseId <= 0)
                {
                    MessageBox.Show("Could not determine default warehouse for this count.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Save
                bool saved = sf.Save_Physical_Count(
                    _physicalCountId,
                    countTypeId,
                    listTypeId,
                    _responsibleUserId,
                    userId,        // verificator
                    1,             // number of checks
                    description,
                    startDate,
                    finishDate,
                    100.0M,        // goal
                    warehouseId
                );

                if (!saved)
                {
                    sf.Show_Error_Message(sf.G.Error_Message);
                    return false;
                }

                // Recuperar ID generado si es nuevo
                if (_physicalCountId == 0)
                {
                    _physicalCountId = _repo.GetLastCountIdByUser(userId);
                }

                if (_physicalCountId <= 0)
                {
                    MessageBox.Show("Inventory saved but could not retrieve the generated ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                MessageBox.Show(
                    "Inventory created successfully. ID: " + _physicalCountId,
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Critical error saving header: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Carga el combo de tipos de lista de conteo desde el repositorio.
        /// </summary>
        private void LoadPhysicalCountListTypes()
        {
            try
            {
                var dt = _repo.LoadListTypes();
                Combo_Physical_Count_List_Type.DataSource    = dt;
                Combo_Physical_Count_List_Type.DisplayMember = "Physical_Count_List_Type";
                Combo_Physical_Count_List_Type.ValueMember   = "Physical_Count_List_Type_ID";
                Combo_Physical_Count_List_Type.SelectedIndex = 0;
                Combo_Physical_Count_List_Type.Enabled       = false; // fijo en 1 solo tipo
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar tipos de lista: " + ex.Message);
            }
        }
        // --- Responsible search ---

        private void txtResponsibleEmpNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                OpenResponsibleLookup();
                return;
            }

            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SearchResponsibleByText(txtResponsibleEmpNo.Text.Trim());
            }
        }

        private void button_Select_Physical_Count_Responsible_ID_Click(object sender, EventArgs e)
        {
            OpenResponsibleLookup();
        }

        private void WireStep1ResponsibleEvents()
        {
            if (txtResponsibleEmpNo != null)
            {
                txtResponsibleEmpNo.KeyDown -= txtResponsibleEmpNo_KeyDown;
                txtResponsibleEmpNo.KeyDown += txtResponsibleEmpNo_KeyDown;

                txtResponsibleEmpNo.Enter -= txtResponsibleEmpNo_Enter;
                txtResponsibleEmpNo.Enter += txtResponsibleEmpNo_Enter;

                txtResponsibleEmpNo.Leave -= txtResponsibleEmpNo_Leave;
                txtResponsibleEmpNo.Leave += txtResponsibleEmpNo_Leave;
            }

            if (button_Select_Physical_Count_Responsible_ID != null)
            {
                button_Select_Physical_Count_Responsible_ID.Click -= button_Select_Physical_Count_Responsible_ID_Click;
                button_Select_Physical_Count_Responsible_ID.Click += button_Select_Physical_Count_Responsible_ID_Click;
            }
        }

        /// <summary>Opens the F3 user lookup dialog.</summary>
        private void OpenResponsibleLookup()
        {
            this.Cursor = Cursors.WaitCursor;

            try
            {
                Persal.Select_User M = new Persal.Select_User();
                M.ShowDialog();

                if (M.User_ID > 0)
                {
                    LoadResponsibleByUserId(M.User_ID);

                    txtResponsibleEmpNo.Clear();
                    txtResponsibleEmpNo.Focus();
                }
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        /// <summary>Carga los datos del responsable a partir de su User_ID.</summary>
        private bool LoadResponsibleByUserId(int userId)
        {
            int  respId;
            string empNo, nombre, puesto;

            bool found = _repo.LoadResponsibleById(userId, out respId, out empNo, out nombre, out puesto);

            if (!found)
            {
                MessageBox.Show(
                    "El usuario seleccionado no fue encontrado.",
                    "No encontrado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return false;
            }

            _responsibleUserId      = respId;
            txtResponsibleEmpNo.Text = empNo;
            lblempleadonum.Text      = empNo;
            lblResponsibleName.Text  = nombre;
            lblResponsibleJob.Text   = puesto;
            return true;
        }

        /// <summary>Busca un responsable activo por número de empleado, login o nombre.</summary>
        private bool SearchResponsibleByText(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;

            int    respId;
            string empNo, nombre, puesto;

            bool found = _repo.SearchResponsibleByText(
                value.Trim(), out respId, out empNo, out nombre, out puesto);

            if (!found)
            {
                MessageBox.Show(
                    "Usuario activo no encontrado. Prueba con número de empleado, login, nombre o usa F3.",
                    "No encontrado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                txtResponsibleEmpNo.Clear();
                txtResponsibleEmpNo.Focus();
                return false;
            }

            _responsibleUserId      = respId;
            lblempleadonum.Text     = empNo;
            lblResponsibleName.Text = nombre;
            lblResponsibleJob.Text  = puesto;
            txtResponsibleEmpNo.Clear();
            txtResponsibleEmpNo.Focus();
            return true;
        }

        // --- Designer events for Step 1 controls ---



        private void dtpStartDate_ValueChanged(object sender, EventArgs e)
        {
            if (dtpEndDate != null) dtpEndDate.MinDate = dtpStartDate.Value;
        }

        private void dtpenddate_ValueChanged(object sender, EventArgs e)
        {
            if (dtpStartDate != null) dtpStartDate.MaxDate = dtpEndDate.Value;
        }

        private void txtDescription_keydown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
        }


        #endregion

        // =====================================================================
        // #region PASO 2 — SELECCIONAR UBICACIONES
        // =====================================================================
        #region PASO 2 — SELECCIONAR UBICACIONES

        /// <summary>
        /// Configures the columns for the locations grid (Step 2).
        /// </summary>

        // --- Step 2 (Hierarchy Locations) ---
        private bool _step2UiReady = false;

        private TextBox _txtSearchAlm;
        private TextBox _txtSearchArea;
        private TextBox _txtSearchSub;

        private DataGridView _dgvAlm;
        private DataGridView _dgvArea;
        private DataGridView _dgvSub;

        private Label _lblAlmCount;
        private Label _lblAreaCount;
        private Label _lblSubCount;

        private int? _focusedAlmacenId = null;  // “seleccionado” para navegar (no necesariamente checked)
        private int? _focusedAreaId = null;

        private readonly HashSet<int> _checkedAlmacenIds = new HashSet<int>();
        private readonly HashSet<int> _checkedAreaIds = new HashSet<int>();
        private readonly HashSet<int> _checkedLocationIds = new HashSet<int>();
        private CheckBox _chkSelectAllSubLocations;
        private bool _updatingSelectAllSubLocations = false;


        // Sub-áreas quedan siempre auto-incluidas según áreas checked
        // (no se permite des-seleccionar)
        private readonly HashSet<int> _effectiveSelectedLocationIds = new HashSet<int>();


        //Locacion 3 
        private const int Step2FixedAlmacenId = 3; // WAREHOUSE SUPPLIES




        private Step2ColumnUi BuildStep2ColumnModern(string title, string iconText)
        {
            // Card root
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(8),
            };

            // Outer border (simula card)
            var border = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(229, 231, 235),
                Padding = new Padding(1),
            };
            card.Controls.Add(border);

            var inner = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(12),
            };
            border.Controls.Add(inner);

            // Header row
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 32,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.White,
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 26));  // icon
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // title
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 38)); // badge

            var lblIcon = new Label
            {
                Text = iconText, // puedes usar "🏭" "▦" "📍" o un glyph
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(107, 114, 128),
            };

            var lblTitle = new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(55, 65, 81),
            };

            var lblCount = new Label
            {
                Text = "0",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                ForeColor = Color.FromArgb(55, 65, 81),
                BackColor = Color.FromArgb(243, 244, 246),
                Margin = new Padding(0, 6, 0, 6),
            };

            // Para que parezca "pill"
            lblCount.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                var r = new Rectangle(0, 0, lblCount.Width - 1, lblCount.Height - 1);
                int radius = 12;

                using (var b = new SolidBrush(lblCount.BackColor))
                using (var p = new Pen(Color.FromArgb(229, 231, 235)))
                using (var path = BuildRoundRectPath(r, radius))
                {
                    g.FillPath(b, path);
                    g.DrawPath(p, path);

                    TextRenderer.DrawText(
                        g,
                        lblCount.Text,
                        lblCount.Font,
                        r,
                        lblCount.ForeColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                    );
                }
            };

            header.Controls.Add(lblIcon, 0, 0);
            header.Controls.Add(lblTitle, 1, 0);
            header.Controls.Add(lblCount, 2, 0);

            // Search box (panel con icon + textbox)
            var searchWrap = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = Color.FromArgb(243, 244, 246),
                Padding = new Padding(10, 8, 10, 8),
                Margin = new Padding(0, 10, 0, 10),
            };
            searchWrap.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(229, 231, 235)))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, searchWrap.Width - 1, searchWrap.Height - 1);
                }
            };

            var lblSearchIcon = new Label
            {
                Text = "🔍",
                AutoSize = false,
                Width = 24,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(107, 114, 128),
            };

            var txtSearch = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F),
                BackColor = searchWrap.BackColor,
                ForeColor = Color.FromArgb(17, 24, 39),
            };

            string placeholder = $"Buscar {title.ToLowerInvariant()}...";
            txtSearch.Text = placeholder;
            txtSearch.ForeColor = Color.FromArgb(156, 163, 175);

            txtSearch.GotFocus += (s, e) =>
            {
                if (txtSearch.Text == placeholder)
                {
                    txtSearch.Text = "";
                    txtSearch.ForeColor = Color.FromArgb(17, 24, 39);
                }
            };
            txtSearch.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtSearch.Text))
                {
                    txtSearch.Text = placeholder;
                    txtSearch.ForeColor = Color.FromArgb(156, 163, 175);
                }
            };

            searchWrap.Controls.Add(txtSearch);
            searchWrap.Controls.Add(lblSearchIcon);

            // Grid list
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                ColumnHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                ScrollBars = ScrollBars.Vertical,
            };

            StyleStep2ListGrid(dgv);

            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 34,
                BackColor = Color.White,
                Padding = new Padding(4, 2, 4, 0),
                Margin = new Padding(0)
            };

            footer.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(229, 231, 235)))
                {
                    e.Graphics.DrawLine(pen, 0, 0, footer.Width, 0);
                }
            };

            inner.Controls.Add(dgv);
            inner.Controls.Add(footer);
            inner.Controls.Add(searchWrap);
            inner.Controls.Add(header);

            return new Step2ColumnUi
            {
                Root = card,
                CountLabel = lblCount,
                SearchBox = txtSearch,
                Grid = dgv,
                FooterHost = footer
            };
        }

        private void EnsureStep2HierarchyUI()
        {
            if (_step2UiReady) return;
            if (pnlStep2_Locations == null) return;

            pnlStep2_Locations.Controls.Clear();
            pnlStep2_Locations.Dock = DockStyle.Fill;
            pnlStep2_Locations.BackColor = Color.White;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.White,
                Padding = new Padding(12),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // Iconos aproximados (puedes cambiarlos)
            var colAlm = BuildStep2ColumnModern("ALMACÉN", "🏭");
            var colArea = BuildStep2ColumnModern("ÁREA", "▦");
            var colSub = BuildStep2ColumnModern("LOCATION", "📍");


            _chkSelectAllSubLocations = new CheckBox
            {
                Text = "Selecciona todas las locaciones",
                AutoSize = false,
                Dock = DockStyle.Left,
                Width = 220,
                Height = 24,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.75F),
                ForeColor = Color.FromArgb(75, 85, 99),
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(4, 0, 0, 0)
            };

            _chkSelectAllSubLocations.FlatAppearance.BorderSize = 0;
            _chkSelectAllSubLocations.CheckedChanged -= ChkSelectAllSubLocations_CheckedChanged;
            _chkSelectAllSubLocations.CheckedChanged += ChkSelectAllSubLocations_CheckedChanged;

            if (colSub.FooterHost != null)
                colSub.FooterHost.Controls.Add(_chkSelectAllSubLocations);

            _lblAlmCount = colAlm.CountLabel;
            _txtSearchAlm = colAlm.SearchBox;
            _dgvAlm = colAlm.Grid;

            _lblAreaCount = colArea.CountLabel;
            _txtSearchArea = colArea.SearchBox;
            _dgvArea = colArea.Grid;

            _lblSubCount = colSub.CountLabel;
            _txtSearchSub = colSub.SearchBox;
            _dgvSub = colSub.Grid;

            root.Controls.Add(colAlm.Root, 0, 0);
            root.Controls.Add(colArea.Root, 1, 0);
            root.Controls.Add(colSub.Root, 2, 0);

            pnlStep2_Locations.Controls.Add(root);

            // Search
            _txtSearchAlm.TextChanged += (s, e) => ApplyStep2Filter(_dgvAlm, _txtSearchAlm.Text);
            _txtSearchArea.TextChanged += (s, e) => ApplyStep2Filter(_dgvArea, _txtSearchArea.Text);
            _txtSearchSub.TextChanged += (s, e) => ApplyStep2Filter(_dgvSub, _txtSearchSub.Text);

            // Check handling (mantienes tu lógica)
            _dgvAlm.CellContentClick += (s, e) => Step2_ToggleCheck_Almacen(e);
            _dgvArea.CellContentClick += (s, e) => Step2_ToggleCheck_Area(e);
            _dgvSub.CellContentClick += (s, e) => Step2_ToggleCheck_Location(e);

            _checkedAlmacenIds.Clear();
            _checkedAreaIds.Clear();
            _checkedLocationIds.Clear();
            _effectiveSelectedLocationIds.Clear();
            _focusedAlmacenId = null;
            _focusedAreaId = null;

            _step2UiReady = true;
            ConfigureStep2Grid(_dgvSub, allowCheck: true);
            _dgvSub.AutoGenerateColumns = false;

            BindEmptyMessage(_dgvArea, "Selecciona un almacén");
            BindEmptyMessage(_dgvSub, "Selecciona un área");
        }
        private void StyleStep2ListGrid(DataGridView dgv)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.GridColor = Color.White;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.None;
            dgv.RowTemplate.Height = 44;
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(17, 24, 39);
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 231, 255); // azul suave
            dgv.DefaultCellStyle.SelectionForeColor = Color.FromArgb(17, 24, 39);
            dgv.DefaultCellStyle.Padding = new Padding(6, 0, 6, 0);

            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.White;

            dgv.Columns.Clear();

            // Checkbox
            dgv.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Selected",
                DataPropertyName = "Selected",
                Width = 34,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });

            // Hidden ID
            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Location_ID",
                DataPropertyName = "Location_ID",
                Visible = false
            });

            // Name (fill)
            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Location_Name",
                DataPropertyName = "Location_Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                ReadOnly = true
            });

            // Chevron ">"
            //dgv.Columns.Add(new DataGridViewTextBoxColumn
            //{
            //    Name = "Chevron",
            //    HeaderText = "",
            //    ReadOnly = true,
            //    Width = 26,
            //    AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            //});

            //dgv.CellFormatting -= Step2Grid_CellFormattingChevron;
            //dgv.CellFormatting += Step2Grid_CellFormattingChevron;
        }

        private void Step2Grid_CellFormattingChevron(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null) return;

            if (dgv.Columns[e.ColumnIndex].Name == "Chevron")
            {
                e.Value = "›";
                e.FormattingApplied = true;
                dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.FromArgb(156, 163, 175);
                dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            }
        }


        private sealed class Step2ColumnUi
        {
            public Control Root { get; set; }
            public Label CountLabel { get; set; }
            public TextBox SearchBox { get; set; }
            public DataGridView Grid { get; set; }
            public Panel FooterHost { get; set; }
        }

        private void ConfigureStep2Grid(DataGridView dgv, bool allowCheck)
        {
            dgv.Columns.Clear();

            if (allowCheck)
            {
                dgv.Columns.Add(new DataGridViewCheckBoxColumn
                {
                    Name = "Selected",
                    HeaderText = "",
                    DataPropertyName = "Selected",
                    Width = 34
                });
            }

            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Location_ID",
                HeaderText = "ID",
                DataPropertyName = "Location_ID",
                Visible = false
            });

            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Location_Name",
                HeaderText = "Nombre",
                DataPropertyName = "Location_Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                ReadOnly = true
            });

            dgv.ReadOnly = false;
            dgv.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void BindEmptyMessage(DataGridView dgv, string msg)
        {
            var dt = new DataTable();
            dt.Columns.Add("Location_ID", typeof(int));
            dt.Columns.Add("Location_Name", typeof(string));
            dt.Columns.Add("Selected", typeof(bool));
            dt.Rows.Add(0, msg, false);
            dgv.DataSource = dt;
        }

        private void LoadAlmacenes()
        {
            EnsureStep2HierarchyUI();

            var dt = _repo.LoadWarehouses(Step2FixedAlmacenId);

            foreach (DataRow r in dt.Rows)
                r["Selected"] = _checkedAlmacenIds.Contains(Convert.ToInt32(r["Location_ID"]));

            _dgvAlm.DataSource     = dt;
            _lblAlmCount.Text      = dt.Rows.Count.ToString();
            _focusedAlmacenId      = dt.Rows.Count > 0 ? Step2FixedAlmacenId : (int?)null;
            LoadAreasForFocusedAlmacen();
            RebuildEffectiveSelection();
        }

        private void LoadAreasForFocusedAlmacen()
        {
            if (_focusedAlmacenId == null || _focusedAlmacenId <= 0)
            {
                BindEmptyMessage(_dgvArea, "Selecciona un almacén");
                _lblAreaCount.Text = "0";
                return;
            }

            var dt = _repo.LoadAreas(_focusedAlmacenId.Value);

            foreach (DataRow r in dt.Rows)
                r["Selected"] = _checkedAreaIds.Contains(Convert.ToInt32(r["Location_ID"]));

            _dgvArea.DataSource = dt;
            _lblAreaCount.Text  = dt.Rows.Count.ToString();
        }

        private void RebuildEffectiveSelection()
        {
            _effectiveSelectedLocationIds.Clear();

            // 1) Locaciones marcadas manualmente
            foreach (int locId in _checkedLocationIds)
                _effectiveSelectedLocationIds.Add(locId);

            // 2) Áreas checked sin locaciones hijas
            foreach (int areaId in _checkedAreaIds)
            {
                if (_repo.CountChildLocations(areaId) == 0)
                    _effectiveSelectedLocationIds.Add(areaId);
            }

            // 3) Almacenes checked sin áreas hijas
            foreach (int almId in _checkedAlmacenIds)
            {
                if (_repo.CountChildLocations(almId) == 0)
                    _effectiveSelectedLocationIds.Add(almId);
            }
        }

        private bool SaveSelectedLocations()
        {
            try
            {
                if (_physicalCountId <= 0)
                {
                    MessageBox.Show("ID de inventario inválido. Guarde el encabezado primero.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                RebuildEffectiveSelection();

                if (_effectiveSelectedLocationIds.Count == 0)
                {
                    MessageBox.Show("Seleccione al menos una locación.", "Validación",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Elimina las locaciones seleccionadas anteriores
                _repo.DeleteSelectedLocations(_physicalCountId);

                int userId = GetCurrentUserId();

                // 1) Guardar LOCACIONES marcadas manualmente
                foreach (int locId in _checkedLocationIds)
                {
                    int? parentAreaId = null;
                    try { parentAreaId = _repo.GetLocationParentId(locId); } catch { }

                    _repo.SaveSelectedLocationRow(
                        _physicalCountId,
                        selectedId: locId, selectedLevel: 3,
                        effectiveId: locId, effectiveLevel: 3,
                        parentId: parentAreaId,
                        userId: userId);
                }

                // 2) Guardar ÁREAS sin locaciones hijas
                foreach (int areaId in _checkedAreaIds)
                {
                    if (_repo.CountChildLocations(areaId) == 0)
                    {
                        _repo.SaveSelectedLocationRow(
                            _physicalCountId,
                            selectedId: areaId, selectedLevel: 2,
                            effectiveId: areaId, effectiveLevel: 2,
                            parentId: _focusedAlmacenId,
                            userId: userId);
                    }
                }

                // 3) Guardar ALMACÉN solo si no tiene áreas
                foreach (int almId in _checkedAlmacenIds)
                {
                    if (_repo.CountChildLocations(almId) == 0)
                    {
                        _repo.SaveSelectedLocationRow(
                            _physicalCountId,
                            selectedId: almId, selectedLevel: 1,
                            effectiveId: almId, effectiveLevel: 1,
                            parentId: null,
                            userId: userId);
                    }
                }

                int saved = _repo.CountSavedLocations(_physicalCountId);

                if (saved <= 0)
                {
                    MessageBox.Show(
                        "No se encontraron partes para las locaciones seleccionadas.\n" +
                        "Seleccione otra locación.",
                        "Sin existencias",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return false;
                }

                MessageBox.Show(
                    "Locaciones guardadas correctamente.\n\n" +
                    "El conteo se generará en la plataforma web.\n" +
                    "Haga clic en 'Verificar Conteo' cuando esté completo.",
                    "Éxito",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar locaciones: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }


        private void ApplyStep2Filter(DataGridView dgv, string text)
        {
            if (dgv?.DataSource == null) return;

            // soporta DataTable directo o DataView
            DataView view = null;

            if (dgv.DataSource is DataTable dt)
                view = dt.DefaultView;
            else if (dgv.DataSource is DataView dv)
                view = dv;

            if (view == null) return;

            string q = (text ?? "").Trim();

            // ignora placeholder
            if (string.IsNullOrWhiteSpace(q) || q.StartsWith("Buscar ", StringComparison.OrdinalIgnoreCase))
            {
                view.RowFilter = "";
                return;
            }

            q = q.Replace("'", "''");
            // columnas esperadas: Location_Name, Location_ID
            view.RowFilter =
                $"Location_Name LIKE '%{q}%' OR Convert(Location_ID, 'System.String') LIKE '%{q}%'";
        }

        private void ChkSelectAllSubLocations_CheckedChanged(object sender, EventArgs e)
        {
            if (_updatingSelectAllSubLocations) return;
            if (_dgvSub == null) return;
            if (!_dgvSub.Columns.Contains("Selected")) return;

            bool checkAll = ((CheckBox)sender).Checked;

            _dgvSub.SuspendLayout();
            try
            {
                foreach (DataGridViewRow row in _dgvSub.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (row.Cells["Location_ID"].Value == DBNull.Value) continue;

                    int id = Convert.ToInt32(row.Cells["Location_ID"].Value);
                    if (id <= 0) continue;

                    row.Cells["Selected"].Value = checkAll;

                    if (checkAll)
                        _checkedLocationIds.Add(id);
                    else
                        _checkedLocationIds.Remove(id);
                }

                RebuildEffectiveSelection();
            }
            finally
            {
                _dgvSub.ResumeLayout();
                _dgvSub.Refresh();
            }
        }

        private void UpdateSelectAllSubLocationsState()
        {
            if (_chkSelectAllSubLocations == null) return;

            _updatingSelectAllSubLocations = true;
            try
            {
                bool hasRows = _dgvSub != null && _dgvSub.Rows.Count > 0;
                _chkSelectAllSubLocations.Enabled = hasRows;

                if (!hasRows)
                {
                    _chkSelectAllSubLocations.Checked = false;
                    return;
                }

                bool allChecked = true;

                foreach (DataGridViewRow row in _dgvSub.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool isChecked = false;
                    if (row.Cells["Selected"].Value != null &&
                        row.Cells["Selected"].Value != DBNull.Value)
                    {
                        isChecked = Convert.ToBoolean(row.Cells["Selected"].Value);
                    }

                    if (!isChecked)
                    {
                        allChecked = false;
                        break;
                    }
                }

                _chkSelectAllSubLocations.Checked = allChecked;
            }
            finally
            {
                _updatingSelectAllSubLocations = false;
            }
        }

        private void Step2_ToggleCheck_Almacen(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!_dgvAlm.Columns.Contains("Selected")) return;
            if (e.ColumnIndex != _dgvAlm.Columns["Selected"].Index) return;

            int id = Convert.ToInt32(_dgvAlm.Rows[e.RowIndex].Cells["Location_ID"].Value);
            if (id <= 0) return;

            bool current = Convert.ToBoolean(_dgvAlm.Rows[e.RowIndex].Cells["Selected"].Value);
            bool next = !current;
            _dgvAlm.Rows[e.RowIndex].Cells["Selected"].Value = next;

            if (next) _checkedAlmacenIds.Add(id);
            else _checkedAlmacenIds.Remove(id);

            // Regla: solo nos interesa el almacén 3
            bool almacen3Checked = _checkedAlmacenIds.Contains(Step2FixedAlmacenId);

            if (!almacen3Checked)
            {
                // Limpia todo downstream
                _checkedAreaIds.Clear();
                _checkedLocationIds.Clear();
                BindEmptyMessage(_dgvArea, "Selecciona un almacén");
                _lblAreaCount.Text = "0";
                BindEmptyMessage(_dgvSub, "Selecciona un área");
                _lblSubCount.Text = "0";


            }
            else
            {
                // Cargar áreas del almacén 3
                _focusedAlmacenId = Step2FixedAlmacenId;
                LoadAreasForFocusedAlmacen();
                // Al cargar áreas no cargamos SUB hasta que el usuario marque áreas
                _checkedLocationIds.Clear();
                BindEmptyMessage(_dgvSub, "Selecciona un área");
                _lblSubCount.Text = "0";
            }

            RebuildEffectiveSelection();
        }

        private void Step2_ToggleCheck_Area(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!_dgvArea.Columns.Contains("Selected")) return;
            if (e.ColumnIndex != _dgvArea.Columns["Selected"].Index) return;

            int id = Convert.ToInt32(_dgvArea.Rows[e.RowIndex].Cells["Location_ID"].Value);
            if (id <= 0) return;

            bool current = Convert.ToBoolean(_dgvArea.Rows[e.RowIndex].Cells["Selected"].Value);
            bool next = !current;
            _dgvArea.Rows[e.RowIndex].Cells["Selected"].Value = next;

            if (next) _checkedAreaIds.Add(id);
            else _checkedAreaIds.Remove(id);

            if (_checkedAreaIds.Count == 0)
            {
                _checkedLocationIds.Clear();
                _effectiveSelectedLocationIds.Clear();
                BindEmptyMessage(_dgvSub, "Selecciona un área");
                _lblSubCount.Text = "0";
                UpdateSelectAllSubLocationsState();
            }
            else
            {
                LoadSubAreasForCheckedAreas();
            }

            RebuildEffectiveSelection();
        }

        private void Step2_ToggleCheck_Location(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!_dgvSub.Columns.Contains("Selected")) return;
            if (e.ColumnIndex != _dgvSub.Columns["Selected"].Index) return;

            var row = _dgvSub.Rows[e.RowIndex];

            object rawId = row.Cells["Location_ID"].Value;
            if (rawId == null || rawId == DBNull.Value) return;

            int id = Convert.ToInt32(rawId);
            if (id <= 0) return;

            bool current = false;
            if (row.Cells["Selected"].Value != null && row.Cells["Selected"].Value != DBNull.Value)
                current = Convert.ToBoolean(row.Cells["Selected"].Value);

            bool next = !current;
            row.Cells["Selected"].Value = next;

            var drv = row.DataBoundItem as DataRowView;
            if (drv != null && drv.Row.Table.Columns.Contains("Selected"))
                drv.Row["Selected"] = next;

            if (next) _checkedLocationIds.Add(id);
            else _checkedLocationIds.Remove(id);

            RebuildEffectiveSelection();

            _dgvSub.Refresh();
            _dgvSub.Invalidate();
        }

        private void LoadSubAreasForCheckedAreas()
        {
            if (_checkedAreaIds.Count == 0)
            {
                _checkedLocationIds.Clear();
                _effectiveSelectedLocationIds.Clear();
                BindEmptyMessage(_dgvSub, "Selecciona un área");
                _lblSubCount.Text = "0";
                UpdateSelectAllSubLocationsState();
                return;
            }

            var dt = _repo.LoadSubLocations(_checkedAreaIds.ToList());

            bool selectAll = _chkSelectAllSubLocations != null && _chkSelectAllSubLocations.Checked;

            foreach (DataRow r in dt.Rows)
            {
                int id = Convert.ToInt32(r["Location_ID"]);

                if (selectAll)
                {
                    r["Selected"] = true;
                    _checkedLocationIds.Add(id);
                }
                else
                {
                    r["Selected"] = _checkedLocationIds.Contains(id);
                }
            }

            _dgvSub.DataSource = dt;
            _lblSubCount.Text  = dt.Rows.Count.ToString();

            UpdateSelectAllSubLocationsState();
            RebuildEffectiveSelection();
            _dgvSub.Refresh();
        }








        #endregion


        // =====================================================================
        // #region STEP 3 
        // =====================================================================
        #region PASO 3 — CONTEO EN PLATAFORMA WEB

        // --- Step 3 Manual Refresh Button ---
        private Button _btnStep3ManualRefresh;


        private void EnsureStep3TabsUI()
        {
            pnlStep3_Confirm.Dock = DockStyle.Fill;
            pnlStep3_Confirm.BringToFront();
            if (_tabStep3 != null) return;

            // Collect and detach extra designer labels
            var exclude = new HashSet<Control>
            {
                lblConfirmDescription, lblConfirmType,
                lblConfirmLocations,   lblStep3Info
            };
            _step3ExtraLabels = CollectLabels(pnlStep3_Confirm, exclude);
            DetachControls(_step3ExtraLabels);

            pnlStep3_Confirm.Controls.Clear();

            _tabStep3 = new TabControl { Dock = DockStyle.Fill };
            _tabStep3Preview = new TabPage("Count Preview");
            _tabStep3Attachments = new TabPage("Attachments");

            _tabStep3.TabPages.Add(_tabStep3Preview);
            _tabStep3.TabPages.Add(_tabStep3Attachments);

            pnlStep3_Confirm.Controls.Add(_tabStep3);

            BuildStep3PreviewTab();
            BuildStep3AttachmentsTab();
        }

        /// <summary>
        /// Builds the Count Preview tab.
        /// No editing controls — strictly a monitoring view.
        /// </summary>
        private void BuildStep3PreviewTab()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.White
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // Top info panel
            var top = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            var left = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true
            };

            NormalizeLabel(lblConfirmDescription);
            NormalizeLabel(lblConfirmType);
            NormalizeLabel(lblConfirmLocations);
            NormalizeLabel(lblStep3Info);

            left.Controls.Add(lblConfirmDescription);
            left.Controls.Add(lblConfirmType);
            left.Controls.Add(lblConfirmLocations);
            left.Controls.Add(lblStep3Info);

            // Extra designer labels (no editing controls added)
            if (_step3ExtraLabels != null)
                foreach (var lbl in _step3ExtraLabels)
                {
                    NormalizeLabel(lbl);
                    left.Controls.Add(lbl);
                }

            top.Controls.Add(left);

            // Grid + botones
            dgvStep3Preview.Parent = null;
            dgvStep3Preview.Dock = DockStyle.Fill;

            // ── Botón Export (ya existente) ──────────────────────────────────
            btnStep3ExportExcel.Parent = null;
            btnStep3ExportExcel.Width = 180;
            btnStep3ExportExcel.Height = 36;
            StyleButtonSecondary(btnStep3ExportExcel);

            // ── Botón Refresh Manual (nuevo) ─────────────────────────────────
            _btnStep3ManualRefresh = new Button
            {
                Text = "⟳  Refresh",
                Width = 130,
                Height = 36,

            };
            StyleButtonPrimary(_btnStep3ManualRefresh);

            // Tooltip informativo
            var tip = new ToolTip();
            tip.SetToolTip(
                _btnStep3ManualRefresh,
                "Manually refresh the count data.\n" +
                "The automatic refresh every 3 seconds is still active.");

            // Handler del botón manual
            _btnStep3ManualRefresh.Click -= BtnStep3ManualRefresh_Click;
            _btnStep3ManualRefresh.Click += BtnStep3ManualRefresh_Click;

            // ── Barra inferior ────────────────────────────────────────────────
            var exportBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                BackColor = Color.White
            };

            exportBar.Controls.Add(btnStep3ExportExcel);
            exportBar.Controls.Add(_btnStep3ManualRefresh);

            // Posicionar ambos botones al hacer resize
            exportBar.Resize += (s, e) =>
            {
                // Export → extremo derecho
                btnStep3ExportExcel.Left = exportBar.Width - btnStep3ExportExcel.Width - 8;
                btnStep3ExportExcel.Top = 8;

                // Refresh → extremo izquierdo
                _btnStep3ManualRefresh.Left = 8;
                _btnStep3ManualRefresh.Top = 8;
            };

            var gridWrap = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            gridWrap.Controls.Add(dgvStep3Preview);
            gridWrap.Controls.Add(exportBar);

            root.Controls.Add(top, 0, 0);
            root.Controls.Add(gridWrap, 0, 1);

            _tabStep3Preview.Controls.Clear();
            _tabStep3Preview.Controls.Add(root);
        }
        /// <summary>
        /// Configures the Step 3 preview grid columns.
        /// All columns are READ-ONLY — no editing allowed from APS.
        /// </summary>
        /// 


        /// <summary>Builds the "Attachments" tab layout.</summary>
        private void BuildStep3AttachmentsTab()
        {
            if (_tabStep3Attachments == null) return;
            _tabStep3Attachments.Controls.Clear();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1,
                BackColor = Color.White,
                Padding = new Padding(8)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));

            // Toolbar
            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.White
            };

            _btnAttachAdd = new Button { Text = "Add File", Width = 140, Height = 32 };
            _btnAttachOpen = new Button { Text = "Open", Width = 90, Height = 32 };
            _btnAttachRemove = new Button { Text = "Remove", Width = 90, Height = 32 };

            StyleButtonPrimary(_btnAttachAdd);
            StyleButtonSecondary(_btnAttachOpen);
            StyleButtonGhost(_btnAttachRemove);

            _btnAttachAdd.Click -= AttachAdd_Click; _btnAttachAdd.Click += AttachAdd_Click;
            _btnAttachOpen.Click -= AttachOpen_Click; _btnAttachOpen.Click += AttachOpen_Click;
            _btnAttachRemove.Click -= AttachRemove_Click; _btnAttachRemove.Click += AttachRemove_Click;

            toolbar.Controls.Add(_btnAttachAdd);
            toolbar.Controls.Add(_btnAttachOpen);
            toolbar.Controls.Add(_btnAttachRemove);

            // List
            _lvAttachments = new ListView
            {
                Dock = DockStyle.Fill,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false,
                View = View.Details,
                GridLines = true
            };
            _lvAttachments.Columns.Add("File", 280);
            _lvAttachments.Columns.Add("Type", 120);
            _lvAttachments.Columns.Add("Size", 110);
            _lvAttachments.Columns.Add("Modified", 170);
            _lvAttachments.DoubleClick -= AttachOpen_Click;
            _lvAttachments.DoubleClick += AttachOpen_Click;

            // Info label
            _lblAttachInfo = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(75, 85, 99),
                Text = "You can attach images, PDFs, Excel files or any file related to this count."
            };
            NormalizeLabel(_lblAttachInfo);

            root.Controls.Add(toolbar, 0, 0);
            root.Controls.Add(_lvAttachments, 0, 1);
            root.Controls.Add(_lblAttachInfo, 0, 2);

            _tabStep3Attachments.Controls.Add(root);
            LoadAttachments();
        }

        /// <summary>
        /// Refresh manual del Step 3.
        /// Recarga los datos sin detener el timer automático.
        /// </summary>
        private void BtnStep3ManualRefresh_Click(object sender, EventArgs e)
        {
            if (_step3Loading) return; // evitar doble clic mientras carga

            _step3Loading = true;

            if (_btnStep3ManualRefresh != null)
            {
                _btnStep3ManualRefresh.Enabled = false;
                _btnStep3ManualRefresh.Text = "Refreshing...";
            }

            try
            {
                RefreshStep3Grid();
                RefreshStep3Summary();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error refreshing data: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _step3Loading = false;

                if (_btnStep3ManualRefresh != null)
                {
                    _btnStep3ManualRefresh.Enabled = true;
                    _btnStep3ManualRefresh.Text = "⟳  Refresh";
                }
            }
        }


        // --- Attachment helpers ---
        /// <summary>Carga los adjuntos del conteo en el ListView.</summary>
        private void LoadAttachments()
        {
            if (_lvAttachments == null) return;

            _lvAttachments.BeginUpdate();
            _lvAttachments.Items.Clear();

            if (_physicalCountId <= 0)
            {
                _lvAttachments.EndUpdate();
                return;
            }

            try
            {
                var dt = _repo.LoadAttachments(_physicalCountId);

                foreach (DataRow r in dt.Rows)
                {
                    int      attachmentId = Convert.ToInt32(r["Physical_Count_Attachment_ID"]);
                    string   fileName     = Convert.ToString(r["File_Name"]);
                    string   ext          = Convert.ToString(r["File_Extension"]);
                    long     size         = Convert.ToInt64(r["File_Size"]);
                    DateTime modified     = Convert.ToDateTime(r["Created_Date"]);

                    var item = new ListViewItem(fileName);
                    item.SubItems.Add((ext ?? "").Replace(".", "").ToUpperInvariant());
                    item.SubItems.Add((size / 1024.0).ToString("N1") + " KB");
                    item.SubItems.Add(modified.ToString("dd/MM/yyyy HH:mm"));
                    item.Tag = attachmentId;
                    _lvAttachments.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar adjuntos: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _lvAttachments.EndUpdate();
            }
        }

        /// <summary>Agrega uno o varios archivos como adjuntos del conteo.</summary>
        private void AttachAdd_Click(object sender, EventArgs e)
        {
            if (_physicalCountId <= 0)
            {
                MessageBox.Show("Guarde primero el inventario para poder adjuntar archivos.",
                    "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var ofd = new OpenFileDialog { Multiselect = true, Filter = "Todos los archivos|*.*" })
            {
                if (ofd.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    int userId = GetCurrentUserId();

                    foreach (string src in ofd.FileNames)
                    {
                        byte[]   fileBytes = File.ReadAllBytes(src);
                        FileInfo fi        = new FileInfo(src);

                        _repo.SaveAttachment(
                            _physicalCountId,
                            fi.Name,
                            fi.Extension,
                            GetMimeType(fi.Extension),
                            fi.Length,
                            fileBytes,
                            userId);
                    }

                    LoadAttachments();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar adjunto: " + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>Abre el adjunto seleccionado en la aplicación predeterminada del SO.</summary>
        private void AttachOpen_Click(object sender, EventArgs e)
        {
            if (_lvAttachments?.SelectedItems.Count == 0)
            {
                MessageBox.Show("Seleccione un archivo para abrir.",
                    "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int    attachmentId = Convert.ToInt32(_lvAttachments.SelectedItems[0].Tag);
            string fileName;
            byte[] fileData;

            try
            {
                if (!_repo.GetAttachmentData(attachmentId, out fileName, out fileData))
                {
                    MessageBox.Show("Adjunto no encontrado.",
                        "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string tempPath = Path.Combine(
                    Path.GetTempPath(),
                    Guid.NewGuid().ToString() + "_" + fileName);

                File.WriteAllBytes(tempPath, fileData);
                System.Diagnostics.Process.Start(tempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al abrir adjunto: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>Devuelve el MIME type según la extensión del archivo.</summary>
        private string GetMimeType(string extension)
        {
            switch ((extension ?? "").ToLowerInvariant())
            {
                case ".pdf":  return "application/pdf";
                case ".xls":  return "application/vnd.ms-excel";
                case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case ".doc":  return "application/msword";
                case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case ".png":  return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".txt":  return "text/plain";
                default:      return "application/octet-stream";
            }
        }

        /// <summary>Elimina el adjunto seleccionado del conteo.</summary>
        private void AttachRemove_Click(object sender, EventArgs e)
        {
            if (_lvAttachments?.SelectedItems.Count == 0)
            {
                MessageBox.Show("Seleccione un archivo para eliminar.",
                    "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int    attachmentId = Convert.ToInt32(_lvAttachments.SelectedItems[0].Tag);
            string name         = _lvAttachments.SelectedItems[0].Text;

            if (MessageBox.Show(
                    "¿Eliminar el archivo '" + name + "'?",
                    "Confirmar",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                _repo.DeleteAttachment(attachmentId);
                LoadAttachments();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al eliminar adjunto: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        // --- Step 3 grid configuration ---

        /// <summary>Configures columns for the Step 3 preview grid.</summary>
        private void ConfigureStep3PreviewGrid()
        {
            if (dgvStep3Preview == null) return;

            dgvStep3Preview.AutoGenerateColumns = false;
            dgvStep3Preview.AllowUserToAddRows = false;
            dgvStep3Preview.RowHeadersVisible = false;
            dgvStep3Preview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvStep3Preview.MultiSelect = false;

            // ✅ Always read-only — web platform owns the data
            dgvStep3Preview.ReadOnly = true;
            dgvStep3Preview.EditMode = DataGridViewEditMode.EditProgrammatically;

            dgvStep3Preview.Columns.Clear();

            void AddCol(string name, string header, string prop,
                        int width = 100,
                        DataGridViewAutoSizeColumnMode fill = DataGridViewAutoSizeColumnMode.None,
                        string format = null,
                        bool visible = true,
                        DataGridViewContentAlignment align = DataGridViewContentAlignment.MiddleLeft)
            {
                var col = new DataGridViewTextBoxColumn
                {
                    Name = name,
                    HeaderText = header,
                    DataPropertyName = prop,
                    ReadOnly = true,   // always read-only
                    Width = width,
                    AutoSizeMode = fill,
                    Visible = visible
                };
                if (format != null)
                    col.DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Format = format,
                        Alignment = align
                    };
                dgvStep3Preview.Columns.Add(col);
            }

            AddCol("Physical_Count_Snapshot_ID", "ID", "Physical_Count_Snapshot_ID", visible: false);
            AddCol("Location", "Location", "Location", width: 120);
            AddCol("PartNumber", "Part Number", "PartNumber", width: 130);
            AddCol("Description", "Description", "Description", fill: DataGridViewAutoSizeColumnMode.Fill);
            AddCol("Dimension", "Dimension", "Dimension", width: 100);
            AddCol("HeatNumber", "Heat Number", "HeatNumber", width: 120);
            AddCol("UM", "UM", "UM", width: 60);
            AddCol("SystemStock", "System Stock", "SystemStock", width: 110,
                   format: "N2", align: DataGridViewContentAlignment.MiddleRight);
            AddCol("CountedQty", "Counted Qty", "CountedQty", width: 110,
                   format: "N2", align: DataGridViewContentAlignment.MiddleRight);
            AddCol("Difference", "Difference", "Difference", width: 110,
                   format: "N2", align: DataGridViewContentAlignment.MiddleRight);

            // ✅ Only wire DataError — NO CellEndEdit, NO CellBeginEdit edit handlers
            dgvStep3Preview.DataError -= Step3Grid_DataError;
            dgvStep3Preview.DataError += Step3Grid_DataError;
            WireStep3DifferenceCellFormatting();
        }


        private void Step3Grid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }
        /// <summary>
        /// Sets the Step 3 grid visual style to read-only (blue tint).
        /// Called every time the grid refreshes.
        /// </summary>
        /// 
        private void Step3Grid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Editing is always read-only in web mode
            e.Cancel = true;
        }

        private void Step3Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e) { }

        // --- Step 3 timer (real-time refresh) ---

        private void StartStep3RefreshTimer()
        {
            if (_step3RefreshTimer != null) return;

            _step3RefreshTimer = new System.Windows.Forms.Timer();
            _step3RefreshTimer.Interval = Step3RefreshMs;
            _step3RefreshTimer.Tick += Step3Timer_Tick;
            _step3RefreshTimer.Start();
        }

        private void StopStep3RefreshTimer()
        {
            if (_step3RefreshTimer == null) return;
            _step3RefreshTimer.Stop();
            _step3RefreshTimer.Dispose();
            _step3RefreshTimer = null;
        }

        private void Step3Timer_Tick(object sender, EventArgs e)
        {
            try { RefreshStep3Grid(); }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Step 3 timer error: " + ex.Message);
            }
        }

        /// <summary>
        /// Verifies that the web platform has set the count status to Finished (3).
        /// APS cannot advance to Step 4 until the web count is complete.
        /// </summary>

        // --- Carga de datos del Paso 3 ---

        /// <summary>Actualiza las etiquetas de resumen del Paso 3.</summary>
        private void RefreshStep3Summary()
        {
            // Nombre del conteo
            try
            {
                string nombre = _repo.GetCountName(_physicalCountId);
                if (lblConfirmDescription != null)
                    lblConfirmDescription.Text = "Nombre del conteo: " + (nombre ?? "").Trim();
            }
            catch
            {
                if (lblConfirmDescription != null)
                    lblConfirmDescription.Text = "Nombre del conteo: -";
            }

            if (lblConfirmType != null)
                lblConfirmType.Text = "Tipo: " + (cboCountType?.Text ?? "");

            try
            {
                int c = _repo.CountSelectedLocations(_physicalCountId);
                if (lblConfirmLocations != null)
                    lblConfirmLocations.Text = "Locaciones seleccionadas: " + c;
            }
            catch
            {
                if (lblConfirmLocations != null)
                    lblConfirmLocations.Text = "Locaciones seleccionadas: -";
            }
        }

        /// <summary>
        /// Initializes Step 3 preview — ensures snapshots exist and loads the grid.
        /// </summary>
        private void InitStep3Preview()
        {
            _step3ReportReady = false;
            _dtStep3Preview = null;
            btnStep3ExportExcel.Enabled = false;
            dgvStep3Preview.DataSource = null;
            btnNext.Text = "Verify Count";

            ConfigureStep3PreviewGrid();
            SetStep3GridReadOnly();

            if (!EnsureSnapshotsExist())
            {
                if (lblStep3Info != null)
                    lblStep3Info.Text =
                        "No movements found for selected locations.\n" +
                        "Please verify stock exists in those locations.";
                return;
            }

            RefreshStep3Grid();
            StartStep3RefreshTimer();
        }

        /// <summary>
        /// Refresca la grilla del Paso 3 desde Physical_Count_Snapshots.
        /// Solo lectura — la plataforma web posee el campo Found.
        /// </summary>
        private void RefreshStep3Grid()
        {
            if (_physicalCountId <= 0)
            {
                dgvStep3Preview.DataSource = null;
                return;
            }

            try
            {
                var dtNew = _repo.LoadStep3Grid(_physicalCountId);

                if (!HasDataChanged(_dtStep3Preview, dtNew))
                    return;

                _dtStep3Preview = dtNew;

                if (dgvStep3Preview.Columns.Count == 0)
                    ConfigureStep3PreviewGrid();

                int firstRowIndex    = dgvStep3Preview.FirstDisplayedScrollingRowIndex;
                int selectedRowIndex = dgvStep3Preview.CurrentCell?.RowIndex ?? -1;
                int horizontalOffset = dgvStep3Preview.HorizontalScrollingOffset;

                dgvStep3Preview.DataSource = null;
                dgvStep3Preview.DataSource = _dtStep3Preview;

                if (dgvStep3Preview.Columns.Contains("Physical_Count_Snapshot_ID"))
                    dgvStep3Preview.Columns["Physical_Count_Snapshot_ID"].Visible = false;

                SetStep3GridReadOnly();

                if (horizontalOffset > 0)
                    dgvStep3Preview.HorizontalScrollingOffset = horizontalOffset;

                if (firstRowIndex > 0 && firstRowIndex < dgvStep3Preview.RowCount)
                {
                    try { dgvStep3Preview.FirstDisplayedScrollingRowIndex = firstRowIndex; }
                    catch { }
                }

                if (selectedRowIndex >= 0 && selectedRowIndex < dgvStep3Preview.RowCount)
                {
                    try
                    {
                        dgvStep3Preview.CurrentCell =
                            dgvStep3Preview.Rows[selectedRowIndex].Cells[
                                dgvStep3Preview.CurrentCell?.ColumnIndex ?? 0];
                    }
                    catch { }
                }

                btnStep3ExportExcel.Enabled = (_dtStep3Preview.Rows.Count > 0);

                int total   = _dtStep3Preview.Rows.Count;
                int counted = 0;

                foreach (DataRow r in _dtStep3Preview.Rows)
                {
                    object found = r["CountedQty"];
                    if (found != null && found != DBNull.Value &&
                        Convert.ToDecimal(found) != 0)
                        counted++;
                }

                if (lblStep3Info != null)
                {
                    int pct = (total > 0) ? (counted * 100 / total) : 0;
                    lblStep3Info.Text =
                        "Conteo en progreso en plataforma web...\n" +
                        "Total líneas: " + total + "  |  Contadas: " +
                        counted + " / " + total + " (" + pct + "%)";
                }

                WireStep3DifferenceCellFormatting();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al refrescar la grilla del Paso 3: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// Sets the Step 3 grid to read-only mode (web count, no editing in APS).
        /// </summary>
        private void SetStep3GridReadOnly()
        {
            if (dgvStep3Preview == null) return;

            dgvStep3Preview.ReadOnly = true;
            dgvStep3Preview.EditMode = DataGridViewEditMode.EditProgrammatically;
            dgvStep3Preview.DefaultCellStyle.BackColor = Color.FromArgb(240, 244, 255);
            dgvStep3Preview.DefaultCellStyle.ForeColor = Color.FromArgb(17, 24, 39);

            foreach (DataGridViewColumn col in dgvStep3Preview.Columns)
                col.ReadOnly = true;
        }

        /// <summary>
        /// Verifica si el conteo en la plataforma web está en estado Finalizado.
        /// APS no puede avanzar al Paso 4 hasta que el conteo web esté completo.
        /// </summary>
        private bool ValidateCountIsFinished()
        {
            if (_physicalCountId <= 0)
            {
                MessageBox.Show("No hay un ID de inventario válido.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {
                int    statusId;
                string statusName;
                bool   isFinished;

                bool ok = _repo.GetCountStatus(_physicalCountId,
                    out statusId, out statusName, out isFinished);

                if (!ok)
                {
                    MessageBox.Show("No se pudo obtener el estado del conteo.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                if (!isFinished)
                {
                    MessageBox.Show(
                        "El conteo aún no ha sido completado en la plataforma web.\n\n" +
                        "Estado actual: " + statusName + " (ID: " + statusId + ")\n\n" +
                        "Espere a que la plataforma web termine el conteo.",
                        "Conteo no finalizado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return false;
                }

                MessageBox.Show(
                    "Conteo completado correctamente.\n" +
                    "Puede revisar los resultados en el siguiente paso.",
                    "Éxito",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al validar el conteo: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        /// <summary>
        /// Compares two DataTables to detect changes.
        /// Used to avoid unnecessary grid refreshes on the timer tick.
        /// </summary>

        // --- Step 3 export ---

        private void btnStep3ExportExcel_Click(object sender, EventArgs e)
        {
            if (_dtStep3Preview == null || _dtStep3Preview.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog { Filter = "CSV|*.csv", FileName = $"Count_{_physicalCountId}.csv" })
            {
                if (sfd.ShowDialog(this) != DialogResult.OK) return;
                ExportToCsv(_dtStep3Preview, sfd.FileName);
                MessageBox.Show("Exported: " + sfd.FileName);
            }
        }

        /// <summary>
        /// Compares two DataTables to detect changes (used to avoid unnecessary grid refreshes).
        /// </summary>
        private bool HasDataChanged(DataTable oldDt, DataTable newDt)
        {
            if (oldDt == null && newDt == null) return false;
            if (oldDt == null || newDt == null) return true;
            if (oldDt.Rows.Count != newDt.Rows.Count) return true;

            for (int i = 0; i < oldDt.Rows.Count; i++)
                for (int c = 0; c < oldDt.Columns.Count; c++)
                {
                    object a = oldDt.Rows[i][c];
                    object b = newDt.Rows[i][c];
                    if ((a == null) != (b == null)) return true;
                    if (a != null && !a.Equals(b)) return true;
                }

            return false;
        }

        //Flechas Grid  
        private void WireStep3DifferenceCellFormatting()
        {
            dgvStep3Preview.CellPainting -= Step3Grid_DifferenceCellPainting;
            dgvStep3Preview.CellPainting += Step3Grid_DifferenceCellPainting;
        }

        private void Step3Grid_DifferenceCellPainting(
            object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!dgvStep3Preview.Columns.Contains("Difference")) return;
            if (e.ColumnIndex != dgvStep3Preview.Columns["Difference"].Index) return;

            // NULL = aún no contado → celda normal sin badge
            if (e.Value == null || e.Value == DBNull.Value)
            {
                e.Handled = false;
                return;
            }

            decimal diff;
            if (!decimal.TryParse(Convert.ToString(e.Value), out diff))
            {
                e.Handled = false;
                return;
            }

            // ── Leer Found (CountedQty) de la fila ───────────────────────────────
            decimal found = 0;
            if (dgvStep3Preview.Columns.Contains("CountedQty"))
            {
                object fVal = dgvStep3Preview.Rows[e.RowIndex]
                                  .Cells["CountedQty"].Value;
                if (fVal != null && fVal != DBNull.Value)
                    found = Convert.ToDecimal(fVal);
            }

            // ── Colores y texto del badge según caso ──────────────────────────────
            Color rowBack, badgeBack, badgeFore, badgeBorder;
            string badgeText;

            if (diff > 0 && found == 0)
            {
                // Found = 0, Quantity > 0 → NO CONTADO → 🔴 ROJO sin flecha
                rowBack = e.CellStyle.BackColor;
                badgeBack = Color.FromArgb(254, 226, 226);
                badgeFore = Color.FromArgb(185, 28, 28);
                badgeBorder = Color.FromArgb(252, 165, 165);
                badgeText = $"{SpacedNumber(diff)}";
            }
            else if (diff > 0 && found > 0)
            {
                // Found > Quantity → 🟢 VERDE ^ +
                rowBack = e.CellStyle.BackColor;
                badgeBack = Color.FromArgb(220, 252, 231);
                badgeFore = Color.FromArgb(21, 128, 61);
                badgeBorder = Color.FromArgb(134, 239, 172);
                badgeText = $"↑ +{SpacedNumber(diff)}";

            }
            else if (diff == 0)
            {
                // Found = Quantity → 🟡 BEIGE
                rowBack = e.CellStyle.BackColor;
                badgeBack = Color.FromArgb(245, 245, 245);
                badgeFore = Color.FromArgb(100, 100, 100);
                badgeBorder = Color.FromArgb(210, 210, 210);
                badgeText = $"{SpacedNumber(diff)}";
            }
            else if (diff < 0 && found == 0)
            {
                // Found = 0, Quantity > 0 → 🔴 ROJO sin flecha
                rowBack = e.CellStyle.BackColor;
                badgeBack = Color.FromArgb(254, 226, 226);
                badgeFore = Color.FromArgb(185, 28, 28);
                badgeBorder = Color.FromArgb(252, 165, 165);
                badgeText = $"{SpacedNumber(diff)}";
            }
            else
            {
                // Found < Quantity → 🔴 ROJO v -
                rowBack = e.CellStyle.BackColor;
                badgeBack = Color.FromArgb(254, 226, 226);
                badgeFore = Color.FromArgb(185, 28, 28);
                badgeBorder = Color.FromArgb(252, 165, 165);
                badgeText = $"↓ - {SpacedNumber(diff)}";
            }

            // ── Pintar fondo de la celda (igual al de la fila) ────────────────────
            using (var rowBrush = new SolidBrush(rowBack))
                e.Graphics.FillRectangle(rowBrush, e.CellBounds);

            // ── Dimensiones del badge ─────────────────────────────────────────────
            const int padH = 14;
            const int padV = 6;
            const int radius = 6;

            using (var fntBadge = new Font("Segoe UI", 8F, FontStyle.Bold))
            {
                SizeF textSize = e.Graphics.MeasureString(badgeText, fntBadge);

                int badgeW = (int)textSize.Width + padH * 2 + 6;
                int badgeH = (int)textSize.Height + padV * 2;

                // Centrar badge en la celda
                int bx = e.CellBounds.X + (e.CellBounds.Width - badgeW) / 2;
                int by = e.CellBounds.Y + (e.CellBounds.Height - badgeH) / 2;

                // Clamp para que no desborde
                bx = Math.Max(e.CellBounds.X + 4, bx);
                by = Math.Max(e.CellBounds.Y + 2, by);
                badgeW = Math.Min(badgeW, e.CellBounds.Width - 8);
                badgeH = Math.Min(badgeH, e.CellBounds.Height - 4);

                var badgeRect = new Rectangle(bx, by, badgeW, badgeH);

                e.Graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                // ── Fondo y borde del badge ───────────────────────────────────────
                using (var path = BuildBadgeRoundRect(badgeRect, radius))
                using (var fillBrush = new SolidBrush(badgeBack))
                using (var borderPen = new Pen(badgeBorder, 1f))
                {
                    e.Graphics.FillPath(fillBrush, path);
                    e.Graphics.DrawPath(borderPen, path);
                }

                // ── Texto del badge ───────────────────────────────────────────────
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter
                };

                using (var textBrush = new SolidBrush(badgeFore))
                    e.Graphics.DrawString(badgeText, fntBadge, textBrush, badgeRect, sf);
            }

            // ── Línea de borde inferior (estilo del grid) ─────────────────────────
            using (var gridLinePen = new Pen(dgvStep3Preview.GridColor))
                e.Graphics.DrawLine(gridLinePen,
                    e.CellBounds.Left, e.CellBounds.Bottom - 1,
                    e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);

            e.Handled = true;
        }
        // ── Helper para formatear número con espacios entre dígitos ──────────
        private string SpacedNumber(decimal value)
        {
            string raw = Math.Abs(value).ToString("N2");
            return string.Join("\u2009", raw.ToCharArray()); // \u2009 = thin space (espacio fino)
        }

        /// <summary>Builds a rounded rectangle GraphicsPath for the badge.</summary>
        private System.Drawing.Drawing2D.GraphicsPath BuildBadgeRoundRect(
            Rectangle r, int radius)
        {
            int d = radius * 2;
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }


        #endregion

        // =====================================================================
        // INICIALIZACIÓN DE GRILLA (pnlStep4_Capture permanece oculto)
        // ConfigureStep4Grid se llama desde Load; el panel no está en uso.
        // =====================================================================
        #region INICIALIZACIÓN GRILLA PASO 4

        /// <summary>
        /// Configura la grilla del Paso 4 (solo lectura, sin edición).
        /// El panel pnlStep4_Capture se mantiene oculto; este método se llama
        /// desde Load para inicializar la grilla antes de que el paso sea visible.
        /// </summary>
        private void ConfigureStep4Grid()
        {
            if (dgvStep4 == null) return;

            StyleGrid(dgvStep4);
            dgvStep4.AllowUserToAddRows    = false;
            dgvStep4.AllowUserToDeleteRows = false;
            dgvStep4.RowHeadersVisible     = false;
            dgvStep4.SelectionMode         = DataGridViewSelectionMode.FullRowSelect;
            dgvStep4.ReadOnly              = true;
            dgvStep4.ScrollBars            = ScrollBars.Both;

            dgvStep4.DataError -= Step4Grid_DataError;
            dgvStep4.DataError += Step4Grid_DataError;
        }

        private void Step4Grid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }

        #endregion


        // =====================================================================
        // #region STEP 5 — REVIEW & FINALIZE
        // =====================================================================
        #region PASO 4 — REVISIÓN Y CIERRE

        /// <summary>
        /// Ensures the Step 5 UI is created (lazy init).
        /// </summary>

        private ComboBox _cboStep5LocationFilter;
        private DataTable _dtStep5AllDetail;
        private DataTable _dtStep5AllExecutive;
        private Label _lblStep5HeaderTitle;
        private Label _lblStep5HeaderLine2;

        private void EnsureStep5UI()
        {
            if (_tabStep5 != null) return;
            if (pnlStep5_Review == null) return;

            pnlStep5_Review.Controls.Clear();
            pnlStep5_Review.BackColor = Color.White;
            pnlStep5_Review.Dock = DockStyle.Fill;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.White
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));  // header + filter/kpi
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // tabs

            // ── TOP: header info (row 0) | filter+kpi (row 1) ──
            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.White,
                Padding = new Padding(9, 4, 9, 4),
                Margin = new Padding(0)
            };
            top.RowStyles.Add(new RowStyle(SizeType.Absolute, 34)); // header line
            top.RowStyles.Add(new RowStyle(SizeType.Absolute, 34)); // filter + kpi line
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // Header info panel (line 1)
            var headerInfo = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            headerInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55)); // left
            headerInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45)); // right

            _lblStep5HeaderTitle = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 34,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Text = "Inventory",
                Margin = new Padding(0)
            };

            _lblStep5HeaderLine2 = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 34,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(55, 65, 81),
                Text = "",
                Margin = new Padding(0)
            };

            headerInfo.Controls.Add(_lblStep5HeaderTitle, 0, 0);
            headerInfo.Controls.Add(_lblStep5HeaderLine2, 1, 0);

            // Filter + KPI row (line 2)
            var filterAndKpi = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            filterAndKpi.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320)); // filter
            filterAndKpi.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // kpi

            // ── Location filter (left) ──
            var filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,               // ✅ no wrap
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0),
                AutoSize = false,
                Height = 34                          // ✅ fuerza una sola línea con altura fija
            };

            var lblFilter = new Label
            {
                Text = "Location:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                AutoSize = false,
                Width = 70,
                Height = 34,                         // ✅ misma altura de la fila
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 6, 0)     // ✅ sin offset vertical
            };

            _cboStep5LocationFilter = new ComboBox
            {
                Width = 220,
                Height = 34,                         // ✅ misma altura de la fila
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F),
                BackColor = Color.FromArgb(249, 250, 251),
                ForeColor = Color.FromArgb(17, 24, 39),
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0)              // ✅ sin offset vertical
            };

            // ⚠️ ComboBox a veces se “dibuja” 1px más abajo según tema.
            // Si lo notas, cambia a new Padding(0, 1, 0, 0) o (0, 2, 0, 0).
            _cboStep5LocationFilter.Margin = new Padding(0, 2, 0, 0);

            filterPanel.Controls.Add(lblFilter);
            filterPanel.Controls.Add(_cboStep5LocationFilter);

            // ── KPI bar (right) ──
            _lblStep5Kpis = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 34,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Black,
                Padding = new Padding(10, 0, 10, 0),
                Text = "",
                Margin = new Padding(0)
            };

            filterAndKpi.Controls.Add(filterPanel, 0, 0);
            filterAndKpi.Controls.Add(_lblStep5Kpis, 1, 0);

            // Put into top
            top.Controls.Add(headerInfo, 0, 0);
            top.Controls.Add(filterAndKpi, 0, 1);


            // ── TABS ──
            _tabStep5 = new TabControl
            {
                Dock = DockStyle.Fill,
                Appearance = TabAppearance.Normal,
                Multiline = false
            };

            _tabStep5Detail = new TabPage("Detail") { BackColor = Color.White };
            _tabStep5Charts = new TabPage("Charts") { BackColor = Color.White };
            _tabStep5Executive = new TabPage("Executive Summary") { BackColor = Color.White };

            _tabStep5.TabPages.Add(_tabStep5Detail);
            _tabStep5.TabPages.Add(_tabStep5Charts);
            _tabStep5.TabPages.Add(_tabStep5Executive);

            BuildStep5DetailTab();
            BuildStep5ChartsTab();
            BuildStep5ExecutiveTab();

            // Default open tab = Charts
            _tabStep5.SelectedTab = _tabStep5Charts;

            BuildStep5DetailTab();
            BuildStep5ExecutiveTab();

            root.Controls.Add(top, 0, 0);
            root.Controls.Add(_tabStep5, 0, 1);

            pnlStep5_Review.Controls.Add(root);

            // Init print documents
            InitStep5PrintDocuments();
        }

        private void BuildStep5ChartsTab()
        {
            if (_tabStep5Charts == null) return;

            _tabStep5Charts.Controls.Clear();
            _tabStep5Charts.Padding = new Padding(0);
            _tabStep5Charts.BackColor = Color.FromArgb(243, 244, 246);

            // Scroll container
            var scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(243, 244, 246),
                Padding = new Padding(12)
            };

            // Real content panel (THIS is the one we want to capture/print)
            var page = new Panel
            {
                Dock = DockStyle.Top,
                Height = 730,
                Width = 1120,
                BackColor = scroll.BackColor
            };

            // IMPORTANT:
            // Store the real charts page in the field so Preview / Print can capture it
            _pnlStep5ChartsHost = page;

            scroll.Resize += (s, e) =>
            {
                page.Width = Math.Max(1120, scroll.ClientSize.Width - 24);
            };

            // =========================
            // Toolbar
            // =========================
            var toolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = scroll.BackColor
            };

            _btnStep5PreviewCharts = new Button
            {
                Text = "Preview / Print",
                Width = 150,
                Height = 36,
                Top = 4
            };

            _btnStep5ExportPdfCharts = new Button
            {
                Text = "Export PDF",
                Width = 130,
                Height = 36,
                Top = 4
            };

            StyleButtonSecondary(_btnStep5PreviewCharts);
            StyleButtonPrimary(_btnStep5ExportPdfCharts);

            _btnStep5PreviewCharts.Enabled = true;
            _btnStep5ExportPdfCharts.Enabled = true;

            _btnStep5PreviewCharts.Click -= Step5PreviewCharts_Click;
            _btnStep5PreviewCharts.Click += Step5PreviewCharts_Click;

            _btnStep5ExportPdfCharts.Click -= Step5ExportPdfCharts_Click;
            _btnStep5ExportPdfCharts.Click += Step5ExportPdfCharts_Click;

            // Solo agrega esto si YA tienes el método implementado
            // _btnStep5ExportPdfCharts.Click -= Step5ExportPdfCharts_Click;
            // _btnStep5ExportPdfCharts.Click += Step5ExportPdfCharts_Click;

            toolbar.Controls.Add(_btnStep5PreviewCharts);
            toolbar.Controls.Add(_btnStep5ExportPdfCharts);

            toolbar.Resize += (s, e) =>
            {
                _btnStep5ExportPdfCharts.Left = toolbar.Width - _btnStep5ExportPdfCharts.Width;
                _btnStep5PreviewCharts.Left = _btnStep5ExportPdfCharts.Left - _btnStep5PreviewCharts.Width - 10;
            };

            // =========================
            // Top row (3 cards)
            // =========================
            var topRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 372,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = scroll.BackColor,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            topRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3333F));
            topRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3333F));
            topRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3333F));
            topRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var cardItems = CreateStep5ChartCard(
                "Items by Location",
                "Distribution of counted lines",
                out _pnlStep5ChartItemsByLocation);

            var cardAccuracy = CreateStep5ChartCard(
                "Accuracy by Location",
                "Green = exact / Red = discrepancy",
                out _pnlStep5ChartAccuracyByLocation);

            var cardTopDiff = CreateStep5ChartCard(
                "Top Discrepancies (Units)",
                "Highest absolute differences",
                out _pnlStep5ChartTopDiscrepancies);

            cardItems.Margin = new Padding(0, 0, 12, 0);

            cardTopDiff.Margin = new Padding(0);

            topRow.Controls.Add(cardItems, 0, 0);

            topRow.Controls.Add(cardTopDiff, 1, 0);

            // =========================
            // Bottom row (2 cards)
            // =========================
            var bottomRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 290,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = scroll.BackColor,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            bottomRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            bottomRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            bottomRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var cardStockVsCount = CreateStep5ChartCard(
                "System Stock vs Counted Qty",
                "Bubble size = difference magnitude",
                out _pnlStep5ChartStockVsCount);

            var cardDistribution = CreateStep5ChartCard(
                "Stock Distribution (Top 20)",
                "Red = with discrepancy",
                out _pnlStep5ChartStockDistribution);

            cardStockVsCount.Margin = new Padding(0, 0, 12, 0);
            cardDistribution.Margin = new Padding(0);

            bottomRow.Controls.Add(cardStockVsCount, 0, 0);
            bottomRow.Controls.Add(cardDistribution, 1, 0);

            // =========================
            // Real chart controls
            // =========================
            _chtStep5ItemsByLocation = CreateStep5DoughnutChart();
            _chtStep5TopDiscrepancies = CreateStep5TopDiscrepanciesChart();
            _chtStep5StockVsCount = CreateStep5BubbleChart();

            _chtStep5TopDiscrepancies.MouseMove -= Step5TopDiffChart_MouseMove;
            _chtStep5TopDiscrepancies.MouseMove += Step5TopDiffChart_MouseMove;

            _chtStep5TopDiscrepancies.MouseLeave -= Step5TopDiffChart_MouseLeave;
            _chtStep5TopDiscrepancies.MouseLeave += Step5TopDiffChart_MouseLeave;

            _pnlStep5ChartItemsByLocation.Controls.Clear();
            _pnlStep5ChartAccuracyByLocation.Controls.Clear();
            _pnlStep5ChartTopDiscrepancies.Controls.Clear();
            _pnlStep5ChartStockVsCount.Controls.Clear();
            _pnlStep5ChartStockDistribution.Controls.Clear();

            _pnlStep5ChartItemsByLocation.Controls.Add(_chtStep5ItemsByLocation);
            _pnlStep5ChartTopDiscrepancies.Controls.Add(_chtStep5TopDiscrepancies);
            _pnlStep5ChartStockVsCount.Controls.Add(_chtStep5StockVsCount);

            BuildStep5StockDistributionPanel();

            // Add in reverse order because Dock=Top
            page.Controls.Add(bottomRow);
            page.Controls.Add(CreateStep5ChartsSpacer(12));
            page.Controls.Add(topRow);
            page.Controls.Add(CreateStep5ChartsSpacer(12));
            page.Controls.Add(toolbar);

            scroll.Controls.Add(page);
            _tabStep5Charts.Controls.Add(scroll);

            BuildStep5TopDiffHoverCard();
        }


        private void Step5PreviewCharts_Click(object sender, EventArgs e)
        {
            try
            {
                if (_pnlStep5ChartsHost == null)
                {
                    MessageBox.Show("Charts area is not available.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (_step5ChartsPrintBitmap != null)
                {
                    _step5ChartsPrintBitmap.Dispose();
                    _step5ChartsPrintBitmap = null;
                }

                _step5ChartsPrintBitmap = CaptureControlToBitmap(_pnlStep5ChartsHost);

                if (_step5ChartsPrintBitmap == null)
                {
                    MessageBox.Show("Unable to capture charts for printing.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                EnsureStep5ChartsPrintObjects();
                _ppdStep5Charts.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating chart print preview: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BuildStep5StockDistributionPanel()
        {
            if (_pnlStep5ChartStockDistribution == null) return;

            _pnlStep5ChartStockDistribution.Controls.Clear();

            _pnlStep5TreemapSurface = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            _pnlStep5TreemapSurface.Resize += (s, e) =>
            {
                RenderStep5Treemap();
            };

            _pnlStep5ChartStockDistribution.Controls.Add(_pnlStep5TreemapSurface);
        }

        private sealed class Step5TreemapItem
        {
            public string Label { get; set; }
            public double Value { get; set; }
            public bool HasDifference { get; set; }
        }

        private Chart CreateStep5ChartBase(bool showLegend)
        {
            var chart = new Chart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Palette = ChartColorPalette.None,
                BorderlineColor = Color.White,
                BorderlineDashStyle = ChartDashStyle.Solid
            };

            var area = new ChartArea("MainArea");
            area.BackColor = Color.White;
            area.BorderColor = Color.White;

            area.Position.Auto = false;
            area.Position.X = 6;
            area.Position.Y = 8;
            area.Position.Width = showLegend ? 68 : 88;
            area.Position.Height = 82;

            area.InnerPlotPosition.Auto = false;
            area.InnerPlotPosition.X = 10;
            area.InnerPlotPosition.Y = 8;
            area.InnerPlotPosition.Width = 82;
            area.InnerPlotPosition.Height = 82;

            area.AxisX.LineColor = Color.FromArgb(156, 163, 175);
            area.AxisY.LineColor = Color.FromArgb(156, 163, 175);

            area.AxisX.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            area.AxisY.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);

            area.AxisX.LabelStyle.ForeColor = Color.FromArgb(107, 114, 128);
            area.AxisY.LabelStyle.ForeColor = Color.FromArgb(107, 114, 128);

            area.AxisX.LabelStyle.Font = new Font("Segoe UI", 8F);
            area.AxisY.LabelStyle.Font = new Font("Segoe UI", 8F);

            area.AxisX.TitleForeColor = Color.FromArgb(107, 114, 128);
            area.AxisY.TitleForeColor = Color.FromArgb(107, 114, 128);

            area.AxisX.TitleFont = new Font("Segoe UI", 8.5F);
            area.AxisY.TitleFont = new Font("Segoe UI", 8.5F);

            area.AxisX.LabelStyle.Format = "";  // ← ANTES: "0" — cada chart define su propio formato
            area.AxisY.LabelStyle.Format = "";  // ← ANTES: "0" — cada chart define su propio formato

            area.AxisX.IsMarginVisible = false; // ← ANTES: true — eliminamos ticks extra
            area.AxisY.IsMarginVisible = false; // ← ANTES: true — eliminamos ticks extra

            chart.ChartAreas.Add(area);

            if (showLegend)
            {
                var legend = new Legend("MainLegend");
                legend.Docking = Docking.Right;
                legend.Alignment = StringAlignment.Near;
                legend.BackColor = Color.White;
                legend.ForeColor = Color.FromArgb(75, 85, 99);
                legend.Font = new Font("Segoe UI", 8.5F);
                legend.BorderColor = Color.White;
                legend.IsTextAutoFit = false;
                legend.MaximumAutoSize = 28;
                chart.Legends.Add(legend);
            }

            return chart;
        }

        private Chart CreateStep5DoughnutChart()
        {
            var chart = CreateStep5ChartBase(true);
            var area = chart.ChartAreas["MainArea"];

            area.AxisX.Enabled = AxisEnabled.False;
            area.AxisY.Enabled = AxisEnabled.False;
            area.Area3DStyle.Enable3D = false;

            area.Position.Auto = false;
            area.Position.X = 4;
            area.Position.Y = 10;
            area.Position.Width = 62;
            area.Position.Height = 78;

            area.InnerPlotPosition.Auto = false;
            area.InnerPlotPosition.X = 6;
            area.InnerPlotPosition.Y = 6;
            area.InnerPlotPosition.Width = 88;
            area.InnerPlotPosition.Height = 88;

            var s = new Series("Locations");
            s.ChartType = SeriesChartType.Doughnut;
            s.ChartArea = "MainArea";
            s.Legend = "MainLegend";
            s.IsValueShownAsLabel = false;
            s["DoughnutRadius"] = "68";
            s["PieStartAngle"] = "270";
            s.BorderColor = Color.White;
            s.BorderWidth = 2;

            chart.Series.Add(s);
            return chart;
        }

        private Chart CreateStep5AccuracyChart()
        {
            var chart = new Chart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Palette = ChartColorPalette.None,
                BorderlineColor = Color.White,
                BorderlineDashStyle = ChartDashStyle.Solid
            };

            var area = new ChartArea("MainArea");
            area.BackColor = Color.White;
            area.BorderColor = Color.Transparent;

            // ── Eje X: porcentaje 0–100 ───────────────────────────────────────────
            area.AxisX.Minimum = 0;
            area.AxisX.Maximum = 100;
            area.AxisX.Interval = 25;
            area.AxisX.LabelStyle.Format = "0'%'";
            area.AxisX.LabelStyle.Font = new Font("Segoe UI", 8F);
            area.AxisX.LabelStyle.ForeColor = Color.FromArgb(107, 114, 128);
            area.AxisX.LineColor = Color.FromArgb(209, 213, 219);
            area.AxisX.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            area.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            area.AxisX.MajorGrid.Enabled = true;
            area.AxisX.IsMarginVisible = false;

            // ── Eje Y: categorías (nombres de ubicación) ──────────────────────────
            area.AxisY.Interval = 1;
            area.AxisY.LabelStyle.Font = new Font("Segoe UI", 7.5F);
            area.AxisY.LabelStyle.ForeColor = Color.FromArgb(55, 65, 81);
            area.AxisY.LabelStyle.Enabled = true;
            area.AxisY.LineColor = Color.Transparent;
            area.AxisY.MajorGrid.Enabled = false;
            area.AxisY.MajorTickMark.Enabled = false;
            area.AxisY.MinorTickMark.Enabled = false;
            area.AxisY.IsMarginVisible = true;

            // ── Plot area: espacio izquierdo para nombres ─────────────────────────
            area.Position.Auto = false;
            area.Position.X = 0;
            area.Position.Y = 0;
            area.Position.Width = 100;
            area.Position.Height = 88;

            area.InnerPlotPosition.Auto = false;
            area.InnerPlotPosition.X = 25;  // espacio para etiquetas Y
            area.InnerPlotPosition.Y = 2;
            area.InnerPlotPosition.Width = 72;
            area.InnerPlotPosition.Height = 90;

            chart.ChartAreas.Add(area);

            // ── Serie Exact (verde) ───────────────────────────────────────────────
            var sExact = new Series("Exact")
            {
                ChartType = SeriesChartType.StackedBar100,
                ChartArea = "MainArea",
                Color = Color.FromArgb(16, 185, 129),
                IsVisibleInLegend = false,
                IsValueShownAsLabel = false,
                BorderWidth = 0
            };
            sExact["PointWidth"] = "0.9";
            sExact.SmartLabelStyle.Enabled = false;

            // ── Serie Difference (rojo) ───────────────────────────────────────────
            var sDiff = new Series("Difference")
            {
                ChartType = SeriesChartType.StackedBar100,
                ChartArea = "MainArea",
                Color = Color.FromArgb(239, 68, 68),
                IsVisibleInLegend = false,
                IsValueShownAsLabel = false,
                BorderWidth = 0
            };
            sDiff["PointWidth"] = "0.9";
            sDiff.SmartLabelStyle.Enabled = false;

            chart.Series.Add(sExact);
            chart.Series.Add(sDiff);

            // ── Leyenda abajo ─────────────────────────────────────────────────────
            var legend = new Legend("MainLegend");
            legend.Docking = Docking.Bottom;
            legend.Alignment = StringAlignment.Center;
            legend.BackColor = Color.White;
            legend.ForeColor = Color.FromArgb(75, 85, 99);
            legend.Font = new Font("Segoe UI", 8.5F);
            legend.BorderColor = Color.White;
            legend.IsTextAutoFit = false;
            legend.MaximumAutoSize = 12;

            var liExact = new LegendItem
            {
                ImageStyle = LegendImageStyle.Rectangle,
                Color = Color.FromArgb(16, 185, 129),
                BorderColor = Color.Transparent,
                Name = "Exact"
            };
            var liDiff = new LegendItem
            {
                ImageStyle = LegendImageStyle.Rectangle,
                Color = Color.FromArgb(239, 68, 68),
                BorderColor = Color.Transparent,
                Name = "Difference"
            };

            legend.CustomItems.Add(liExact);
            legend.CustomItems.Add(liDiff);
            chart.Legends.Add(legend);

            // ── Tooltip ───────────────────────────────────────────────────────────
            var tt = new ToolTip();
            tt.UseAnimation = false;
            tt.UseFading = false;
            tt.AutomaticDelay = 1;
            tt.InitialDelay = 1;
            tt.AutoPopDelay = 6000;
            string lastTip = "";

            chart.MouseMove += (s, e) =>
            {
                var hit = chart.HitTest(e.X, e.Y);
                string tip = "";
                if (hit?.ChartElementType == ChartElementType.DataPoint
                    && hit.PointIndex >= 0 && hit.Series != null)
                {
                    object tag = hit.Series.Points[hit.PointIndex].Tag;
                    if (tag != null) tip = Convert.ToString(tag);
                }
                if (tip != lastTip) { lastTip = tip; tt.SetToolTip(chart, tip); }
            };
            chart.MouseLeave += (s, e) =>
            {
                if (lastTip != "") { lastTip = ""; tt.SetToolTip(chart, ""); }
            };

            return chart;
        }

        private Chart CreateStep5TopDiscrepanciesChart()
        {
            var chart = CreateNoTooltipChartBase(false); // ← bloquea cuadro blanco nativo
            var area = chart.ChartAreas["MainArea"];

            area.AxisX.LabelStyle.Format = "0";
            area.AxisX.MajorGrid.Enabled = true;

            area.AxisY.MajorGrid.Enabled = false;
            area.AxisY.Interval = 1;
            area.AxisY.LabelStyle.Font = new Font("Consolas", 8F);

            var s = new Series("TopDiff");
            s.ChartType = SeriesChartType.Bar;
            s.ChartArea = "MainArea";
            s.Color = Color.FromArgb(239, 68, 68);
            s.IsValueShownAsLabel = false;
            s["PointWidth"] = "0.68";
            s.BorderWidth = 0;
            s.IsVisibleInLegend = false;
            s.SmartLabelStyle.Enabled = false;

            chart.Series.Add(s);

            // ── Tooltip estilizado sin parpadeo ───────────────────────────────────

            string lastTip = "";

            chart.MouseMove += (sender, e) =>
            {
                var hitTest = chart.HitTest(e.X, e.Y);
                string newTip = "";

                if (hitTest != null
                    && hitTest.ChartElementType == ChartElementType.DataPoint
                    && hitTest.PointIndex >= 0
                    && hitTest.Series != null)
                {
                    var tag = hitTest.Series.Points[hitTest.PointIndex].Tag;
                    newTip = tag != null ? Convert.ToString(tag) : "";
                }

                if (newTip != lastTip)
                {
                    lastTip = newTip;

                }
            };

            chart.MouseLeave += (sender, e) =>
            {
                if (lastTip != "")
                {
                    lastTip = "";

                }
            };

            return chart;
        }

        private class NoTooltipChart : Chart
        {
            private const int WM_NOTIFY = 0x004E;

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_NOTIFY)
                    return; // ← bloquea el tooltip nativo de Windows
                base.WndProc(ref m);
            }
        }
        private Chart CreateNoTooltipChartBase(bool showLegend)
        {
            var chart = new NoTooltipChart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Palette = ChartColorPalette.None,
                BorderlineColor = Color.White,
                BorderlineDashStyle = ChartDashStyle.Solid
            };

            var area = new ChartArea("MainArea");
            area.BackColor = Color.White;
            area.BorderColor = Color.White;

            area.Position.Auto = false;
            area.Position.X = 6;
            area.Position.Y = 8;
            area.Position.Width = showLegend ? 68 : 88;
            area.Position.Height = 82;

            area.InnerPlotPosition.Auto = false;
            area.InnerPlotPosition.X = 10;
            area.InnerPlotPosition.Y = 8;
            area.InnerPlotPosition.Width = 82;
            area.InnerPlotPosition.Height = 82;

            area.AxisX.LineColor = Color.FromArgb(156, 163, 175);
            area.AxisY.LineColor = Color.FromArgb(156, 163, 175);

            area.AxisX.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            area.AxisY.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);

            area.AxisX.LabelStyle.ForeColor = Color.FromArgb(107, 114, 128);
            area.AxisY.LabelStyle.ForeColor = Color.FromArgb(107, 114, 128);

            area.AxisX.LabelStyle.Font = new Font("Segoe UI", 8F);
            area.AxisY.LabelStyle.Font = new Font("Segoe UI", 8F);

            area.AxisX.TitleForeColor = Color.FromArgb(107, 114, 128);
            area.AxisY.TitleForeColor = Color.FromArgb(107, 114, 128);

            area.AxisX.TitleFont = new Font("Segoe UI", 8.5F);
            area.AxisY.TitleFont = new Font("Segoe UI", 8.5F);

            area.AxisX.LabelStyle.Format = "";
            area.AxisY.LabelStyle.Format = "";

            area.AxisX.IsMarginVisible = false;
            area.AxisY.IsMarginVisible = false;

            chart.ChartAreas.Add(area);

            if (showLegend)
            {
                var legend = new Legend("MainLegend");
                legend.Docking = Docking.Right;
                legend.Alignment = StringAlignment.Near;
                legend.BackColor = Color.White;
                legend.ForeColor = Color.FromArgb(75, 85, 99);
                legend.Font = new Font("Segoe UI", 8.5F);
                legend.BorderColor = Color.White;
                legend.IsTextAutoFit = false;
                legend.MaximumAutoSize = 28;
                chart.Legends.Add(legend);
            }

            return chart;
        }


        private Chart CreateStep5BubbleChart()
        {
            var chart = CreateStep5ChartBase(false); // ← sin leyenda
            var area = chart.ChartAreas["MainArea"];

            // ── Eje X — rank de Stock ─────────────────────────────────────────────
            area.AxisX.Minimum = 0;
            area.AxisX.Maximum = 100;
            area.AxisX.Interval = 25;
            area.AxisX.LabelStyle.Format = "0'%'";  // ← 0% 25% 50% 75% 100%
            area.AxisX.LabelStyle.Enabled = true;
            area.AxisX.MajorGrid.Enabled = true;
            area.AxisX.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            area.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            area.AxisX.IsMarginVisible = false;      // ← sin margen extra que genera ticks extra

            // ── Eje Y — rank de Counted ───────────────────────────────────────────
            area.AxisY.Minimum = 0;
            area.AxisY.Maximum = 100;
            area.AxisY.Interval = 25;
            area.AxisY.LabelStyle.Format = "0'%'";  // ← 0% 25% 50% 75% 100%
            area.AxisY.LabelStyle.Enabled = true;
            area.AxisY.MajorGrid.Enabled = true;
            area.AxisY.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            area.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            area.AxisY.IsMarginVisible = false;      // ← sin margen extra

            // ── Títulos ───────────────────────────────────────────────────────────
            area.AxisX.Title = "Stock rank";
            area.AxisY.Title = "Counted rank";
            area.AxisX.TitleFont = new Font("Segoe UI", 8F);
            area.AxisY.TitleFont = new Font("Segoe UI", 8F);
            area.AxisX.TitleForeColor = Color.FromArgb(107, 114, 128);
            area.AxisY.TitleForeColor = Color.FromArgb(107, 114, 128);

            // ── Posición del área de plot ─────────────────────────────────────────
            area.InnerPlotPosition.Auto = false;
            area.InnerPlotPosition.X = 10;
            area.InnerPlotPosition.Y = 6;
            area.InnerPlotPosition.Width = 82;
            area.InnerPlotPosition.Height = 82;

            // ── Serie Bubble ──────────────────────────────────────────────────────
            var s = new Series("Bubble");
            s.ChartType = SeriesChartType.Bubble;
            s.MarkerBorderColor = Color.White;
            s.MarkerBorderWidth = 1;
            s.Color = Color.FromArgb(96, 165, 250);
            s.YValuesPerPoint = 2;
            s.IsValueShownAsLabel = false;          // ← sin números encima de burbujas
            s.SmartLabelStyle.Enabled = false;       // ← desactiva smart labels
            s["BubbleMaxSize"] = "18";
            s["BubbleMinSize"] = "5";            // ← consistente con ApplyVisualPolish
            s.IsVisibleInLegend = false;          // ← sin leyenda

            chart.Series.Add(s);

            // ── Sin leyenda ───────────────────────────────────────────────────────
            chart.Legends.Clear();

            return chart;
        }

        private void AddPiePoint(Series s, string name, double value)
        {
            int idx = s.Points.AddY(value);
            var p = s.Points[idx];
            p.LegendText = name + "  " + value.ToString("0");
            p.Label = "";
        }

        private void AddAccuracyPoint(Series exact, Series diff, string location, double exactValue, double diffValue)
        {
            int i1 = exact.Points.AddXY(location, exactValue);
            exact.Points[i1].AxisLabel = location;

            int i2 = diff.Points.AddXY(location, diffValue);
            diff.Points[i2].AxisLabel = location;
        }

        private void AddBarPoint(Series s, string name, double value)
        {
            int idx = s.Points.AddXY(name, value);
            s.Points[idx].AxisLabel = name;
        }

        private void AddBubblePoint(Series s, string label, double stock, double counted, double diff)
        {
            var p = new DataPoint();
            p.AxisLabel = label;
            p.XValue = stock;
            p.YValues = new double[] { counted, Math.Max(diff, 1) };
            p.ToolTip = label + "\nStock: " + stock.ToString("0.##") +
                        "\nCounted: " + counted.ToString("0.##") +
                        "\nDiff: " + diff.ToString("0.##");
            s.Points.Add(p);
        }

        private Panel CreateStep5ChartCard(string title, string subtitle, out Panel host)
        {
            var border = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(209, 213, 219),
                Padding = new Padding(1),
                Margin = new Padding(0)
            };

            var card = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(16, 14, 16, 16),
                Margin = new Padding(0)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var lblTitle = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                Text = title,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Margin = new Padding(0)
            };

            var lblSubtitle = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                Text = subtitle,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = Color.FromArgb(107, 114, 128),
                Margin = new Padding(0, 6, 0, 0)
            };

            host = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Margin = new Padding(0, 12, 0, 0),
                Padding = new Padding(0)
            };

            layout.Controls.Add(lblTitle, 0, 0);
            layout.Controls.Add(lblSubtitle, 0, 1);
            layout.Controls.Add(host, 0, 2);

            card.Controls.Add(layout);
            border.Controls.Add(card);

            return border;
        }

        private Control CreateStep5ChartPlaceholder(string text)
        {
            var border = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(229, 231, 235),
                Padding = new Padding(1),
                Margin = new Padding(0)
            };

            var inner = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(249, 250, 251)
            };

            var lbl = new Label
            {
                Dock = DockStyle.Fill,
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(107, 114, 128)
            };

            inner.Controls.Add(lbl);
            border.Controls.Add(inner);

            return border;
        }

        private Control CreateStep5ChartsSpacer(int height)
        {
            return new Panel
            {
                Dock = DockStyle.Top,
                Height = height,
                BackColor = Color.Transparent
            };
        }

        private void Step5LocationFilter_Changed(object sender, EventArgs e)
        {
            if (_cboStep5LocationFilter == null) return;

            try
            {
                // ✅ Leer directamente el item seleccionado — no usar SelectedValue
                var item = _cboStep5LocationFilter.SelectedItem as LocationFilterItem;
                if (item == null) return;

                int locId = item.LocationId;

                // ✅ SQL directo a BD — sin RowFilter en memoria
                DataTable dtDetail = LoadStep5DetailForLocation(locId);
                DataTable dtExec = LoadStep5ExecutiveForLocation(locId);

                if (_dgvStep5Detail != null) _dgvStep5Detail.DataSource = dtDetail;
                if (_dgvStep5Executive != null) _dgvStep5Executive.DataSource = dtExec;

                ApplyStep5GridFormatting();
                UpdateStep5KpisFromTable(dtExec);
                RefreshStep5ChartsFromTables(dtDetail, dtExec);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error filtering locations: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BuildStep5TopDiffHoverCard()
        {

        }

        private void ShowStep5TopDiffHover(string title, decimal diffValue, int x, int y)
        {

        }

        private void HideStep5TopDiffHover()
        {

        }

        private void Step5TopDiffChart_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void Step5TopDiffChart_MouseLeave(object sender, EventArgs e)
        {

        }


        private void RenderStep5Treemap()
        {
            if (_pnlStep5TreemapSurface == null) return;
            if (_step5TreemapItems == null || _step5TreemapItems.Count == 0) return;

            _pnlStep5TreemapSurface.SuspendLayout();
            try
            {
                _pnlStep5TreemapSurface.Controls.Clear();

                int totalW = Math.Max(200, _pnlStep5TreemapSurface.ClientSize.Width);
                int totalH = Math.Max(180, _pnlStep5TreemapSurface.ClientSize.Height);

                var items = _step5TreemapItems
                    .Where(x => x.Value > 0)
                    .OrderByDescending(x => x.Value)
                    .ToList();

                double total = items.Sum(x => x.Value);
                if (total <= 0) return;

                int leftW = (int)(totalW * 0.58);
                int rightW = totalW - leftW - 8;

                var leftItems = items.Take(4).ToList();
                var rightItems = items.Skip(4).ToList();

                var leftPanel = new Panel
                {
                    Left = 0,
                    Top = 0,
                    Width = leftW,
                    Height = totalH,
                    BackColor = Color.White
                };

                var rightPanel = new Panel
                {
                    Left = leftW + 8,
                    Top = 0,
                    Width = rightW,
                    Height = totalH,
                    BackColor = Color.White
                };

                _pnlStep5TreemapSurface.Controls.Add(leftPanel);
                _pnlStep5TreemapSurface.Controls.Add(rightPanel);

                RenderVerticalTiles(leftPanel, leftItems);
                RenderGridTiles(rightPanel, rightItems, 4);
            }
            finally
            {
                _pnlStep5TreemapSurface.ResumeLayout();
            }
        }

        private decimal Step5ToDecimal(object value)
        {
            if (value == null || value == DBNull.Value) return 0m;
            decimal d;
            return decimal.TryParse(Convert.ToString(value), out d) ? d : 0m;
        }

        private int Step5ToInt(object value)
        {
            if (value == null || value == DBNull.Value) return 0;
            int n;
            return int.TryParse(Convert.ToString(value), out n) ? n : 0;
        }

        private string Step5ToText(object value, string fallback = "")
        {
            if (value == null || value == DBNull.Value) return fallback;
            var s = Convert.ToString(value);
            return string.IsNullOrWhiteSpace(s) ? fallback : s.Trim();
        }

        private void RefreshStep5ChartsFromTables(DataTable dtDetail, DataTable dtExec)
        {
            BindStep5ItemsByLocationFromReal(dtExec);
            BindStep5AccuracyByLocationFromReal(dtExec);
            BindStep5TopDiscrepanciesFromReal(dtDetail);
            BindStep5StockVsCountFromReal(dtDetail);
            BindStep5StockDistributionFromReal(dtDetail);
            ApplyStep5ChartsVisualPolish();
        }

        // =====================================================================
        // REPLACE these three methods in your existing file
        // =====================================================================

        private void ApplyStep5ChartsVisualPolish()
        {
            //// ── Items by Location ────────────────────────────────────────────────
            //if (_chtStep5ItemsByLocation != null)
            //{
            //    _chtStep5ItemsByLocation.Palette = ChartColorPalette.BrightPastel;
            //}

            //// ── REPLACE the foreach inside the "Accuracy by Location" block ─────────────
            //foreach (var serie in _chtStep5AccuracyByLocation.Series)
            //{
            //    serie.ChartType = SeriesChartType.StackedBar100;
            //    serie["PointWidth"] = "0.9";
            //    serie.BorderWidth = 0;

            //    if (serie.Name == "Difference")
            //    {
            //        serie.IsValueShownAsLabel = true;
            //        serie.LabelFormat = "0.#'%'";      // renders "4.2%"
            //        serie.SmartLabelStyle.Enabled = false;         // use our manual labels
            //    }
            //    else
            //    {
            //        serie.IsValueShownAsLabel = false;
            //        serie.SmartLabelStyle.Enabled = false;
            //    }
            //}

            //// ── Top Discrepancies ─────────────────────────────────────────────────
            //if (_chtStep5TopDiscrepancies != null)
            //{
            //    var area = _chtStep5TopDiscrepancies.ChartAreas["MainArea"];

            //    // ── Eje X — porcentaje automático ─────────────────────────────────
            //    area.AxisX.Minimum = 0;
            //    area.AxisX.Maximum = double.NaN;
            //    area.AxisX.Interval = double.NaN;
            //    area.AxisX.LabelStyle.Format = "0.#'%'";
            //    area.AxisX.LabelStyle.Enabled = true;
            //    area.AxisX.MajorGrid.Enabled = true;
            //    area.AxisX.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            //    area.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            //    area.AxisX.IsMarginVisible = false;

            //    // ── Eje Y — completamente deshabilitado ───────────────────────────
            //    area.AxisY.LabelStyle.Enabled = false;
            //    area.AxisY.LabelStyle.Format = "";
            //    area.AxisY.MajorGrid.Enabled = false;
            //    area.AxisY.MajorTickMark.Enabled = false;
            //    area.AxisY.MinorTickMark.Enabled = false;
            //    area.AxisY.IsMarginVisible = false;
            //    area.AxisY.LineColor = Color.Transparent;
            //    area.AxisY.Enabled = AxisEnabled.False; // ← kill switch

            //    // ── Área de plot — más ancha porque no hay eje Y ──────────────────
            //    area.InnerPlotPosition.Auto = false;
            //    area.InnerPlotPosition.X = 4;
            //    area.InnerPlotPosition.Y = 4;
            //    area.InnerPlotPosition.Width = 92;
            //    area.InnerPlotPosition.Height = 82;

            //    foreach (var serie in _chtStep5TopDiscrepancies.Series)
            //    {
            //        serie["PointWidth"] = "0.65";
            //        serie.IsValueShownAsLabel = false;
            //        serie.SmartLabelStyle.Enabled = false;
            //        serie.BorderWidth = 0;
            //    }

            //    area.RecalculateAxesScale();
            //}

            //// ── Stock vs Count (Percentil Rank) ───────────────────────────────────
            //if (_chtStep5StockVsCount != null)
            //{
            //    var area = _chtStep5StockVsCount.ChartAreas["MainArea"];

            //    // ── Eje X — rank de Stock (0% a 100%) ────────────────────────────
            //    area.AxisX.Minimum = 0;
            //    area.AxisX.Maximum = 100;
            //    area.AxisX.Interval = 25;
            //    area.AxisX.LabelStyle.Format = "0'%'";
            //    area.AxisX.LabelStyle.Enabled = true;   // ← muestra 0% 25% 50% 75% 100%
            //    area.AxisX.MajorGrid.Enabled = true;
            //    area.AxisX.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            //    area.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            //    // ── Eje Y — rank de Counted (0% a 100%) ──────────────────────────
            //    area.AxisY.Minimum = 0;
            //    area.AxisY.Maximum = 100;
            //    area.AxisY.Interval = 25;
            //    area.AxisY.LabelStyle.Format = "0'%'";
            //    area.AxisY.LabelStyle.Enabled = true;   // ← muestra 0% 25% 50% 75% 100%
            //    area.AxisY.MajorGrid.Enabled = true;
            //    area.AxisY.MajorGrid.LineColor = Color.FromArgb(229, 231, 235);
            //    area.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            //    // ── Títulos de ejes ───────────────────────────────────────────────
            //    area.AxisX.Title = "Stock rank";
            //    area.AxisY.Title = "Counted rank";

            //    area.AxisX.TitleFont = new Font("Segoe UI", 8F);
            //    area.AxisY.TitleFont = new Font("Segoe UI", 8F);
            //    area.AxisX.TitleForeColor = Color.FromArgb(107, 114, 128);
            //    area.AxisY.TitleForeColor = Color.FromArgb(107, 114, 128);

            //    // ── Sin etiquetas de puntos ni leyenda ────────────────────────────
            //    foreach (var serie in _chtStep5StockVsCount.Series)
            //    {
            //        serie.IsValueShownAsLabel = false;  // ← sin números encima de burbujas
            //        serie.SmartLabelStyle.Enabled = false; // ← desactiva smart labels
            //        serie["BubbleMaxSize"] = "18";
            //        serie["BubbleMinSize"] = "5";
            //    }

            //    // ── Sin leyenda en este chart ─────────────────────────────────────
            //    _chtStep5StockVsCount.Legends.Clear();

            //    area.InnerPlotPosition.Auto = false;
            //    area.InnerPlotPosition.X = 10;
            //    area.InnerPlotPosition.Y = 6;
            //    area.InnerPlotPosition.Width = 82;
            //    area.InnerPlotPosition.Height = 82;

            //    area.RecalculateAxesScale();
            //}
        }

        // ── Accuracy: convert to complementary ──────

        private void BindStep5AccuracyByLocationFromReal(DataTable dtExec)
        {
            if (_chtStep5AccuracyByLocation == null) return;

            var sExact = _chtStep5AccuracyByLocation.Series["Exact"];
            var sDiff = _chtStep5AccuracyByLocation.Series["Difference"];

            sExact.Points.Clear();
            sDiff.Points.Clear();

            if (dtExec == null || dtExec.Rows.Count == 0) return;

            // ── Top 10 por mayor discrepancia ────────────────────────────────────────
            var top10 = dtExec.AsEnumerable()
                .GroupBy(r => Step5ToText(r["Location"], "EMPTY"))
                .Select(g => new
                {
                    LocationFull = g.Key,
                    TotalLines = g.Sum(r => Step5ToDecimal(r["Lines"])),
                    LinesWithDiff = g.Sum(r => Step5ToDecimal(r["Lines w/ Diff"]))
                })
                .Where(x => x.TotalLines > 0)
                .OrderByDescending(x => x.LinesWithDiff / x.TotalLines)
                .Take(10)
                .ToList();

            if (top10.Count == 0) return;

            // ── Configurar área — sin ejes, sin grids, sin etiquetas ─────────────────
            var area = _chtStep5AccuracyByLocation.ChartAreas["MainArea"];

            // Eje X (valores 0–100%): completamente oculto
            area.AxisX.LabelStyle.Enabled = false;
            area.AxisX.MajorGrid.Enabled = false;
            area.AxisX.MinorGrid.Enabled = false;
            area.AxisX.MajorTickMark.Enabled = false;
            area.AxisX.MinorTickMark.Enabled = false;
            area.AxisX.LineColor = Color.Transparent;
            area.AxisX.IsMarginVisible = false;

            // Eje Y (nombres de ubicación): solo las etiquetas, sin grid
            area.AxisY.Interval = 1;
            area.AxisY.LabelStyle.Enabled = true;
            area.AxisY.LabelStyle.Angle = 0;
            area.AxisY.LabelStyle.Font = new Font("Segoe UI", 8.5F);
            area.AxisY.LabelStyle.ForeColor = Color.FromArgb(31, 41, 55);
            area.AxisY.LabelStyle.IsEndLabelVisible = true;
            area.AxisY.MajorGrid.Enabled = false;
            area.AxisY.MinorGrid.Enabled = false;
            area.AxisY.MajorTickMark.Enabled = false;
            area.AxisY.MinorTickMark.Enabled = false;
            area.AxisY.LineColor = Color.Transparent;
            area.AxisY.IsMarginVisible = true;

            // Plot area — espacio generoso a la izquierda para los nombres
            area.InnerPlotPosition.Auto = false;
            area.InnerPlotPosition.X = 30;
            area.InnerPlotPosition.Y = 2;
            area.InnerPlotPosition.Width = 68;
            area.InnerPlotPosition.Height = 94;

            // ── Configurar ambas series — barras gruesas, SIN etiquetas ──────────────
            foreach (var serie in _chtStep5AccuracyByLocation.Series)
            {
                serie.ChartType = SeriesChartType.StackedBar100;
                serie["PointWidth"] = "0.85";   // barras gruesas
                serie.BorderWidth = 0;
                serie.IsValueShownAsLabel = false;    // sin porcentajes
                serie.LabelFormat = "";
                serie.Label = "";
                serie.SmartLabelStyle.Enabled = false;
            }

            // ── Agregar puntos ────────────────────────────────────────────────────────
            foreach (var loc in top10)
            {
                string locationLabel = TruncateLabel(loc.LocationFull, 20);

                double pctExact = Math.Round(
                    (double)((loc.TotalLines - loc.LinesWithDiff) / loc.TotalLines * 100m), 1);
                double pctDiff = Math.Round(100.0 - pctExact, 1);

                // Mínimo visual para que la franja roja siempre sea visible
                if (pctDiff > 0 && pctDiff < 3.0)
                    pctDiff = 3.0;

                // ── Barra verde (Exact) ───────────────────────────────────────────────
                int iExact = sExact.Points.AddXY(locationLabel, pctExact);
                sExact.Points[iExact].Color = Color.FromArgb(16, 185, 129);
                sExact.Points[iExact].IsValueShownAsLabel = false;
                sExact.Points[iExact].Label = "";
                sExact.Points[iExact].ToolTip = "";

                // ── Barra roja (Difference) ───────────────────────────────────────────
                int iDiff = sDiff.Points.AddXY(locationLabel, pctDiff);
                sDiff.Points[iDiff].Color = Color.FromArgb(239, 68, 68);
                sDiff.Points[iDiff].IsValueShownAsLabel = false;
                sDiff.Points[iDiff].Label = "";    // sin porcentaje
                sDiff.Points[iDiff].ToolTip = "";
            }

            // ── Redimensionar panel para barras con altura mínima de 36px ────────────
            ResizeChartForRowCount(
                _chtStep5AccuracyByLocation,
                _pnlStep5ChartAccuracyByLocation,
                top10.Count,
                minRowPx: 36);
        }


        // ═══════════════════════════════════════════════════════════════════════════
        // FIX 2 — Top Discrepancies: agrupa por Part Number, Top 5, sin etiquetas
        // ═══════════════════════════════════════════════════════════════════════════
        private void BindStep5TopDiscrepanciesFromReal(DataTable dtDetail)
        {
            if (_chtStep5TopDiscrepancies == null) return;

            var s = _chtStep5TopDiscrepancies.Series["TopDiff"];
            s.Points.Clear();
            s.SmartLabelStyle.Enabled = false;

            if (dtDetail == null || dtDetail.Rows.Count == 0) return;

            // ── PASO 1: Agrupar 10,000 productos por Part Number (LINQ GroupBy) ───────
            var grouped = dtDetail.AsEnumerable()
                .GroupBy(r => Step5ToText(r["Part Number"], "N/A"))
                .Select(g => new
                {
                    PartNumberFull = g.Key,
                    TotalDiff = g.Sum(x => Step5ToDecimal(x["Difference"])),
                    TotalStock = g.Sum(x => Step5ToDecimal(x["SystemStock"]))
                })
                .Where(x => x.TotalDiff > 0)
                .ToList();

            // ── PASO 2: Top 5 productos con más unidades de diferencia ────────────────
            var top5 = grouped
                .OrderByDescending(x => x.TotalDiff)
                .Take(5)
                .ToList();

            if (top5.Count == 0) return;

            // ── PASO 3: Configurar ejes ────────────────────────────────────────────────
            var area = _chtStep5TopDiscrepancies.ChartAreas["MainArea"];

            // Quitar grid del eje X
            area.AxisX.MajorGrid.Enabled = false;
            area.AxisX.MajorGrid.LineColor = Color.Transparent;
            area.AxisX.LabelStyle.Enabled = true;
            area.AxisX.LabelStyle.Format = "N0";

            // Deshabilitar eje Y completamente (nombres van en AxisLabel de cada punto)
            area.AxisY.LabelStyle.Enabled = false;
            area.AxisY.MajorGrid.Enabled = false;
            area.AxisY.MajorTickMark.Enabled = false;
            area.AxisY.LineColor = Color.Transparent;
            area.AxisY.Enabled = AxisEnabled.False;

            // PointWidth fijo 0.6 para 5 barras
            s["PointWidth"] = "0.60";
            s.IsValueShownAsLabel = false;  // sin número encima, usamos Tag para tooltip

            decimal totalDiff = top5.Sum(x => x.TotalDiff);

            // ── PASO 4: Agregar barras ─────────────────────────────────────────────────
            foreach (var item in top5)
            {
                // % del total de discrepancias (share)
                double pct = totalDiff > 0
                    ? Math.Round((double)(item.TotalDiff / totalDiff * 100m), 1)
                    : 0;

                string tipText =
                    item.PartNumberFull +
                    "\nShare of total diff: " + pct.ToString("0.#") + "%" +
                    "\nAbsolute difference: " + item.TotalDiff.ToString("N0") + " units" +
                    "\nSystem stock:        " + item.TotalStock.ToString("N0") + " units";

                var pt = new DataPoint();
                pt.SetValueXY("", pct);             // X vacío — la barra usa AxisLabel
                pt.AxisLabel = TruncateLabel(item.PartNumberFull, 12);
                pt.IsValueShownAsLabel = false;
                pt.ToolTip = tipText;
                pt.Tag = tipText;
                pt.Color = Color.FromArgb(239, 68, 68);

                s.Points.Add(pt);
            }
        }

        // ── ADD a helper method to the class ────────────────────────────────────────
        /// <summary>
        /// Resizes a Chart control's parent panel so each bar row gets minRowPx pixels.
        /// Call this after binding data, before Refresh().
        /// </summary>
        private void ResizeChartForRowCount(Chart chart, Panel host, int rowCount, int minRowPx = 28)
        {
            if (chart == null || host == null || rowCount <= 0) return;

            // header + axis title + legend ≈ 80 px overhead
            int desiredH = Math.Max(host.Height, rowCount * minRowPx + 80);
            if (host.Height < desiredH)
                host.Height = desiredH;    // host panel grows → parent scroll takes over
        }


        // ── Helper: truncate long names so Y-axis labels don't overflow ───────
        private static string TruncateLabel(string text, int maxChars)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxChars)
                return text;
            return text.Substring(0, maxChars - 1) + "…";
        }

        private void BindStep5ItemsByLocationFromReal(DataTable dtExec)
        {
            if (_chtStep5ItemsByLocation == null) return;

            var s = _chtStep5ItemsByLocation.Series["Locations"];
            s.Points.Clear();
            _chtStep5ItemsByLocation.Palette = ChartColorPalette.BrightPastel;

            if (dtExec == null || dtExec.Rows.Count == 0) return;

            // ✅ Calcular total de líneas para obtener porcentaje
            decimal totalLines = 0m;
            foreach (DataRow r in dtExec.Rows)
                totalLines += Step5ToDecimal(r["Lines"]);

            if (totalLines <= 0) return;

            foreach (DataRow r in dtExec.Rows)
            {
                string location = Step5ToText(r["Location"], "EMPTY");
                decimal lines = Step5ToDecimal(r["Lines"]);
                if (lines <= 0) continue;

                // ✅ Convertir a porcentaje redondeado a 1 decimal
                double pct = Math.Round((double)(lines / totalLines * 100m), 1);

                int idx = s.Points.AddY(pct);
                // ✅ Mostrar % en la leyenda
                s.Points[idx].LegendText = $"{location}  {pct:0.#}%";
                s.Points[idx].Label = "";
                // ✅ Tooltip con detalle
                s.Points[idx].ToolTip = $"{location}\n{lines:0} líneas ({pct:0.#}%)";
            }
        }


        private void BindStep5StockVsCountFromReal(DataTable dtDetail)
        {
            if (_chtStep5StockVsCount == null) return;

            var s = _chtStep5StockVsCount.Series["Bubble"];
            s.Points.Clear();

            if (dtDetail == null || dtDetail.Rows.Count == 0) return;

            // ── Tomar Top 20 por stock ───────────────────────────────────────────
            var rows = dtDetail.AsEnumerable()
                .Select(r => new
                {
                    PartNumberFull = Step5ToText(r["Part Number"], "N/A"),
                    Stock = Step5ToDecimal(r["SystemStock"]),
                    Counted = Step5ToDecimal(r["CountedQty"]),
                    Diff = Step5ToDecimal(r["Difference"])
                })
                .Where(x => x.Stock > 0 || x.Counted > 0)
                .OrderByDescending(x => x.Stock)
                .Take(20)
                .ToList();

            if (rows.Count == 0) return;

            int total = rows.Count;

            // ── Ordenar por Stock para calcular rank de Stock ────────────────────
            var rankedByStock = rows
                .OrderBy(x => x.Stock)
                .ToList();

            // ── Ordenar por Counted para calcular rank de Counted ────────────────
            var rankedByCounted = rows
                .OrderBy(x => x.Counted)
                .ToList();

            foreach (var item in rows)
            {
                // ── Rank de Stock (posición relativa 5% → 100%) ──────────────────
                int stockRankIndex = rankedByStock.FindIndex(x => x.PartNumberFull == item.PartNumberFull);
                int countedRankIndex = rankedByCounted.FindIndex(x => x.PartNumberFull == item.PartNumberFull);

                // Convertir rank a porcentaje — mínimo 5% para que no quede en el borde
                double pctStock = total == 1 ? 50.0
                    : Math.Round(5.0 + (stockRankIndex / (double)(total - 1)) * 95.0, 1);

                double pctCounted = total == 1 ? 50.0
                    : Math.Round(5.0 + (countedRankIndex / (double)(total - 1)) * 95.0, 1);

                // ── Tamaño de burbuja según diferencia relativa ──────────────────
                // Si no hay diferencia → burbuja mínima (5)
                // Si hay diferencia → proporcional al máximo
                decimal maxDiff = rows.Max(x => x.Diff);
                double bubbleSize = maxDiff > 0
                    ? Math.Round(5.0 + (double)(item.Diff / maxDiff) * 15.0, 1)
                    : 5.0;

                // ── Color: rojo si tiene diferencia, azul si está exacto ─────────
                var point = new DataPoint();
                point.XValue = pctStock;
                point.YValues = new double[] { pctCounted, bubbleSize };
                point.AxisLabel = "";  // ← sin etiqueta en el punto

                // Color por diferencia
                point.Color = item.Diff > 0
                    ? Color.FromArgb(160, 239, 68, 68)    // rojo semitransparente
                    : Color.FromArgb(160, 96, 165, 250);  // azul semitransparente

                point.MarkerBorderColor = item.Diff > 0
                    ? Color.FromArgb(220, 38, 38)
                    : Color.FromArgb(37, 99, 235);

                // ── Tooltip completo al hover ─────────────────────────────────────
                point.ToolTip =
                    item.PartNumberFull +
                    "\nSystem Stock:  " + item.Stock.ToString("N0") +
                    "  (rank " + (stockRankIndex + 1).ToString() + "/" + total + ")" +
                    "\nCounted Qty:   " + item.Counted.ToString("N0") +
                    "  (rank " + (countedRankIndex + 1).ToString() + "/" + total + ")" +
                    "\nDifference:    " + item.Diff.ToString("N0") + " units" +
                    (item.Diff == 0 ? "  ✓ Exact" : "  ✗ Discrepancy");

                s.Points.Add(point);
            }
        }

        private void BindStep5StockDistributionFromReal(DataTable dtDetail)
        {
            _step5TreemapItems = new List<Step5TreemapItem>();

            if (dtDetail == null || dtDetail.Rows.Count == 0)
            {
                RenderStep5Treemap();
                return;
            }

            var items =
                dtDetail.AsEnumerable()
                .Select(r => new Step5TreemapItem
                {
                    Label = Step5ToText(r["Part Number"], "N/A") + "\n" +
                            Step5ToDecimal(r["SystemStock"]).ToString("0"),
                    Value = (double)Step5ToDecimal(r["SystemStock"]),
                    HasDifference = Step5ToDecimal(r["Difference"]) > 0
                })
                .Where(x => x.Value > 0)
                .OrderByDescending(x => x.Value)
                .Take(20)
                .ToList();

            _step5TreemapItems = items;
            RenderStep5Treemap();
        }


        private void RenderVerticalTiles(Panel host, List<Step5TreemapItem> items)
        {
            host.Controls.Clear();
            if (items == null || items.Count == 0) return;

            double total = items.Sum(x => x.Value);
            if (total <= 0) return;

            int y = 0;
            int gap = 6;

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];

                int h;
                if (i == items.Count - 1)
                    h = host.Height - y;
                else
                    h = Math.Max(40, (int)Math.Round((item.Value / total) * (host.Height - gap * (items.Count - 1))));

                var tile = CreateTreemapTile(item);
                tile.Left = 0;
                tile.Top = y;
                tile.Width = host.Width;
                tile.Height = Math.Max(40, h);

                host.Controls.Add(tile);
                y += tile.Height + gap;
            }
        }

        private void RenderGridTiles(Panel host, List<Step5TreemapItem> items, int columns)
        {
            host.Controls.Clear();
            if (items == null || items.Count == 0) return;

            int gap = 6;
            columns = Math.Max(2, columns);

            int rows = (int)Math.Ceiling(items.Count / (double)columns);
            int cellW = Math.Max(40, (host.Width - gap * (columns - 1)) / columns);
            int cellH = Math.Max(36, (host.Height - gap * (rows - 1)) / Math.Max(1, rows));

            for (int i = 0; i < items.Count; i++)
            {
                int row = i / columns;
                int col = i % columns;

                var tile = CreateTreemapTile(items[i]);
                tile.Left = col * (cellW + gap);
                tile.Top = row * (cellH + gap);
                tile.Width = cellW;
                tile.Height = cellH;

                host.Controls.Add(tile);
            }
        }

        private Control CreateTreemapTile(Step5TreemapItem item)
        {
            Color backColor;

            if (item.HasDifference)
                backColor = Color.FromArgb(209, 83, 100);   // red-ish
            else
                backColor = Color.FromArgb(99, 124, 163);   // blue-ish

            if (!item.HasDifference && item.Value <= 150)
                backColor = Color.FromArgb(67, 201, 179);   // teal-ish

            if (!item.HasDifference && item.Value >= 250)
                backColor = Color.FromArgb(216, 161, 58);   // amber-ish

            var tile = new Panel
            {
                BackColor = backColor,
                Padding = new Padding(8),
                Margin = new Padding(0)
            };

            var lbl = new Label
            {
                Dock = DockStyle.Fill,
                Text = item.Label,
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Font = new Font("Consolas", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.TopLeft
            };

            tile.Controls.Add(lbl);

            var tip = new ToolTip();
            tip.SetToolTip(tile, item.Label.Replace("\n", " | "));
            tip.SetToolTip(lbl, item.Label.Replace("\n", " | "));

            return tile;
        }



        /// <summary>Carga el encabezado del Paso 5: ID, nombre, fechas y responsable.</summary>
        private void LoadStep5HeaderInfo()
        {
            if (_physicalCountId <= 0) return;
            if (_lblStep5HeaderTitle == null || _lblStep5HeaderLine2 == null) return;

            try
            {
                var dt = _repo.LoadReviewHeader(_physicalCountId);

                if (dt.Rows.Count == 0)
                {
                    _lblStep5HeaderTitle.Text = "ID de inventario: " + _physicalCountId;
                    _lblStep5HeaderLine2.Text = "";
                    return;
                }

                var row = dt.Rows[0];

                int      id        = Convert.ToInt32(row["Physical_Count_ID"]);
                string   nombre    = Convert.ToString(row["Physical_Count_Name"] ?? "");
                DateTime inicio    = Convert.ToDateTime(row["Start_Date"]);
                DateTime fin       = Convert.ToDateTime(row["End_Date"]);
                string   empNo     = Convert.ToString(row["Employee_Number"] ?? "");
                string   respNombre= Convert.ToString(row["Name"] ?? "");
                string   puesto    = Convert.ToString(row["Job_Position"] ?? "");

                _lblStep5HeaderTitle.Text =
                    "ID del conteo: " + id + "   Nombre: " + nombre;

                _lblStep5HeaderLine2.Text =
                    "Inicio: " + inicio.ToString("MM/dd/yyyy") +
                    "   Fin: " + fin.ToString("MM/dd/yyyy") +
                    "   Responsable: " + empNo + " - " + respNombre + " (" + puesto + ")";
            }
            catch (Exception ex)
            {
                _lblStep5HeaderTitle.Text = "ID de inventario: " + _physicalCountId;
                _lblStep5HeaderLine2.Text = "Error al cargar encabezado: " + ex.Message;
            }
        }


        private int GetStep5SelectedLocationId()
        {
            if (_cboStep5LocationFilter == null) return -1;

            // ✅ Leer directamente el item — no usar SelectedValue
            var item = _cboStep5LocationFilter.SelectedItem as LocationFilterItem;
            if (item == null) return -1;

            return item.LocationId;
        }

        /// <summary>
        /// Initializes print documents for Step 5 Detail and Executive tabs.
        /// Called once after EnsureStep5UI builds the layout.
        /// </summary>
        private void InitStep5PrintDocuments()
        {
            // Wire Detail print button
            if (_btnStep5PreviewDetail != null)
            {
                _btnStep5PreviewDetail.Click -= Step5PreviewDetail_Click;
                _btnStep5PreviewDetail.Click += Step5PreviewDetail_Click;
            }

            if (_btnStep5ExportPdfDetail != null)
            {
                _btnStep5ExportPdfDetail.Click -= Step5ExportPdfDetail_Click;
                _btnStep5ExportPdfDetail.Click += Step5ExportPdfDetail_Click;
            }

            // Wire Executive print button
            if (_btnStep5PreviewExecutive != null)
            {
                _btnStep5PreviewExecutive.Click -= Step5PreviewExecutive_Click;
                _btnStep5PreviewExecutive.Click += Step5PreviewExecutive_Click;
            }

            if (_btnStep5ExportPdfExecutive != null)
            {
                _btnStep5ExportPdfExecutive.Click -= Step5ExportPdfExecutive_Click;
                _btnStep5ExportPdfExecutive.Click += Step5ExportPdfExecutive_Click;
            }
        }
        /// <summary>
        /// Updates the KPI bar using the given executive DataTable.
        /// Works for both full data and filtered subsets.
        /// Diff % is capped at 100%.
        /// </summary>
        private void UpdateStep5KpisFromTable(DataTable dtExec)
        {
            if (_lblStep5Kpis == null) return;

            if (dtExec == null || dtExec.Rows.Count == 0)
            {
                _lblStep5Kpis.Text = "No data for selected location.";
                return;
            }

            decimal totalAbsDiff = 0m;
            decimal totalBase = 0m;
            int linesCounted = 0;
            int linesTotal = 0;

            foreach (DataRow r in dtExec.Rows)
            {
                if (dtExec.Columns.Contains("Abs Diff"))
                    totalAbsDiff += Convert.ToDecimal(r["Abs Diff"]);

                if (dtExec.Columns.Contains("Lines Counted"))
                    linesCounted += Convert.ToInt32(r["Lines Counted"]);

                if (dtExec.Columns.Contains("Lines"))
                    linesTotal += Convert.ToInt32(r["Lines"]);

                // Recover base for weighted % calculation
                if (dtExec.Columns.Contains("Abs Diff") &&
                    dtExec.Columns.Contains("Diff %"))
                {
                    decimal absDiff = Convert.ToDecimal(r["Abs Diff"]);
                    decimal pct = Convert.ToDecimal(r["Diff %"]);

                    if (pct > 0 && pct < 100)
                        totalBase += absDiff * 100m / pct;
                    else if (pct >= 100)
                        totalBase += absDiff;
                }
            }

            // Overall Diff % capped at 100%
            decimal overallPct = 0m;
            if (linesCounted > 0 && totalBase > 0)
                overallPct = totalAbsDiff * 100m / totalBase;

            if (overallPct > 100m) overallPct = 100m;
            if (linesCounted == 0) { totalAbsDiff = 0m; overallPct = 0m; }

            _lblStep5Kpis.Text =
                $"Locations: {dtExec.Rows.Count}" +
                $"  |  Lines counted: {linesCounted} / {linesTotal}" +
                $"  |  Total diff (abs): {totalAbsDiff:0.##}" +
                $"  |  Diff %: {overallPct:0.##}%";
        }

        /// <summary>
        /// Backward compatibility — uses full executive dataset.
        /// </summary>
        private void UpdateStep5Kpis()
        {
            UpdateStep5KpisFromTable(_dtStep5AllExecutive ?? _dtStep5Executive);
        }

        private void BuildStep5DetailTab()
        {
            _tabStep5Detail.Controls.Clear();

            var root = CreateTwoRowRoot(56);

            var bar = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };

            _btnStep5PreviewDetail = new Button { Text = "Preview / Print", Width = 150, Height = 36, Top = 10 };
            _btnStep5ExportDetail = new Button { Text = "Export", Width = 115, Height = 36, Top = 10 };

            StyleButtonSecondary(_btnStep5PreviewDetail);
            StyleButtonGhost(_btnStep5ExportDetail);

            bar.Controls.Add(_btnStep5PreviewDetail);
            bar.Controls.Add(_btnStep5ExportDetail);

            bar.Resize += (s, e) =>
            {
                int gap = 10;
                _btnStep5ExportDetail.Left = bar.Width - _btnStep5ExportDetail.Width - 10;
                _btnStep5PreviewDetail.Left = _btnStep5ExportDetail.Left - _btnStep5PreviewDetail.Width - gap;
            };

            _dgvStep5Detail = CreateStep5Grid();

            root.Controls.Add(bar, 0, 0);
            root.Controls.Add(_dgvStep5Detail, 0, 1);
            _tabStep5Detail.Controls.Add(root);

            // Events
            _btnStep5PreviewDetail.Click -= Step5PreviewDetail_Click;
            _btnStep5PreviewDetail.Click += Step5PreviewDetail_Click;

            BuildStep5DetailExportMenu();

            _btnStep5ExportDetail.Click -= Step5ExportDetail_Click;
            _btnStep5ExportDetail.Click += Step5ExportDetail_Click;
        }

        private void Step5ExportDetail_Click(object sender, EventArgs e)
        {
            if (_btnStep5ExportDetail == null || _cmsStep5ExportDetail == null) return;

            _cmsStep5ExportDetail.Show(
                _btnStep5ExportDetail,
                0,
                _btnStep5ExportDetail.Height);
        }

        private void Step5ExportExecutive_Click(object sender, EventArgs e)
        {
            if (_btnStep5ExportExecutive == null || _cmsStep5ExportExecutive == null) return;

            _cmsStep5ExportExecutive.Show(
                _btnStep5ExportExecutive,
                0,
                _btnStep5ExportExecutive.Height);
        }


        private void BuildStep5ExecutiveTab()
        {
            _tabStep5Executive.Controls.Clear();

            var root = CreateTwoRowRoot(56);

            var bar = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };

            _btnStep5PreviewExecutive = new Button { Text = "Preview / Print", Width = 150, Height = 36, Top = 10 };
            _btnStep5ExportExecutive = new Button { Text = "Export", Width = 115, Height = 36, Top = 10 };

            StyleButtonSecondary(_btnStep5PreviewExecutive);
            StyleButtonGhost(_btnStep5ExportExecutive);

            bar.Controls.Add(_btnStep5PreviewExecutive);
            bar.Controls.Add(_btnStep5ExportExecutive);

            bar.Resize += (s, e) =>
            {
                int gap = 10;
                _btnStep5ExportExecutive.Left = bar.Width - _btnStep5ExportExecutive.Width - 10;
                _btnStep5PreviewExecutive.Left = _btnStep5ExportExecutive.Left - _btnStep5PreviewExecutive.Width - gap;
            };

            _dgvStep5Executive = CreateStep5Grid();

            root.Controls.Add(bar, 0, 0);
            root.Controls.Add(_dgvStep5Executive, 0, 1);
            _tabStep5Executive.Controls.Add(root);

            // Events
            _btnStep5PreviewExecutive.Click -= Step5PreviewExecutive_Click;
            _btnStep5PreviewExecutive.Click += Step5PreviewExecutive_Click;

            BuildStep5ExecutiveExportMenu();

            _btnStep5ExportExecutive.Click -= Step5ExportExecutive_Click;
            _btnStep5ExportExecutive.Click += Step5ExportExecutive_Click;
        }


        private void BuildStep5DetailExportMenu()
        {
            if (_cmsStep5ExportDetail != null) return;

            _cmsStep5ExportDetail = new ContextMenuStrip();

            _miStep5ExportPdfDetail = new ToolStripMenuItem("Export PDF");
            _miStep5ExportExcelDetail = new ToolStripMenuItem("Export Excel");
            _miStep5ExportCsvDetail = new ToolStripMenuItem("Export CSV");

            _miStep5ExportPdfDetail.Click += Step5ExportPdfDetail_Click;
            _miStep5ExportExcelDetail.Click += Step5ExportExcelDetail_Click;
            _miStep5ExportCsvDetail.Click += Step5ExportCsvDetail_Click;

            _cmsStep5ExportDetail.Items.Add(_miStep5ExportPdfDetail);
            _cmsStep5ExportDetail.Items.Add(_miStep5ExportExcelDetail);
            _cmsStep5ExportDetail.Items.Add(_miStep5ExportCsvDetail);
        }

        private void BuildStep5ExecutiveExportMenu()
        {
            if (_cmsStep5ExportExecutive != null) return;

            _cmsStep5ExportExecutive = new ContextMenuStrip();

            _miStep5ExportPdfExecutive = new ToolStripMenuItem("Export PDF");
            _miStep5ExportExcelExecutive = new ToolStripMenuItem("Export Excel");
            _miStep5ExportCsvExecutive = new ToolStripMenuItem("Export CSV");

            _miStep5ExportPdfExecutive.Click += Step5ExportPdfExecutive_Click;
            _miStep5ExportExcelExecutive.Click += Step5ExportExcelExecutive_Click;
            _miStep5ExportCsvExecutive.Click += Step5ExportCsvExecutive_Click;

            _cmsStep5ExportExecutive.Items.Add(_miStep5ExportPdfExecutive);
            _cmsStep5ExportExecutive.Items.Add(_miStep5ExportExcelExecutive);
            _cmsStep5ExportExecutive.Items.Add(_miStep5ExportCsvExecutive);
        }



        private void Step5ExportExcelDetail_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentDetailData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No detail data to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExportDataTableToExcel(
                dt,
                "Detail",
                $"Detail_Differences_{_physicalCountId}.xlsx");
        }


        private void Step5ExportCsvExecutive_Click(object sender, EventArgs e)
        {
            MessageBox.Show("CSV export is not implemented yet.", "Step 5",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Step5ExportExcelExecutive_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentExecutiveData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No executive data to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExportDataTableToExcel(
                dt,
                "Executive Summary",
                $"Executive_Summary_{_physicalCountId}.xlsx");
        }

        private void ReleaseComObjectSafe(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        private void ExportDataTableToExcel(DataTable dt, string sheetName, string defaultFileName)
        {
            if (dt == null || dt.Columns.Count == 0)
            {
                MessageBox.Show("No data available to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = defaultFileName;
                sfd.Title = "Export Excel";

                if (sfd.ShowDialog(this) != DialogResult.OK)
                    return;

                Excel.Application xlApp = null;
                Excel.Workbook xlBook = null;
                Excel.Worksheet xlSheet = null;
                Excel.Range headerRange = null;
                Excel.Range usedRange = null;

                try
                {
                    xlApp = new Excel.Application();
                    xlApp.DisplayAlerts = false;
                    xlApp.Visible = false;

                    xlBook = xlApp.Workbooks.Add();
                    xlSheet = (Excel.Worksheet)xlBook.Worksheets[1];
                    xlSheet.Name = sheetName;

                    // Headers
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        xlSheet.Cells[1, c + 1] = dt.Columns[c].ColumnName;
                    }

                    headerRange = xlSheet.Range[
                        xlSheet.Cells[1, 1],
                        xlSheet.Cells[1, dt.Columns.Count]
                    ];

                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(230, 238, 250));
                    headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    // Data
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            object value = dt.Rows[r][c];
                            xlSheet.Cells[r + 2, c + 1] = value == DBNull.Value ? "" : Convert.ToString(value);
                        }
                    }

                    usedRange = xlSheet.Range[
                        xlSheet.Cells[1, 1],
                        xlSheet.Cells[Math.Max(1, dt.Rows.Count + 1), dt.Columns.Count]
                    ];

                    usedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    usedRange.Columns.AutoFit();

                    // Freeze header
                    xlApp.ActiveWindow.SplitRow = 1;
                    xlApp.ActiveWindow.FreezePanes = true;

                    xlBook.SaveAs(
                        sfd.FileName,
                        Excel.XlFileFormat.xlOpenXMLWorkbook);

                    MessageBox.Show("Excel generated: " + sfd.FileName, "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error generating Excel: " + ex.Message, "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (xlBook != null)
                    {
                        try { xlBook.Close(false); } catch { }
                    }
                    if (xlApp != null)
                    {
                        try { xlApp.Quit(); } catch { }
                    }

                    ReleaseComObjectSafe(usedRange);
                    ReleaseComObjectSafe(headerRange);
                    ReleaseComObjectSafe(xlSheet);
                    ReleaseComObjectSafe(xlBook);
                    ReleaseComObjectSafe(xlApp);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }


        private DataGridView CreateStep5Grid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
        }

        private TableLayoutPanel CreateTwoRowRoot(int topHeight)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.White
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, topHeight));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            return root;
        }

        /// <summary>
        /// Carga los datos de detalle y resumen ejecutivo para el Paso 5.
        /// </summary>
        private void LoadStep5Data()
        {
            if (_physicalCountId <= 0) return;

            try
            {
                _repo.LoadReviewData(_physicalCountId,
                    out _dtStep5Detail, out _dtStep5Executive);

                _dtStep5AllDetail    = _dtStep5Detail;
                _dtStep5AllExecutive = _dtStep5Executive;

                if (_dgvStep5Detail    != null) _dgvStep5Detail.DataSource    = _dtStep5Detail;
                if (_dgvStep5Executive != null) _dgvStep5Executive.DataSource = _dtStep5Executive;

                ApplyStep5GridFormatting();
                UpdateStep5Kpis();
                PopulateStep5LocationFilter();

                RefreshStep5ChartsFromTables(
                    _dtStep5AllDetail    ?? _dtStep5Detail,
                    _dtStep5AllExecutive ?? _dtStep5Executive);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar datos del Paso 5: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Llena el ComboBox con las ubicaciones distintas del inventario ────
        private void PopulateStep5LocationFilter()
        {
            if (_cboStep5LocationFilter == null) return;
            if (_dtStep5AllExecutive == null) return;

            // Temporalmente desconectar el evento para evitar dispararlo al cargar
            _cboStep5LocationFilter.SelectedIndexChanged -= Step5LocationFilter_Changed;

            _cboStep5LocationFilter.Items.Clear();

            // ✅ Primera opción: todas las ubicaciones
            _cboStep5LocationFilter.Items.Add(
                new LocationFilterItem { LocationId = -1, LocationName = "— All Locations —" });

            // ✅ Una entrada por ubicación, tomada del Executive DataTable
            foreach (DataRow r in _dtStep5AllExecutive.Rows)
            {
                int locId = Convert.ToInt32(r["Location_ID"]);
                string locName = Convert.ToString(r["Location"] ?? "EMPTY");

                _cboStep5LocationFilter.Items.Add(
                    new LocationFilterItem { LocationId = locId, LocationName = locName });
            }

            _cboStep5LocationFilter.DisplayMember = "LocationName";
            _cboStep5LocationFilter.ValueMember = "LocationId";
            _cboStep5LocationFilter.SelectedIndex = 0; // selecciona "All Locations"

            // Reconectar evento
            _cboStep5LocationFilter.SelectedIndexChanged += Step5LocationFilter_Changed;
        }

        // ── Clase auxiliar para los items del ComboBox ───────────────────────
        private class LocationFilterItem
        {
            public int LocationId { get; set; }
            public string LocationName { get; set; }

            public override string ToString() { return LocationName; }
        }

        /// <summary>
        /// Updates the KPI bar at the top of Step 5.
        /// Shows 0% if nothing has been counted yet.
        /// </summary>


        private void ApplyStep5GridFormatting()
        {
            // Detail grid
            if (_dgvStep5Detail != null)
            {
                HideColumn(_dgvStep5Detail, "Location_ID");

                SetNumberFormat(_dgvStep5Detail, "SystemStock", "N2");
                SetNumberFormat(_dgvStep5Detail, "CountedQty", "N2");
                SetNumberFormat(_dgvStep5Detail, "Difference", "N2");
                SetNumberFormat(_dgvStep5Detail, "Diff %", "N2");

                if (_dgvStep5Detail.Columns.Contains("Heat Number"))
                {
                    _dgvStep5Detail.Columns["Heat Number"].HeaderText = "Heat Number";
                    _dgvStep5Detail.Columns["Heat Number"].DisplayIndex = 5;

                }
            }

            // Executive grid
            if (_dgvStep5Executive != null)
            {
                HideColumn(_dgvStep5Executive, "Location_ID");
                SetNumberFormat(_dgvStep5Executive, "Lines", "N0");
                SetNumberFormat(_dgvStep5Executive, "Lines w/ Diff", "N0");
                SetNumberFormat(_dgvStep5Executive, "Net Diff", "N2");
                SetNumberFormat(_dgvStep5Executive, "Abs Diff", "N2");
                SetNumberFormat(_dgvStep5Executive, "Diff %", "N2");
            }
        }

        private void HideColumn(DataGridView dgv, string name)
        {
            if (dgv?.Columns.Contains(name) == true)
                dgv.Columns[name].Visible = false;
        }

        private void SetNumberFormat(DataGridView dgv, string name, string format)
        {
            if (dgv?.Columns.Contains(name) == true)
                dgv.Columns[name].DefaultCellStyle.Format = format;
        }


        // --- RDLC Preview & PDF Export ---

        private string GetReportPath(string fileName) =>
            Path.Combine(Application.StartupPath, "Reports", fileName);

        private Dictionary<string, object> BuildStep5ReportParams()
        {
            decimal totalDiff = 0m, sumPct = 0m;

            if (_dtStep5Executive != null)
                foreach (DataRow r in _dtStep5Executive.Rows)
                {
                    totalDiff += Convert.ToDecimal(r["Abs Diff"]);
                    sumPct += Convert.ToDecimal(r["Diff %"]);
                }

            if (sumPct > 100m) sumPct = 100m;

            return new Dictionary<string, object>
            {
                { "pFolio",      _physicalCountId },
                { "pTotalDiff",  totalDiff.ToString("0.##") },
                { "pTotalPct",   sumPct.ToString("0.##") }
            };
        }

        private void OpenReportViewer(string title, string rdlcPath, DataTable data,
      string dataSetName, Dictionary<string, object> parameters)
        {
            var f = new Form
            {
                Text = title,
                StartPosition = FormStartPosition.CenterParent,
                Width = 1200,
                Height = 800
            };

            var rv = new Microsoft.Reporting.WinForms.ReportViewer
            {
                Dock = DockStyle.Fill,
                ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local
            };

            rv.LocalReport.DataSources.Clear();

            if (File.Exists(rdlcPath))
                rv.LocalReport.ReportPath = rdlcPath;
            else
                rv.LocalReport.ReportEmbeddedResource = rdlcPath;

            rv.LocalReport.DataSources.Add(
                new Microsoft.Reporting.WinForms.ReportDataSource(dataSetName, data));

            if (parameters?.Count > 0)
            {
                var ps = new List<Microsoft.Reporting.WinForms.ReportParameter>();
                foreach (var kv in parameters)
                    ps.Add(new Microsoft.Reporting.WinForms.ReportParameter(kv.Key, Convert.ToString(kv.Value)));
                rv.LocalReport.SetParameters(ps);
            }

            f.Controls.Add(rv);
            rv.RefreshReport();
            f.ShowDialog(this);
        }

        private void ExportReportToPdf(string rdlcPath, DataTable data, string dataSetName,
            Dictionary<string, object> parameters, string defaultFileName)
        {
            var report = new Microsoft.Reporting.WinForms.LocalReport();
            report.ReportPath = rdlcPath;
            report.DataSources.Clear();
            report.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource(dataSetName, data));

            if (parameters?.Count > 0)
            {
                var ps = new List<Microsoft.Reporting.WinForms.ReportParameter>();
                foreach (var kv in parameters)
                    ps.Add(new Microsoft.Reporting.WinForms.ReportParameter(kv.Key, Convert.ToString(kv.Value)));
                report.SetParameters(ps);
            }

            byte[] bytes = report.Render("PDF", null,
                out string mimeType, out string encoding, out string extension,
                out string[] streams, out Microsoft.Reporting.WinForms.Warning[] warnings);

            using (var sfd = new SaveFileDialog { Filter = "PDF|*.pdf", FileName = defaultFileName })
            {
                if (sfd.ShowDialog(this) != DialogResult.OK) return;
                File.WriteAllBytes(sfd.FileName, bytes);
                MessageBox.Show("PDF generated: " + sfd.FileName, "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // --- Step 5 button handlers ---

        private void Step5PreviewDetail_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentDetailData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No detail data available.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ShowStep5PrintPreview("Differences Detail Report", dt, isExecutive: false);
        }

        private void Step5PreviewExecutive_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentExecutiveData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No executive data available.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ShowStep5PrintPreview("Executive Summary Report", dt, isExecutive: true);
        }

        private void Step5ExportPdfDetail_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentDetailData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No detail data to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ExportStep5ToPdf(dt, isExecutive: false,
                fileName: $"Detail_Differences_{_physicalCountId}.pdf");
        }

        private void Step5ExportPdfExecutive_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentExecutiveData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No executive data to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ExportStep5ToPdf(dt, isExecutive: true,
                fileName: $"Executive_Summary_{_physicalCountId}.pdf");
        }

        private DataTable GetStep5CurrentDetailData()
        {
            if (_dtStep5AllDetail == null) return _dtStep5Detail;
            int locId = GetStep5SelectedLocationId();
            if (locId == -1) return _dtStep5AllDetail;
            return LoadStep5DetailForLocation(locId);
        }

        private DataTable GetStep5CurrentExecutiveData()
        {
            if (_dtStep5AllExecutive == null) return _dtStep5Executive;
            int locId = GetStep5SelectedLocationId();
            if (locId == -1) return _dtStep5AllExecutive;
            return LoadStep5ExecutiveForLocation(locId);
        }

        /// <summary>Carga el detalle de diferencias por locación para el Paso 5.</summary>
        private DataTable LoadStep5DetailForLocation(int locationId)
        {
            try
            {
                return _repo.LoadDetailByLocation(_physicalCountId, locationId);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar detalle por locación: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new DataTable();
            }
        }

        /// <summary>Carga el resumen ejecutivo por locación para el Paso 5.</summary>
        private DataTable LoadStep5ExecutiveForLocation(int locationId)
        {
            try
            {
                return _repo.LoadExecutiveByLocation(_physicalCountId, locationId);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar resumen ejecutivo por locación: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new DataTable();
            }
        }

        // ── GDI+ Print Preview (no RDLC) ─────────────────────────────────────

        private void ShowStep5PrintPreview(string title, DataTable dt, bool isExecutive)
        {
            var pd = BuildStep5PrintDocument(title, dt, isExecutive);

            // ✅ PrintPreviewDialog está en System.Windows.Forms, no en System.Drawing.Printing
            using (var ppd = new System.Windows.Forms.PrintPreviewDialog())
            {
                ppd.Document = pd;
                ppd.Text = title;
                ppd.WindowState = FormWindowState.Maximized;
                ppd.ShowDialog(this);
            }
        }


        private void ExportStep5ToPdf(DataTable dt, bool isExecutive, string fileName)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Save PDF";
                dlg.Filter = "PDF files (*.pdf)|*.pdf";
                dlg.FileName = fileName;
                dlg.InitialDirectory =
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                if (dlg.ShowDialog() != DialogResult.OK) return;

                string path = dlg.FileName;

                try
                {
                    // ✅ GDI+ → print to file via PrintDocument
                    var pd = BuildStep5PrintDocument(
                        isExecutive ? "Executive Summary" : "Differences Detail",
                        dt, isExecutive);

                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = path;
                    pd.Print();

                    MessageBox.Show($"File saved:\n{path}", "Export OK",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Open folder
                    System.Diagnostics.Process.Start("explorer.exe",
                        $"/select,\"{path}\"");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error exporting: " + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ── Core: builds PrintDocument from DataTable ─────────────────────────
        /// <summary>
        /// Suma los elementos de un array int[] sin necesitar System.Linq.
        /// </summary>
        private int SumArray(int[] arr)
        {
            int total = 0;
            foreach (int v in arr) total += v;
            return total;
        }


        private System.Drawing.Printing.PrintDocument BuildStep5PrintDocument(
      string title, DataTable dt, bool isExecutive)
        {
            var pd = new System.Drawing.Printing.PrintDocument();

            // Columns to print
            string[] cols = isExecutive
                ? new[] { "Location", "Lines", "Lines Counted",
                  "Lines w/ Diff", "Abs Diff", "Diff %" }
                : new[] { "Location", "Part Number", "Description",
                  "Dimension", "SystemStock", "CountedQty",
                  "Difference", "Diff %" };

            int[] widths = isExecutive
                ? new[] { 160, 60, 90, 90, 100, 60 }
                : new[] { 90, 100, 160, 70, 80, 80, 80, 60 };

            // ✅ Declarar DENTRO del método pero FUERA del evento
            //    para que se reseteen cada vez que se llama BuildStep5PrintDocument
            int pageNum = 0;
            int rowIdx = 0;

            pd.PrintPage += (s, e) =>
            {
                pageNum++;
                var g = e.Graphics;
                var pageRect = e.MarginBounds;
                float x = pageRect.Left;
                float y = pageRect.Top;

                var fTitle = new Font("Segoe UI", 12, FontStyle.Bold);
                var fSub = new Font("Segoe UI", 8, FontStyle.Regular);
                var fHead = new Font("Segoe UI", 8, FontStyle.Bold);
                var fCell = new Font("Segoe UI", 8, FontStyle.Regular);
                float rowH = 18f;
                int totalW = SumArray(widths);

                // ── Page header ──────────────────────────────────────────────
                g.DrawString(title, fTitle, Brushes.Black, x, y);
                y += fTitle.GetHeight(g) + 2;

                g.DrawString(
                    "Inventory ID: " + _physicalCountId +
                    "    Date: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") +
                    "    Page: " + pageNum,
                    fSub, Brushes.Gray, x, y);
                y += fSub.GetHeight(g) + 8;

                // ── Column headers ───────────────────────────────────────────
                float cx = x;
                for (int i = 0; i < cols.Length; i++)
                {
                    var rect = new RectangleF(cx, y, widths[i], rowH);
                    g.FillRectangle(new SolidBrush(Color.FromArgb(31, 41, 55)), rect);
                    g.DrawString(cols[i], fHead, Brushes.White,
                        new RectangleF(cx + 2, y + 2, widths[i] - 4, rowH - 2));
                    cx += widths[i];
                }
                y += rowH;

                // ── Data rows ────────────────────────────────────────────────
                bool alt = false;
                while (rowIdx < dt.Rows.Count)
                {
                    if (y + rowH > pageRect.Bottom)
                    {
                        e.HasMorePages = true;
                        goto Cleanup;   // sal limpio sin disponer fuentes aún
                    }

                    var row = dt.Rows[rowIdx];
                    cx = x;
                    alt = !alt;

                    // Fondo alterno
                    if (alt)
                        g.FillRectangle(
                            new SolidBrush(Color.FromArgb(245, 247, 250)),
                            new RectangleF(x, y, totalW, rowH));

                    // Celdas
                    for (int i = 0; i < cols.Length; i++)
                    {
                        string val = dt.Columns.Contains(cols[i])
                            ? Convert.ToString(row[cols[i]] ?? "")
                            : "?";

                        g.DrawString(val, fCell, Brushes.Black,
                            new RectangleF(cx + 2, y + 2, widths[i] - 4, rowH - 2));
                        cx += widths[i];
                    }

                    // ✅ Línea inferior — solo UNA vez por fila
                    g.DrawLine(Pens.LightGray, x, y + rowH, x + totalW, y + rowH);

                    rowIdx++;
                    y += rowH;
                }

                e.HasMorePages = false;

            Cleanup:
                fTitle.Dispose();
                fSub.Dispose();
                fHead.Dispose();
                fCell.Dispose();
            };

            return pd;
        }



        // ── Helpers: devuelve datos filtrados o completos según ComboBox ──────

        // --- Finalize wizard ---

        /// <summary>
        /// Finalizes the inventory — sets status to Finished (3) and disables navigation.
        /// </summary>
        private bool _finalizeWarningOpen = false;
        /// <summary>Finaliza el inventario marcándolo como completado.</summary>
        private void FinalizeWizard()
        {
            if (_physicalCountId <= 0)
            {
                MessageBox.Show("No hay un ID de inventario para finalizar.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int noContados = _repo.CountUncountedItems(_physicalCountId);

            if (noContados > 0)
            {
                if (_finalizeWarningOpen) return;

                _finalizeWarningOpen = true;
                try
                {
                    if (MessageBox.Show(
                            "¿Está seguro de que desea cerrar el inventario?",
                            "Advertencia",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) == DialogResult.No)
                        return;
                }
                finally
                {
                    _finalizeWarningOpen = false;
                }
            }

            try
            {
                var sf = new Persal.System_Functions();
                bool ok = sf.Set_Physical_Count_Status(STATUS_FINISHED, _physicalCountId);

                if (!ok)
                {
                    MessageBox.Show(
                        "No se pudo finalizar el inventario:\n" + sf.G.Error_Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                StopStep3RefreshTimer();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error al finalizar el inventario: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// <summary>
        /// Asegura que existen snapshots para el conteo actual.
        /// Si no existen los genera llamando al repositorio.
        /// Devuelve false si no se encontraron movimientos.
        /// </summary>
        private bool EnsureSnapshotsExist()
        {
            if (_physicalCountId <= 0) return false;

            try
            {
                bool   existentes;
                int    generados;
                string mensaje;

                bool ok = _repo.EnsureSnapshots(_physicalCountId,
                    out existentes, out generados, out mensaje);

                if (existentes)
                    return true;

                if (!ok)
                {
                    MessageBox.Show(
                        mensaje ?? "Error al generar los snapshots.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                if (generados <= 0)
                {
                    MessageBox.Show(
                        "No se generaron snapshots.\n\n" +
                        "Posibles causas:\n" +
                        "• No hay movimientos para las locaciones seleccionadas.\n" +
                        "• Todas las cantidades son cero.",
                        "Advertencia",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                MessageBox.Show(
                    mensaje ?? (generados + " líneas generadas para el conteo."),
                    "Éxito",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error al generar snapshots:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private Bitmap CaptureControlToBitmap(Control ctrl)
        {
            if (ctrl == null) return null;
            if (ctrl.Width <= 0 || ctrl.Height <= 0) return null;

            var bmp = new Bitmap(ctrl.Width, ctrl.Height);
            ctrl.DrawToBitmap(bmp, new Rectangle(0, 0, ctrl.Width, ctrl.Height));
            return bmp;
        }

        private void EnsureStep5ChartsPrintObjects()
        {
            if (_pdStep5Charts == null)
            {
                _pdStep5Charts = new System.Drawing.Printing.PrintDocument();
                _pdStep5Charts.DefaultPageSettings.Landscape = true;
                _pdStep5Charts.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(25, 25, 25, 25);
                _pdStep5Charts.PrintPage -= Step5Charts_PrintPage;
                _pdStep5Charts.PrintPage += Step5Charts_PrintPage;
            }

            if (_ppdStep5Charts == null)
            {
                _ppdStep5Charts = new PrintPreviewDialog();
                _ppdStep5Charts.Document = _pdStep5Charts;
                _ppdStep5Charts.Width = 1200;
                _ppdStep5Charts.Height = 850;
            }
            else
            {
                _ppdStep5Charts.Document = _pdStep5Charts;
            }
        }

        private void Step5ExportPdfCharts_Click(object sender, EventArgs e)
        {
            try
            {
                if (_pnlStep5ChartsHost == null)
                {
                    MessageBox.Show("Charts area is not available.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (_step5ChartsPrintBitmap != null)
                {
                    _step5ChartsPrintBitmap.Dispose();
                    _step5ChartsPrintBitmap = null;
                }

                _step5ChartsPrintBitmap = CaptureControlToBitmap(_pnlStep5ChartsHost);

                if (_step5ChartsPrintBitmap == null)
                {
                    MessageBox.Show("Unable to capture charts for PDF export.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                EnsureStep5ChartsPrintObjects();

                using (var dlg = new SaveFileDialog())
                {
                    dlg.Title = "Save PDF";
                    dlg.Filter = "PDF files (*.pdf)|*.pdf";
                    dlg.FileName = $"Charts_{_physicalCountId}.pdf";
                    dlg.InitialDirectory =
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    if (dlg.ShowDialog(this) != DialogResult.OK) return;

                    string path = dlg.FileName;

                    _pdStep5Charts.PrinterSettings.PrintToFile = true;
                    _pdStep5Charts.PrinterSettings.PrintFileName = path;
                    _pdStep5Charts.DefaultPageSettings.Landscape = true;

                    _pdStep5Charts.Print();

                    MessageBox.Show("PDF saved:\n" + path, "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting charts PDF: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void Step5Charts_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            if (_step5ChartsPrintBitmap == null)
            {
                e.HasMorePages = false;
                return;
            }

            e.Graphics.Clear(Color.White);

            Rectangle margin = e.MarginBounds;

            float ratioX = (float)margin.Width / _step5ChartsPrintBitmap.Width;
            float ratioY = (float)margin.Height / _step5ChartsPrintBitmap.Height;
            float ratio = Math.Min(ratioX, ratioY);

            int drawW = (int)(_step5ChartsPrintBitmap.Width * ratio);
            int drawH = (int)(_step5ChartsPrintBitmap.Height * ratio);

            int drawX = margin.Left + ((margin.Width - drawW) / 2);
            int drawY = margin.Top + ((margin.Height - drawH) / 2);

            e.Graphics.DrawImage(
                _step5ChartsPrintBitmap,
                new Rectangle(drawX, drawY, drawW, drawH),
                new Rectangle(0, 0, _step5ChartsPrintBitmap.Width, _step5ChartsPrintBitmap.Height),
                GraphicsUnit.Pixel);

            e.HasMorePages = false;
        }


        #endregion

        // =====================================================================
        // HELPERS DE UI — Utilidades de Form (conexión, usuario, combos)
        // =====================================================================
        #region HELPERS DE UI

        /// <summary>Returns the active connection string.</summary>
        private string GetConnectionString()
        {
            if (!string.IsNullOrWhiteSpace(_connectionString))
                return _connectionString;
            try { return new con().ConnectionString; }
            catch { return ""; }
        }

        /// <summary>Returns the current logged-in user ID.</summary>
        private int GetCurrentUserId()
        {
            try { return Convert.ToInt32(Persal.GlobalVars.User_ID); }
            catch { return 0; }
        }

        /// <summary>Returns the selected integer value from a ComboBox.</summary>
        private int GetComboSelectedValueAsInt(ComboBox cbo)
        {
            try
            {
                if (cbo?.SelectedValue != null &&
                    int.TryParse(cbo.SelectedValue.ToString(), out int v))
                    return v;
            }
            catch { }
            return cbo?.SelectedIndex ?? 0;
        }

        #endregion

        // =====================================================================
        // #region EXPORTACIÓN CSV
        // =====================================================================
        #region EXPORTACIÓN CSV

        /// <summary>
        /// Exports a DataTable to a semicolon-delimited CSV file (UTF-8).
        /// Semicolon separator works better with Spanish/European Excel settings.
        /// </summary>

        private string EscapeCsvValue(object value)
        {
            if (value == null || value == DBNull.Value) return "";

            string s = Convert.ToString(value) ?? "";

            bool mustQuote =
                s.Contains(",") ||
                s.Contains("\"") ||
                s.Contains("\r") ||
                s.Contains("\n");

            s = s.Replace("\"", "\"\"");

            return mustQuote ? "\"" + s + "\"" : s;
        }

        private void ExportDataTableToCsv(DataTable dt, string defaultFileName)
        {
            if (dt == null || dt.Columns.Count == 0)
            {
                MessageBox.Show("No data available to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                sfd.FileName = defaultFileName;
                sfd.Title = "Export CSV";

                if (sfd.ShowDialog(this) != DialogResult.OK)
                    return;

                var sb = new StringBuilder();

                // Header row
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    if (c > 0) sb.Append(",");
                    sb.Append(EscapeCsvValue(dt.Columns[c].ColumnName));
                }
                sb.AppendLine();

                // Data rows
                foreach (DataRow row in dt.Rows)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        if (c > 0) sb.Append(",");
                        sb.Append(EscapeCsvValue(row[c]));
                    }
                    sb.AppendLine();
                }

                // UTF-8 with BOM so Excel opens accents correctly
                File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));

                MessageBox.Show("CSV generated: " + sfd.FileName, "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void Step5ExportCsvDetail_Click(object sender, EventArgs e)
        {
            DataTable dt = GetStep5CurrentDetailData();
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No detail data to export.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExportDataTableToCsv(dt, $"Detail_Differences_{_physicalCountId}.csv");
        }


        private void ExportToCsv(DataTable dt, string filePath)
        {
            const string sep = ";";

            using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // Header row
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i > 0) sw.Write(sep);
                    sw.Write(EscapeCsvValue(dt.Columns[i].ColumnName));
                }
                sw.WriteLine();

                // Data rows
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        if (c > 0) sw.Write(sep);
                        string val = dt.Rows[r][c] == DBNull.Value
                            ? ""
                            : dt.Rows[r][c].ToString();
                        sw.Write(EscapeCsvValue(val));
                    }
                    sw.WriteLine();
                }
            }
        }

        private string EscapeCsvValue(string s)
        {
            if (s == null) return "";
            bool mustQuote = s.Contains(";") || s.Contains("\"") ||
                             s.Contains("\n") || s.Contains("\r");
            if (mustQuote) s = "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        #endregion

        // =====================================================================
        // #region SHELL DE UI MODERNO
        // =====================================================================
        #region SHELL DE UI MODERNO

        /// <summary>
        /// Applies the full modern UI shell:
        /// borderless chrome, header bar, stepper bar, card content panel, footer buttons.
        /// Called once on Load.
        /// </summary>
        private void ApplyModernUiShell()
        {
            if (_modernUiReady) return;

            // Form-level
            this.Font = new Font("Segoe UI", 9F);
            this.BackColor = Color.FromArgb(243, 244, 246);
            this.Padding = new Padding(10);
            this.MinimumSize = new Size(1180, 760);
            this.StartPosition = FormStartPosition.CenterScreen;

            if (UseCustomChrome)
            {
                this.FormBorderStyle = FormBorderStyle.None;
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                ApplyRoundedRegion(this, FormCornerRadius);
            }

            // Root panel
            _pnlShellRoot = new Panel { Dock = DockStyle.Fill, BackColor = this.BackColor };
            this.Controls.Add(_pnlShellRoot);
            _pnlShellRoot.BringToFront();

            // --- Header bar ---
            _pnlHeaderBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = Color.FromArgb(31, 41, 55),
                Padding = new Padding(16, 10, 12, 10)
            };

            _lblHeaderTitle = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 22,
                Text = "Cyclic Inventory",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };

            _lblHeaderSubtitle = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = "Create physical count",
                ForeColor = Color.FromArgb(156, 163, 175),
                Font = new Font("Segoe UI", 9F),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var headerLeft = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = _pnlHeaderBar.BackColor
            };
            headerLeft.Controls.Add(_lblHeaderSubtitle);
            headerLeft.Controls.Add(_lblHeaderTitle);

            // Chrome window buttons
            var headerRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 172,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = _pnlHeaderBar.BackColor,
                Padding = new Padding(0)
            };

            _btnClose = CreateChromeButton("✕");
            _btnMax = CreateChromeButton("□");
            _btnMin = CreateChromeButton("—");

            _btnClose.Click += (s, e) => this.Close();
            _btnMax.Click += (s, e) => ToggleMaximize();
            _btnMin.Click += (s, e) => this.WindowState = FormWindowState.Minimized;

            headerRight.Controls.Add(_btnClose);
            headerRight.Controls.Add(_btnMax);
            headerRight.Controls.Add(_btnMin);

            // Drag support
            _pnlHeaderBar.MouseDown += Header_MouseDown;
            headerLeft.MouseDown += Header_MouseDown;
            _lblHeaderTitle.MouseDown += Header_MouseDown;
            _lblHeaderSubtitle.MouseDown += Header_MouseDown;

            _pnlHeaderBar.Controls.Add(headerLeft);
            _pnlHeaderBar.Controls.Add(headerRight);

            // --- Stepper bar ---
            _pnlStepperBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = Color.White,
                Padding = new Padding(16, 10, 16, 8)
            };

            _stepper = new StepIndicatorBar(new[]
                 {
                    "Create", "Locations", "Count", "Finalize"
                });
            _stepper.Dock = DockStyle.Fill;
            _pnlStepperBar.Controls.Add(_stepper);

            // --- Footer ---
            _pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 64,
                BackColor = Color.White,
                Padding = new Padding(16, 12, 16, 12)
            };

            // Re-parent designer buttons
            if (btnBack != null) btnBack.Parent = null;
            // if (btnCancel != null) btnCancel.Parent = null;
            if (btnNext != null) btnNext.Parent = null;

            StyleButtonSecondary(btnBack);
            StyleButtonGhost(btnCancel);
            StyleButtonPrimary(btnNext);

            var footerLeft = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = 220,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = _pnlFooter.BackColor
            };

            var footerRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 360,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = _pnlFooter.BackColor
            };

            if (btnBack != null) { btnBack.Margin = new Padding(0, 0, 12, 0); footerLeft.Controls.Add(btnBack); }
            // if (btnCancel != null) { btnCancel.Margin = new Padding(0, 0, 12, 0); footerLeft.Controls.Add(btnCancel); }
            if (btnNext != null) { btnNext.Margin = new Padding(0); footerRight.Controls.Add(btnNext); }

            _pnlFooter.Controls.Add(footerLeft);
            _pnlFooter.Controls.Add(footerRight);

            // --- Content card ---
            _pnlContent = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = this.BackColor,
                Padding = new Padding(0, 12, 0, 12)
            };

            _cardContent = new ModernCardPanel
            {
                Dock = DockStyle.Fill,
                Radius = 10,
                BorderColor = Color.FromArgb(229, 231, 235),
                BorderThickness = 1,
                Padding = new Padding(16),
                BackColor = Color.White
            };

            _pnlStepHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };

            // Move designer step panels into the host
            ReparentStepPanels();

            // Hide legacy controls
            if (lblStepTitle != null) lblStepTitle.Visible = false;
            if (prgSteps != null) prgSteps.Visible = false;

            _cardContent.Controls.Add(_pnlStepHost);
            _pnlContent.Controls.Add(_cardContent);

            // Assemble shell (add in reverse dock order)
            _pnlShellRoot.Controls.Add(_pnlContent);
            _pnlShellRoot.Controls.Add(_pnlFooter);
            _pnlShellRoot.Controls.Add(_pnlStepperBar);
            _pnlShellRoot.Controls.Add(_pnlHeaderBar);

            // Style existing grids
            StyleGrid(dgvLocations);
            StyleGrid(dgvStep3Preview);
            StyleGrid(dgvStep4);

            _modernUiReady = true;
            UpdateChromeForWindowState();
            UpdateStepper(_step);
            UpdateHeaderSubtitle(_step);
        }

        /// <summary>Moves designer step panels into the step host panel.</summary>
        private void ReparentStepPanels()
        {
            if (_pnlStepHost == null) return;

            foreach (var panel in new Control[]
            {
                pnlStep1_Create, pnlStep2_Locations,
                pnlStep3_Confirm, pnlStep4_Capture, pnlStep5_Review
            })
            {
                if (panel == null) continue;
                panel.Parent = null;
                panel.Dock = DockStyle.Fill;
                panel.Margin = new Padding(0);
                panel.Padding = new Padding(0);
                panel.BackColor = Color.White;
                _pnlStepHost.Controls.Add(panel);
            }
        }

        private void UpdateStepper(int step)
        {
            _stepper?.SetStep(step);
        }

        private void UpdateHeaderSubtitle(int step)
        {
            if (_lblHeaderSubtitle == null) return;
            switch (step)
            {
                case 1: _lblHeaderSubtitle.Text = "Step 1 of 4 — Create cyclic inventory"; break;
                case 2: _lblHeaderSubtitle.Text = "Step 2 of 4 — Select locations"; break;
                case 3: _lblHeaderSubtitle.Text = "Step 3 of 4 — Count report"; break;
                case 4: _lblHeaderSubtitle.Text = "Step 4 of 4 — Review & Finalize"; break;
                default: _lblHeaderSubtitle.Text = ""; break;
            }
        }

        private void UpdateButtonVisualStyle(Button btn, bool enabled)
        {
            if (btn == null) return;

            if (enabled)
            {
                btn.ForeColor = Color.FromArgb(31, 41, 55);
                btn.BackColor = Color.White;
                btn.Cursor = Cursors.Hand;
                btn.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(249, 250, 251);
                btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(243, 244, 246);
                btn.Enabled = true;
            }
            else
            {
                btn.ForeColor = Color.FromArgb(156, 163, 175);
                btn.BackColor = Color.FromArgb(249, 250, 251);
                btn.Cursor = Cursors.No;
                btn.FlatAppearance.BorderColor = Color.FromArgb(229, 231, 235);
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(249, 250, 251);
                btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(249, 250, 251);
                btn.Enabled = false;
            }
        }

        #endregion

        // =====================================================================
        // #region ESTILOS Y CONTROLES
        // =====================================================================
        #region ESTILOS Y CONTROLES

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

        private void StyleButtonSecondary(Button b)
        {
            if (b == null) return;
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            b.BackColor = Color.White;
            b.ForeColor = Color.FromArgb(31, 41, 55);
            b.Height = 40;
            b.Width = Math.Max(120, b.Width);
            b.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            b.Cursor = Cursors.Hand;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(249, 250, 251);
            b.FlatAppearance.MouseDownBackColor = Color.FromArgb(243, 244, 246);
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

        private void StyleControl(Control control)
        {
            if (control is TextBox tb)
            {
                tb.Font = new Font("Segoe UI", 10F);
                tb.BackColor = Color.FromArgb(249, 250, 251);
                tb.ForeColor = Color.FromArgb(17, 24, 39);
                tb.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is ComboBox cb)
            {
                cb.Font = new Font("Segoe UI", 10F);
                cb.BackColor = Color.FromArgb(249, 250, 251);
                cb.ForeColor = Color.FromArgb(17, 24, 39);
                cb.FlatStyle = FlatStyle.Flat;
            }
            else if (control is DateTimePicker dtp)
            {
                dtp.Font = new Font("Segoe UI", 10F);
                dtp.BackColor = Color.FromArgb(249, 250, 251);
                dtp.ForeColor = Color.FromArgb(17, 24, 39);
            }
        }

        private void NormalizeLabel(Label lbl)
        {
            if (lbl == null) return;
            lbl.AutoSize = true;
            lbl.ForeColor = Color.FromArgb(17, 24, 39);
            lbl.Margin = new Padding(0, 0, 0, 6);
        }

        #endregion

        // =====================================================================
        // #region HELPERS DE LAYOUT (reusable builders)
        // =====================================================================
        #region HELPERS DE LAYOUT

        private Panel CreateSection() => new Panel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            Padding = new Padding(0, 0, 0, 20),
            Margin = new Padding(0, 0, 0, 20)
        };

        private Label CreateSectionTitle(string text) => new Label
        {
            Text = text,
            Font = new Font("Segoe UI", 11F, FontStyle.Bold),
            ForeColor = Color.FromArgb(31, 41, 55),
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 12),
            Dock = DockStyle.Fill
        };

        private Panel CreateSeparator() => new Panel
        {
            Dock = DockStyle.Fill,
            Height = 1,
            BackColor = Color.FromArgb(229, 231, 235),
            Margin = new Padding(0, 12, 0, 0),
            AutoSize = true
        };

        private TableLayoutPanel CreateTableContainer(int rows, int cols) => new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = rows,
            ColumnCount = cols,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };

        private TableLayoutPanel CreateTwoColumnRow()
        {
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            row.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return row;
        }

        private void AddColumnStyle(TableLayoutPanel tlp, SizeType type, float value)
        {
            tlp.ColumnStyles.Clear();
            tlp.ColumnStyles.Add(new ColumnStyle(type, value));
        }

        private void AddRowStyles(TableLayoutPanel tlp, int count)
        {
            tlp.RowStyles.Clear();
            for (int i = 0; i < count; i++)
                tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        /// <summary>Creates a label + control field row (horizontal layout).</summary>
        private Panel CreateFieldRow(string label, Control control)
        {
            var wrapper = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0, 0, 12, 16)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lbl = new Label
            {
                Text = label,
                Font = new Font("Segoe UI", 10F),
                ForeColor = Color.FromArgb(55, 65, 81),
                AutoSize = true,
                Padding = new Padding(0, 8, 16, 0),
                TextAlign = ContentAlignment.MiddleLeft
            };

            control.Parent = null;
            control.Height = 36;
            control.Margin = new Padding(0);
            StyleControl(control);

            grid.Controls.Add(lbl, 0, 0);
            grid.Controls.Add(control, 1, 0);
            wrapper.Controls.Add(grid);

            return wrapper;
        }

        /// <summary>Collects all Label descendants of root, excluding those in the exclude set.</summary>
        private List<Label> CollectLabels(Control root, HashSet<Control> exclude)
        {
            var list = new List<Label>();
            var stack = new Stack<Control>();
            stack.Push(root);

            while (stack.Count > 0)
            {
                var cur = stack.Pop();
                foreach (Control child in cur.Controls)
                {
                    stack.Push(child);
                    if (child is Label lbl && (exclude == null || !exclude.Contains(lbl)))
                        list.Add(lbl);
                }
            }

            return list;
        }

        /// <summary>Detaches a list of controls from their parents.</summary>
        private void DetachControls(IEnumerable<Control> controls)
        {
            if (controls == null) return;
            foreach (var c in controls)
            {
                if (c?.Parent != null) c.Parent.Controls.Remove(c);
                if (c != null) c.Parent = null;
            }
        }

        #endregion

        // =====================================================================
        // #region CHROME DE VENTANA (borderless resize / drag / rounded)
        // =====================================================================
        #region CHROME DE VENTANA

        private void ConfigureFormResizeBehavior()
        {
            if (!UseCustomChrome) return;
            this.Resize += (s, e) => UpdateChromeForWindowState();
            this.SizeChanged += (s, e) => UpdateChromeForWindowState();
            UpdateChromeForWindowState();
        }

        private void ToggleMaximize()
        {
            if (this.WindowState == FormWindowState.Maximized)
                this.WindowState = FormWindowState.Normal;
            else
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Maximized;
            }
            UpdateChromeForWindowState();
        }

        private void UpdateChromeForWindowState()
        {
            if (!UseCustomChrome) return;

            bool maximized = (this.WindowState == FormWindowState.Maximized);

            if (_btnMax != null) _btnMax.Text = maximized ? "❐" : "□";

            if (maximized)
            {
                this.Padding = new Padding(0);
                this.Region = null;
            }
            else
            {
                this.Padding = new Padding(0);
                ApplyRoundedRegion(this, FormCornerRadius);
            }

            if (_pnlShellRoot != null)
                _pnlShellRoot.Padding = maximized ? new Padding(6) : new Padding(0);

            if (_cardContent != null)
                _cardContent.Margin = maximized ? new Padding(2) : new Padding(0);
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

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
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

        // =====================================================================
        // #region CONTROL STEPPER (inner class)
        // =====================================================================
        #region CONTROL STEPPER

        private class StepIndicatorBar : Panel
        {
            private readonly List<StepItem> _items = new List<StepItem>();

            public StepIndicatorBar(string[] titles)
            {
                this.DoubleBuffered = true;
                this.BackColor = Color.White;

                var flow = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = false,
                    AutoScroll = true,
                    BackColor = Color.White,
                    Padding = new Padding(0, 2, 0, 0)
                };

                for (int i = 0; i < titles.Length; i++)
                {
                    var item = new StepItem(i + 1, titles[i]);
                    item.Margin = new Padding(0, 0, 12, 0);
                    _items.Add(item);
                    flow.Controls.Add(item);
                }

                this.Controls.Add(flow);
                SetStep(1);
            }

            public void SetStep(int step)
            {
                step = Math.Max(1, Math.Min(step, _items.Count));
                for (int i = 0; i < _items.Count; i++)
                {
                    int s = i + 1;
                    if (s < step) _items[i].SetState(StepState.Done);
                    else if (s == step) _items[i].SetState(StepState.Active);
                    else _items[i].SetState(StepState.Todo);
                }
            }

            private enum StepState { Todo, Active, Done }

            private class StepItem : Panel
            {
                private readonly Label _dot;
                private readonly Label _text;

                public StepItem(int index, string title)
                {
                    this.Height = 28;
                    this.Width = 150;
                    this.BackColor = Color.White;

                    _dot = new Label
                    {
                        Width = 24,
                        Height = 24,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                        Margin = new Padding(0, 2, 8, 0)
                    };

                    _text = new Label
                    {
                        AutoSize = false,
                        Dock = DockStyle.Fill,
                        Text = title,
                        TextAlign = ContentAlignment.MiddleLeft,
                        Font = new Font("Segoe UI", 9F)
                    };

                    var row = new FlowLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        FlowDirection = FlowDirection.LeftToRight,
                        WrapContents = false,
                        BackColor = Color.White
                    };
                    row.Controls.Add(_dot);
                    row.Controls.Add(_text);
                    this.Controls.Add(row);

                    SetState(StepState.Todo);
                }

                public void SetState(StepState state)
                {
                    switch (state)
                    {
                        case StepState.Active:
                            _dot.Text = "●";
                            _dot.ForeColor = Color.FromArgb(37, 99, 235);
                            _text.ForeColor = Color.FromArgb(17, 24, 39);
                            _text.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                            break;

                        case StepState.Done:
                            _dot.Text = "✔";
                            _dot.ForeColor = Color.FromArgb(16, 185, 129);
                            _text.ForeColor = Color.FromArgb(107, 114, 128);
                            _text.Font = new Font("Segoe UI", 9F);
                            break;

                        default: // Todo
                            _dot.Text = "○";
                            _dot.ForeColor = Color.FromArgb(156, 163, 175);
                            _text.ForeColor = Color.FromArgb(107, 114, 128);
                            _text.Font = new Font("Segoe UI", 9F);
                            break;
                    }
                }
            }
        }

        #endregion

        // =====================================================================
        // #region TARJETA MODERNA (inner class)
        // =====================================================================
        #region TARJETA MODERNA

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

        // =====================================================================
        // #region DIÁLOGO DE BÚSQUEDA DE USUARIO (F3)
        // =====================================================================
        #region DIÁLOGO DE BÚSQUEDA DE USUARIO

        private sealed class UserLookupDialog : Form
        {
            private readonly string _cs;
            private TextBox _txt;
            private DataGridView _grid;
            private Button _btnOk;
            //    private Button _btnCancel;

            public int SelectedUserId { get; private set; }
            public string SelectedName { get; private set; }
            public string SelectedEmployeeNumber { get; private set; }
            public string SelectedJobPosition { get; private set; }

            public UserLookupDialog(string connectionString)
            {
                _cs = connectionString ?? "";

                this.Text = "Search Responsible (F3)";
                this.Width = 800;
                this.Height = 450;
                this.StartPosition = FormStartPosition.CenterParent;

                _txt = new TextBox { Dock = DockStyle.Top };

                _grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    MultiSelect = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                };

                var pnlButtons = new Panel { Dock = DockStyle.Bottom, Height = 44 };
                _btnOk = new Button { Text = "OK", Dock = DockStyle.Right, Width = 100 };
                // _btnCancel = new Button { Text = "Cancel", Dock = DockStyle.Right, Width = 100 };

                //  pnlButtons.Controls.Add(_btnCancel);
                pnlButtons.Controls.Add(_btnOk);

                this.Controls.Add(_grid);
                this.Controls.Add(_txt);
                this.Controls.Add(pnlButtons);

                this.Load += (_, __) => LoadUsers("");
                _txt.TextChanged += (_, __) => LoadUsers(_txt.Text.Trim());
                _grid.CellDoubleClick += (_, __) => AcceptSelection();
                _btnOk.Click += (_, __) => AcceptSelection();
                //   _btnCancel.Click += (_, __) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            }

            private void LoadUsers(string filter)
            {
                const string sql = @"
                    SELECT TOP 200
                        User_ID, Name, Employee_Number, Job_Position
                    FROM Users
                    WHERE (@p = '' OR Employee_Number LIKE @like OR Name LIKE @like)
                    ORDER BY Name;";

                using (var conn = new SqlConnection(_cs))
                using (var cmd = new SqlCommand(sql, conn))
                {
                    string p = filter ?? "";
                    cmd.Parameters.AddWithValue("@p", p);
                    cmd.Parameters.AddWithValue("@like", "%" + p + "%");

                    var dt = new DataTable();
                    conn.Open();
                    using (var da = new SqlDataAdapter(cmd)) da.Fill(dt);
                    _grid.DataSource = dt;
                }
            }

            private void AcceptSelection()
            {
                if (_grid.CurrentRow == null) return;

                SelectedUserId = Convert.ToInt32(_grid.CurrentRow.Cells["User_ID"].Value);
                SelectedName = Convert.ToString(_grid.CurrentRow.Cells["Name"].Value);
                SelectedEmployeeNumber = Convert.ToString(_grid.CurrentRow.Cells["Employee_Number"].Value);
                SelectedJobPosition = Convert.ToString(_grid.CurrentRow.Cells["Job_Position"].Value);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        #endregion

        //Conexion Con monitor Physical_Counts_Managemet
        #region CONEXIÓN CON MONITOR

        private int _startStep = 1;
        public PhysicalCountCyclicForm(string connectionString, int physicalCountId, int startStep)
        : this(connectionString)
        {
            _physicalCountId = physicalCountId;
            _startStep = startStep;
        }


        #endregion

    } // end class PhysicalCountCyclicForm
} // end namespace 
