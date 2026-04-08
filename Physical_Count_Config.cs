/*
 * =====================================================================
 * SCRIPT DE CREACIÓN DE TABLAS — Physical Count Config
 * Ejecutar en la base de datos antes de usar este módulo
 * =====================================================================
 *
 * -- Tabla principal de programaciones
 * CREATE TABLE dbo.PC_Config (
 *     Id                 INT IDENTITY(1,1) PRIMARY KEY,
 *     Nombre             NVARCHAR(150)  NOT NULL,
 *     Intervalo_Dias     INT            NOT NULL DEFAULT 365,
 *     Color_Hex          NVARCHAR(7)    NOT NULL DEFAULT '#3B82F6',
 *     Fecha_Base         DATE           NOT NULL,
 *     Proximo_Inventario DATE           NOT NULL,
 *     Ultimo_PC_Id       INT            NULL REFERENCES dbo.Physical_Counts(Physical_Count_ID),
 *     Activo             BIT            NOT NULL DEFAULT 1,
 *     Creado_Por         INT            NOT NULL,
 *     Creado_En          DATETIME       NOT NULL DEFAULT GETDATE()
 * );
 *
 * -- Locaciones asignadas a cada programación
 * CREATE TABLE dbo.PC_Config_Locaciones (
 *     Id            INT IDENTITY(1,1) PRIMARY KEY,
 *     PC_Config_Id  INT NOT NULL REFERENCES dbo.PC_Config(Id) ON DELETE CASCADE,
 *     Almacen_Id    INT NOT NULL,
 *     Area_Id       INT NULL,
 *     Locacion_Id   INT NULL
 * );
 *
 * -- Log de acciones (1=Creado, 2=Pospuesto)
 * CREATE TABLE dbo.PC_Config_Log (
 *     Id             INT IDENTITY(1,1) PRIMARY KEY,
 *     PC_Config_Id   INT      NOT NULL REFERENCES dbo.PC_Config(Id),
 *     Accion         TINYINT  NOT NULL,
 *     Usuario_Id     INT      NOT NULL,
 *     Fecha_Accion   DATETIME NOT NULL DEFAULT GETDATE(),
 *     PC_Id_Generado INT      NULL
 * );
 */

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace Persal._003_Physical_Counting_2
{
    // =========================================================================
    // Physical_Count_Config — Pantalla de configuración de programación de inventarios
    // =========================================================================
    public partial class Physical_Count_Config : Form
    {
        // =====================================================================
        // #region CAMPOS PRIVADOS
        // =====================================================================
        #region CAMPOS PRIVADOS

        // --- Conexión y repositorios ---
        private string _connectionString;
        private PC_ConfigRepository _repoConfig;

        // --- Estado ---
        private int _idSeleccionado = 0;   // 0 = nueva programación
        private int _usuarioActual;

        // --- Árbol de ubicaciones ---
        private DataGridView _dgvAlmacen;
        private DataGridView _dgvArea;
        private DataGridView _dgvLocacion;
        private HashSet<int> _checkedAreas = new HashSet<int>();
        private HashSet<int> _checkedLocaciones = new HashSet<int>();

        // ID fijo del almacén raíz (WAREHOUSE SUPPLIES). Se define aquí para evitar
        // el uso de literales "magic number" dispersos en el código.
        private const int DEFAULT_WAREHOUSE_ID = 3;
        private readonly int _almacenFijoId = DEFAULT_WAREHOUSE_ID;

        // --- Controles del detalle ---
        private TextBox _txtNombre;
        private NumericUpDown _nudIntervalo;
        private ComboBox _cboColor;
        private DateTimePicker _dtpProximo;
        private DataGridView _dgvLista;
        private Panel _pnlAlerta;
        private Label _lblAlerta;

        // --- Chrome ---
        private Panel _pnlHeader;
        private Button _btnCerrar;
        private Button _btnMaximizar;
        private Button _btnMinimizar;

        #endregion

        // =====================================================================
        // #region CONSTRUCTORES
        // =====================================================================
        #region CONSTRUCTORES

        /// <summary>Constructor principal — recibe la cadena de conexión.</summary>
        public Physical_Count_Config(string connectionString)
        {
            // Cuando connectionString es nulo o vacío, el repositorio y los helpers
            // intentarán obtener la conexión desde el objeto global 'con()' del proyecto.
            _connectionString = connectionString ?? "";
            _repoConfig = new PC_ConfigRepository(_connectionString);
            _usuarioActual = ObtenerUsuarioActual();

            SetupUI();
            this.Load += OnFormLoad;
        }

        /// <summary>Constructor sin parámetros (requiere conexión válida en el entorno).</summary>
        public Physical_Count_Config() : this("") { }

        #endregion

        // =====================================================================
        // #region EVENTOS DEL FORM
        // =====================================================================
        #region EVENTOS DEL FORM

        private void OnFormLoad(object sender, EventArgs e)
        {
            CargarLista();
            MostrarAlertaSiHayVencidas();
        }

        #endregion

        // =====================================================================
        // #region SETUP DE UI (todo el layout construido por código)
        // =====================================================================
        #region SETUP UI

        private void SetupUI()
        {
            // --- Propiedades del form ---
            this.Text = "Inventory Schedule";
            this.Font = new Font("Segoe UI", 9F);
            this.BackColor = Color.FromArgb(243, 244, 246);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(1100, 700);
            this.Size = new Size(1200, 750);
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizedBounds = Screen.PrimaryScreen.WorkingArea;

            ApplyRoundedRegion(this, 10);

            // === PANEL RAÍZ ===
            var pnlRaiz = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = this.BackColor
            };
            this.Controls.Add(pnlRaiz);

            // === HEADER ===
            _pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = Color.FromArgb(31, 41, 55),
                Padding = new Padding(16, 0, 8, 0)
            };

            var lblTitulo = new Label
            {
                Text = "Inventory Schedule",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false
            };

            // Botones chrome
            var pnlChromeRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 140,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.FromArgb(31, 41, 55),
                Padding = new Padding(0, 12, 0, 0)
            };

            _btnCerrar = CrearBotonChrome("✕");
            _btnMaximizar = CrearBotonChrome("□");
            _btnMinimizar = CrearBotonChrome("—");

            _btnCerrar.Click += (s, e) => this.Close();
            _btnMaximizar.Click += (s, e) => AlternarMaximizado();
            _btnMinimizar.Click += (s, e) => this.WindowState = FormWindowState.Minimized;

            pnlChromeRight.Controls.Add(_btnCerrar);
            pnlChromeRight.Controls.Add(_btnMaximizar);
            pnlChromeRight.Controls.Add(_btnMinimizar);

            // Soporte drag
            _pnlHeader.MouseDown += Header_MouseDown;
            lblTitulo.MouseDown += Header_MouseDown;

            _pnlHeader.Controls.Add(pnlChromeRight);
            _pnlHeader.Controls.Add(lblTitulo);

            // === PANEL DE ALERTA (vencidas) ===
            _pnlAlerta = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(254, 243, 199),
                Padding = new Padding(16, 0, 8, 0),
                Visible = false
            };

            _lblAlerta = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(120, 53, 15),
                AutoSize = false
            };

            var btnDismiss = new Button
            {
                Text = "Dismiss",
                Dock = DockStyle.Right,
                Width = 80,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(254, 243, 199),
                ForeColor = Color.FromArgb(120, 53, 15),
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDismiss.FlatAppearance.BorderSize = 0;
            btnDismiss.Click += (s, e) => _pnlAlerta.Visible = false;

            _pnlAlerta.Controls.Add(_lblAlerta);
            _pnlAlerta.Controls.Add(btnDismiss);

            // === CUERPO PRINCIPAL (panel izquierdo + panel derecho) ===
            var pnlCuerpo = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(243, 244, 246),
                Padding = new Padding(10)
            };

            // --- Panel izquierdo (lista) ---
            var pnlIzquierdo = new Panel
            {
                Dock = DockStyle.Left,
                Width = 340,
                BackColor = Color.White,
                Padding = new Padding(8)
            };
            ApplyRoundedRegion(pnlIzquierdo, 6);

            // Botón nueva programación
            var btnNueva = new Button
            {
                Text = "+ New Schedule",
                Dock = DockStyle.Top,
                Height = 40,
                Margin = new Padding(0, 0, 0, 8)
            };
            StyleButtonPrimary(btnNueva);
            btnNueva.Click += BtnNueva_Click;

            // Separador
            var pnlSep = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = Color.FromArgb(229, 231, 235),
                Margin = new Padding(0, 0, 0, 8)
            };

            // Grid de lista
            _dgvLista = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false
            };
            StyleGrid(_dgvLista);
            ConfigurarColumnasLista();
            _dgvLista.CellClick += DgvLista_CellClick;
            _dgvLista.CellPainting += DgvLista_CellPainting;

            pnlIzquierdo.Controls.Add(_dgvLista);
            pnlIzquierdo.Controls.Add(pnlSep);
            pnlIzquierdo.Controls.Add(btnNueva);

            // --- Panel derecho (detalle + ubicaciones) ---
            var pnlDerecho = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(243, 244, 246),
                Padding = new Padding(10, 0, 0, 0)
            };

            // Sección de detalle
            var cardDetalle = new ModernCardPanel
            {
                Dock = DockStyle.Top,
                Height = 160,
                Padding = new Padding(16)
            };
            ConstruirPanelDetalle(cardDetalle);

            // Sección de ubicaciones
            var cardUbicaciones = new ModernCardPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16)
            };
            ConstruirArbolUbicaciones(cardUbicaciones);

            // Footer con botones de acción
            var pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 56,
                BackColor = Color.White,
                Padding = new Padding(0, 8, 0, 0)
            };

            var btnGuardar = new Button { Text = "💾  Save Schedule", Width = 160 };
            StyleButtonPrimary(btnGuardar);
            btnGuardar.Click += BtnGuardar_Click;
            btnGuardar.Dock = DockStyle.Left;

            var btnEliminar = new Button { Text = "🗑  Delete", Width = 120 };
            StyleButtonGhost(btnEliminar);
            btnEliminar.Click += BtnEliminar_Click;
            btnEliminar.Dock = DockStyle.Right;

            pnlFooter.Controls.Add(btnEliminar);
            pnlFooter.Controls.Add(btnGuardar);

            pnlDerecho.Controls.Add(pnlFooter);
            pnlDerecho.Controls.Add(cardUbicaciones);
            pnlDerecho.Controls.Add(cardDetalle);

            // Ensamblar cuerpo
            pnlCuerpo.Controls.Add(pnlDerecho);
            pnlCuerpo.Controls.Add(pnlIzquierdo);

            // Ensamblar raíz (orden inverso por Dock)
            pnlRaiz.Controls.Add(pnlCuerpo);
            pnlRaiz.Controls.Add(_pnlAlerta);
            pnlRaiz.Controls.Add(_pnlHeader);
        }

        // =====================================================================
        // Construcción del panel de detalle
        // =====================================================================
        private void ConstruirPanelDetalle(Panel padre)
        {
            var lblTitulo = new Label
            {
                Text = "Schedule Details",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Dock = DockStyle.Top,
                Height = 24,
                AutoSize = false
            };

            var tbl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2,
                BackColor = Color.White,
                Padding = new Padding(0, 8, 0, 0)
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));

            // --- Nombre ---
            tbl.Controls.Add(CrearLabel("Name:"), 0, 0);
            _txtNombre = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 8, 4) };
            StyleControl(_txtNombre);
            tbl.Controls.Add(_txtNombre, 1, 0);

            // --- Intervalo ---
            tbl.Controls.Add(CrearLabel("Interval (days):"), 2, 0);
            _nudIntervalo = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 9999,
                Value = 365,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 8, 4),
                Font = new Font("Segoe UI", 10F)
            };
            tbl.Controls.Add(_nudIntervalo, 3, 0);

            // --- Color ---
            tbl.Controls.Add(CrearLabel("Color:"), 0, 1);
            _cboColor = new ComboBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 8, 4),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            CargarOpcionesColor(_cboColor);
            StyleControl(_cboColor);
            tbl.Controls.Add(_cboColor, 1, 1);

            // --- Próximo inventario ---
            tbl.Controls.Add(CrearLabel("Next Inventory:"), 2, 1);
            _dtpProximo = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today.AddDays(365),
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 8, 4),
                Font = new Font("Segoe UI", 10F)
            };
            tbl.Controls.Add(_dtpProximo, 3, 1);

            padre.Controls.Add(tbl);
            padre.Controls.Add(lblTitulo);
        }

        // =====================================================================
        // Construcción del árbol de ubicaciones (3 columnas)
        // =====================================================================
        private void ConstruirArbolUbicaciones(Panel padre)
        {
            var lblTitulo = new Label
            {
                Text = "Locations",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Dock = DockStyle.Top,
                Height = 24,
                AutoSize = false
            };

            var tbl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.White,
                Padding = new Padding(0, 8, 0, 0)
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.4F));
            tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // Columna 1: Almacén
            _dgvAlmacen = new DataGridView();
            var colAlmacen = ConstruirColumnaUbicacion("Warehouse", _dgvAlmacen, false);
            _dgvAlmacen.CellClick += DgvAlmacen_CellClick;
            tbl.Controls.Add(colAlmacen, 0, 0);

            // Columna 2: Área
            _dgvArea = new DataGridView();
            var colArea = ConstruirColumnaUbicacion("Area", _dgvArea, true);
            _dgvArea.CellClick += DgvArea_CellClick;
            tbl.Controls.Add(colArea, 1, 0);

            // Columna 3: Locaciones
            _dgvLocacion = new DataGridView();
            var colLocacion = ConstruirColumnaUbicacion("Locations", _dgvLocacion, true);
            _dgvLocacion.CellClick += DgvLocacion_CellClick;
            tbl.Controls.Add(colLocacion, 2, 0);

            padre.Controls.Add(tbl);
            padre.Controls.Add(lblTitulo);
        }

        // Construye una columna del árbol de ubicaciones
        private Panel ConstruirColumnaUbicacion(string titulo, DataGridView dgv, bool conCheck)
        {
            var pnl = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(4)
            };

            // Encabezado con badge
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 32,
                BackColor = Color.FromArgb(243, 244, 246)
            };

            var lblTitulo = new Label
            {
                Text = titulo,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0)
            };

            var lblBadge = new Label
            {
                Text = "0",
                Dock = DockStyle.Right,
                Width = 40,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                ForeColor = Color.FromArgb(37, 99, 235),
                BackColor = Color.FromArgb(219, 234, 254)
            };
            // Guardar referencia al badge en el Tag del dgv
            dgv.Tag = lblBadge;

            pnlHeader.Controls.Add(lblTitulo);
            pnlHeader.Controls.Add(lblBadge);

            // Caja de búsqueda
            var txtBuscar = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 28,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(156, 163, 175),
                Text = "Search..."
            };
            txtBuscar.GotFocus += (s, e) => { if (txtBuscar.Text == "Search...") { txtBuscar.Text = ""; txtBuscar.ForeColor = Color.FromArgb(17, 24, 39); } };
            txtBuscar.LostFocus += (s, e) => { if (txtBuscar.Text == "") { txtBuscar.Text = "Search..."; txtBuscar.ForeColor = Color.FromArgb(156, 163, 175); } };
            txtBuscar.TextChanged += (s, e) => FiltrarGrid(dgv, txtBuscar.Text == "Search..." ? "" : txtBuscar.Text);

            // Grid
            ConfigurarGridUbicacion(dgv, conCheck);
            dgv.Dock = DockStyle.Fill;
            StyleGrid(dgv);

            pnl.Controls.Add(dgv);
            pnl.Controls.Add(txtBuscar);
            pnl.Controls.Add(pnlHeader);

            return pnl;
        }

        private void ConfigurarGridUbicacion(DataGridView dgv, bool conCheck)
        {
            dgv.Columns.Clear();
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.MultiSelect = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.EditMode = DataGridViewEditMode.EditOnEnter;
            dgv.ReadOnly = false;

            if (conCheck)
            {
                dgv.Columns.Add(new DataGridViewCheckBoxColumn
                {
                    Name = "Selected",
                    HeaderText = "",
                    DataPropertyName = "Selected",
                    Width = 32,
                    ReadOnly = false
                });
            }

            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Location_ID",
                HeaderText = "ID",
                DataPropertyName = "Location_ID",
                Visible = false,
                ReadOnly = true
            });

            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Location_Name",
                HeaderText = "Name",
                DataPropertyName = "Location_Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                ReadOnly = true
            });
        }

        private void ConfigurarColumnasLista()
        {
            _dgvLista.Columns.Clear();

            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Id",
                HeaderText = "ID",
                DataPropertyName = "Id",
                Visible = false
            });

            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Nombre",
                HeaderText = "Name",
                DataPropertyName = "Nombre",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Proximo_Inventario",
                HeaderText = "Next Date",
                DataPropertyName = "Proximo_Inventario",
                Width = 90
            });

            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Intervalo_Dias",
                HeaderText = "Interval",
                DataPropertyName = "Intervalo_Dias",
                Width = 60
            });

            // Columna de color (pill)
            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Color_Hex",
                HeaderText = "Color",
                DataPropertyName = "Color_Hex",
                Width = 60
            });

            // Columna oculta para vencido
            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Vencido",
                HeaderText = "Overdue",
                DataPropertyName = "Vencido",
                Visible = false
            });

            // Columna oculta Ultimo_PC_Id
            _dgvLista.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Ultimo_PC_Id",
                DataPropertyName = "Ultimo_PC_Id",
                Visible = false
            });
        }

        #endregion

        // =====================================================================
        // #region CARGA DE DATOS
        // =====================================================================
        #region CARGA DE DATOS

        private void CargarLista()
        {
            try
            {
                var dt = _repoConfig.CargarProgramaciones(_usuarioActual);
                _dgvLista.DataSource = dt;

                // Formatear fecha
                foreach (DataGridViewRow row in _dgvLista.Rows)
                {
                    if (row.DataBoundItem == null) continue;
                    var dr = ((DataRowView)row.DataBoundItem).Row;
                    if (dr["Proximo_Inventario"] != DBNull.Value)
                    {
                        DateTime fecha = Convert.ToDateTime(dr["Proximo_Inventario"]);
                        row.Cells["Proximo_Inventario"].Value = fecha.ToString("MM/dd/yyyy");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading schedules: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarAlmacenes()
        {
            try
            {
                var dt = _repoConfig.CargarAlmacenes();
                _dgvAlmacen.DataSource = dt;
                ActualizarBadge(_dgvAlmacen, dt.Rows.Count);

                // Enfocar automáticamente el primer almacén para cargar áreas
                if (_dgvAlmacen.Rows.Count > 0)
                {
                    _dgvAlmacen.ClearSelection();
                    _dgvAlmacen.Rows[0].Selected = true;
                    CargarAreas(_almacenFijoId);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading warehouses: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarAreas(int almacenId)
        {
            try
            {
                _checkedAreas.Clear();
                var dt = _repoConfig.CargarAreas(almacenId);

                // Restaurar checks si hay una programación seleccionada
                if (_idSeleccionado > 0)
                {
                    var locs = _repoConfig.CargarLocaciones(_idSeleccionado);
                    foreach (DataRow lr in locs.Rows)
                    {
                        if (lr["Area_Id"] != DBNull.Value)
                            _checkedAreas.Add(Convert.ToInt32(lr["Area_Id"]));
                    }
                }

                foreach (DataRow r in dt.Rows)
                {
                    int id = Convert.ToInt32(r["Location_ID"]);
                    r["Selected"] = _checkedAreas.Contains(id);
                }

                _dgvArea.DataSource = dt;
                ActualizarBadge(_dgvArea, dt.Rows.Count);

                // Cargar locaciones de las áreas marcadas
                if (_checkedAreas.Count > 0)
                    CargarLocaciones(new List<int>(_checkedAreas));
                else
                {
                    _dgvLocacion.DataSource = null;
                    ActualizarBadge(_dgvLocacion, 0);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading areas: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarLocaciones(List<int> areaIds)
        {
            try
            {
                var dt = _repoConfig.CargarLocacionesPorAreas(areaIds);

                // Restaurar checks si hay una programación seleccionada
                if (_idSeleccionado > 0)
                {
                    var locs = _repoConfig.CargarLocaciones(_idSeleccionado);
                    _checkedLocaciones.Clear();
                    foreach (DataRow lr in locs.Rows)
                    {
                        if (lr["Locacion_Id"] != DBNull.Value)
                            _checkedLocaciones.Add(Convert.ToInt32(lr["Locacion_Id"]));
                    }
                }

                foreach (DataRow r in dt.Rows)
                {
                    int id = Convert.ToInt32(r["Location_ID"]);
                    r["Selected"] = _checkedLocaciones.Contains(id);
                }

                _dgvLocacion.DataSource = dt;
                ActualizarBadge(_dgvLocacion, dt.Rows.Count);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading locations: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CargarProgramacionEnFormulario(int id)
        {
            try
            {
                var row = _repoConfig.CargarProgramacion(id);
                if (row == null) return;

                _idSeleccionado = id;

                _txtNombre.Text = Convert.ToString(row["Nombre"]);
                _nudIntervalo.Value = Convert.ToDecimal(row["Intervalo_Dias"]);

                string colorHex = Convert.ToString(row["Color_Hex"]);
                SeleccionarColorEnCombo(colorHex);

                if (row["Proximo_Inventario"] != DBNull.Value)
                    _dtpProximo.Value = Convert.ToDateTime(row["Proximo_Inventario"]);

                // Si ya tiene un último inventario, la fecha próxima no es editable:
                // se calcula automáticamente como FechaUltimo + Intervalo_Dias al crear inventario
                bool tienePcAnterior = row["Ultimo_PC_Id"] != DBNull.Value;
                _dtpProximo.Enabled = !tienePcAnterior;

                // Recargar árbol con las locaciones de esta programación
                CargarAlmacenes();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading schedule: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MostrarAlertaSiHayVencidas()
        {
            try
            {
                var dt = _repoConfig.ObtenerVencidas(_usuarioActual);
                int n = dt.Rows.Count;

                if (n > 0)
                {
                    _lblAlerta.Text = string.Format("⚠  You have {0} scheduled inventory(ies) due. Select one and click Save.", n);
                    _pnlAlerta.Visible = true;
                }
                else
                {
                    _pnlAlerta.Visible = false;
                }

                // Marcar filas vencidas en la lista
                foreach (DataGridViewRow row in _dgvLista.Rows)
                {
                    if (row.DataBoundItem == null) continue;
                    var dr = ((DataRowView)row.DataBoundItem).Row;
                    bool vencido = dr["Vencido"] != DBNull.Value && Convert.ToInt32(dr["Vencido"]) == 1;
                    if (vencido)
                        row.DefaultCellStyle.BackColor = Color.FromArgb(254, 226, 226);
                }
            }
            catch { /* No interrumpir la carga si falla la alerta */ }
        }

        #endregion

        // =====================================================================
        // #region EVENTOS DE CONTROLES
        // =====================================================================
        #region EVENTOS DE CONTROLES

        private void BtnNueva_Click(object sender, EventArgs e)
        {
            // Limpiar formulario para nueva programación
            _idSeleccionado = 0;
            _txtNombre.Text = GenerarNombreAuto();
            _nudIntervalo.Value = 365;
            if (_cboColor.Items.Count > 0) _cboColor.SelectedIndex = 1; // Blue por defecto
            _dtpProximo.Value = DateTime.Today.AddDays(365);
            _dtpProximo.Enabled = true;

            _checkedAreas.Clear();
            _checkedLocaciones.Clear();
            CargarAlmacenes();

            _dgvLista.ClearSelection();
        }

        private void DgvLista_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = _dgvLista.Rows[e.RowIndex];
            if (row.DataBoundItem == null) return;

            int id = Convert.ToInt32(row.Cells["Id"].Value);
            if (id > 0)
                CargarProgramacionEnFormulario(id);
        }

        private void DgvLista_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // Pintar pill de color en columna Color_Hex
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_dgvLista.Columns[e.ColumnIndex].Name != "Color_Hex") return;

            e.Handled = true;
            e.PaintBackground(e.ClipBounds, true);

            string hex = Convert.ToString(e.Value);
            if (string.IsNullOrEmpty(hex)) return;

            try
            {
                Color c = ColorDesdeHex(hex);
                var rect = new Rectangle(e.CellBounds.X + 8, e.CellBounds.Y + 8, 16, 16);
                using (var brush = new SolidBrush(c))
                    e.Graphics.FillEllipse(brush, rect);
            }
            catch { }
        }

        private void DgvAlmacen_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            // Solo cargar áreas (el almacén es fijo, no hay check)
            CargarAreas(_almacenFijoId);
        }

        private void DgvArea_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!_dgvArea.Columns.Contains("Selected")) return;
            if (e.ColumnIndex != _dgvArea.Columns["Selected"].Index) return;

            var row = _dgvArea.Rows[e.RowIndex];
            int id = Convert.ToInt32(row.Cells["Location_ID"].Value);
            if (id <= 0) return;

            bool actual = Convert.ToBoolean(row.Cells["Selected"].Value);
            bool siguiente = !actual;
            row.Cells["Selected"].Value = siguiente;

            if (siguiente) _checkedAreas.Add(id);
            else _checkedAreas.Remove(id);

            // Recargar locaciones de las áreas marcadas
            if (_checkedAreas.Count > 0)
                CargarLocaciones(new List<int>(_checkedAreas));
            else
            {
                _dgvLocacion.DataSource = null;
                ActualizarBadge(_dgvLocacion, 0);
                _checkedLocaciones.Clear();
            }
        }

        private void DgvLocacion_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!_dgvLocacion.Columns.Contains("Selected")) return;
            if (e.ColumnIndex != _dgvLocacion.Columns["Selected"].Index) return;

            var row = _dgvLocacion.Rows[e.RowIndex];
            int id = Convert.ToInt32(row.Cells["Location_ID"].Value);
            if (id <= 0) return;

            bool actual = Convert.ToBoolean(row.Cells["Selected"].Value);
            bool siguiente = !actual;
            row.Cells["Selected"].Value = siguiente;

            if (siguiente) _checkedLocaciones.Add(id);
            else _checkedLocaciones.Remove(id);
        }

        private void BtnGuardar_Click(object sender, EventArgs e)
        {
            GuardarProgramacion();
        }

        private void BtnEliminar_Click(object sender, EventArgs e)
        {
            if (_idSeleccionado <= 0)
            {
                MessageBox.Show("Please select a schedule to delete.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var res = MessageBox.Show(
                "Are you sure you want to delete this schedule?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (res != DialogResult.Yes) return;

            string error;
            bool ok = _repoConfig.Eliminar(_idSeleccionado, out error);

            if (!ok)
            {
                MessageBox.Show("Error deleting schedule: " + error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("Schedule deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _idSeleccionado = 0;
            CargarLista();
            LimpiarFormulario();
        }

        #endregion

        // =====================================================================
        // #region LÓGICA DE NEGOCIO
        // =====================================================================
        #region LÓGICA DE NEGOCIO

        private void GuardarProgramacion()
        {
            // 1. Validar campos
            string nombre = (_txtNombre.Text ?? "").Trim();
            if (nombre.Length == 0)
            {
                MessageBox.Show("Please enter a name for the schedule.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtNombre.Focus();
                return;
            }

            int intervalo = (int)_nudIntervalo.Value;
            if (intervalo <= 0)
            {
                MessageBox.Show("Interval must be greater than 0.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _nudIntervalo.Focus();
                return;
            }

            // Verificar que haya al menos una locación
            var locaciones = ObtenerLocacionesSeleccionadas();
            if (locaciones.Count == 0)
            {
                MessageBox.Show("Please select at least one location.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Construir DTO
            string colorHex = ObtenerColorSeleccionado();
            var dto = new PC_ConfigDto
            {
                Id = _idSeleccionado,
                Nombre = nombre,
                IntervaloDias = intervalo,
                ColorHex = colorHex,
                FechaBase = DateTime.Today,
                ProximoInventario = _dtpProximo.Value.Date,
                CreadoPor = _usuarioActual,
                Locaciones = locaciones
            };

            // 3. Guardar
            int nuevoId;
            string error;
            bool ok = _repoConfig.Guardar(dto, out nuevoId, out error);

            if (!ok)
            {
                MessageBox.Show("Error saving schedule: " + error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int idGuardado = (_idSeleccionado > 0) ? _idSeleccionado : nuevoId;

            // 4. Verificar si la fecha ya venció o es hoy
            if (dto.ProximoInventario <= DateTime.Today)
            {
                var dlgRes = MessageBox.Show(
                    "The next inventory date is due. What would you like to do?\n\n" +
                    "• Yes — Create inventory now\n" +
                    "• No  — Postpone\n" +
                    "• Cancel — Do nothing",
                    "Inventory Due",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (dlgRes == DialogResult.Yes)
                {
                    int pcId;
                    string errInv;
                    bool creado = _repoConfig.CrearInventarioAutomatico(idGuardado, _usuarioActual, out pcId, out errInv);

                    if (!creado)
                    {
                        MessageBox.Show("Error creating inventory: " + errInv, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        MessageBox.Show("Inventory created successfully. Opening wizard...", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        try
                        {
                            // Se abre en el paso 1 para que el usuario designe el responsable
                            const int PASO_INICIAL_WIZARD = 1;
                            var wizard = new PhysicalCountCyclicForm(_connectionString, pcId, PASO_INICIAL_WIZARD);
                            wizard.ShowDialog(this);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error opening wizard: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else if (dlgRes == DialogResult.No)
                {
                    _repoConfig.RegistrarLog(idGuardado, 2, _usuarioActual, null);
                    MessageBox.Show("Schedule postponed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                // Cancel → no hacer nada
            }
            else
            {
                MessageBox.Show("Schedule saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // 5. Recargar lista
            _idSeleccionado = idGuardado;
            CargarLista();
            MostrarAlertaSiHayVencidas();
        }

        private void LimpiarFormulario()
        {
            _txtNombre.Text = "";
            _nudIntervalo.Value = 365;
            if (_cboColor.Items.Count > 0) _cboColor.SelectedIndex = 1;
            _dtpProximo.Value = DateTime.Today.AddDays(365);
            _dtpProximo.Enabled = true;
            _checkedAreas.Clear();
            _checkedLocaciones.Clear();
            CargarAlmacenes();
        }

        #endregion

        // =====================================================================
        // #region HELPERS
        // =====================================================================
        #region HELPERS

        private int ObtenerUsuarioActual()
        {
            try { return Convert.ToInt32(Persal.GlobalVars.User_ID); }
            catch { return 0; }
        }

        private string ObtenerConnectionString()
        {
            if (!string.IsNullOrEmpty(_connectionString)) return _connectionString;
            try { return new con().ConnectionString; }
            catch { return ""; }
        }

        private string GenerarNombreAuto()
        {
            return string.Format("INV-{0}-{1}", "SUPPLY", DateTime.Today.Year);
        }

        private void CargarOpcionesColor(ComboBox cbo)
        {
            cbo.Items.Clear();
            cbo.Items.Add(new ColorItem("Red", "#EF4444"));
            cbo.Items.Add(new ColorItem("Blue", "#3B82F6"));
            cbo.Items.Add(new ColorItem("Green", "#10B981"));
            cbo.Items.Add(new ColorItem("Orange", "#F97316"));
            cbo.Items.Add(new ColorItem("Purple", "#8B5CF6"));
            cbo.Items.Add(new ColorItem("Gray", "#6B7280"));
            cbo.DisplayMember = "Nombre";
            cbo.ValueMember = "Hex";
            if (cbo.Items.Count > 1) cbo.SelectedIndex = 1; // Blue por defecto
        }

        private void SeleccionarColorEnCombo(string hex)
        {
            for (int i = 0; i < _cboColor.Items.Count; i++)
            {
                var item = _cboColor.Items[i] as ColorItem;
                if (item != null && item.Hex.Equals(hex, StringComparison.OrdinalIgnoreCase))
                {
                    _cboColor.SelectedIndex = i;
                    return;
                }
            }
        }

        private string ObtenerColorSeleccionado()
        {
            var item = _cboColor.SelectedItem as ColorItem;
            return item != null ? item.Hex : "#3B82F6";
        }

        private List<LocacionDto> ObtenerLocacionesSeleccionadas()
        {
            var lista = new List<LocacionDto>();

            if (_checkedLocaciones.Count > 0)
            {
                // Locaciones específicas seleccionadas
                foreach (int locId in _checkedLocaciones)
                {
                    // Determinar a qué área pertenece
                    int areaId = ObtenerAreaDeLocacion(locId);
                    lista.Add(new LocacionDto
                    {
                        AlmacenId = _almacenFijoId,
                        AreaId = areaId > 0 ? areaId : (int?)null,
                        LocacionId = locId
                    });
                }
            }
            else if (_checkedAreas.Count > 0)
            {
                // Áreas completas seleccionadas
                foreach (int areaId in _checkedAreas)
                {
                    lista.Add(new LocacionDto
                    {
                        AlmacenId = _almacenFijoId,
                        AreaId = areaId,
                        LocacionId = null
                    });
                }
            }

            return lista;
        }

        private int ObtenerAreaDeLocacion(int locacionId)
        {
            // Siempre consultar la BD para obtener el área padre real de la locación,
            // evitando suposiciones incorrectas cuando hay múltiples áreas marcadas
            try
            {
                string sql = "SELECT Parent_Location_ID FROM dbo.Locations WHERE Location_ID = @p1";
                using (var conn = new SqlConnection(ObtenerConnectionString()))
                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@p1", locacionId);
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o);
                }
            }
            catch { return 0; }
        }

        private void ActualizarBadge(DataGridView dgv, int count)
        {
            var badge = dgv.Tag as Label;
            if (badge != null) badge.Text = count.ToString();
        }

        private void FiltrarGrid(DataGridView dgv, string texto)
        {
            if (dgv.DataSource == null) return;
            try
            {
                var dt = dgv.DataSource as DataTable;
                if (dt == null) return;

                if (string.IsNullOrEmpty(texto))
                {
                    dt.DefaultView.RowFilter = "";
                }
                else
                {
                    // Escapar caracteres especiales para el filtro de DataView:
                    // se escapan comilla simple, comodines [ ] % y el comodín de un carácter _
                    string textoBusqueda = texto
                        .Replace("'", "''")
                        .Replace("[", "[[]")
                        .Replace("]", "[]]")
                        .Replace("%", "[%]")
                        .Replace("_", "[_]");
                    dt.DefaultView.RowFilter = string.Format("Location_Name LIKE '%{0}%'", textoBusqueda);
                }
            }
            catch { /* No interrumpir la UI si falla el filtro */ }
        }

        private Color ColorDesdeHex(string hex)
        {
            hex = hex.TrimStart('#');
            if (hex.Length == 6)
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                return Color.FromArgb(r, g, b);
            }
            return Color.FromArgb(59, 130, 246); // Blue por defecto
        }

        private Label CrearLabel(string texto)
        {
            return new Label
            {
                Text = texto,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(55, 65, 81),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                AutoSize = false
            };
        }

        /// <summary>Clase interna para ítems del combo de color.</summary>
        private class ColorItem
        {
            public string Nombre { get; set; }
            public string Hex { get; set; }
            public ColorItem(string nombre, string hex) { Nombre = nombre; Hex = hex; }
            public override string ToString() { return Nombre; }
        }

        #endregion

        // =====================================================================
        // #region ESTILOS VISUALES (mismos que PhysicalCountCyclicForm)
        // =====================================================================
        #region ESTILOS VISUALES

        private Button CrearBotonChrome(string texto)
        {
            var b = new Button
            {
                Text = texto,
                Width = 36,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(31, 41, 55),
                Cursor = Cursors.Hand,
                Margin = new Padding(4, 0, 0, 0),
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
            dgv.ColumnHeadersHeight = 34;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(17, 24, 39);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgv.DefaultCellStyle.SelectionForeColor = Color.FromArgb(17, 24, 39);
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 9F);
            dgv.RowTemplate.Height = 30;
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

        private void AlternarMaximizado()
        {
            if (this.WindowState == FormWindowState.Maximized)
                this.WindowState = FormWindowState.Normal;
            else
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Maximized;
            }
            if (_btnMaximizar != null)
                _btnMaximizar.Text = (this.WindowState == FormWindowState.Maximized) ? "❐" : "□";
        }

        #endregion

        // =====================================================================
        // #region WINDOW CHROME (drag / resize)
        // =====================================================================
        #region WINDOW CHROME

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized) return;
            if (e.Button != MouseButtons.Left) return;
            ReleaseCapture();
            SendMessage(this.Handle, 0xA1, 0x2, 0);
        }

        #endregion

        // =====================================================================
        // #region CLASE INTERNA ModernCardPanel
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

    } // end class Physical_Count_Config

    // =========================================================================
    // DTO interno para la programación de inventario
    // =========================================================================
    internal class PC_ConfigDto
    {
        /// <summary>0 = nueva programación</summary>
        public int Id { get; set; }
        public string Nombre { get; set; }
        public int IntervaloDias { get; set; }
        public string ColorHex { get; set; }
        public DateTime FechaBase { get; set; }
        public DateTime ProximoInventario { get; set; }
        public int? UltimoPcId { get; set; }
        public int CreadoPor { get; set; }
        public List<LocacionDto> Locaciones { get; set; }
    }

    // =========================================================================
    // DTO de locación
    // =========================================================================
    internal class LocacionDto
    {
        public int AlmacenId { get; set; }
        public int? AreaId { get; set; }
        public int? LocacionId { get; set; }
    }

    // =========================================================================
    // PC_ConfigRepository — acceso a datos de programaciones
    // =========================================================================
    internal sealed class PC_ConfigRepository
    {
        private readonly string _cs;

        // ID del almacén raíz fijo (WAREHOUSE SUPPLIES).
        // Se define localmente para que el repositorio no dependa de la clase form.
        private const int DEFAULT_ROOT_WAREHOUSE_ID = 3;

        public PC_ConfigRepository(string connectionString)
        {
            _cs = connectionString ?? "";
        }

        // =====================================================================
        // Obtiene la cadena de conexión (fallback al helper global)
        // =====================================================================
        private string ObtenerCs()
        {
            if (!string.IsNullOrEmpty(_cs)) return _cs;
            try { return new con().ConnectionString; }
            catch { return ""; }
        }

        // =====================================================================
        // Carga todas las programaciones activas del usuario
        // =====================================================================
        public DataTable CargarProgramaciones(int creadoPor)
        {
            const string sql = @"
                SELECT c.Id, c.Nombre, c.Intervalo_Dias, c.Color_Hex,
                       c.Fecha_Base, c.Proximo_Inventario, c.Ultimo_PC_Id,
                       c.Creado_Por, c.Creado_En,
                       -- Comparación con DATE para evitar diferencias de hora (Proximo_Inventario es DATE)
                       CASE WHEN c.Proximo_Inventario <= CAST(GETDATE() AS DATE) THEN 1 ELSE 0 END AS Vencido
                FROM dbo.PC_Config c
                WHERE c.Activo = 1 AND c.Creado_Por = @p1
                ORDER BY c.Proximo_Inventario ASC";

            return EjecutarTabla(sql, creadoPor);
        }

        // =====================================================================
        // Carga una programación específica por Id
        // =====================================================================
        public DataRow CargarProgramacion(int id)
        {
            const string sql = @"
                SELECT Id, Nombre, Intervalo_Dias, Color_Hex,
                       Fecha_Base, Proximo_Inventario, Ultimo_PC_Id,
                       Creado_Por, Creado_En
                FROM dbo.PC_Config
                WHERE Id = @p1 AND Activo = 1";

            var dt = EjecutarTabla(sql, id);
            return (dt.Rows.Count > 0) ? dt.Rows[0] : null;
        }

        // =====================================================================
        // Carga las locaciones asignadas a una programación
        // =====================================================================
        public DataTable CargarLocaciones(int pcConfigId)
        {
            const string sql = @"
                SELECT Id, PC_Config_Id, Almacen_Id, Area_Id, Locacion_Id
                FROM dbo.PC_Config_Locaciones
                WHERE PC_Config_Id = @p1";

            return EjecutarTabla(sql, pcConfigId);
        }

        // =====================================================================
        // Carga almacenes (fijo ID=3 — WAREHOUSE SUPPLIES)
        // =====================================================================
        public DataTable CargarAlmacenes()
        {
            const string sql = @"
                SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected
                FROM dbo.Locations
                WHERE Location_ID = 3";

            return EjecutarTabla(sql);
        }

        // =====================================================================
        // Carga áreas hijas del almacén dado
        // =====================================================================
        public DataTable CargarAreas(int almacenId)
        {
            const string sql = @"
                SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected
                FROM dbo.Locations
                WHERE Parent_Location_ID = @p1
                ORDER BY Location_Name";

            return EjecutarTabla(sql, almacenId);
        }

        // =====================================================================
        // Carga locaciones hijas de varias áreas
        // =====================================================================
        public DataTable CargarLocacionesPorAreas(List<int> areaIds)
        {
            if (areaIds == null || areaIds.Count == 0) return new DataTable();

            // Construir parámetros dinámicamente
            var paramNames = new List<string>();
            for (int i = 0; i < areaIds.Count; i++)
                paramNames.Add("@a" + i);

            string sql = string.Format(@"
                SELECT Location_ID, Location_Name, CAST(0 AS bit) AS Selected
                FROM dbo.Locations
                WHERE Parent_Location_ID IN ({0})
                ORDER BY Location_Name",
                string.Join(",", paramNames));

            var dt = new DataTable();
            using (var conn = new SqlConnection(ObtenerCs()))
            using (var cmd = new SqlCommand(sql, conn))
            {
                for (int i = 0; i < areaIds.Count; i++)
                    cmd.Parameters.AddWithValue("@a" + i, areaIds[i]);

                using (var da = new SqlDataAdapter(cmd))
                    da.Fill(dt);
            }
            return dt;
        }

        // =====================================================================
        // Cuenta hijos de una locación
        // =====================================================================
        public int ContarHijos(int locationId)
        {
            const string sql = "SELECT COUNT(*) FROM dbo.Locations WHERE Parent_Location_ID = @p1";
            return EjecutarEscalar(sql, locationId);
        }

        // =====================================================================
        // INSERT o UPDATE según dto.Id == 0
        // =====================================================================
        public bool Guardar(PC_ConfigDto dto, out int nuevoId, out string error)
        {
            nuevoId = 0;
            error = "";

            try
            {
                using (var conn = new SqlConnection(ObtenerCs()))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        try
                        {
                            if (dto.Id == 0)
                            {
                                // INSERT
                                const string sqlInsert = @"
                                    INSERT INTO dbo.PC_Config
                                        (Nombre, Intervalo_Dias, Color_Hex, Fecha_Base, Proximo_Inventario,
                                         Ultimo_PC_Id, Activo, Creado_Por, Creado_En)
                                    VALUES (@p1,@p2,@p3,@p4,@p5,NULL,1,@p6,GETDATE());
                                    SELECT SCOPE_IDENTITY();";

                                using (var cmd = new SqlCommand(sqlInsert, conn, tx))
                                {
                                    cmd.Parameters.AddWithValue("@p1", dto.Nombre);
                                    cmd.Parameters.AddWithValue("@p2", dto.IntervaloDias);
                                    cmd.Parameters.AddWithValue("@p3", dto.ColorHex);
                                    cmd.Parameters.AddWithValue("@p4", dto.FechaBase);
                                    cmd.Parameters.AddWithValue("@p5", dto.ProximoInventario);
                                    cmd.Parameters.AddWithValue("@p6", dto.CreadoPor);

                                    object o = cmd.ExecuteScalar();
                                    nuevoId = (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o);
                                }

                                dto.Id = nuevoId;
                            }
                            else
                            {
                                // UPDATE
                                const string sqlUpdate = @"
                                    UPDATE dbo.PC_Config
                                    SET Nombre=@p1, Intervalo_Dias=@p2, Color_Hex=@p3,
                                        Proximo_Inventario=@p4
                                    WHERE Id=@p5 AND Creado_Por=@p6";

                                using (var cmd = new SqlCommand(sqlUpdate, conn, tx))
                                {
                                    cmd.Parameters.AddWithValue("@p1", dto.Nombre);
                                    cmd.Parameters.AddWithValue("@p2", dto.IntervaloDias);
                                    cmd.Parameters.AddWithValue("@p3", dto.ColorHex);
                                    cmd.Parameters.AddWithValue("@p4", dto.ProximoInventario);
                                    cmd.Parameters.AddWithValue("@p5", dto.Id);
                                    cmd.Parameters.AddWithValue("@p6", dto.CreadoPor);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            // Guardar locaciones: borrar anteriores e insertar nuevas
                            GuardarLocaciones(conn, tx, dto.Id, dto.Locaciones);

                            // Registrar log de creación
                            RegistrarLogInterno(conn, tx, dto.Id, 1, dto.CreadoPor, null);

                            tx.Commit();
                            return true;
                        }
                        catch (Exception ex)
                        {
                            tx.Rollback();
                            error = ex.Message;
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // =====================================================================
        // Guarda las locaciones de una programación (borra y vuelve a insertar)
        // =====================================================================
        private void GuardarLocaciones(SqlConnection conn, SqlTransaction tx, int pcConfigId, List<LocacionDto> locaciones)
        {
            // Borrar locaciones anteriores
            using (var cmd = new SqlCommand("DELETE FROM dbo.PC_Config_Locaciones WHERE PC_Config_Id = @p1", conn, tx))
            {
                cmd.Parameters.AddWithValue("@p1", pcConfigId);
                cmd.ExecuteNonQuery();
            }

            if (locaciones == null || locaciones.Count == 0) return;

            // Insertar nuevas locaciones
            const string sqlInsert = @"
                INSERT INTO dbo.PC_Config_Locaciones (PC_Config_Id, Almacen_Id, Area_Id, Locacion_Id)
                VALUES (@p1, @p2, @p3, @p4)";

            foreach (var loc in locaciones)
            {
                using (var cmd = new SqlCommand(sqlInsert, conn, tx))
                {
                    cmd.Parameters.AddWithValue("@p1", pcConfigId);
                    cmd.Parameters.AddWithValue("@p2", loc.AlmacenId);
                    cmd.Parameters.AddWithValue("@p3", loc.AreaId.HasValue ? (object)loc.AreaId.Value : DBNull.Value);
                    cmd.Parameters.AddWithValue("@p4", loc.LocacionId.HasValue ? (object)loc.LocacionId.Value : DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // =====================================================================
        // Soft delete (Activo = 0)
        // =====================================================================
        public bool Eliminar(int id, out string error)
        {
            error = "";
            try
            {
                const string sql = "UPDATE dbo.PC_Config SET Activo = 0 WHERE Id = @p1";
                EjecutarSinResultado(sql, id);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // =====================================================================
        // Registra una acción en el log
        // =====================================================================
        public void RegistrarLog(int pcConfigId, byte accion, int usuarioId, int? pcIdGenerado)
        {
            try
            {
                const string sql = @"
                    INSERT INTO dbo.PC_Config_Log (PC_Config_Id, Accion, Usuario_Id, Fecha_Accion, PC_Id_Generado)
                    VALUES (@p1, @p2, @p3, GETDATE(), @p4)";

                using (var conn = new SqlConnection(ObtenerCs()))
                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@p1", pcConfigId);
                    cmd.Parameters.AddWithValue("@p2", accion);
                    cmd.Parameters.AddWithValue("@p3", usuarioId);
                    cmd.Parameters.AddWithValue("@p4", pcIdGenerado.HasValue ? (object)pcIdGenerado.Value : DBNull.Value);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch { /* No interrumpir flujo principal si falla el log */ }
        }

        private void RegistrarLogInterno(SqlConnection conn, SqlTransaction tx, int pcConfigId, byte accion, int usuarioId, int? pcIdGenerado)
        {
            const string sql = @"
                INSERT INTO dbo.PC_Config_Log (PC_Config_Id, Accion, Usuario_Id, Fecha_Accion, PC_Id_Generado)
                VALUES (@p1, @p2, @p3, GETDATE(), @p4)";

            using (var cmd = new SqlCommand(sql, conn, tx))
            {
                cmd.Parameters.AddWithValue("@p1", pcConfigId);
                cmd.Parameters.AddWithValue("@p2", accion);
                cmd.Parameters.AddWithValue("@p3", usuarioId);
                cmd.Parameters.AddWithValue("@p4", pcIdGenerado.HasValue ? (object)pcIdGenerado.Value : DBNull.Value);
                cmd.ExecuteNonQuery();
            }
        }

        // =====================================================================
        // Verifica programaciones vencidas del usuario
        // =====================================================================
        public DataTable ObtenerVencidas(int usuarioId)
        {
            const string sql = @"
                SELECT Id, Nombre, Proximo_Inventario
                FROM dbo.PC_Config
                WHERE Activo = 1 AND Creado_Por = @p1
                  -- Comparación con DATE para evitar diferencias de hora (Proximo_Inventario es DATE)
                  AND Proximo_Inventario <= CAST(GETDATE() AS DATE)
                ORDER BY Proximo_Inventario ASC";

            return EjecutarTabla(sql, usuarioId);
        }

        // =====================================================================
        // Crea el inventario automáticamente a partir de una programación
        // =====================================================================
        public bool CrearInventarioAutomatico(int pcConfigId, int usuarioId, out int pcId, out string error)
        {
            pcId = 0;
            error = "";

            try
            {
                // 1. Leer la programación
                var row = CargarProgramacion(pcConfigId);
                if (row == null) { error = "Schedule not found."; return false; }

                string nombre = Convert.ToString(row["Nombre"]);
                int intervaloDias = Convert.ToInt32(row["Intervalo_Dias"]);

                // 2. Crear el encabezado del inventario usando Save_Physical_Count
                var sf = new Persal.System_Functions();
                int warehouseId = sf.Return_Default_Warehouse_Location_ID_For_Count();
                // Fallback al ID del almacén raíz (WAREHOUSE SUPPLIES = 3) cuando el sistema
                // no puede determinar el almacén por defecto
                if (warehouseId <= 0) warehouseId = DEFAULT_ROOT_WAREHOUSE_ID;

                // Fechas del inventario
                DateTime fechaInicio = DateTime.Today;
                DateTime fechaFin = DateTime.Today.AddDays(intervaloDias);

                // Guardar encabezado
                bool guardado = sf.Save_Physical_Count(
                    0,              // nuevo inventario
                    1,              // countTypeId (cíclico)
                    1,              // listTypeId
                    usuarioId,      // responsable
                    usuarioId,      // verificador
                    1,              // número de chequeos
                    nombre,         // descripción
                    fechaInicio,
                    fechaFin,
                    100.0M,         // meta
                    warehouseId
                );

                if (!guardado)
                {
                    error = sf.G.Error_Message;
                    return false;
                }

                // 3. Recuperar el ID generado — se agrega filtro de tiempo (últimos 2 minutos)
                //    para reducir el riesgo en escenarios concurrentes donde múltiples
                //    usuarios crean inventarios al mismo tiempo
                using (var conn = new SqlConnection(ObtenerCs()))
                using (var cmd = new SqlCommand(
                    "SELECT TOP 1 Physical_Count_ID FROM Physical_Counts " +
                    "WHERE Created_By_User_ID = @p1 AND Physical_Count_DateTime >= DATEADD(minute, -2, GETDATE()) " +
                    "ORDER BY Physical_Count_ID DESC",
                    conn))
                {
                    cmd.Parameters.AddWithValue("@p1", usuarioId);
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    pcId = (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o);
                }

                if (pcId <= 0) { error = "Could not retrieve generated inventory ID."; return false; }

                // 4. Guardar locaciones usando sp_PC_SaveSelectedLocationRow
                var locaciones = CargarLocaciones(pcConfigId);
                using (var conn = new SqlConnection(ObtenerCs()))
                {
                    conn.Open();
                    foreach (DataRow loc in locaciones.Rows)
                    {
                        // Determinar el ID de locación efectivo según la jerarquía disponible
                        int locId = ObtenerLocacionIdEfectivo(loc);

                        using (var cmd = new SqlCommand("dbo.sp_PC_SaveSelectedLocationRow", conn))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@Physical_Count_ID", pcId);
                            cmd.Parameters.AddWithValue("@Location_ID", locId);
                            cmd.Parameters.AddWithValue("@User_ID", usuarioId);
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                // 5. Actualizar Ultimo_PC_Id y Proximo_Inventario en la programación
                // Se usa CAST(GETDATE() AS DATE) para trabajar con fechas sin hora
                using (var conn = new SqlConnection(ObtenerCs()))
                using (var cmd = new SqlCommand(
                    "UPDATE dbo.PC_Config SET Ultimo_PC_Id = @p1, Proximo_Inventario = DATEADD(day, Intervalo_Dias, CAST(GETDATE() AS DATE)) WHERE Id = @p2",
                    conn))
                {
                    cmd.Parameters.AddWithValue("@p1", pcId);
                    cmd.Parameters.AddWithValue("@p2", pcConfigId);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }

                // 6. Registrar log con Accion=1 (Creado)
                RegistrarLog(pcConfigId, 1, usuarioId, pcId);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // =====================================================================
        // Helpers internos
        // =====================================================================

        /// <summary>
        /// Determina el ID de locación efectivo de una fila de PC_Config_Locaciones.
        /// Prioridad: Locacion_Id > Area_Id > Almacen_Id
        /// </summary>
        private static int ObtenerLocacionIdEfectivo(DataRow loc)
        {
            if (loc["Locacion_Id"] != DBNull.Value)
                return Convert.ToInt32(loc["Locacion_Id"]);

            if (loc["Area_Id"] != DBNull.Value)
                return Convert.ToInt32(loc["Area_Id"]);

            return Convert.ToInt32(loc["Almacen_Id"]);
        }

        private DataTable EjecutarTabla(string sql, params object[] args)
        {
            var dt = new DataTable();
            using (var conn = new SqlConnection(ObtenerCs()))
            using (var cmd = conn.CreateCommand())
            using (var da = new SqlDataAdapter(cmd))
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                da.Fill(dt);
            }
            return dt;
        }

        private int EjecutarEscalar(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(ObtenerCs()))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                conn.Open();
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o);
            }
        }

        private void EjecutarSinResultado(string sql, params object[] args)
        {
            using (var conn = new SqlConnection(ObtenerCs()))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i < args.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + (i + 1), args[i] ?? DBNull.Value);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

    } // end class PC_ConfigRepository

} // end namespace Persal._003_Physical_Counting_2
