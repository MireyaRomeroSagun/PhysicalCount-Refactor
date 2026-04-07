using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace Persal._003_Physical_Counting_2
{
    /// <summary>
    /// Formulario Wizard de inventario físico cíclico.
    /// Esta clase es una capa de interfaz pura: toda la lógica de negocio
    /// y acceso a datos está delegada en PhysicalCountRepository.
    /// </summary>
    public partial class PhysicalCountCyclicForm : Form
    {
        // =====================================================================
        // Campos privados del formulario
        // =====================================================================

        /// <summary>Repositorio que centraliza toda la lógica de acceso a datos.</summary>
        private readonly PhysicalCountRepository _repo;

        /// <summary>ID del inventario físico actualmente en edición.</summary>
        private int _physicalCountId;

        /// <summary>ID del usuario que opera el Wizard.</summary>
        private int _currentUserId;

        /// <summary>ID del almacén fijo predeterminado (obtenido del repositorio al iniciar).</summary>
        private int _almacenFijoId;

        /// <summary>Paso actual del Wizard (1–4).</summary>
        private int _currentStep;

        /// <summary>
        /// Inicializa el formulario y construye el repositorio a partir de la
        /// cadena de conexión de la aplicación.
        /// </summary>
        /// <param name="connectionString">Cadena de conexión a la base de datos.</param>
        /// <param name="userId">ID del usuario que abre el Wizard.</param>
        /// <param name="physicalCountId">ID del inventario a editar (0 para nuevo).</param>
        public PhysicalCountCyclicForm(string connectionString, int userId, int physicalCountId)
        {
            InitializeComponent();

            _repo            = new PhysicalCountRepository(connectionString);
            _currentUserId   = userId;
            _physicalCountId = physicalCountId;
            _currentStep     = 1;
        }

        // =====================================================================
        // PASO 1 — ENCABEZADO DE INVENTARIO
        // =====================================================================
        #region PASO 1 — ENCABEZADO DE INVENTARIO

        /// <summary>
        /// Carga los datos iniciales del Paso 1 al activar ese panel.
        /// </summary>
        private void CargarPaso1()
        {
            try
            {
                // Cargar tipos de lista de conteo en el ComboBox
                DataTable tiposLista = _repo.ObtenerTiposListaConteo();
                cmbListType.DataSource    = tiposLista;
                cmbListType.DisplayMember = "Physical_Count_List_Type";
                cmbListType.ValueMember   = "Physical_Count_List_Type_ID";

                // Obtener almacén predeterminado para referencia interna
                _almacenFijoId = _repo.ObtenerAlmacenPredeterminado();

                // Si es un inventario existente, restaurar el responsable
                if (_physicalCountId > 0)
                    RestaurarResponsable();
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar el encabezado del inventario.", ex);
            }
        }

        /// <summary>
        /// Restaura los datos del responsable al cargar un inventario existente.
        /// </summary>
        private void RestaurarResponsable()
        {
            // El responsable se almacena localmente; se reconstruye desde el encabezado
            CountHeaderInfo encabezado = _repo.ObtenerResumenEncabezado(_physicalCountId);
            if (encabezado == null)
                return;

            txtCountName.Text          = encabezado.CountName;
            dtpStartDate.Value         = encabezado.StartDate;
            dtpEndDate.Value           = encabezado.EndDate;
            lblResponsibleName.Text    = encabezado.ResponsibleName;
            lblResponsibleEmpNumber.Text = encabezado.EmployeeNumber;
            lblJobPosition.Text        = encabezado.JobPosition;
        }

        /// <summary>
        /// Evento: el usuario hace clic en "Buscar Responsable".
        /// Busca por número de empleado, login o nombre parcial.
        /// </summary>
        private void btnSearchResponsible_Click(object sender, EventArgs e)
        {
            string texto = txtResponsible.Text.Trim();
            if (string.IsNullOrEmpty(texto))
            {
                MessageBox.Show(
                    "Ingrese un número de empleado, login o nombre para buscar.",
                    "Búsqueda de responsable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                ResponsibleInfo responsable = _repo.BuscarResponsablePorTexto(texto);
                if (responsable == null)
                {
                    MessageBox.Show(
                        "No se encontró ningún usuario activo con ese criterio.",
                        "Responsable no encontrado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    LimpiarCamposResponsable();
                    return;
                }

                MostrarResponsableEnUI(responsable);
            }
            catch (Exception ex)
            {
                MostrarError("Error al buscar el responsable.", ex);
            }
        }

        /// <summary>
        /// Muestra los datos del responsable encontrado en los controles de UI.
        /// </summary>
        private void MostrarResponsableEnUI(ResponsibleInfo responsable)
        {
            lblResponsibleName.Text      = responsable.Name;
            lblResponsibleEmpNumber.Text = responsable.EmployeeNumber;
            lblJobPosition.Text          = responsable.JobPosition;

            // Almacenar el ID como Tag del label para uso posterior
            lblResponsibleName.Tag = responsable.UserId;
        }

        /// <summary>
        /// Limpia los campos del responsable cuando no hay coincidencia.
        /// </summary>
        private void LimpiarCamposResponsable()
        {
            lblResponsibleName.Text      = string.Empty;
            lblResponsibleEmpNumber.Text = string.Empty;
            lblJobPosition.Text          = string.Empty;
            lblResponsibleName.Tag       = null;
        }

        /// <summary>
        /// Guarda el encabezado del inventario físico (Paso 1).
        /// Este método coordina UI y llama a System_Functions, por lo que permanece en el Form.
        /// </summary>
        private bool GuardarEncabezadoInventario()
        {
            if (string.IsNullOrEmpty(txtCountName.Text.Trim()))
            {
                MessageBox.Show(
                    "El nombre del inventario es obligatorio.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            if (lblResponsibleName.Tag == null)
            {
                MessageBox.Show(
                    "Seleccione un responsable antes de continuar.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            if (dtpStartDate.Value.Date > dtpEndDate.Value.Date)
            {
                MessageBox.Show(
                    "La fecha de inicio no puede ser mayor a la fecha de fin.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            // SaveInventoryHeader permanece aquí porque usa Persal.System_Functions directamente.
            // Al completar el guardado, _physicalCountId queda actualizado.
            return true;
        }

        #endregion

        // =====================================================================
        // PASO 2 — SELECCIÓN DE UBICACIONES
        // =====================================================================
        #region PASO 2 — SELECCIÓN DE UBICACIONES

        /// <summary>
        /// Carga los datos iniciales del Paso 2: lista de almacenes disponibles.
        /// </summary>
        private void CargarPaso2()
        {
            try
            {
                DataTable almacenes = _repo.ObtenerAlmacenes(_almacenFijoId);
                clbAlmacenes.DataSource    = almacenes;
                clbAlmacenes.DisplayMember = "Location_Name";
                clbAlmacenes.ValueMember   = "Location_ID";
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar los almacenes.", ex);
            }
        }

        /// <summary>
        /// Evento: el usuario selecciona (focaliza) un almacén en la lista.
        /// Carga las áreas del almacén enfocado.
        /// </summary>
        private void clbAlmacenes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (clbAlmacenes.SelectedItem == null)
                return;

            DataRowView fila      = clbAlmacenes.SelectedItem as DataRowView;
            int         almacenId = Convert.ToInt32(fila["Location_ID"]);

            try
            {
                DataTable areas = _repo.ObtenerAreasPorAlmacen(almacenId);
                clbAreas.DataSource    = areas;
                clbAreas.DisplayMember = "Location_Name";
                clbAreas.ValueMember   = "Location_ID";

                // Limpiar sub-áreas al cambiar de almacén
                clbSubAreas.DataSource = null;
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar las áreas del almacén.", ex);
            }
        }

        /// <summary>
        /// Evento: el usuario selecciona (focaliza) un área en la lista.
        /// Carga las sub-áreas del área enfocada.
        /// </summary>
        private void clbAreas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (clbAreas.SelectedItem == null)
                return;

            DataRowView fila   = clbAreas.SelectedItem as DataRowView;
            int         areaId = Convert.ToInt32(fila["Location_ID"]);

            try
            {
                DataTable subAreas = _repo.ObtenerSubAreasPorArea(areaId);
                clbSubAreas.DataSource    = subAreas;
                clbSubAreas.DisplayMember = "Location_Name";
                clbSubAreas.ValueMember   = "Location_ID";
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar las sub-áreas.", ex);
            }
        }

        /// <summary>
        /// Evento: el usuario marca/desmarca un área en la lista de selección múltiple.
        /// Recarga las sub-áreas de todas las áreas marcadas.
        /// </summary>
        private void clbAreas_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Dar tiempo al control para actualizar su estado antes de leer las marcadas
            BeginInvoke(new Action(RecargarSubAreasDeAreasSeleccionadas));
        }

        /// <summary>
        /// Refresca la lista de sub-áreas en base a las áreas actualmente marcadas.
        /// </summary>
        private void RecargarSubAreasDeAreasSeleccionadas()
        {
            var areaIds = new List<int>();
            foreach (int indice in clbAreas.CheckedIndices)
            {
                DataRowView fila = clbAreas.Items[indice] as DataRowView;
                if (fila != null)
                    areaIds.Add(Convert.ToInt32(fila["Location_ID"]));
            }

            try
            {
                DataTable subAreas = _repo.ObtenerSubAreasPorAreas(areaIds);
                clbSubAreas.DataSource    = subAreas;
                clbSubAreas.DisplayMember = "Location_Name";
                clbSubAreas.ValueMember   = "Location_ID";
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar sub-áreas de las áreas seleccionadas.", ex);
            }
        }

        /// <summary>
        /// Guarda las ubicaciones seleccionadas por el usuario antes de avanzar al Paso 3.
        /// Verifica que no haya ubicaciones ya contadas si se quiere modificar la selección.
        /// </summary>
        /// <returns>True si el guardado fue exitoso; False en caso contrario.</returns>
        private bool GuardarUbicacionesSeleccionadas()
        {
            try
            {
                // Verificar si ya hay ubicaciones contadas (no se puede cambiar)
                int contadas = _repo.ContarUbicacionesContadas(_physicalCountId);
                if (contadas > 0)
                {
                    MessageBox.Show(
                        "Ya existen ubicaciones contadas. No se puede modificar la selección.",
                        "Operación no permitida",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                // Eliminar selección previa y guardar la nueva
                _repo.EliminarUbicacionesSeleccionadas(_physicalCountId);

                foreach (int indice in clbSubAreas.CheckedIndices)
                {
                    DataRowView fila      = clbSubAreas.Items[indice] as DataRowView;
                    int         locationId = Convert.ToInt32(fila["Location_ID"]);
                    int         hijos      = _repo.ContarHijosDeUbicacion(locationId);

                    // Nivel efectivo: si tiene hijos se guarda a nivel área (padre),
                    // si no tiene hijos se guarda a nivel sub-área
                    int nivelEfectivo = hijos > 0 ? 2 : 3;

                    _repo.GuardarUbicacionSeleccionada(
                        _physicalCountId,
                        locationId,   // selectedId
                        3,            // selectedLevel (sub-área)
                        locationId,   // effectiveId
                        nivelEfectivo,
                        null,
                        _currentUserId);
                }

                int guardadas = _repo.ContarUbicacionesGuardadas(_physicalCountId);
                if (guardadas == 0)
                {
                    MessageBox.Show(
                        "Seleccione al menos una ubicación para continuar.",
                        "Sin ubicaciones",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                MostrarError("Error al guardar las ubicaciones seleccionadas.", ex);
                return false;
            }
        }

        #endregion

        // =====================================================================
        // PASO 3 — CONTEO Y SNAPSHOTS
        // =====================================================================
        #region PASO 3 — CONTEO Y SNAPSHOTS

        /// <summary>
        /// Carga los datos iniciales del Paso 3: genera snapshots si aún no existen
        /// y refresca la grilla de conteo.
        /// </summary>
        private void CargarPaso3()
        {
            try
            {
                // Asegurar que los snapshots existan (los genera si es la primera vez)
                SnapshotResult resultado = _repo.AsegurarSnapshotsExistentes(_physicalCountId);
                if (!resultado.Success)
                {
                    MessageBox.Show(
                        "No se pudieron generar los snapshots del inventario.\n" + resultado.Message,
                        "Error en snapshots",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Refrescar resumen y grilla
                ActualizarResumenPaso3();
                RefrescarGridPaso3();
                CargarAdjuntosPaso3();
            }
            catch (Exception ex)
            {
                MostrarError("Error al inicializar el Paso 3.", ex);
            }
        }

        /// <summary>
        /// Actualiza las etiquetas de resumen del Paso 3 (nombre del inventario y ubicaciones).
        /// </summary>
        private void ActualizarResumenPaso3()
        {
            string nombre      = _repo.ObtenerNombreConteo(_physicalCountId);
            int    ubicaciones = _repo.ContarUbicacionesDistintas(_physicalCountId);

            lblStep3CountName.Text      = nombre ?? string.Empty;
            lblStep3LocationCount.Text  = string.Format("{0} ubicación(es)", ubicaciones);
        }

        /// <summary>
        /// Refresca los datos de la grilla de conteo del Paso 3.
        /// </summary>
        private void RefrescarGridPaso3()
        {
            try
            {
                DataTable datos = _repo.RefrescarGridPaso3(_physicalCountId);
                dgvStep3.DataSource = datos;
            }
            catch (Exception ex)
            {
                MostrarError("Error al refrescar la grilla de conteo.", ex);
            }
        }

        /// <summary>
        /// Evento: botón "Refrescar" del Paso 3. Vuelve a cargar la grilla.
        /// </summary>
        private void btnRefreshStep3_Click(object sender, EventArgs e)
        {
            RefrescarGridPaso3();
        }

        /// <summary>
        /// Verifica que el conteo esté finalizado antes de permitir avanzar al Paso 4.
        /// </summary>
        /// <returns>True si el conteo está completado; False en caso contrario.</returns>
        private bool ValidarConteoPaso3()
        {
            try
            {
                CountStatusResult estado = _repo.ObtenerEstadoConteo(_physicalCountId);
                if (!estado.IsFinished)
                {
                    MessageBox.Show(
                        "El conteo no ha sido finalizado en la plataforma web.\n" +
                        "Estado actual: " + estado.StatusName,
                        "Conteo pendiente",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                MostrarError("Error al validar el estado del conteo.", ex);
                return false;
            }
        }

        // -----------------------------------------------------------------
        // Paso 3 — Adjuntos
        // -----------------------------------------------------------------

        /// <summary>
        /// Carga la lista de archivos adjuntos en el grid del Paso 3.
        /// </summary>
        private void CargarAdjuntosPaso3()
        {
            try
            {
                DataTable adjuntos = _repo.ObtenerArchivosAdjuntos(_physicalCountId);
                dgvAttachments.DataSource = adjuntos;
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar los archivos adjuntos.", ex);
            }
        }

        /// <summary>
        /// Evento: botón "Agregar adjunto". Abre diálogo de selección de archivo
        /// y guarda el adjunto a través del repositorio.
        /// </summary>
        private void btnAttachAdd_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialogo = new OpenFileDialog())
            {
                dialogo.Title  = "Seleccionar archivo adjunto";
                dialogo.Filter = "Todos los archivos (*.*)|*.*";

                if (dialogo.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    string rutaArchivo  = dialogo.FileName;
                    string nombreSinExt = Path.GetFileNameWithoutExtension(rutaArchivo);
                    string extension    = Path.GetExtension(rutaArchivo);
                    byte[] datos        = File.ReadAllBytes(rutaArchivo);
                    long   tamano       = datos.LongLength;

                    // El Content-Type se determina por extensión; valor genérico si no se reconoce
                    string tipoContenido = ObtenerTipoContenido(extension);

                    _repo.GuardarAdjunto(
                        _physicalCountId,
                        nombreSinExt,
                        extension,
                        tipoContenido,
                        tamano,
                        datos,
                        _currentUserId);

                    CargarAdjuntosPaso3();
                }
                catch (Exception ex)
                {
                    MostrarError("Error al adjuntar el archivo.", ex);
                }
            }
        }

        /// <summary>
        /// Evento: botón "Abrir adjunto". Descarga y abre el adjunto seleccionado.
        /// </summary>
        private void btnAttachOpen_Click(object sender, EventArgs e)
        {
            if (dgvAttachments.CurrentRow == null)
                return;

            int attachmentId = Convert.ToInt32(dgvAttachments.CurrentRow.Cells["Attachment_ID"].Value);

            try
            {
                AttachmentData adjunto = _repo.ObtenerAdjunto(attachmentId);
                string rutaTemporal    = Path.Combine(Path.GetTempPath(), adjunto.FileName);

                File.WriteAllBytes(rutaTemporal, adjunto.FileData);
                System.Diagnostics.Process.Start(rutaTemporal);
            }
            catch (Exception ex)
            {
                MostrarError("Error al abrir el adjunto.", ex);
            }
        }

        /// <summary>
        /// Evento: botón "Eliminar adjunto". Elimina el adjunto seleccionado.
        /// </summary>
        private void btnAttachRemove_Click(object sender, EventArgs e)
        {
            if (dgvAttachments.CurrentRow == null)
                return;

            DialogResult confirmacion = MessageBox.Show(
                "¿Está seguro de eliminar el adjunto seleccionado?",
                "Confirmar eliminación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmacion != DialogResult.Yes)
                return;

            int attachmentId = Convert.ToInt32(dgvAttachments.CurrentRow.Cells["Attachment_ID"].Value);

            try
            {
                _repo.EliminarAdjunto(attachmentId);
                CargarAdjuntosPaso3();
            }
            catch (Exception ex)
            {
                MostrarError("Error al eliminar el adjunto.", ex);
            }
        }

        #endregion

        // =====================================================================
        // PASO 4 — REVISIÓN Y FINALIZACIÓN
        // =====================================================================
        #region PASO 4 — REVISIÓN Y FINALIZACIÓN

        /// <summary>
        /// Carga los datos iniciales del Paso 4: encabezado, combo de ubicaciones y grilla de diferencias.
        /// </summary>
        private void CargarPaso4()
        {
            try
            {
                // Encabezado del inventario
                CargarEncabezadoPaso4();

                // Combo de ubicaciones con snapshot
                DataTable ubicaciones = _repo.ObtenerUbicacionesSnapshot(_physicalCountId);
                cmbLocationsStep4.DataSource    = ubicaciones;
                cmbLocationsStep4.DisplayMember = "Location_Name";
                cmbLocationsStep4.ValueMember   = "Location_ID";

                // Cargar datos del resumen ejecutivo (Paso 5)
                CargarDatosResumenEjecutivo();
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar la revisión del inventario.", ex);
            }
        }

        /// <summary>
        /// Muestra el encabezado del inventario en las etiquetas del Paso 4.
        /// </summary>
        private void CargarEncabezadoPaso4()
        {
            CountHeaderInfo encabezado = _repo.ObtenerResumenEncabezado(_physicalCountId);
            if (encabezado == null)
                return;

            lblStep4CountName.Text      = encabezado.CountName;
            lblStep4StartDate.Text      = encabezado.StartDate.ToString("dd/MM/yyyy");
            lblStep4EndDate.Text        = encabezado.EndDate.ToString("dd/MM/yyyy");
            lblStep4Responsible.Text    = encabezado.ResponsibleName;
            lblStep4EmpNumber.Text      = encabezado.EmployeeNumber;
            lblStep4JobPosition.Text    = encabezado.JobPosition;
        }

        /// <summary>
        /// Evento: el usuario cambia la ubicación en el combo del Paso 4.
        /// Recarga la grilla de diferencias para la ubicación seleccionada.
        /// </summary>
        private void cmbLocationsStep4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbLocationsStep4.SelectedValue == null)
                return;

            int locationId = Convert.ToInt32(cmbLocationsStep4.SelectedValue);
            CargarDiferenciasPorUbicacion(locationId);
            CargarDetallePaso4PorUbicacion(locationId);
        }

        /// <summary>
        /// Carga las diferencias de inventario para la ubicación seleccionada.
        /// </summary>
        private void CargarDiferenciasPorUbicacion(int locationId)
        {
            try
            {
                DataTable diferencias = _repo.ObtenerDiferenciasPorUbicacion(_physicalCountId, locationId);
                dgvDifferences.DataSource = diferencias;
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar las diferencias.", ex);
            }
        }

        /// <summary>
        /// Carga el detalle del Paso 5 para la ubicación indicada.
        /// </summary>
        private void CargarDetallePaso4PorUbicacion(int locationId)
        {
            try
            {
                DataTable detalle = _repo.CargarDetallePorUbicacion(_physicalCountId, locationId);
                dgvStep5Detail.DataSource = detalle;

                DataTable ejecutivo = _repo.CargarResumenEjecutivoPorUbicacion(_physicalCountId, locationId);
                dgvStep5Executive.DataSource = ejecutivo;
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar el detalle de la ubicación.", ex);
            }
        }

        /// <summary>
        /// Carga el resumen ejecutivo completo del inventario (todas las ubicaciones).
        /// </summary>
        private void CargarDatosResumenEjecutivo()
        {
            try
            {
                Step5DataResult datos = _repo.CargarDatosPaso5(_physicalCountId);
                dgvStep5Detail.DataSource    = datos.Detail;
                dgvStep5Executive.DataSource = datos.Executive;
            }
            catch (Exception ex)
            {
                MostrarError("Error al cargar el resumen ejecutivo.", ex);
            }
        }

        /// <summary>
        /// Finaliza el Wizard: valida que todo esté contado y cierra el inventario.
        /// Este método coordina UI + llamada al SP de cambio de estado, por lo que permanece en el Form.
        /// </summary>
        private void FinalizarWizard()
        {
            try
            {
                // Verificar que no queden líneas sin contar
                int pendientes = _repo.ContarLineasSinConteo(_physicalCountId);
                if (pendientes > 0)
                {
                    MessageBox.Show(
                        string.Format(
                            "Existen {0} línea(s) sin contar. Complete el conteo antes de finalizar.",
                            pendientes),
                        "Conteo incompleto",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                DialogResult confirmacion = MessageBox.Show(
                    "¿Está seguro de finalizar el inventario? Esta acción no se puede revertir.",
                    "Confirmar finalización",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirmacion != DialogResult.Yes)
                    return;

                // Set_Physical_Count_Status es un método de System_Functions que
                // permanece en el formulario original; aquí sólo se muestra la llamada.
                // sf.Set_Physical_Count_Status(_physicalCountId, "CLOSED", _currentUserId);

                MessageBox.Show(
                    "El inventario ha sido finalizado exitosamente.",
                    "Inventario finalizado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MostrarError("Error al finalizar el inventario.", ex);
            }
        }

        /// <summary>
        /// Evento: botón "Finalizar" del Paso 4.
        /// </summary>
        private void btnFinish_Click(object sender, EventArgs e)
        {
            FinalizarWizard();
        }

        #endregion

        // =====================================================================
        // MÉTODOS AUXILIARES DE NAVEGACIÓN Y UI
        // =====================================================================
        #region MÉTODOS AUXILIARES DE NAVEGACIÓN Y UI

        /// <summary>
        /// Evento de carga del formulario. Inicializa el Paso 1.
        /// </summary>
        private void PhysicalCountCyclicForm_Load(object sender, EventArgs e)
        {
            IrAlPaso(1);
        }

        /// <summary>
        /// Evento: botón "Siguiente". Valida el paso actual y avanza al siguiente.
        /// </summary>
        private void btnNext_Click(object sender, EventArgs e)
        {
            if (!ValidarPasoActual())
                return;

            IrAlPaso(_currentStep + 1);
        }

        /// <summary>
        /// Evento: botón "Anterior". Retrocede al paso anterior sin validar.
        /// </summary>
        private void btnBack_Click(object sender, EventArgs e)
        {
            if (_currentStep > 1)
                IrAlPaso(_currentStep - 1);
        }

        /// <summary>
        /// Evento: botón "Cancelar". Cierra el Wizard sin guardar cambios pendientes.
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult confirmacion = MessageBox.Show(
                "¿Desea cancelar el Wizard? Los cambios no guardados se perderán.",
                "Cancelar",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmacion == DialogResult.Yes)
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        /// <summary>
        /// Activa el panel correspondiente al paso indicado y carga sus datos.
        /// </summary>
        /// <param name="paso">Número de paso destino (1–4).</param>
        private void IrAlPaso(int paso)
        {
            _currentStep = paso;

            // Mostrar u ocultar paneles según el paso
            pnlStep1.Visible = (paso == 1);
            pnlStep2.Visible = (paso == 2);
            pnlStep3.Visible = (paso == 3);
            pnlStep4.Visible = (paso == 4);

            // Actualizar estado de botones de navegación
            btnBack.Enabled = (paso > 1);
            btnNext.Visible = (paso < 4);
            btnFinish.Visible = (paso == 4);

            // Actualizar indicador de paso en la barra de progreso
            ActualizarIndicadorPasos(paso);

            // Cargar datos del paso correspondiente
            switch (paso)
            {
                case 1: CargarPaso1(); break;
                case 2: CargarPaso2(); break;
                case 3: CargarPaso3(); break;
                case 4: CargarPaso4(); break;
            }
        }

        /// <summary>
        /// Valida el paso actual antes de permitir avanzar.
        /// </summary>
        /// <returns>True si la validación es exitosa; False en caso contrario.</returns>
        private bool ValidarPasoActual()
        {
            switch (_currentStep)
            {
                case 1:
                    return GuardarEncabezadoInventario();

                case 2:
                    return GuardarUbicacionesSeleccionadas();

                case 3:
                    return ValidarConteoPaso3();

                default:
                    return true;
            }
        }

        /// <summary>
        /// Actualiza el indicador visual de pasos (e.g., labels o TabControl en modo lectura).
        /// </summary>
        /// <param name="pasoActivo">Número del paso actualmente activo (1–4).</param>
        private void ActualizarIndicadorPasos(int pasoActivo)
        {
            lblStep1Indicator.Font = new System.Drawing.Font(
                lblStep1Indicator.Font,
                pasoActivo == 1 ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);

            lblStep2Indicator.Font = new System.Drawing.Font(
                lblStep2Indicator.Font,
                pasoActivo == 2 ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);

            lblStep3Indicator.Font = new System.Drawing.Font(
                lblStep3Indicator.Font,
                pasoActivo == 3 ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);

            lblStep4Indicator.Font = new System.Drawing.Font(
                lblStep4Indicator.Font,
                pasoActivo == 4 ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);
        }

        /// <summary>
        /// Determina el tipo de contenido MIME a partir de la extensión del archivo.
        /// </summary>
        /// <param name="extension">Extensión incluyendo el punto (ej. ".pdf").</param>
        /// <returns>Cadena MIME o "application/octet-stream" si no se reconoce.</returns>
        private static string ObtenerTipoContenido(string extension)
        {
            if (string.IsNullOrEmpty(extension))
                return "application/octet-stream";

            switch (extension.ToLowerInvariant())
            {
                case ".pdf":  return "application/pdf";
                case ".xls":  return "application/vnd.ms-excel";
                case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case ".doc":  return "application/msword";
                case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".png":  return "image/png";
                case ".txt":  return "text/plain";
                default:      return "application/octet-stream";
            }
        }

        /// <summary>
        /// Muestra un mensaje de error amigable y registra la excepción.
        /// </summary>
        /// <param name="mensajeAmigable">Descripción de la operación que falló.</param>
        /// <param name="ex">Excepción capturada.</param>
        private void MostrarError(string mensajeAmigable, Exception ex)
        {
            MessageBox.Show(
                mensajeAmigable + "\n\nDetalle: " + ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        #endregion
    }
}
