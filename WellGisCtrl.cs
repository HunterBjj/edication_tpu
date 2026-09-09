using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Docking;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DotLib.util;
using pxp.client.common.controls;
using PXP.Common.Navigation;
using Sibir.Framework.Client.Shell.Common.util;
using Sibir.Framework.Client.Shell.Common.util.controls;
using Sibir.Framework.Client.Shell.Common.util.forms;
using Sibir.Framework.Common;
using pxp.client.common.util;
using pxp.common;
using Sibir.Framework.Common.util;
using sibir.pxp.client.core.components.WellInfo.forms;
using sibir.pxp.client.core.Properties;
using System.IO.Compression;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;


namespace sibir.pxp.client.core.controls {
    /// <summary>
    /// Класс отображения данных по ГИС
    /// </summary>
    public partial class WellGisCtrl : BaseGridPanelCtrl {
        #region Данные

        /// <summary>
        /// Id скважины
        /// </summary>
        private long _wellId;

        /// <summary>
        /// Нужно ли загружать данные
        /// </summary>
        private bool _isNeedLoad;

        /// <summary>
        /// Форма для отображения прикрепленных файлов
        /// </summary>
        private FileUploadFrom _uploadForm;

        /// <summary>
        /// Строки помеченые на удаление из таблицы в базе
        /// </summary>
        private Stack<long> deleteList;

        /// <summary>
        /// Строки помеченые для обновления
        /// </summary>
        private Stack<long?> updateList = new Stack<long?>();

        /// <summary>
        /// Контрол для проверки правильности внесенного значения пласта при вставке
        /// </summary>
        private readonly LayerSelectCtrl _layerCtrlForCheck = new LayerSelectCtrl();

        /// <summary>
        /// Контрол для проверки правильности внесенного значения пласта при вставке
        /// </summary>
        private List<string> _uniqColumns = new List<string>();

        /// <summary>
        /// Контрол для проверки правильности внесенного значения пласта при вставке
        /// </summary>
        private DataTable _template = null;

        /// <summary>
        /// Название столбца в РИГИС
        /// </summary>
        private string ValueFieldName;

        /// <summary>
        /// Id заголовка РИГИС проектные документы
        /// </summary>
        private long dirRigisSourceHdrId = 0;

        /// <summary>
        /// Нужно ли загружать данные
        /// </summary>
        private bool isNeedLoad;

        /// <summary>
        /// Это данные версионные для РИГИС
        /// </summary>
        private static bool isWellLogHist;

        public static bool IsWellLogHist
        {
            get
            {
                return isWellLogHist;
            }
            set
            {
                isWellLogHist = value;
            }
        }



        #endregion

        #region Свойства

        /// <summary>
        /// Возвращает имя файла в который будет экспортироваться грид(без пути)
        /// </summary>
        public override string DefExportFileName {
            get {
                string navStr = string.Empty;
                if (Component != null &&
                    Component.NavigationControl != null) {
                    navStr = Component.NavigationControl.getSelectionViewStringS();
                }
                return "РИГИС " + navStr + ".xls";
            }
        }

        #endregion

        #region Конструкторы и инициализация

        /// <summary>
        /// Конструктор
        /// </summary>
        public WellGisCtrl() {
            InitializeComponent();
        }

        /// <summary> Инициализация
        /// </summary>
        /// <param name="wellId"></param>
        /// <param name="isNeedLoad">Нужно ли перезагружать данные на контроле</param>
// ReSharper disable ParameterHidesMember
        public void init(long wellId, bool isNeedLoad) {
            _template = null;
            _wellId = wellId;
            dirRigisSourceHdrId = 0;
            ValueFieldName = "DATE_LOAD";
            _layerCtrlForCheck.init(Well.getFieldId(wellId), 0);
            _isNeedLoad = isNeedLoad;
            pnlData.Hide();
            barBtnCreate.Visibility = BarItemVisibility.Never;
            ActiveView = gisGridView;
            DefExportFileName =
                string.Format(Resources.ExportToExcel_FileName_F,
                              PXPClientCommonUtil.getProjResStrByLang(PXPConstants.PROJ_RES_STR_HYDR_RESEARCH),
                              ((GEONavigationCtrl)NavigationControl).getCurWell().Title,
                              ((GEONavigationCtrl)NavigationControl).getCurField().Title);

            updateBarItemGisVersion(wellId);

            init();

            _uniqColumns.Add("TOP");
            _uniqColumns.Add("BASE");
            _uniqColumns.Add("SOURCE_S");
            _uniqColumns.Add("WELL_S");
        }

        #endregion

        private void updateBarItemGisVersion(long wellId) {
            barItemWellGisVersion.EditValueChanged -= inklSelectSourceRigisCtrlEditValueChanged;
            WellGisVersionEditCtrlRepItem1.SelectCtrl.init(wellId);
            barItemWellGisVersion.EditValue = WellGisVersionEditCtrlRepItem1.SelectCtrl.getActual();
            barItemWellGisVersion.EditValueChanged += inklSelectSourceRigisCtrlEditValueChanged;
        }

        #region Реализация
        public override void setRights(Sibir.Framework.Client.Shell.Common.rights.RightsDescriptor right) {
            base.setRights(right);
            if (!ShellGateway.IsCurrUserAdmin) {
                dockPanel1.Visibility = DockVisibility.Hidden;
            }
        }

        /// <summary>
        /// Возвращает выбранную в Grid'e скважину 
        /// </summary>
        private DataRow getDataRow() {
            DataRow result = null;
            if (gisGridView.SelectedRowsCount > 0) {
                int[] selRowId = gisGridView.GetSelectedRows();
                DataRow dataRow = gisGridView.GetDataRow(selRowId[0]);
                result = dataRow;
            }
            return result;
        }

        /// <summary>
        /// Показывет форму редактирования интерпретации
        /// </summary>
        private void showWellGisForm(DataRow dataRow) {
            var wellGisForm1 = new WellGisForm(Component, Rights, dataRow, gisGridView);
            if (wellGisForm1.ShowDialog(this) ==
                DialogResult.OK) {
                // Обновляем данные в Grid'e
                gridControl1.FireChanged();

            }
            wellGisForm1.Dispose();
        }

        /// <summary>
        /// Обновление данные контрола - перезагружаем данные если это нужно
        /// </summary>
        public override void Refresh() {
            if (_isNeedLoad) {
                Module.ShellGateway.ShowMessageInStatusString(String.Format(
                    PXPClientCommonUtil.getProjResStrByLang(PXPConstants.PROJ_RES_FSTR_LOADING),
                    PXPClientCommonUtil.getProjResStrByLang(PXPConstants.PROJ_RES_STR_LAYER_INTERPR)));
                loadWellGis();
                _isNeedLoad = false;
            }
        }

        /// <summary>
        /// Загрузка данных по ГИС
        /// </summary>
        public void loadWellGis(bool is_layer_hist = false, long? sourceId = null) {
            try {
                setStatus(Status.Loading);

                if (is_layer_hist && sourceId != 0) {

                    callService(((ModuleCore)Module).getWellInfoService(), "getWellVersionHistInterpretationDataSet",
                    new object[] { _wellId, sourceId }, endLoadWellGis);
                }
                else {
                    callService(((ModuleCore)Module).getWellInfoService(), "getWellInterpretationDataSet",
                    new object[] { _wellId }, endLoadWellGis);
                }

            } catch (Exception ex) {
                Module.ShellGateway.commonShowError(
                    PXPClientCommonUtil.getProjResStrByLang(PXPConstants.PROJ_RES_STR_FAIL_READ_DATA), ex);
            }
            Module.ShellGateway.showReadyInStatusString();
        }

        /// <summary>
        /// Инициализация грида
        /// </summary>
        /// <param name="ds"></param>
        private void loadWellGis(DataSet ds) {
            gridControl1.BeginUpdate();
            gridControl1.DataSource = ds;
            gridControl1.DataMember = ds.Tables[0].TableName;
            if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0) {
                hideOrShowGridColumns(ds.Tables[0]);
            }
            //gisGridView.BestFitColumns();
            init();

            // === ПОДКЛЮЧАЕМ ПОДСВЕТКУ ПЕРЕСЕЧЕНИЙ ===
            gisGridView.RowCellStyle -= gisGridView_RowCellStyle; 
            gisGridView.RowCellStyle += gisGridView_RowCellStyle; 

            columnFilterCtrl.init(gisGridView);
            if (_template != null) {
                columnFilterCtrl.DataTable = _template;
                columnFilterCtrl.loadVisabilitys(_template);
            }
            deleteList = new Stack<long>();
            gridControl1.EndUpdate();

            initGridColorHighlighting();

            // Заставляем грид принудительно перерисовать ячейки прямо сейчас
            gisGridView.RefreshData();
        }

        private void initGridColorHighlighting() {
            gisGridView.RowCellStyle -= gisGridView_RowCellStyle;
            gisGridView.RowCellStyle += gisGridView_RowCellStyle;
        }

        /// <summary>
        /// Метод скрывает/отображает дополнительные колонки в зависимости от наличия в них данных
        /// </summary>
        /// <param name="table"></param>
        private void hideOrShowGridColumns(DataTable table) {
            foreach (GridColumn column in gisGridView.Columns) {
                column.Visible = false;
            }
            foreach (DataRow row in table.Rows) {
                foreach (GridColumn column in gisGridView.Columns) {
                    if (row.Table.Columns.Contains(column.FieldName) &&
                        !ConvertUtil.isNullOrEmptyTrimString(row[column.FieldName].ToString())) {
                        column.Visible = true;
                    }
                }
            }
            foreach (GridBand band in gisGridView.Bands) {
                hideOrShowGridBand(band);
            }
        }

        /// <summary>
        /// Закончилась загрузка данных по ГИС
        /// </summary>		
        /// <param name="args"></param>
        private void endLoadWellGis(CallServiceCompleteHandlerArgs args) {
            try {
                if (args.Status ==
                    LoadOverInfo.Ok) {
                    ShellGateway.ShowMessageInStatusString(
                        PXPClientCommonUtil.getProjResStrByLang(PXPConstants.PROJ_RES_STR_PREP_DATA_TO_SHOW));
                    DataSet ds = CommonUtil.decompressData(args.Result as byte[]);
                    loadWellGis(ds);
                    setStatus(Status.Ok);
                    Module.ShellGateway.showReadyInStatusString();
                }
            } catch (Exception ex) {
                Module.ShellGateway.commonShowError(
                    Module.ShellGateway.getShellResourcesStringLang(FwConstants.SHELL_RESOURCE_STR_ERROR_LOAD_DATA),
                    ex);
            }
        }

        /// <summary>
        /// Задает параметры для экспорта в Excel    
        /// </summary>
        /// <returns></returns>
        public override ShellUtil.xlsGridExportParams getDefExportParams() {
            ShellUtil.xlsGridExportParams exportParams = base.getDefExportParams();
            exportParams.fileName =
                exportParams.sheetName =
                exportParams.tableTitle =
                string.Format(Resources.ExportToExcel_FileName_F,
                              PXPClientCommonUtil.getProjResStrByLang(PXPConstants.PROJ_RES_STR_LAYER_INTERPR),
                              ((GEONavigationCtrl)NavigationControl).getCurWell().Title,
                              ((GEONavigationCtrl)NavigationControl).getCurField().Title);
            return exportParams;
        }


        /// <summary>
        /// Метод получает список, выбранных id интерпретации
        /// </summary>
        /// <returns></returns>
        private long[] getSelectedRowsId() {
            var selectedRows = gisGridView.GetSelectedRows();
            var rowsIds = new long[selectedRows.Length];

            for (int index = 0; index < rowsIds.Length; ++index) {
                var rowIndex = selectedRows[index];
                var row = gisGridView.GetDataRow(rowIndex);

                long selectId = getRowId(row);
                if (selectId != 0) {
                    rowsIds[index] = selectId;
                }
            }
            return rowsIds;
        }

        /// <summary>
        /// Метод формирует список с description'ов для отображения в форме
        /// </summary>
        /// <param name="isAdd"></param>
        /// <remarks>
        /// Формирование description'ов нужно чтобы отличать одну запись с интерпретацией от другой
        /// </remarks>
        /// <returns></returns>
        private string[] getSelectedRowsDesc(bool isAdd) {
            var indexes = gisGridView.GetSelectedRows();
            var rows = new Dictionary<int, DataRow>();
            foreach (int index in indexes) {
                rows.Add(index, gisGridView.GetDataRow(index));
            }
            var paramColumns = new[] { "TOP", "BASE" };
            var res = FileUploadFrom.createRowsDesc(rows, paramColumns, isAdd);
            return res;
        }

        /// <summary>
        /// Метод возвращает id записи с интерпретацией
        /// </summary>
        /// <param name="row">Строка с данными по интерпретации</param>
        /// <returns></returns>
        private static long getRowId(DataRow row) {
            if (row == null) {
                return 0;
            }
            if (row.Table.Columns.Contains("WELL_LOG_RESULT_LAYER_S")) {
                if (row["WELL_LOG_RESULT_LAYER_S"] == DBNull.Value ||
                    row["WELL_LOG_RESULT_LAYER_S"] == null) {
                    return 0;
                }
                return CommonUtil.toLong(row["WELL_LOG_RESULT_LAYER_S"]);
            }
            return 0;
        }

        /// <summary>
        /// Метод возввращает тело файла для интерпретации ГИС
        /// </summary>
        /// <param name="selectRow">Строка с файлом для интерпретации ГИС</param>
        /// <returns></returns>
        private static byte[] getFileBody(DataRow selectRow) {
            long fileId = CommonUtil.toLong(selectRow["WELL_LOG_RESULT_LAYER_SOURCE_S"]);
            long regionId = CommonUtil.toLong(selectRow["ADM_OBJECT_S"]);
            var bytes = PXPClientCommonUtil.getDictService().getWellLogResultLayerSourceFileBody(fileId, regionId);
            try {
                byte[] decompressedBytes = decompressGzip(bytes);
                return decompressedBytes;
            } catch (InvalidDataException) {
                return bytes;
            }
        }


        static byte[] decompressGzip(byte[] gzip) {
            var result = gzip;
            if (gzip != null) {
                using (GZipStream stream = new GZipStream(new MemoryStream(gzip), CompressionMode.Decompress)) {
                    const int size = 4096;
                    byte[] buffer = new byte[size];
                    using (MemoryStream memory = new MemoryStream()) {
                        int count = 0;
                        do {
                            count = stream.Read(buffer, 0, size);
                            if (count > 0) {
                                memory.Write(buffer, 0, count);
                            }
                        }
                        while (count > 0);
                        result = memory.ToArray();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Метод отсоединеяет файл от интерпретации ГИС
        /// </summary>
        /// <param name="selectRow"></param>
        private void deattachFile(DataRow selectRow) {
            DataRow wellInterpretationSelectedRow = gisGridView.GetDataRow(gisGridView.FocusedRowHandle);
            long wellLogResultLayerId = CommonUtil.toLong(wellInterpretationSelectedRow["WELL_LOG_RESULT_LAYER_S"]);
            long fileId = CommonUtil.toLong(selectRow["WELL_LOG_RESULT_LAYER_SOURCE_S"]);
            PXPClientCommonUtil.getLoaderSvc().detachFileWellGis(wellLogResultLayerId, fileId);
        }

        /// <summary>
        /// Метод скрывает/отображает соответствующие банды в зависимости от колонки и от дочерних бандов
        /// </summary>
        /// <param name="gridBand"></param>
        /// <returns></returns>
        public static bool hideOrShowGridBand(GridBand gridBand) {
            bool result = false;
            if (gridBand.HasChildren) {
                foreach (GridBand child in gridBand.Children) {
                    result |= hideOrShowGridBand(child);
                }
                gridBand.Visible = result;
            } else if (gridBand.Columns.Count > 0) {
                foreach (GridColumn column in gridBand.Columns) {
                    result |= column.Visible;
                }
                gridBand.Visible = result;
            }
            return gridBand.Visible;
        }

        private void barButtonPaste_ItemClick(object sender, ItemClickEventArgs e) {
            pasteFromBuffer();
        }

        private void pasteFromBuffer() {
            var data = getClipboardData();
            var textTable = createTableFromText(data);
            var rowIndex = gisGridView.FocusedRowHandle;
            //var table = Source.Tables["PTS"];
            var table = ((DataView)gisGridView.DataSource).Table;
            var columnIndex = table.Columns.IndexOf(gisGridView.FocusedColumn.FieldName);

            if (table.Rows.Count == 0)
                columnIndex = rowIndex = 0;
            if (rowIndex > -1 && columnIndex > -1) {
                try {
                    mergeTables(textTable, table, rowIndex, columnIndex);
                    Source = table.DataSet;
                } catch (Exception ex) {
                    ShellMessageBox.Show(ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string getClipboardData() {
            IDataObject iData = null;
            try {
                iData = Clipboard.GetDataObject();
            } catch (Exception ex) {
                iData = Clipboard.GetDataObject();
            }
            if (iData != null && iData.GetDataPresent(DataFormats.Text))
                return (string)iData.GetData(DataFormats.Text);
            return "";
        }

        private DataTable createTableFromText(string data) {
            var rows = data.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            DataTable table = new DataTable();
            int maxColumns = 0;
            for (int i = 0; i < rows.Length; i++) {
                var columns = rows[i].Replace("\r", "").Split('\t');
                if (columns.Length > maxColumns)
                    maxColumns = columns.Length;
            }
            while (table.Columns.Count < maxColumns)
                table.Columns.Add(new DataColumn((table.Columns.Count + 1).ToString()));
            for (int i = 0; i < rows.Length; i++) {

                var columns = rows[i].Replace("\r", "").Split(new[] { '\t' });
                DataRow dataRow = table.NewRow();
                for (int indexColumn = 0; indexColumn < columns.Length; ++indexColumn) {
                    dataRow[indexColumn] = columns[indexColumn];
                }
                table.Rows.Add(dataRow);
            }
            return table;
        }

        public DataSet Source {
            get {
                return ((DataTable)gridControl1.DataSource).DataSet;
            }
            private set {
                if (value != null && value.Tables.Contains("PTS")) {
                    var table = value.Tables["PTS"];
                    setStatus(Status.Ok);
                    if (table.Columns.Contains("CURVING") == false)
                        value.Tables["PTS"].Columns.Add("CURVING", typeof(double));
                    for (int i = 1; i < table.Rows.Count; i++) {
                        DataRow rows1 = table.Rows[i - 1];
                        DataRow rows2 = table.Rows[i];
                        double angle1 = CommonUtil.toDouble(rows1["DEVIATION_ANGLE"]);
                        double azimuth1 = CommonUtil.toDouble(rows1["AZIMUTH"]);
                        double angle2 = CommonUtil.toDouble(rows2["DEVIATION_ANGLE"]);
                        double azimuth2 = CommonUtil.toDouble(rows2["AZIMUTH"]);
                        var radAngle1 = CommonUtil.toRadians(angle1);
                        var radAngle2 = CommonUtil.toRadians(angle2);
                        var radAzimuth1 = CommonUtil.toRadians(azimuth1);
                        var radAzimuth2 = CommonUtil.toRadians(azimuth2);
                        double x = Math.Acos(Math.Cos(radAngle2 - radAngle1) - Math.Sin(radAngle1) * Math.Sin(radAngle2) * (1 - Math.Cos(radAzimuth2 - radAzimuth1)));
                        x = CommonUtil.toDegree(x);
                        rows2["CURVING"] = x;
                    }
                    gridControl1.DataSource = table;

                }
            }
        }

        private bool checkData(DataTable sourceTable, DataTable mainTable, int startRow, int startColumn) {
            if (sourceTable.Columns.Count != gisGridView.VisibleColumns.Count) {
                MessageBox.Show(
                    "Ошибка корректности данных. \nКоличество колонок в шаблоне несовпадает с количеством копируемых колонок",
                    "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                return false;
            }

            string message = string.Empty;
            int rowCount = sourceTable.Rows.Count;
            if (((DataView)gisGridView.DataSource).Table.Rows.Count < sourceTable.Rows.Count && barCheckItem1.Checked == true) {
                rowCount = ((DataView)gisGridView.DataSource).Table.Rows.Count;
            }
            for (int row = 0; row < rowCount; row++) {
                int rowNumber = row + 1;
                for (int column = 0; column < sourceTable.Columns.Count; column++) {
                    int columnNumber = column + 1;
                    bool check = true;
                    var elem = sourceTable.Rows[row].ItemArray[column];

                    switch (gisGridView.VisibleColumns[startColumn + column].FieldName) {
                        case "APS":
                        case "CLAY":
                        case "CLAY_GK":
                        case "H":
                        case "KG":
                        case "KNG":
                        case "KS":
                        case "KVO":
                        case "KV":
                        case "OG_SATURATION_PODV":
                        case "OG_SATURATION_SV":
                        case "POROSITY_AK":
                        case "POROSITY":
                        case "POROSITY_EL_M":
                        case "POROSITY_E":
                        case "POROSITY_GGK":
                        case "POROSITY_GK":
                        case "POROSITY_KV":
                        case "POROSITY_NGK":
                        case "POROSITY_NNKT":
                        case "POROSITY_O":
                        case "POROSITY_RA_M":
                        case "POROSITY_SP":
                        case "POROSITY_TR":
                        case "POROSITY_NKT":
                        case "POROSITY_SV":
                        case "WAT_SATURATION_PODV": {
                                if (!isNumber(elem.ToString())) {
                                    check = false;
                                    message += string.Format("В строке номер {0} колонка {1} указано не число\n", rowNumber, columnNumber);
                                } else if (CommonUtil.toDoubleNull(elem) == null) {
                                    check = true;
                                    break;
                                } else if (double.Parse((string)elem) > 0 &&
                                    double.Parse((string)elem) <= 1) {
                                    check = true;
                                    break;
                                } else {
                                    check = false;
                                    message += string.Format("В строке номер {0} колонка {1} не выполнены условия 0 < знач и знач <= 1.\n", rowNumber, columnNumber);
                                }
                            }
                            break;
                        case "WELL_LOG_RESULT_LAYER_S":
                        case "SOURCE_S":
                        case "CURRENT_SATURATION_S":
                        case "LOADER_USER_S":
                        case "LOAD_DATE":
                        case "LOADER_TYPE_S":
                            check = elem == DBNull.Value ? false : true;
                            if (!check) {
                                message += string.Format("В строке номер {0} колонка {1} значение пустое.\n", rowNumber, columnNumber);
                            }
                            break;
                        case "RESULT_DATE":
                            if (!isDateTime(elem.ToString())) {
                                check = false;
                                message += string.Format("В строке номер {0} колонка {1} неверно указана дата\n", rowNumber, columnNumber);
                            }
                            break;
                        case "DIF_PAR_G":
                        case "DIF_PAR_N":
                        case "REL_PAR_G":
                        case "REL_PAR_N":
                        case "REL_PAR_SP":
                            if (!isNumber(elem.ToString())) {
                                check = false;
                                message += string.Format("В строке номер {0} колонка {1} указано не число\n", rowNumber, columnNumber);
                            } else if (CommonUtil.toDoubleNull(elem) == null) {
                                check = true;
                            } else if (double.Parse((string)elem) > 0 &&
                                double.Parse((string)elem) <= 100) {
                                check = true;
                            } else {
                                check = false;
                                message += string.Format("В строке номер {0} колонка {1} не выполнены условия 0 < знач и знач <= 100.\n", rowNumber, columnNumber);
                            }
                            break;
                        case "LAYER_S":
                            check = _layerCtrlForCheck.getIdByName((string)elem) != null;
                            if (!check) {
                                message += string.Format("В строке номер {0} колонка {1} значение пласта пустое.\n", rowNumber, columnNumber);
                            }
                            break;
                    }
                    if (!check) {
                        MessageBox.Show(
                           "Ошибка корректности данных. \n" + message,
                           "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                        return false;
                    }
                }
            }
            if (barCheckItem1.Checked) {
                int visibleFocusedColumnIndex = gisGridView.VisibleColumns.View.FocusedColumn.VisibleIndex;
                Dictionary<string, string> notNullColumns = new Dictionary<string, string>();
                notNullColumns.Add("SOURCE_NAME", "SOURCE_S");
                notNullColumns.Add("LITHOLOGY_NAME", "LITHOLOGY_S");
                notNullColumns.Add("SATURATION_NAME", "SATURATION_S");
                notNullColumns.Add("CURRENT_SATURATION_NAME", "CURRENT_SATURATION_S");

                foreach (var notNullColumn in notNullColumns) {

                    for (int rowIndex = startRow; rowIndex < rowCount; rowIndex++) {
                        if (CommonUtil.toLongNull(mainTable.Rows[rowIndex][notNullColumn.Value]) == null) {
                            int columnIndex = 0;
                            int sourceBeginColumnIndex = 0;
                            foreach (GridColumn visibleColumn in gisGridView.VisibleColumns) {

                                if (
                                    string.Compare(gisGridView.VisibleColumns[columnIndex].FieldName,
                                                   gisGridView.VisibleColumns[visibleFocusedColumnIndex].FieldName) == 0) {
                                    sourceBeginColumnIndex = columnIndex;
                                }
                                if (
                                    string.Compare(gisGridView.VisibleColumns[columnIndex].FieldName,
                                                   notNullColumn.Key) == 0)
                                {
                                    int rowCounter = 0;
                                    foreach (DataRow sourceRow in sourceTable.Rows) {
                                        try {
                                            var variable = Dict.getIdByName((long)PXPConstants.SourcesData.sdParent,
                                                                            CommonUtil.toStr(sourceRow[columnIndex - sourceBeginColumnIndex]),
                                                                            false);
                                            rowCounter++;
                                        } catch (Exception ex) {
                                            MessageBox.Show(
                                                "Ошибка в " + (rowCounter + 1).ToString() + " cтроке, " +
                                                " поле " + notNullColumn.Key + " не должно быть пустыми, а также данные должны содержаться в соответствующих справочниках!",
                                                "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                                            return false;
                                        }
                                    }
                                    return true;
                                }
                                columnIndex++;
                            }

                            MessageBox.Show(
                                "Ошибка в поле " + notNullColumn.Key + " не должно быть пустыми!",
                                "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                            return false;
                        }
                    }
                }


                for (int rowIndex = startRow; rowIndex < rowCount; rowIndex++) {
                    if (CommonUtil.toLongNull(mainTable.Rows[rowIndex]["LAYER_S"]) == null) {
                        int columnIndex = 0;
                        int sourceBeginColumnIndex = 0;
                        foreach (GridColumn visibleColumn in gisGridView.VisibleColumns) {
                            if (string.Compare(gisGridView.VisibleColumns[columnIndex].FieldName,
                                                   gisGridView.VisibleColumns[visibleFocusedColumnIndex].FieldName) == 0) {
                                sourceBeginColumnIndex = columnIndex;
                            }
                            if (string.Compare(gisGridView.VisibleColumns[columnIndex].FieldName, "LAYER_NAME") == 0) {
                                int rowCounter = 0;
                                foreach (DataRow sourceRow in sourceTable.Rows) {
                                    try {
                                        var variable = _layerCtrlForCheck.getIdByName(
                                        CommonUtil.toStr(sourceRow[columnIndex - sourceBeginColumnIndex]));
                                        rowCounter++;
                                    } catch (Exception ex) {
                                        MessageBox.Show(
                               "Ошибка в " + (rowCounter + 1).ToString() + " cтроке, " +
                               " поле \"Пласт\" не должно быть пустыми, а также данные должны содержаться в соответствующих справочниках!",
                               "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                                        return false;
                                    }
                                }
                                return true;
                            }
                            columnIndex++;
                        }

                        MessageBox.Show(
                               "Ошибка в поле \"Пласт\" не должно быть пустыми, а также данные должны содержаться в соответствующих справочниках!",
                               "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                        return false;
                    }
                }
            }
            //for (int column = 0; column < gisGridView.VisibleColumns.Count; column++) {
            //    if(String.Compare(gisGridView.VisibleColumns[column].FieldName, ""))
            //}
            return true;
        }

        /// <summary>
        /// Метод проверяет является ли строка числом
        /// </summary>
        /// <param name="value">Строка для проверки</param>
        /// <returns></returns>
        private static bool isNumber(string value) {
            const string pattern = @"([\d.,]+)";
            return string.IsNullOrEmpty(value) || Regex.Match(value, pattern).Success;
        }

        /// <summary>
        /// Метод проверяет является ли строка датой
        /// </summary>
        /// <param name="value">Строка для проверки</param>
        /// <returns></returns>
        private static bool isDateTime(string value) {
            var formats = new[] { "M-d-yyyy", "dd-MM-yyyy", "MM-dd-yyyy", "M.d.yyyy", "dd.MM.yyyy", "MM.dd.yyyy", "MM.dd.yy" }
                .Union(CultureInfo.CurrentCulture.DateTimeFormat.GetAllDateTimePatterns())
                .ToList();
            DateTime result = new DateTime();
            return string.IsNullOrEmpty(value) || formats.Any(f => DateTime.TryParseExact(value, f, CultureInfo.InvariantCulture, DateTimeStyles.None, out result));
        }

        private void mergeTables(DataTable sourceTable, DataTable destTable, int beginRow, int beginColumn) {
            bool oneMoreCheckPermeability = true;
            int sourceRowIndex = 0;
            int startRowIndex = 0;
            int startColumn = 0;
            bool start = true;
            DataSet copyDs = ((DataSet)gridControl1.DataSource).Copy();
            if (barCheckItem1.Checked) {
                //startRowIndex = gisGridView.VisibleColumns.View.FocusedRowHandle;
                startRowIndex = gisGridView.FocusedRowHandle;
                startColumn = gisGridView.VisibleColumns.View.FocusedColumn.VisibleIndex;
            }

            if (!checkData(sourceTable, destTable, startRowIndex, startColumn)) {
                return;
            }

            gridControl1.BeginUpdate();
            try {
                int rowCount = sourceTable.Rows.Count;
                if (barCheckItem1.Checked) {
                    rowCount = sourceTable.Rows.Count > ((DataView)gisGridView.DataSource).Table.Rows.Count
                                       ? ((DataView)gisGridView.DataSource).Table.Rows.Count
                                       : sourceTable.Rows.Count;
                }

                for (int row = 0; row < rowCount; row++) {
                    DataRow destRow = destTable.NewRow();
                    if (!barCheckItem1.Checked) {
                        destRow["WELL_S"] = _wellId;
                        destRow["LAYER_FIELD_S"] = destRow["WELL_FIELD_S"] = Well.getFieldId(_wellId);
                        destRow["FIELD_NAME"] = Field.getName(Well.getFieldId(_wellId));
                        destRow["WELL_NAME"] = Well.getName(_wellId);
                        destRow["LOADER_TYPE_S"] = 148011703;
                        destTable.Rows.InsertAt(destRow, startRowIndex);
                    } else {
                        start = false;
                    }
                    int indexSource = 0;
                    foreach (GridColumn elem in gisGridView.VisibleColumns) {
                        try {
                            if (elem.VisibleIndex == startColumn || start == true) {
                                start = true;
                                var value = sourceTable.Rows[sourceRowIndex].ItemArray[indexSource];
                                if (string.IsNullOrWhiteSpace(value as string))
                                    value = DBNull.Value;

                                DataRow nRow = gisGridView.GetDataRow(startRowIndex);
                                nRow[elem.FieldName] = value;

                                indexSource++;
                                if (indexSource == sourceTable.Columns.Count) {
                                    sourceRowIndex++;
                                    startRowIndex++;
                                    break;
                                }
                            }
                        } catch (Exception ex) {
                            if (!barCheckItem1.Checked) {
                                int row1 = row;
                                while (row1 >= 0) {
                                    gisGridView.DeleteRow(0);
                                    row1--;
                                }
                            } else {
                                gridControl1.BeginUpdate();
                                gridControl1.DataSource = copyDs;
                                gridControl1.DataMember = copyDs.Tables[0].TableName;
                                gridControl1.EndUpdate();
                            }
                            MessageBox.Show(
                                "Ошибка в корректности данных, проверьте соотвествую ли данные шаблону!",
                                "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                            return;
                        }

                    }
                    if (barCheckItem1.Checked) {
                        long? id = CommonUtil.toLongNull(gisGridView.GetDataRow(startRowIndex - 1)["WELL_LOG_RESULT_LAYER_S"]);
                        if (id != null)
                            updateList.Push(id);
                        destRow = gisGridView.GetDataRow(startRowIndex - 1);
                    }
                    try {
                        destRow["SOURCE_S"] = CommonUtil.toDbId(Dict.getIdByName((long)PXPConstants.SourcesData.sdParent, CommonUtil.toStr(destRow["SOURCE_NAME"]), false));
                        destRow["LAYER_S"] = CommonUtil.toDbId(_layerCtrlForCheck.getIdByName(CommonUtil.toStr(destRow["LAYER_NAME"])));
                        if (destRow["LAYER_S"] != null && destRow["LAYER_S"] != DBNull.Value) {
                            destRow["LAYER_FIELD_S"] = Well.getFieldId(_wellId);
                        }
                        destRow["LITHOLOGY_S"] = CommonUtil.toDbId(Dict.getIdByName(PXPConstants.LITHOLOGY_PARENT, CommonUtil.toStr(destRow["LITHOLOGY_NAME"]), false));
                        destRow["SATURATION_S"] = CommonUtil.toDbId(Dict.getIdByName(PXPConstants.SATURATION_PARENT, CommonUtil.toStr(destRow["SATURATION_NAME"]), false));
                        destRow["CURRENT_SATURATION_S"] = CommonUtil.toDbId(Dict.getIdByName(PXPConstants.SATURATION_PARENT, CommonUtil.toStr(destRow["CURRENT_SATURATION_NAME"]), false));

                        int rowNumber = row + 1;

                        if (destRow["SOURCE_S"] == null || destRow["SOURCE_S"] == DBNull.Value) {
                            throw new Exception("Пустое поле 'Источник' в строке номер " + rowNumber);
                        }
                        if (destRow["LAYER_S"] == null || destRow["LAYER_S"] == DBNull.Value) {
                            throw new Exception("Пустое поле 'Пласт' в строке номер " + rowNumber);
                        }
                        if (destRow["LITHOLOGY_S"] == null || destRow["LITHOLOGY_S"] == DBNull.Value) {
                            throw new Exception("Пустое поле 'Литология' в строке номер " + rowNumber);
                        }
                        if (destRow["SATURATION_S"] == null || destRow["SATURATION_S"] == DBNull.Value) {
                            throw new Exception("Пустое поле 'Текущее насыщение' в строке номер " + rowNumber);
                        }
                        if ((destRow["PERMEABILITY"] != null) && (oneMoreCheckPermeability == true) && destRow["PERMEABILITY"] != DBNull.Value) {
                            if (CommonUtil.toDouble(destRow["PERMEABILITY"]) < 0.01)
                                MessageBox.Show("Внимание", "Проницаемость слишком низкая! Проницаемость = " + CommonUtil.toDouble(destRow["PERMEABILITY"]), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            oneMoreCheckPermeability = false;
                        }
                    }
                    catch (Exception ex) {
                        if (!barCheckItem1.Checked) {
                            int row1 = row;
                            while (row1 >= 0) {
                                gisGridView.DeleteRow(0);
                                row1--;
                            }
                        } else {
                            gridControl1.BeginUpdate();
                            gridControl1.DataSource = copyDs;
                            gridControl1.DataMember = copyDs.Tables[0].TableName;
                            gridControl1.EndUpdate();
                        }
                        MessageBox.Show(ex.Message, "Ошибка в загружаемых данных", MessageBoxButtons.OK);
                        return;
                    }

                }
                if (rowCount < sourceTable.Rows.Count && barCheckItem1.Checked == true) {
                    MessageBox.Show("В вставляемых данных больше строк чем в таблице. Не было вставлено " + (sourceTable.Rows.Count - rowCount).ToString() + " строк!", "Ошибка в загружаемых данных!", MessageBoxButtons.OK);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
            } finally {
                gridControl1.EndUpdate();
            }
            int? notUniqRow = ((DataView)gisGridView.DataSource).isColumnUniq();
            if (notUniqRow != null) {
                int row = sourceTable.Rows.Count;
                if (!barCheckItem1.Checked) {
                    while (row > 0) {
                        gisGridView.DeleteRow(0);
                        row--;
                    }
                } else {
                    gridControl1.BeginUpdate();
                    gridControl1.DataSource = copyDs;
                    gridControl1.DataMember = copyDs.Tables[0].TableName;
                    gridControl1.EndUpdate();
                }
                MessageBox.Show("Ошибка уникальности в " + notUniqRow + " строке", "Ошибка в загружаемых данных", MessageBoxButtons.OK);
            }
        }

        #endregion

        #region Обработчики событий

        /// <summary>
        /// Загрузка компонента
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void wellGisControlLoad(object sender, EventArgs e) {
            //gisGridView.BestFitColumns();
            //columnFilterCtrl.init(Component, gisGridView);  
            /*columnFilterCtrl.setTemplate(GridColumnSelectDesc.read(Component)[0]);
            gisGridView.Invalidate();
            Invalidate();
            Refresh();*/
        }

        /// <summary>
        /// Обрабочик нажатия на кнопку "Edit"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barBtnEditItemClick(object sender, ItemClickEventArgs e) {
            DataRow dataRow = getDataRow();
            if (dataRow != null) {
                showWellGisForm(dataRow);
            }
        }

        /// <summary>
        /// Обработчик двойного клика на grid'e
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gridControl1DoubleClick(object sender, EventArgs e) {
            if (Rights.isAllowUpdate) {
                GridHitInfo info = gisGridView.CalcHitInfo(((MouseEventArgs)e).Location);
                if (info.InRow) {
                    DataRow dataRow = getDataRow();
                    showWellGisForm(dataRow);
                }
            }
        }

        /// <summary>
        /// Обработка добавления файлов с интерпретациями
        /// </summary>
        private void barBtnAddFilesItemClick(object sender, ItemClickEventArgs e) {
            var rowsIds = getSelectedRowsId();
            var rowsDesc = getSelectedRowsDesc(true);

            _uploadForm = new FileUploadFrom { Text = @"Прикрепить файлы с интерпретациями" };
            var result = _uploadForm.showDialog(PXPClientCommonUtil.getLoaderSvc().attachFileWellGis, rowsDesc, rowsIds);
            if (result == FileUploadFrom.UploadResult.urSuccess) {
                gridControl1.BeginUpdate();
                foreach (int rowIndex in gisGridView.GetSelectedRows()) {
                    var row = gisGridView.GetDataRow(rowIndex);
                    row["HAS_FILES"] = "Есть файлы";
                }
                gridControl1.EndUpdate();
            } else if (result == FileUploadFrom.UploadResult.urNeedReload) {
                startLoading();
            }
        }

        /// <summary>
        /// Отрисовка иконки наличия прикрепленных файлов
        /// </summary>
        protected void onGridViewCustomDrawCell(object sender, RowCellCustomDrawEventArgs e) {
            if (string.CompareOrdinal(e.Column.FieldName, "HAS_FILES") != 0 || e.RowHandle < 0) {
                return;
            }
            if (string.CompareOrdinal(CommonUtil.toStr(e.CellValue), "Есть файлы") == 0) {
                Bitmap bmp = Resources.attach;
                var bounds = new Rectangle(e.Bounds.Left - 1, e.Bounds.Top - 1, e.Bounds.Width + 2, e.Bounds.Height + 2);
                e.Graphics.DrawImage(bmp, bounds.Left + (bounds.Width - bmp.Width) / 2,
                    bounds.Top + (bounds.Height - bmp.Width) / 2,
                    bmp.Width, bmp.Width);
            }
            e.Handled = true;
        }

        /// <summary>
        /// Отображение редактора с файлами
        /// </summary>
        private void onShowingEditor(object sender, CancelEventArgs e) {
            if (gisGridView.FocusedColumn == null ||
                string.CompareOrdinal(gisGridView.FocusedColumn.Name, gridColumnHasFiles.Name) != 0 ||
                gisGridView.FocusedRowHandle < 0) {
                e.Cancel = true;
                return;
            }
            if (string.CompareOrdinal(gisGridView.GetFocusedDisplayText(), "Есть файлы") != 0) {
                e.Cancel = true;
                return;
            }
            //получаем список, прикрепленных файлов
            DataRow selectedRow = gisGridView.GetDataRow(gisGridView.FocusedRowHandle);
            long wellLogResultLayerId = CommonUtil.toLong(selectedRow["WELL_LOG_RESULT_LAYER_S"]);
            byte[] filesBytes = PXPClientCommonUtil.getDictService().getWellInterpretationGisFilesList(wellLogResultLayerId);

            DataSet dataSetFiles = CommonUtil.decompressData(filesBytes);
            if (CommonUtil.IsEmptyDataSet(dataSetFiles)) {
                e.Cancel = true;
                return;
            }

            var form = new FilesDownloadForm();
            string[] rowDesc = getSelectedRowsDesc(false);
            if (rowDesc.Length != 0) {
                if (Component.Rights.isAllowDelete) {
                    form.ShowDialog(dataSetFiles, getFileBody, rowDesc[0], deattachFile);
                } else {
                    form.ShowDialog(dataSetFiles, getFileBody);
                }
                if (!form.HasFiles) {
                    var index = gisGridView.GetSelectedRows()[0];
                    var row = gisGridView.GetDataRow(index);
                    row["HAS_FILES"] = "Нет файлов";
                }
            }
            e.Cancel = true;
        }

        private void gisGridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            var view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            if (view == null) return;

            // Приводим имя колонки к нижнему регистру, чтобы избежать ошибок регистра букв
            string fieldName = e.Column.FieldName.ToLower();

            bool isTop = (fieldName == "top" || fieldName == "top_dsrd");
            bool isBase = (fieldName == "base" || fieldName == "base_dsrd");

            // Если это не колонки глубин — ничего не делаем
            if (!isTop && !isBase) return;

            // Для первой строки предыдущей нет
            if (e.RowHandle <= 0) return;

            // Получаем текущую и предыдущую строки
            var currentRowView = view.GetRow(e.RowHandle) as DataRowView;
            var prevRowView = view.GetRow(e.RowHandle - 1) as DataRowView;

            if (currentRowView == null || prevRowView == null) return;

            // Ищем точное имя колонки в объекте данных (так как в DataRowView регистр тоже важен)
            string exactFieldName = e.Column.FieldName;

            double? currentValue = toDoubleNull(currentRowView[exactFieldName]);
            double? prevValue = toDoubleNull(prevRowView[exactFieldName]);

            // Если значения есть и они совпадают с предыдущей строкой — красим
            if (currentValue != null && prevValue != null)
            {
                if (Math.Abs(currentValue.Value - prevValue.Value) < 0.001)
                {
                    e.Appearance.BackColor = Color.MistyRose;
                    e.Appearance.ForeColor = Color.DarkRed;

                    // Включаем флаги DevExpress принудительно
                    e.Appearance.Options.UseBackColor = true;
                    e.Appearance.Options.UseForeColor = true;
                    e.Appearance.Options.UseFont = true;
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
            }
        }

        // Исправленный метод конвертации (без дублирования переменных)
        private double? toDoubleNull(object obj)
        {
            if (obj == null || obj == DBNull.Value) return null;

            string strValue = obj.ToString().Trim();
            double result;

            // 1. Пробуем распарсить с инвариантной культурой (с точкой)
            if (double.TryParse(strValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result))
            {
                return result;
            }

            // 2. Пробуем распарсить с текущей культурой системы (с запятой)
            if (double.TryParse(strValue, out result))
            {
                return result;
            }

            // 3. Если в базе лежит альтернативный разделитель, делаем замену вручную
            strValue = strValue.Contains(".") ? strValue.Replace(".", ",") : strValue.Replace(",", ".");
            if (double.TryParse(strValue, out result))
            {
                return result;
            }

            return null;
        }

        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e) {
            barButtonPaste_ItemClick(sender, null);
        }

        protected override void deleteRecord() {
            //var row = gisGridView.GetDataRow(gisGridView.FocusedRowHandle);
            //long? id = CommonUtil.toLongNull(row["WELL_LOG_RESULT_LAYER_S"]);

            var selectedRows = gisGridView.GetSelectedRows();
            foreach (var selectedRow in selectedRows) {
                var row = gisGridView.GetDataRow(selectedRow);
                long? id = CommonUtil.toLongNull(row["WELL_LOG_RESULT_LAYER_S"]);
                
                if (id != null){
                    deleteList.Push((long)id);
                }
            }
            //if (id != null) {
            //    deleteList.Push((long)id);
            //}
            foreach (var selectedRow in selectedRows) {
                gisGridView.DeleteRow(selectedRows[0]);
            }
           //deleteList.Push(gisGridView. gisGridView.FocusedRowHandle);

        }

        private void barButtonSave_ItemClick(object sender, ItemClickEventArgs e) {
            DataSet set = new DataSet();
            var dt = ((DataView)gisGridView.DataSource).ToTable();
            set.Tables.Add(dt);
            var res = CommonUtil.compressData(set);            

            if(updateList.Count>0) {
                long?[] idList = new long?[updateList.Count];
                updateList.CopyTo(idList,0);
                PXPClientCommonUtil.getDictService().updateWellInterpretationLine(res, idList);
                updateList = new Stack<long?>();
            }
            while (deleteList.Count > 0) {
                PXPClientCommonUtil.getDictService().deleteWellInterpretationLine(deleteList.Pop());
            }
            
            long[] id = PXPClientCommonUtil.getDictService().insertWellInterpretationLine(res);

            int j = 0;
            for (int i = 0; i < gisGridView.RowCount; i++) {
                if (CommonUtil.toLongNull(((DataView)gisGridView.DataSource).Table.Rows[i]["well_log_result_layer_s"]) == null) {
                    ((DataView)gisGridView.DataSource)[i]["well_log_result_layer_s"] = id[j];
                    //gisGridView.SetRowCellValue(i, gisGridView.Columns["well_log_result_layer_s"], ((object)id[j]));
                    j++;
                }
            }   

        }
        #endregion

        private void inklSelectSourceRigisCtrlEditValueChanged(object sender, EventArgs e)
        {
            if (Visible)
            {
                isNeedLoad = true;
                string sourceStr = CommonUtil.toStr(barItemWellGisVersion.EditValue);
                long? sourceId;

                if (!sourceStr.Equals("Основная версия")) {
                    sourceId = Dict.getIdByName(FwConstants.DICT_SOURCE, sourceStr);
                    IsWellLogHist = true;
                }
                else {
                    sourceId = 0;
                    IsWellLogHist = false;
                }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
                loadWellGis(true, sourceId);
            }
        }

        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e) {
            _isNeedLoad = true;
            _template = columnFilterCtrl.DataTable.Copy();
            Refresh();
        }
        
        /// <summary>
        /// Обработка расчета абсолютных отметок
        /// </summary>
        private void barBtnCalcAbsDepthItemClick(object sender, ItemClickEventArgs e) {
            long wellId = ((GEONavigationCtrl)Component.NavigationControl).getCurWellId();
            PXPClientCommonUtil.getWellInfoService().calcWellLogResultLayerAbsDepth(wellId);
            _isNeedLoad = true;
            Refresh();
        }

        private void updateGridWellNodeHist() {
            WellGisVersionEditCtrlRepItem1.SelectCtrl.init(_wellId); 
        }

        /// <summary>
        /// Возвращает тип прокетного документа и дату.
        /// </summary>
        public string getActual()
        {
            DataRow[] rows = null;

            return rows.Length > 0
                    ? rows[0][ValueFieldName].ToString()
                    : null;

        }

        private void mergeRigisVersion(object sender, ItemClickEventArgs e)
        {
            try{
                PXPClientCommonUtil.getWellInfoService().mergeWellLogSourceGis(_wellId);
                MessageBox.Show("Успешно", "Сводный замер расчитан", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updateBarItemGisVersion(_wellId);
            }
            catch (Exception ex) {
                MessageBox.Show("Ошибка", "Ошибка при сохранении РИГИС", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new Exception("Ошибка " + ex);
            }
        }

        private void saveRigisVersion(object sender, ItemClickEventArgs e) {
            try {
                DataSet set = new DataSet();
                var dt = ((DataView)gisGridView.DataSource).ToTable();
                set.Tables.Add(dt);
                var res = CommonUtil.compressData(set);
                long[] id =  PXPClientCommonUtil.getDictService().insertWellLogHistInterpretationLine(res);

                MessageBox.Show("Данные РИГИС успешно сохранены", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updateBarItemGisVersion(_wellId);
            }
            catch (Exception ex) {
                MessageBox.Show("Ошибка", "Ошибка при сохранении РИГИС", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new Exception("Ошибка " + ex);
            }
        }
    }

    public static class DataTableExtension {
        static public int? isColumnUniq(this DataView dt) {
            int? rowIndex = 0;
            foreach (DataRow row in dt.Table.Rows) {
                rowIndex++;
                var equalList = (from DataRow a in dt.Table.Rows
                                 where CommonUtil.toLong(a["WELL_S"]) == CommonUtil.toLong(row["WELL_S"]) && CommonUtil.toLong(a["SOURCE_S"]) == CommonUtil.toLong(row["SOURCE_S"]) &&
                                CommonUtil.toDouble(a["TOP"]) == CommonUtil.toDouble(row["TOP"]) && CommonUtil.toDouble(a["BASE"]) == CommonUtil.toDouble(row["BASE"])
                                select a).ToList();
                if (equalList.Count > 1)
                    return rowIndex;
            }
            return null;
        }
    }
}