using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

using FirebirdSql.Data.FirebirdClient;

namespace RTP3Tools
{
    public partial class FormMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
                
        public string dbconnectionStringIESBK;
        
        public FormMain()
        {
            InitializeComponent();
        
            dbconnectionStringIESBK = "Data Source=;Initial Catalog=iesbk;User ID=;Password=";        
        }
        
        private void ribbonControl1_Merge(object sender, DevExpress.XtraBars.Ribbon.RibbonMergeEventArgs e)
        {
            ribbonControl1.SelectedPage = ribbonControl1.MergedRibbon.Pages[0];
        }

        
        // ручная загрузка данных
        private void barButtonItem8_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            /*FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Загрузка данных";

            form1.Show();*/
        }
               
                                                
        // РТП-3 - Отчеты из РТП-3 - Реестр потребителей
        private void barButtonItemRTP3_ConsumersInfo_ItemClick(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog openFileDialogRTP = new OpenFileDialog();
            openFileDialogRTP.Title = "Выберите файл базы данных РТП-3";
            openFileDialogRTP.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
            openFileDialogRTP.FileName = "";

            if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
            {
                splashScreenManager1.ShowWaitForm();

                splashScreenManager1.SetWaitFormDescription("подключение к базе данных...");

                // создаем соединение и загружаем данные
                FbConnection connection = new FbConnection();

                if (openFileDialogRTP.FileName.Contains("serverpfb"))
                {
                    connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", 
                        openFileDialogRTP.FileName.Replace("\\\\serverpfb", "E:\\"), ";DataSource=", "", ";Charset=UTF8;");                    
                }
                else
                {
                    connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", openFileDialogRTP.FileName, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");
                }
                                                
                FbCommand command = new FbCommand();
                command.Connection = connection;

                // читаем запрос из файла (для теста, потом убрать!!!)
                /*command.CommandText = "";
                string line;
                StreamReader sr = new StreamReader(@"sql_RTP3.txt");
                while ((line = sr.ReadLine()) != null)
                {
                    command.CommandText = String.Concat(command.CommandText, line, Environment.NewLine);
                }*/
                                
                command.CommandText = String.Concat(
                    "SELECT ",
                    "t_LVConsumersInfo.\"GUID\", t_LVConsumersInfo.\"Contract\", t_LVConsumersInfo.\"Customer\", t_LVConsumersInfo.\"PointName\", t_LVConsumersInfo.\"Address\",", " ",
                    "v_LVConsumersA_LVNodes_GUID, v_LVConsumersA_LVNodes_OwnerGUID, v_LVConsumersA_LVNodes_Name,", " ",
                    "v_LVFiders_GUID, v_LVFiders_OwnerGUID, v_LVFiders_Name,", " ",
                    "v_Nodes_GUID, v_Nodes_OwnerGUID, v_Nodes_Name, v_Nodes_onBalance, v_Nodes_Transf,", " ",
                    "v_Transforms2_Ident, v_Transforms2_Unom, v_Transforms2_TypeTR, v_Transforms2_Snom,", " ",
                    "v_Fiders_GUID, v_Fiders_OwnerGUID, v_Fiders_Name,", " ",
                    "v_Sections_GUID, v_Sections_OwnerGUID, v_Sections_Name, v_Sections_Unom,", " ",
                    "v_Centers_GUID, v_Centers_OwnerGUID, v_Centers_Name", " ",

                    "FROM \"LVConsumersInfo\" AS t_LVConsumersInfo", " ",

                    "LEFT JOIN", " ",

                    "(SELECT", " ",
                    "t_LVConsumersA_LVNodes.\"GUID\" AS v_LVConsumersA_LVNodes_GUID, t_LVConsumersA_LVNodes.\"OwnerGUID\" AS v_LVConsumersA_LVNodes_OwnerGUID, t_LVConsumersA_LVNodes.\"Name\" AS v_LVConsumersA_LVNodes_Name,", " ",
                    "v_LVFiders_GUID, v_LVFiders_OwnerGUID, (CASE WHEN v_LVFidersName IS NULL THEN 'подключен к ТП' ELSE v_LVFidersName END) AS v_LVFiders_Name", " ",
                    "FROM", " ",
                    "(SELECT \"GUID\", \"OwnerGUID\", \"Name\" FROM \"LVConsumersA\"", " ",
                    "UNION", " ",
                    "SELECT \"GUID\", \"OwnerGUID\", \"Name\" FROM \"LVNodes\") AS t_LVConsumersA_LVNodes", " ",
                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_LVFiders_GUID, \"OwnerGUID\" AS v_LVFiders_OwnerGUID, \"Name\" AS v_LVFidersName FROM \"LVFiders\") AS t_LVFiders", " ",
                    "ON t_LVConsumersA_LVNodes.\"OwnerGUID\" = t_LVFiders.v_LVFiders_GUID", " ",
                    ")", " ",
                    "AS t_LVConsumersA_LVNodes_Fiders", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Nodes_GUID, \"OwnerGUID\" AS v_Nodes_OwnerGUID, \"Name\" AS v_Nodes_Name,", " ",
                    "(CASE", " ",
                    "WHEN \"Nodes\".\"onBalance\" = 0 THEN 'на балансе'", " ",
                    "WHEN \"Nodes\".\"onBalance\" = 1 THEN 'потребительская'", " ",
                    "WHEN \"Nodes\".\"onBalance\" = 2 THEN 'ССО'", " ",
                    "WHEN \"Nodes\".\"onBalance\" = 3 THEN 'ССП'", " ",
                    "END) AS v_Nodes_onBalance,", " ",
                    "\"Transf\" AS v_Nodes_Transf", " ",
                    "FROM \"Nodes\") AS t_Nodes", " ",
                    "ON(t_LVConsumersA_LVNodes_Fiders.v_LVConsumersA_LVNodes_OwnerGUID = t_Nodes.v_Nodes_GUID) OR(t_LVConsumersA_LVNodes_Fiders.v_LVFiders_OwnerGUID = t_Nodes.v_Nodes_GUID)", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"Ident\" AS v_Transforms2_Ident, \"Unom\" AS v_Transforms2_Unom, \"TypeTR\" AS v_Transforms2_TypeTR, \"Snom\" AS v_Transforms2_Snom", " ",
                    "FROM \"Transforms2\") AS t_Transforms2", " ",
                    "ON t_Nodes.v_Nodes_Transf = t_Transforms2.v_Transforms2_Ident", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Fiders_GUID, \"OwnerGUID\" AS v_Fiders_OwnerGUID, \"Name\" AS v_Fiders_Name", " ",
                    "FROM \"Fiders\") AS t_Fiders", " ",
                    "ON t_Nodes.v_Nodes_OwnerGUID = t_Fiders.v_Fiders_GUID", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Sections_GUID, \"OwnerGUID\" AS v_Sections_OwnerGUID, \"Name\" AS v_Sections_Name, \"Unom\" AS v_Sections_Unom", " ",
                    "FROM \"Sections\") AS t_Sections", " ",
                    "ON t_Fiders.v_Fiders_OwnerGUID = t_Sections.v_Sections_GUID", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Centers_GUID, \"OwnerGUID\" AS v_Centers_OwnerGUID, \"Name\" AS v_Centers_Name", " ",
                    "FROM \"Centers\") AS t_Centers", " ",
                    "ON t_Sections.v_Sections_OwnerGUID = t_Centers.v_Centers_GUID", " ",

                    "ON t_LVConsumersInfo.guid = v_LVConsumersA_LVNodes_GUID"
                    );

                FbDataAdapter dataAdapter = new FbDataAdapter();
                dataAdapter.SelectCommand = command;
                DataTable dt = new DataTable();
                connection.Open();

                splashScreenManager1.SetWaitFormDescription("загрузка данных...");
                dataAdapter.Fill(dt);

                connection.Close();

                // формируем отчет
                splashScreenManager1.SetWaitFormDescription("формирование отчета...");

                FormReportGrid formReport = null;
                formReport = new FormReportGrid();
                formReport.MdiParent = this;
                formReport.Text = String.Concat("Реестр потребителей РТП-3 (", openFileDialogRTP.FileName, ")");
                IWorkbook workbook = formReport.spreadsheetControl1.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                workbook.History.IsEnabled = false;
                formReport.spreadsheetControl1.BeginUpdate();

                int row_columns_group_names = 0;
                int row_start_data = 3;
                int row_columns_names = 1;
                int row_columns_index = 2;

                // выводим и форматируем названия групп столбцов
                worksheet.Rows[row_columns_group_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_group_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[row_columns_group_names].Font.Name = "Arial";
                worksheet.Rows[row_columns_group_names].Font.Size = 8;
                
                worksheet[row_columns_group_names, 0].SetValue("№ п/п"); worksheet.MergeCells(worksheet.Range["A1:A2"]);
                worksheet[row_columns_group_names, 1].SetValue("Данные потребителя"); worksheet.MergeCells(worksheet.Range["B1:F1"]);
                worksheet[row_columns_group_names, 6].SetValue("Наименование на схеме"); worksheet.MergeCells(worksheet.Range["G1:I1"]);
                worksheet[row_columns_group_names, 9].SetValue("Наименование фидера 0,4 кВ"); worksheet.MergeCells(worksheet.Range["J1:L1"]);
                worksheet[row_columns_group_names, 12].SetValue("ТП 6(10)/0,4 кВ"); worksheet.MergeCells(worksheet.Range["M1:Q1"]);
                worksheet[row_columns_group_names, 17].SetValue("Свойства ТП 6(10)/0,4 кВ"); worksheet.MergeCells(worksheet.Range["R1:U1"]);
                worksheet[row_columns_group_names, 21].SetValue("Наименование фидера ВН"); worksheet.MergeCells(worksheet.Range["V1:X1"]);
                worksheet[row_columns_group_names, 24].SetValue("Секция шин"); worksheet.MergeCells(worksheet.Range["Y1:AB1"]);
                worksheet[row_columns_group_names, 28].SetValue("Наименование ПС"); worksheet.MergeCells(worksheet.Range["AC1:AE1"]);

                // выводим названия столбцов
                string[] columns_names = {
                    "", "GUID", "Код точки поставки", "Наименование", "Точка учета", "Адрес",
                    "GUID", "OwnerGUID", "Наименование на схеме",
                    "GUID", "OwnerGUID", "Наименование фидера 0,4 кВ",
                    "GUID", "OwnerGUID", "Наименование ТП", "Балансовая принадлежность", "Номер типа трансформатора",
                    "Номер типа трансформатора", "Напряжение", "Тип", "Мощность",
                    "GUID", "OwnerGUID", "Наименование фидера ВН",
                    "GUID", "OwnerGUID", "Номер секции шин", "Номинальное напряжение",
                    "GUID", "OwnerGUID", "Наименование ПС"
                    };

                for (int col = 0; col < 31; col++)
                {
                    worksheet[row_columns_names, col].SetValue(columns_names[col]);                    
                    worksheet[row_columns_index, col].SetValue((col + 1).ToString());
                }

                // форматируем заголовок таблицы                                
                worksheet.Rows[row_columns_names].Alignment.WrapText = true;                
                worksheet.Rows[row_columns_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[row_columns_names].Font.Name = "Arial";
                worksheet.Rows[row_columns_names].Font.Size = 8;

                worksheet.Rows[row_columns_index].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_index].Font.Name = "Arial";
                worksheet.Rows[row_columns_index].Font.Size = 8;

                worksheet.Range["A1:AE3"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                // выводим данные
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    worksheet.Rows[row + row_start_data].Font.Name = "Arial";
                    worksheet.Rows[row + row_start_data].Font.Size = 8;

                    int row_index = row + 1;
                    worksheet[row + row_start_data, 0].SetValue(row_index.ToString());

                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        worksheet[row + row_start_data, col + 1].SetValue(dt.Rows[row][col].ToString());
                    }
                }

                // выравнивание столбцов
                worksheet.Columns[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                splashScreenManager1.SetWaitFormDescription("выравнивание ширины столбцов...");
                worksheet.Columns.AutoFit(0, dt.Columns.Count);

                formReport.spreadsheetControl1.EndUpdate();

                splashScreenManager1.CloseWaitForm();
                                
                formReport.Show();
                
                this.ribbonControl1.SelectedPage = this.ribbonControl1.MergedPages[0];
            }
                        
        } // private void barButtonItemRTP3_ConsumersInfo_ItemClick(object sender, ItemClickEventArgs e)

        // замена кода точки поставки в выбранной БД РТП-3
        private void barButtonItemRTP3_ReplacePointCode_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (MessageBox.Show("Вы сделали резервную копию базы данных?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                FormLoadDataGrid formLoadDataGrid = null;
                formLoadDataGrid = new FormLoadDataGrid();
                formLoadDataGrid.MdiParent = this;
                formLoadDataGrid.Text = String.Concat("Загрузка данных (замена кода точки поставки)");
                
                formLoadDataGrid.Show();

                this.ribbonControl1.SelectedPage = this.ribbonControl1.MergedPages[0];

                MessageBox.Show("Откройте файл с данными для замены");
            }
        } // private void barButtonItemRTP3_ReplacePointCode_ItemClick(object sender, ItemClickEventArgs e)

        // РТП-3 - Отчеты из РТП-3 - Реестр коммутационных аппаратов
        private void barButtonItemRTP3_KommApparatInfo_ItemClick(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog openFileDialogRTP = new OpenFileDialog();
            openFileDialogRTP.Title = "Выберите файл базы данных РТП-3";
            openFileDialogRTP.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
            openFileDialogRTP.FileName = "";

            if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
            {
                splashScreenManager1.ShowWaitForm();

                splashScreenManager1.SetWaitFormDescription("подключение к базе данных...");

                // создаем соединение и загружаем данные
                FbConnection connection = new FbConnection();

                if (openFileDialogRTP.FileName.Contains("serverpfb"))
                {
                    connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=",
                        openFileDialogRTP.FileName.Replace("\\\\serverpfb", "E:\\"), ";DataSource=", "SERVERPFB", ";Charset=UTF8;");
                }
                else
                {
                    connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", openFileDialogRTP.FileName, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");
                }

                FbCommand command = new FbCommand();
                command.Connection = connection;

                FbDataAdapter dataAdapter = new FbDataAdapter();
                dataAdapter.SelectCommand = command;
                DataTable dtKA = new DataTable();
                connection.Open();

                splashScreenManager1.SetWaitFormDescription("загрузка данных...");

                //---- формируем словари видов КА по уровню напряжения (начало)

                // формируем таблицу видов КА
                command.CommandText = String.Concat(
                    "SELECT ",
                    "DISTINCT v_SwitchGN_Name", " ",

                    "FROM \"Lines\" AS t_Lines", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"Ident\" AS v_SwitchGR_Ident, \"Marka\" AS v_SwitchGR_Marka, \"Unom\" AS v_SwitchGR_Unom, \"SwitchGType\" AS v_SwitchGR_SwitchGType", " ",
                    "FROM \"SwitchGR\") AS t_SwitchGR", " ",
                    "ON t_Lines.\"Provod\" = t_SwitchGR.v_SwitchGR_Ident", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"Ident\" AS v_SwitchGN_Ident, \"Name\" AS v_SwitchGN_Name", " ",
                    "FROM \"SwitchGN\") AS t_SwitchGN", " ",
                    "ON t_SwitchGR.v_SwitchGR_SwitchGType = t_SwitchGN.v_SwitchGN_Ident", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Fiders_GUID, \"OwnerGUID\" AS v_Fiders_OwnerGUID, \"Name\" AS v_Fiders_Name", " ",
                    "FROM \"Fiders\") AS t_Fiders", " ",
                    "ON t_Lines.\"OwnerGUID\" = t_Fiders.v_Fiders_GUID", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Sections_GUID, \"OwnerGUID\" AS v_Sections_OwnerGUID, \"Name\" AS v_Sections_Name, \"Unom\" AS v_Sections_Unom", " ",
                    "FROM \"Sections\") AS t_Sections", " ",
                    "ON t_Fiders.v_Fiders_OwnerGUID = t_Sections.v_Sections_GUID", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Centers_GUID, \"OwnerGUID\" AS v_Centers_OwnerGUID, \"Name\" AS v_Centers_Name", " ",
                    "FROM \"Centers\") AS t_Centers", " ",
                    "ON t_Sections.v_Sections_OwnerGUID = t_Centers.v_Centers_GUID", " ",

                    "WHERE v_SwitchGR_Ident IS NOT null", " ",

                    "ORDER BY v_SwitchGN_Name"
                    );

                dataAdapter.SelectCommand = command;
                dataAdapter.Fill(dtKA);
                                
                Dictionary<string, int> countKA6 = new Dictionary<string, int>();
                Dictionary<string, int> countKA10 = new Dictionary<string, int>();
                Dictionary<string, int> countKA35 = new Dictionary<string, int>();
                Dictionary<string, int> countKA110 = new Dictionary<string, int>();

                for (int row = 0; row < dtKA.Rows.Count; row++)
                {
                    countKA6.Add(dtKA.Rows[row][0].ToString(), 0);
                    countKA10.Add(dtKA.Rows[row][0].ToString(), 0);
                    countKA35.Add(dtKA.Rows[row][0].ToString(), 0);
                    countKA110.Add(dtKA.Rows[row][0].ToString(), 0);
                }
                //---- формируем словари видов КА по уровню напряжения (конец)


                command.CommandText = String.Concat(
                    "SELECT ",
                    "t_Lines.\"OwnerGUID\", t_Lines.\"Provod\",",
                    "v_SwitchGR_Ident, v_SwitchGR_Marka, v_SwitchGR_Unom, v_SwitchGR_SwitchGType,",
                    "v_SwitchGN_Ident, v_SwitchGN_Name,",
                    "v_Fiders_GUID, v_Fiders_OwnerGUID, v_Fiders_Name,",
                    "v_Sections_GUID, v_Sections_OwnerGUID, v_Sections_Name, v_Sections_Unom,",
                    "v_Centers_GUID, v_Centers_OwnerGUID, v_Centers_Name", " ",

                    "FROM \"Lines\" AS t_Lines", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"Ident\" AS v_SwitchGR_Ident, \"Marka\" AS v_SwitchGR_Marka, \"Unom\" AS v_SwitchGR_Unom, \"SwitchGType\" AS v_SwitchGR_SwitchGType", " ",
                    "FROM \"SwitchGR\") AS t_SwitchGR", " ",
                    "ON t_Lines.\"Provod\" = t_SwitchGR.v_SwitchGR_Ident", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"Ident\" AS v_SwitchGN_Ident, \"Name\" AS v_SwitchGN_Name", " ",
                    "FROM \"SwitchGN\") AS t_SwitchGN", " ",
                    "ON t_SwitchGR.v_SwitchGR_SwitchGType = t_SwitchGN.v_SwitchGN_Ident", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Fiders_GUID, \"OwnerGUID\" AS v_Fiders_OwnerGUID, \"Name\" AS v_Fiders_Name", " ",
                    "FROM \"Fiders\") AS t_Fiders", " ",
                    "ON t_Lines.\"OwnerGUID\" = t_Fiders.v_Fiders_GUID", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Sections_GUID, \"OwnerGUID\" AS v_Sections_OwnerGUID, \"Name\" AS v_Sections_Name, \"Unom\" AS v_Sections_Unom", " ",
                    "FROM \"Sections\") AS t_Sections", " ",
                    "ON t_Fiders.v_Fiders_OwnerGUID = t_Sections.v_Sections_GUID", " ",

                    "LEFT JOIN", " ",
                    "(SELECT \"GUID\" AS v_Centers_GUID, \"OwnerGUID\" AS v_Centers_OwnerGUID, \"Name\" AS v_Centers_Name", " ",
                    "FROM \"Centers\") AS t_Centers", " ",
                    "ON t_Sections.v_Sections_OwnerGUID = t_Centers.v_Centers_GUID", " ",

                    "WHERE v_SwitchGR_Ident IS NOT null", " ",

                    "ORDER BY v_Centers_Name, v_Fiders_Name"
                    );

                dataAdapter.SelectCommand = command;
                DataTable dt = new DataTable();
                dataAdapter.Fill(dt);

                connection.Close();

                // формируем отчет
                splashScreenManager1.SetWaitFormDescription("формирование отчета...");

                FormReportGrid formReport = null;
                formReport = new FormReportGrid();
                formReport.MdiParent = this;
                formReport.Text = String.Concat("Реестр коммутационных аппаратов РТП-3 (", openFileDialogRTP.FileName, ")");
                IWorkbook workbook = formReport.spreadsheetControl1.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                workbook.History.IsEnabled = false;
                formReport.spreadsheetControl1.BeginUpdate();

                int row_count_svodform = 13;

                int row_columns_group_names = 0+ row_count_svodform;
                int row_start_data = 3 + row_count_svodform;
                int row_columns_names = 1+ row_count_svodform;
                int row_columns_index = 2+ row_count_svodform;

                // выводим и форматируем названия групп столбцов
                worksheet.Rows[row_columns_group_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_group_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[row_columns_group_names].Font.Name = "Arial";
                worksheet.Rows[row_columns_group_names].Font.Size = 8;

                worksheet[row_columns_group_names, 0].SetValue("№ п/п"); worksheet.MergeCells(worksheet.Range["A14:A15"]);
                worksheet[row_columns_group_names, 1].SetValue("Участки сети"); worksheet.MergeCells(worksheet.Range["B14:C14"]);
                worksheet[row_columns_group_names, 3].SetValue("Коммутационный аппарат"); worksheet.MergeCells(worksheet.Range["D14:G14"]);
                worksheet[row_columns_group_names, 7].SetValue("Вид коммутационного аппарата"); worksheet.MergeCells(worksheet.Range["H14:I14"]);                
                worksheet[row_columns_group_names, 9].SetValue("Наименование фидера ВН"); worksheet.MergeCells(worksheet.Range["J14:L14"]);
                worksheet[row_columns_group_names, 12].SetValue("Секция шин"); worksheet.MergeCells(worksheet.Range["M14:P14"]);
                worksheet[row_columns_group_names, 16].SetValue("Наименование ПС"); worksheet.MergeCells(worksheet.Range["Q14:S14"]);

                // выводим названия столбцов
                string[] columns_names = {
                    "", "Код головного участка", "Код типа участка",
                    "Код типа участка", "Марка", "Номинальное напряжение", "Код вида коммутационного аппарата",
                    "Код вида коммутационного аппарата", "Наименование",                    
                    "GUID", "OwnerGUID", "Наименование фидера ВН",
                    "GUID", "OwnerGUID", "Номер секции шин", "Номинальное напряжение",
                    "GUID", "OwnerGUID", "Наименование ПС"
                    };

                for (int col = 0; col < 19; col++)
                {
                    worksheet[row_columns_names, col].SetValue(columns_names[col]);
                    worksheet[row_columns_index, col].SetValue((col + 1).ToString());
                }

                // форматируем заголовок таблицы                                
                worksheet.Rows[row_columns_names].Alignment.WrapText = true;
                worksheet.Rows[row_columns_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[row_columns_names].Font.Name = "Arial";
                worksheet.Rows[row_columns_names].Font.Size = 8;

                worksheet.Rows[row_columns_index].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_index].Font.Name = "Arial";
                worksheet.Rows[row_columns_index].Font.Size = 8;

                worksheet.Range["A14:S16"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                // выводим данные
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    worksheet.Rows[row + row_start_data].Font.Name = "Arial";
                    worksheet.Rows[row + row_start_data].Font.Size = 8;

                    int row_index = row + 1;
                    worksheet[row + row_start_data, 0].SetValue(row_index.ToString());

                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        worksheet[row + row_start_data, col + 1].SetValue(dt.Rows[row][col].ToString());
                    }

                    // формируем свод по видам коммутационных аппаратов в разрезе уровне напряжения
                    string vidKA = dt.Rows[row][7].ToString();
                    double Unom = Convert.ToDouble(dt.Rows[row][14].ToString().Replace(",", "."));

                    switch (Unom)
                    {
                        case 6.0:
                            countKA6[vidKA]++;
                            break;
                        case 10.0:
                            countKA10[vidKA] = countKA10[vidKA] + 1;
                            break;
                        case 35.0:
                            countKA35[vidKA]++;
                            break;
                        case 110.0:
                            countKA110[vidKA]++;
                            break;
                    }
                    //-----------------------------------------------------------------------------
                }

                // выводим свод по видам коммутационных аппаратов в разрезе уровне напряжения
                worksheet.Range["A1:S9"].Font.Name = "Arial";
                worksheet.Range["A1:S9"].Font.Size = 8;
                worksheet.Range["A4:S9"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Range["A4:S9"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                int col_left = 1;
                                
                worksheet[3, 0 + col_left].SetValue("Уровень напряжения");
                worksheet[3, 1 + col_left].SetValue("Итого по видам");
                
                worksheet[4, 0 + col_left].SetValue("Уровень ВН (110 кВ)");
                worksheet[5, 0 + col_left].SetValue("Уровень СН-I (35 кВ)");
                worksheet[6, 0 + col_left].SetValue("Уровень СН-II (10 кВ)");
                worksheet[7, 0 + col_left].SetValue("Уровень СН-II (6 кВ)");
                worksheet[8, 0 + col_left].SetValue("Итого по базе:");

                int totalKA110 = 0;
                int totalKA35 = 0;
                int totalKA10 = 0;
                int totalKA6 = 0;

                for (int i = 0; i < dtKA.Rows.Count; i++)
                {
                    string vidKA = dtKA.Rows[i][0].ToString();
                    worksheet[3, 2 + col_left + i].SetValue(dtKA.Rows[i][0].ToString());

                    // Уровень ВН (110 кВ)
                    if (countKA110[vidKA] > 0)
                    {
                        worksheet[4, 2 + col_left + i].SetValue(countKA110[vidKA].ToString());
                        totalKA110 += countKA110[vidKA];
                    }

                    // Уровень СН-I (35 кВ)
                    if (countKA35[vidKA] > 0)
                    {
                        worksheet[5, 2 + col_left + i].SetValue(countKA35[vidKA].ToString());
                        totalKA35 += countKA35[vidKA];
                    }

                    // Уровень СН-II (10 кВ)
                    if (countKA10[vidKA] > 0)
                    {
                        worksheet[6, 2 + col_left + i].SetValue(countKA10[vidKA].ToString());
                        totalKA10 += countKA10[vidKA];
                    }

                    // Уровень СН-II (6 кВ)
                    if (countKA6[vidKA] > 0)
                    {
                        worksheet[7, 2 + col_left + i].SetValue(countKA6[vidKA].ToString());
                        totalKA6 += countKA6[vidKA];
                    }

                    // итого по столбцу вида КА
                    worksheet[8, 2 + col_left + i].SetValue((countKA110[vidKA] + countKA35[vidKA] + countKA10[vidKA] + countKA6[vidKA]).ToString());
                }

                worksheet[4, 1 + col_left].SetValue(totalKA110.ToString());
                worksheet[5, 1 + col_left].SetValue(totalKA35.ToString());
                worksheet[6, 1 + col_left].SetValue(totalKA10.ToString());
                worksheet[7, 1 + col_left].SetValue(totalKA6.ToString());

                worksheet[8, 1 + col_left].SetValue((totalKA110 + totalKA35 + totalKA10 + totalKA6).ToString());
                //---------------------------------------------------------------------------

                // выравнивание столбцов
                worksheet.Columns[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                splashScreenManager1.SetWaitFormDescription("выравнивание ширины столбцов...");
                worksheet.Columns.AutoFit(0, dt.Columns.Count);

                // выводим заголовок отчета                
                worksheet.Range["A1:A2"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                worksheet[0, 0].SetValue("Свод по видам коммутационных аппаратов в разрезе уровней напряжения");
                worksheet[1, 0].SetValue(openFileDialogRTP.FileName);

                formReport.spreadsheetControl1.EndUpdate();

                splashScreenManager1.CloseWaitForm();

                formReport.Show();

                this.ribbonControl1.SelectedPage = this.ribbonControl1.MergedPages[0];
            }
        } // private void barButtonItemRTP3_KommApparatInfo_ItemClick(object sender, ItemClickEventArgs e)

        // очистка персональных данных в БД РТП-3
        private void barButtonItemRTP3_ClearPersData_ItemClick(object sender, ItemClickEventArgs e)
        {            
                OpenFileDialog openFileDialogRTP = new OpenFileDialog();
                openFileDialogRTP.Title = "Выберите файл базы данных РТП-3";
                openFileDialogRTP.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
                openFileDialogRTP.FileName = "";

                if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
                {
                    splashScreenManager1.ShowWaitForm();
                                    
                    FbConnection connection = new FbConnection();

                    if (openFileDialogRTP.FileName.Contains("serverpfb"))
                    {
                        connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=",
                            openFileDialogRTP.FileName.Replace("\\\\serverpfb", "E:\\"), ";DataSource=", "SERVERPFB", ";Charset=UTF8;");
                    }
                    else
                    {
                        connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", openFileDialogRTP.FileName, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");
                    }

                    connection.Open();

                    FbCommand commandUpdate = new FbCommand();
                    commandUpdate.Connection = connection;

                    string sqlUpdate = "";                                                
                        sqlUpdate = String.Concat(
                                "UPDATE \"LVConsumersInfo\" ",
                                "SET \"LVConsumersInfo\".\"Contract\" = '', " +
                                    "\"LVConsumersInfo\".\"Customer\" = '', " +
                                    "\"LVConsumersInfo\".\"PointName\" = '', " +
                                    "\"LVConsumersInfo\".\"Address\" = ''"
                                );
                

					commandUpdate.CommandText = sqlUpdate;
                            commandUpdate.ExecuteNonQuery();
                            

                    connection.Close();

                    splashScreenManager1.CloseWaitForm();
                    MessageBox.Show("Обработка завершена");
                } // if (openFileDialogRTP.ShowDialog() == DialogResult.OK)            
        } // private void barButtonItemRTP3_ClearPersData_ItemClick(object sender, ItemClickEventArgs e)
    }
}
