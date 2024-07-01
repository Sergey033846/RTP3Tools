using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

using FirebirdSql.Data.FirebirdClient;

namespace RTP3Tools
{
    public partial class FormLoadDataGrid : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public FormLoadDataGrid()
        {
            InitializeComponent();
        }

        private void FormLoadDataGrid_Load(object sender, EventArgs e)
        {

        }

        private void spreadsheetControl1_BeforeImport(object sender, SpreadsheetBeforeImportEventArgs e)
        {            
            this.Text += String.Concat(" - ", e.Options.SourceUri.ToString());
        }

        private void barButtonItemStartLoading_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormLoadFromXLSParams formLoadFromXLSParams = new FormLoadFromXLSParams();
            if (formLoadFromXLSParams.ShowDialog() == DialogResult.OK) // загрузка (изменение) данных
            {
                OpenFileDialog openFileDialogRTP = new OpenFileDialog();
                openFileDialogRTP.Title = "Выберите файл базы данных РТП-3";
                openFileDialogRTP.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
                openFileDialogRTP.FileName = "";

                if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
                {
                    (this.MdiParent as FormMain).splashScreenManager1.ShowWaitForm();

                    IWorkbook workbook = this.spreadsheetControl1.Document;
                    Worksheet worksheet = workbook.Worksheets[0];

                    workbook.History.IsEnabled = false;

                    int columnCodeOld = Convert.ToInt32(formLoadFromXLSParams.spinEditColumnCodeOld.Value);
                    int columnCodeNew = Convert.ToInt32(formLoadFromXLSParams.spinEditColumnCodeNew.Value);
                    int rowDataFirst = Convert.ToInt32(formLoadFromXLSParams.spinEditDataFirstRow.Value);

                    int rowIndex = rowDataFirst - 1;

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

                    while (!String.IsNullOrWhiteSpace(worksheet[rowIndex, 0].Value.ToString()))
                    {
                        (this.MdiParent as FormMain).splashScreenManager1.SetWaitFormDescription(String.Concat("обработка - строка ", rowIndex.ToString()));

                        string strCodeOld = worksheet[rowIndex, columnCodeOld - 1].Value.ToString();
                        string strCodeNew = worksheet[rowIndex, columnCodeNew - 1].Value.ToString();

                        if (!String.IsNullOrWhiteSpace(strCodeNew))
                        {
                            sqlUpdate = String.Concat(
                                "UPDATE \"LVConsumersInfo\" ",
                                "SET \"LVConsumersInfo\".\"Contract\" = '", strCodeNew, "' ",
                                "WHERE \"LVConsumersInfo\".\"Contract\" = '", strCodeOld, "' "
                                );
                            //FbCommand commandUpdate = new FbCommand(sqlUpdate, connection);
                            //FbTransaction fbt = connection.BeginTransaction();
                            //commandUpdate.Transaction = fbt;

                            commandUpdate.CommandText = sqlUpdate;
                            commandUpdate.ExecuteNonQuery();

                            //fbt.Commit();
                        }

                        rowIndex++;
                    }

                    /*FbCommand commandUpdate = new FbCommand(sqlUpdate, connection);
                    commandUpdate.ExecuteNonQuery();*/

                    connection.Close();

                    (this.MdiParent as FormMain).splashScreenManager1.CloseWaitForm();
                    MessageBox.Show("Загрузка завершена");
                } // if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
            } // if (formLoadFromXLSParams.ShowDialog() == DialogResult.OK)
        }
    }
}