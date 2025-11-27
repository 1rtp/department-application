using ClosedXML.Excel;
using FastReport;
using FastReport.DevComponents.DotNetBar.Controls;
using FastReport.Format;
using FastReport.Preview;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using КП_Кафедра.ReportsBridge;
using static КП_Кафедра.ToastForm;
using КП_Кафедра.IteratorPattern;

namespace КП_Кафедра.Forms
{
    public partial class FormReport : Form
    {
        private string selectedReport = "TeachersReport.frx";
        private readonly string reportsFolderPath;
        private readonly string dbPath;

        private AbstractReport currentReport;

        private ReportCollection reportCollection;
        private IReportIterator reportIterator;

        public FormReport()
        {
            InitializeComponent();
            AppSettings.LoadSettings();
            AppSettings.ApplyStyle(this);
            AppSettings.StyleChanged += () => AppSettings.ApplyStyle(this);

            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string projectRoot = Path.GetFullPath(Path.Combine(exeDir, "..", ".."));
            reportsFolderPath = Path.Combine(projectRoot, "Reports");
            dbPath = Path.Combine(projectRoot, "Data", "department.db");

            reportCollection = new ReportCollection(reportsFolderPath);
            reportIterator = reportCollection.CreateIterator();
        }

        private void FormReport_Load(object sender, EventArgs e)
        {
            rbTeachersReport.CheckedChanged += RadioButtons_CheckedChanged;
            rbSubjectsReport.CheckedChanged += RadioButtons_CheckedChanged;
            rbAssignmentsReport.CheckedChanged += RadioButtons_CheckedChanged;
            rbResearchReport.CheckedChanged += RadioButtons_CheckedChanged;

            rbTeachersReport.Checked = true;
            selectedReport = "TeachersReport.frx";
            UpdateCurrentReport();
            SyncIteratorWithRadio();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            LanguageManager.LanguageChanged += ApplyLocalization;
            ApplyLocalization();
        }

        private void ApplyLocalization()
        {
            rbTeachersReport.Text = LanguageManager.GetString("rbTeachersReport");
            rbSubjectsReport.Text = LanguageManager.GetString("rbSubjectsReport");
            rbResearchReport.Text = LanguageManager.GetString("rbResearchReport");
            rbAssignmentsReport.Text = LanguageManager.GetString("rbAssignmentsReport");
            btnReport.Text = LanguageManager.GetString("btnReport");
            btnOpenExcelReport.Text = LanguageManager.GetString("btnOpenExcelReport");
            btnOpenWordReport.Text = LanguageManager.GetString("btnOpenWordReport");
            btnExportToExcel.Text = LanguageManager.GetString("btnExportToExcel");
            btnExportToWord.Text = LanguageManager.GetString("btnExportToWord");
            btnImportFromExcel.Text = LanguageManager.GetString("btnImportFromExcel");
            btnImportFromWord.Text = LanguageManager.GetString("btnImportFromWord");
            btnConvertToPdf.Text = LanguageManager.GetString("btnConvertToPdf");
        }

        private void RadioButtons_CheckedChanged(object sender, EventArgs e)
        {
            if (sender is RadioButton rb && rb.Checked)
            {
                if (rb == rbTeachersReport) selectedReport = "TeachersReport.frx";
                else if (rb == rbSubjectsReport) selectedReport = "SubjectsReport.frx";
                else if (rb == rbAssignmentsReport) selectedReport = "AssignmentsReport.frx";
                else if (rb == rbResearchReport) selectedReport = "ResearchReport.frx";

                UpdateCurrentReport();
            }
        }

        private void UpdateCurrentReport()
        {
            string query = "";
            string name = "";

            if (rbTeachersReport.Checked)
            {
                query = "SELECT * FROM teacher;";
                name = "Викладачі";
            }
            else if (rbSubjectsReport.Checked)
            {
                query = "SELECT * FROM subjects;";
                name = "Дисципліни";
            }
            else if (rbAssignmentsReport.Checked)
            {
                query = "SELECT * FROM assignment;";
                name = "Навантаження";
            }
            else if (rbResearchReport.Checked)
            {
                query = "SELECT * FROM research;";
                name = "Наукові роботи";
            }

            currentReport = new SqlReport(query, dbPath, new ExcelExporter());
            currentReport.ReportName = name;
        }

        private void SyncIteratorWithRadio()
        {
            string fullPath = Path.Combine(reportsFolderPath, selectedReport);

            int index = reportCollection.Reports.IndexOf(fullPath);
            if (index >= 0)
            {
                reportIterator = reportCollection.CreateIterator();
                for (int i = 0; i < index; i++) reportIterator.Next();
            }
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            OpenReportPreview(selectedReport);
        }

        private void OpenReportPreview(string reportFile)
        {

            panel2.Controls.Clear();

            FormReportViewer reportForm = new FormReportViewer(reportFile);
            reportForm.TopLevel = false;
            reportForm.FormBorderStyle = FormBorderStyle.None;
            reportForm.Dock = DockStyle.Fill;
            panel2.Controls.Add(reportForm);
            reportForm.Show();
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                currentReport.SetExporter(new ExcelExporter());
                currentReport.Export();

                Toast.Show("SUCCESS", "Успішно збережено!");
                LoggerService.LogInfo($"Експорт у Excel: {currentReport.ReportName}");
            }
            catch (Exception ex)
            {
                Toast.Show("ERROR", "Помилка збереження у Excel");
                LoggerService.LogError($"Помилка Excel: {ex.Message}");
            }
        }


        private void btnImportFromExcel_Click(object sender, EventArgs e)
        {
            try
            {
                currentReport.SetExporter(new ExcelExporter());
                DataTable imported = currentReport.Import();

                if (imported == null)
                {
                    LoggerService.LogError("Імпорт Excel скасовано.");
                    return;
                }

                Toast.Show("SUCCESS", "Дані з Excel завантажено.");
                LoggerService.LogInfo("Імпорт Excel виконано.");
            }
            catch (Exception ex)
            {
                Toast.Show("ERROR", "Помилка імпорту з Excel.");
                LoggerService.LogError($"Помилка імпорту Excel: {ex.Message}");
            }
        }

        private void btnOpenExcelReport_Click(object sender, EventArgs e)
        {
            try
            {
                currentReport.SetExporter(new ExcelExporter());
                currentReport.Open();

                LoggerService.LogInfo("Користувач відкрив Excel-звіт.");
            }
            catch (Exception ex)
            {
                Toast.Show("ERROR", "Не вдалося відкрити Excel.");
                LoggerService.LogError($"Помилка відкриття Excel: {ex.Message}");
            }
        }

        private void btnExportToWord_Click(object sender, EventArgs e)
        {
            try
            {
                currentReport.SetExporter(new WordExporter());
                currentReport.Export();

                Toast.Show("SUCCESS", "Успішно збережено у Word!");
                LoggerService.LogInfo($"Експорт у Word: {currentReport.ReportName}");
            }
            catch (Exception ex)
            {
                Toast.Show("ERROR", "Помилка збереження у Word.");
                LoggerService.LogError($"Помилка Word: {ex.Message}");
            }
        }

        private void btnImportFromWord_Click(object sender, EventArgs e)
        {
            try
            {
                currentReport.SetExporter(new WordExporter());
                DataTable imported = currentReport.Import();

                if (imported == null)
                {
                    LoggerService.LogError("Імпорт Word скасовано.");
                    return;
                }

                Toast.Show("SUCCESS", "Дані з Word завантажено.");
                LoggerService.LogInfo("Імпорт Word виконано.");
            }
            catch (Exception ex)
            {
                Toast.Show("ERROR", "Помилка імпорту з Word.");
                LoggerService.LogError($"Помилка імпорту Word: {ex.Message}");
            }
        }

        private void btnOpenWordReport_Click(object sender, EventArgs e)
        {
            try
            {
                currentReport.SetExporter(new WordExporter());
                currentReport.Open();

                LoggerService.LogInfo("Користувач відкрив Word-звіт.");
            }
            catch (Exception ex)
            {
                Toast.Show("ERROR", "Не вдалося відкрити Word.");
                LoggerService.LogError($"Помилка відкриття Word: {ex.Message}");
            }
        }

        private void btnConvertToPdf_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Word файли (*.docx)|*.docx";
                ofd.Title = "Оберіть Word-файл для конвертації в PDF";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string inputPath = ofd.FileName;
                        string pdfPath = Path.ChangeExtension(inputPath, ".pdf");

                        Spire.Doc.Document doc = new Spire.Doc.Document();
                        doc.LoadFromFile(inputPath);
                        doc.SaveToFile(pdfPath, Spire.Doc.FileFormat.PDF);
                        doc.Close();

                        Toast.Show("SUCCESS", "Файл PDF створено!");
                        LoggerService.LogInfo($"Word → PDF: {pdfPath}");
                    }
                    catch (Exception ex)
                    {
                        Toast.Show("ERROR", "Не вдалося створити PDF.");
                        LoggerService.LogError($"Помилка під час створення PDF: {ex.Message}");
                    }
                }
            }
        }

        private void btnNextReport_Click(object sender, EventArgs e)
        {
            if (reportIterator.HasNext())
            {
                string fullPath = reportIterator.Next();
                selectedReport = Path.GetFileName(fullPath);
                OpenReportPreview(selectedReport);
            }
            else { Toast.Show("INFO", "Це перший звіт"); }
        }

        private void btnPreviousReport_Click(object sender, EventArgs e)
        {
            if (reportIterator.HasPrevious())
            {
                string fullPath = reportIterator.Previous();
                selectedReport = Path.GetFileName(fullPath);
                OpenReportPreview(selectedReport);
            }
            else { Toast.Show("INFO", "Це останній звіт"); }
        }
    }
}
