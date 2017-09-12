using Google.Cloud.Vision.V1;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OCR.Test
{
    public partial class frmMainForm : Form
    {
        #region Fields

        private ImageAnnotatorClient myClient;

        #endregion

        #region Properties

        public string Credential_path
        {
            get
            {
                return @"CredentialsFile.json";
            }
        }

        public string ImageText
        {
            get
            {
                return this.txtSource.Text;
            }
        }

        public ImageAnnotatorClient Client
        {
            get
            {
                if (this.myClient == null)
                {
                    this.myClient = ImageAnnotatorClient.Create();
                }

                return this.myClient;
            }
        }

        #endregion

        #region Constructors

        public frmMainForm()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", Credential_path);
            this.SuccesfulyExpressions = 0;
            this.TotalExpressions = 0;

            var tiposDocumento = new string[] { "Telmex", "INE", "IFE", "CDMX", "CFE" };
            this.cboTipoDocumento.Items.AddRange(tiposDocumento);

            var tiposResultado = new string[] { "Exitoso", "No Exitoso", "Todos" };
            this.cboTipoResultado.Items.AddRange(tiposResultado);
            this.cboTipoResultado.SelectedIndex = 2;
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Filter = "Image files (*.jpg) | *.jpg; ";
            var result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                this.SelectFile(ofd.FileName);
            }
        }

        private void txtRegex_TextChanged(object sender, EventArgs e)
        {
            this.ApplyRegex();
        }

        private void cboGrupos_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.ApplyRegex();
        }

        private void btnSelectExcelFile_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            var result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                ApplyRegexFromFilePath(ofd.FileName);
                Cursor.Current = Cursors.Default;
            }
        }

        private void btnTestAll_Click(object sender, EventArgs e)
        {
            if (this.cboTipoDocumento.SelectedItem == null)
            {
                MessageBox.Show("Please selected the document type...");
                return;
            }
            var start = DateTime.Now;
            var ofd = new FolderBrowserDialog();
            var result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                var files = Directory.GetFiles(ofd.SelectedPath);

                foreach (var file in files)
                {
                    this.SelectFile(file);
                    this.lblFileName.Text = "File: " + file;

                    var x = @"C:\Users\ivan\Desktop\expresiones.xlsx";
                    ApplyRegexFromFilePath(x);
                }
            }

            var stop = DateTime.Now;

            var time = stop - start;

            MessageBox.Show(time.ToString());
            MessageBox.Show("Total: " + this.TotalExpressions + " Succesfull: " + this.SuccesfulyExpressions);
        }

        private void SelectFile(string fileName)
        {
            this.lblFileName.Text = "File:" + fileName;
            this.pictureBox1.ImageLocation = fileName;
            Cursor.Current = Cursors.WaitCursor;
            var image = this.GetImageFromPath(fileName);
            this.txtSource.Text = this.GetTextFromImage(image);
            Cursor.Current = Cursors.Default;
        }

        public int TotalExpressions { get; set; }

        public int SuccesfulyExpressions { get; set; }

        private void ApplyRegexFromFilePath(string fileName)
        {
            var dataTable = GetDataFromExcelFile(fileName, "xlsx");
            var resultadoValidacionList = new List<ResultadoValidacion>();

            foreach (DataRow dr in dataTable.Rows)
            {
                var resultadoValidacion = new ResultadoValidacion()
                {
                    Campo = dr.ItemArray[0].ToString(),
                    Expression = dr.ItemArray[1].ToString(),
                    Grupo = dr.ItemArray[2].ToString(),
                    Resultado = this.EvaluateRegex(this.ImageText, dr.ItemArray[1].ToString(), dr.ItemArray[2].ToString())
                };

                TotalExpressions++;
                if (resultadoValidacion.Resultado != "Match not successful")
                {
                    SuccesfulyExpressions++;
                }

                var tipoResultado = this.cboTipoResultado.SelectedItem.ToString();

                if (tipoResultado == "Todos")
                {
                    resultadoValidacionList.Add(resultadoValidacion);
                }
                else
                {
                    if (tipoResultado == "Exitoso" && resultadoValidacion.Resultado != "Match not successful")
                    {
                        resultadoValidacionList.Add(resultadoValidacion);
                    }
                    else
                    {
                        if (tipoResultado == "No Exitoso" && resultadoValidacion.Resultado == "Match not successful")
                        {
                            resultadoValidacionList.Add(resultadoValidacion);
                        }
                    }
                }
            }

            grdResultExpressions.DataSource = resultadoValidacionList;
        }

        private string EvaluateRegex(string text, string pattern, string grupo)
        {
            var result = string.Empty;
            this.txtResult.Text = string.Empty;
            try
            {
                var match = Regex.Match(text, @pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var value = match.Groups[grupo].Value;
                    var answer = new StringBuilder();
                    for (int i = 0; i < value.Length; i++)
                    {
                        byte[] charByte = Encoding.ASCII.GetBytes(new char[] { value[i] });
                        if (charByte[0] == 9)
                        {
                            this.txtResult.AppendText("Tab", Color.Orange);
                            answer.Append("chr(10)");
                        }
                        else if (charByte[0] == 10)
                        {
                            this.txtResult.AppendText("NewLine", Color.Red);
                            answer.Append("chr(10)");
                        }
                        else if (charByte[0] == 13)
                        {
                            this.txtResult.AppendText("CarriageReturn", Color.Red);
                            answer.Append(@"chr(13)");
                        }
                        else if (charByte[0] == 32)
                        {
                            this.txtResult.AppendText("_", Color.Orange);
                            answer.Append(@"chr(13)");
                        }
                        else
                        {
                            this.txtResult.AppendText(value[i].ToString(), Color.Black);
                            answer.Append(value[i]);
                        }
                    }

                    this.LoadGroupsFromMatch(match);
                    result = answer.ToString();
                }
                else
                {
                    result = "Match not successful";
                    this.txtResult.AppendText(result, Color.Red);
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
                this.txtResult.AppendText(result, Color.Red);
            }

            return result;
        }

        private void LoadGroupsFromMatch(Match match)
        {
            this.cboGrupos.Items.Clear();
            var numeroGrupo = 0;
            foreach (var g in match.Groups)
            {
                this.cboGrupos.Items.Add(numeroGrupo);
                numeroGrupo++;
            }
        }

        public DataTable GetDataFromExcelFile(string fileName, string fileExt)
        {
            var documentType = this.cboTipoDocumento.SelectedItem;

            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  

            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + documentType + "$]", con); //here we read data from sheet1 
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception x)
                {
                    MessageBox.Show(x.Message);
                }
            }
            return dtexcel;
        }

        private string GetTextFromImage(Google.Cloud.Vision.V1.Image filePath)
        {
            var response = Client.DetectText(filePath);
            var resultList = new List<ResultElement>();

            foreach (var annotation in response)
            {
                resultList.Add(new ResultElement
                {
                    Description = annotation.Description,
                    Location = annotation.BoundingPoly.ToString()
                });
            }

            return resultList.First().Description;
        }

        private Google.Cloud.Vision.V1.Image GetImageFromPath(string filePath)
        {
            return Google.Cloud.Vision.V1.Image.FromFile(filePath);
        }

        private void ApplyRegex()
        {
            var text = this.ImageText;
            var pattern = this.txtRegex.Text;
            var group = string.Empty;
            if (this.cboGrupos.SelectedItem != null)
            {
                group = this.cboGrupos.SelectedItem.ToString();
            }

            //this.txtResult.Text = this.EvaluateRegex(text, @pattern, group);
            var result = this.EvaluateRegex(text, @pattern, group);
            //this.txtResult.AddText(result, Color.Red);
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            var result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                var datatable = this.GetDataFromExcelFile(ofd.FileName, "xls");
                this.grdResultExpressions.DataSource = datatable;
            }   
        }
    }
}
