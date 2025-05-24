using System;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;
using System.Net.Http;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Proyecto_Final_1
{
    public partial class Form1 : Form
    {
        // Configuración reusable
        private const string ApiKey = ""; // aca se coloca la API Key de OpenAI
        private const string ConnectionString = "Server=MSI\\SQLEXPRESS;Database=InvestigacionesAI;Integrated Security=True;";

        public Form1()
        {
            InitializeComponent();
        }

        private async void btnInvestigar_Click(object sender, EventArgs e)
        {
            string prompt = txtPrompt.Text.Trim(); // Obtener el texto del cuadro de texto

            if (string.IsNullOrWhiteSpace(prompt)) // validar que el texto no este vacio 
            {
                MessageBox.Show("Ingresa un tema válido.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Llamar a la API de OpenAI    
                string respuesta = await ConsultarAPIAsync(prompt);
                txtResultado.Text = respuesta;
           
                await Task.Run(() => InsertarResultadoEnBaseDeDatos(prompt, respuesta));
                await Task.Run(() => GenerarDocumentoWord(prompt, respuesta));
                await Task.Run(() => GenerarPresentacionPowerPoint(prompt, respuesta));

                MessageBox.Show("Proceso completado exitosamente!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task<string> ConsultarAPIAsync(string prompt)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {ApiKey}");

                var requestBody = new
                {
                    model = "gpt-3.5-turbo",
                    messages = new[] { new { role = "user", content = prompt } },
                    max_tokens = 1000
                };

                var content = new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json");
                var response = await client.PostAsync("https://api.openai.com/v1/chat/completions", content);
                response.EnsureSuccessStatusCode();

                var responseString = await response.Content.ReadAsStringAsync();
                dynamic responseJson = JsonConvert.DeserializeObject(responseString);

                return responseJson?.choices[0]?.message?.content?.ToString()
                    ?? throw new Exception("Respuesta de la API no válida.");
            }
        }

        private void InsertarResultadoEnBaseDeDatos(string tema, string resultado)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                string query = @"INSERT INTO Investigaciones (Tema, Resultado, FechaRegistro) 
                               VALUES (@tema, @resultado, @fecha)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@tema", tema);
                    cmd.Parameters.AddWithValue("@resultado", resultado);
                    cmd.Parameters.AddWithValue("@fecha", DateTime.Now);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void GenerarDocumentoWord(string tema, string contenido)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string directorio = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Investigaciones Word");
                Directory.CreateDirectory(directorio);

                string nombreArchivo = $"Investigacion_{SanitizarNombreArchivo(tema)}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                string rutaCompleta = Path.Combine(directorio, nombreArchivo);

                wordApp = new Word.Application();
                doc = wordApp.Documents.Add();

                // Configuración del documento
                Word.Paragraph titulo = doc.Content.Paragraphs.Add();
                titulo.Range.Text = tema;
                titulo.Range.Font.Bold = 1;
                titulo.Range.InsertParagraphAfter();

                Word.Paragraph cuerpo = doc.Content.Paragraphs.Add();
                cuerpo.Range.Text = contenido;
                cuerpo.Range.Font.Size = 12;
                cuerpo.Range.InsertParagraphAfter();

                doc.SaveAs2(rutaCompleta);
            }
            finally
            {
                if (doc != null) Marshal.ReleaseComObject(doc);
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        private void GenerarPresentacionPowerPoint(string tema, string contenido)
        {
            PowerPoint.Application pptApp = null;
            PowerPoint.Presentation presentation = null;

            try
            {
                string directorio = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Investigaciones PowerPoint");
                Directory.CreateDirectory(directorio);

                string nombreArchivo = $"Presentacion_{SanitizarNombreArchivo(tema)}_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";
                string rutaCompleta = Path.Combine(directorio, nombreArchivo);

                pptApp = new PowerPoint.Application();
                presentation = pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

                // Diseño de título y contenido
                PowerPoint.Slide slide = presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText]);

                slide.Shapes.Title.TextFrame.TextRange.Text = tema;
                slide.Shapes[2].TextFrame.TextRange.Text = contenido;

                presentation.SaveAs(rutaCompleta, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Core.MsoTriState.msoTrue);
            }
            finally
            {
                if (presentation != null) Marshal.ReleaseComObject(presentation);
                if (pptApp != null)
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
            }
        }

        private string SanitizarNombreArchivo(string nombre)
        {
            char[] caracteresInvalidos = Path.GetInvalidFileNameChars();
            return string.Join("_", nombre.Split(caracteresInvalidos));
        }

        private void contextMenuStrip2_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void txtResultado_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}


