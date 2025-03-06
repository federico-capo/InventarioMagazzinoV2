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

//per leggere il file INI
using IniParser.Model;
using IniParser;


//libreria execl
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace BoweInventarioV2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        //Lettura File INI
        string LeggiValoreDaIni(string percorsoFileIni, string sezione, string chiave)
        {
            var parser = new FileIniDataParser();
            if (!File.Exists(percorsoFileIni))
            {
                MessageBox.Show("File di configurazione non trovato!");
                return null;
            }

            IniData data = parser.ReadFile(percorsoFileIni);
            return data[sezione][chiave];
        }

        //creazione dizionario e funzione estrapolazione dati
        static Dictionary<string, (string nome, string valore)> EstraiChiaviValoriDaExcel(string filePath, string sheetName)
        {
            var risultati = new Dictionary<string, (string nome, string valore)>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))

            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null) return risultati;

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var chiave = worksheet.Cells[row, 1].Text.Trim();
                    var nome = worksheet.Cells[row, 2].Text;
                    var valore = worksheet.Cells[row, 4].Text;

                    if (!string.IsNullOrEmpty(chiave))
                    {
                        risultati[chiave] = (nome, valore);
                    }
                }
            }
            return risultati;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Apri il file Inventario con OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file csv del magazzino";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //salvataggio percorso inventario preso dal dialogfile
                string percorsoInventario = openFileDialog.FileName;
                //salvo sulla variabile dizionario contenuto excel
                var dizionario = EstraiChiaviValoriDaExcel(percorsoInventario, "Inventario");


                //manda a schermo il dizionario di string string con un ciclo for
                textBox1.Clear();
                foreach (var kvp in dizionario)
                {
                    textBox1.AppendText($"{kvp.Key}: {kvp.Value}{Environment.NewLine}");
                }

                // Leggere il percorso del secondo Excel dal file .ini
                string percorsoFileIni = Path.Combine(Application.StartupPath, "config.ini");
                string percorsoSecondoExcelOriginale = LeggiValoreDaIni(percorsoFileIni, "Impostazioni", "PercorsoSecondoExcel");

                // Creare il percorso della copia
                string percorsoSecondoExcelCopia = percorsoSecondoExcelOriginale + "_agg.xlsx";
                //richiamo la funzione  aggiorno excel e gli passo i percorsi
                AggiornaSecondoExcel(percorsoSecondoExcelOriginale, percorsoSecondoExcelCopia, "Sheet1", dizionario);
               
                //gestisto eccezione ini non configurato
                if (string.IsNullOrEmpty(percorsoSecondoExcelOriginale))
                {
                    MessageBox.Show("Errore: il percorso del secondo file Excel non è presente nel file INI!");
                    return;
                }
                //Aggiorno Textbox e Label 
                // MessageBox.Show($"Inventario Aggiornato con successo \nExcel si trova qui: {percorsoSecondoExcelCopia}");
                textBox1.Clear();
                textBox1.AppendText("Inventario Aggiornato con successo Excel si trova qui\n" + percorsoSecondoExcelCopia);
                label1.Text = "Csv\nAggiornato";
            }
        }

        static List<string> GeneraVariantiDiChiave(string chiave)
        {
            var varianti = new HashSet<string> { chiave };

            if (chiave.Contains(" "))
            {
                string secondaParte = chiave.Substring(chiave.IndexOf(" ") + 1);
                varianti.Add(secondaParte);
            }

            if (chiave.All(char.IsDigit))
            {
                varianti.Add(chiave.TrimStart('0'));
            }

            if (chiave.StartsWith("X"))
            {
                string senzaX = chiave.Substring(1);
                varianti.Add(senzaX);
                varianti.UnionWith(GeneraVariantiDiChiave(senzaX));
            }

            return varianti.ToList();
        }

        static string CercaNelDizionarioConVarianti(string chiave, Dictionary<string, (string nome, string valore)> dizionario, out string nome)
        {
            nome = null;
            foreach (var variante in GeneraVariantiDiChiave(chiave))
            {
                if (dizionario.ContainsKey(variante))
                {
                    nome = dizionario[variante].nome;
                    return variante;
                }
            }
            return null;
        }
        static void AggiornaSecondoExcel(string fileOriginale, string fileCopia, string sheetName, Dictionary<string, (string nome, string valore)> dizionario)
        {
            File.Copy(fileOriginale, fileCopia, true);

            using (var package = new ExcelPackage(new FileInfo(fileCopia)))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null) return;

                int rowCount = worksheet.Dimension.Rows;
                var chiaviTrovate = new HashSet<string>();

                for (int row = 2; row <= rowCount; row++)
                {
                    var cellaA = worksheet.Cells[row, 1].Text.Trim();
                    if (string.IsNullOrEmpty(cellaA)) continue;

                    string nome;
                    string chiaveTrovata = CercaNelDizionarioConVarianti(cellaA, dizionario, out nome);
                    worksheet.Cells[row, 2].Value = nome;

                    worksheet.Cells[row, 3].Value = chiaveTrovata != null && int.TryParse(dizionario[chiaveTrovata].valore, out int valoreInt) ? valoreInt : 0;

                    if (chiaveTrovata != null)
                        chiaviTrovate.Add(chiaveTrovata);
                }

                int newRow = rowCount + 1;
                foreach (var chiaveMancante in dizionario.Keys.Except(chiaviTrovate))
                {
                    worksheet.Cells[newRow, 1].Value = chiaveMancante;
                    worksheet.Cells[newRow, 2].Value = dizionario[chiaveMancante].nome;
                    worksheet.Cells[newRow, 3].Value = int.TryParse(dizionario[chiaveMancante].valore, out int valoreInt) ? valoreInt : 0;
                    newRow++;
                }

                package.Save();

            }
        }
 


        private void Form1_Load(object sender, EventArgs e)
        {
                textBox1.AppendText("Avviato");
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
