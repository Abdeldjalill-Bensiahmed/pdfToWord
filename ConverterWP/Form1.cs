using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;


namespace ConverterWP
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            // Sélectionner le fichier Word à convertir
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Documents Word|*.doc;*.docx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string wordFilePath = openFileDialog.FileName;

                // Créer une nouvelle instance de l'application Microsoft Word
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

                // Ouvrir le document Word
                Document wordDoc = wordApp.Documents.Open(wordFilePath);

                // Créer un SaveFileDialog pour spécifier le chemin de destination du fichier PDF convertir
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Fichiers PDF|*.pdf";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfFilePath = saveFileDialog.FileName;

                    // Enregistrer le document au format PDF
                    wordDoc.SaveAs2(pdfFilePath, WdSaveFormat.wdFormatPDF);

                    // Fermer le document Word et l'application
                    wordDoc.Close();
                    wordApp.Quit();

                    MessageBox.Show("Conversion terminée !");
                }
                else
                {
                    // L'utilisateur a annulé le SaveFileDialog, donc fermer le document Word et l'application
                    wordDoc.Close();
                    wordApp.Quit();
                }
            }
        }
    }
}
