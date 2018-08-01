using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Etiquetas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region Events
        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "(*.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TransformFile(textBox1.Text, double.Parse(textBox2.Text), double.Parse(textBox3.Text));
        }
        #endregion

        #region Methods
        private void TransformFile(object path, double discount1, double discount2)
        {
            var word = new Microsoft.Office.Interop.Word.Application();
            var doc = new Document();

            progressBar1.Show();
            button2.Enabled = false;
            
            try
            {
                object fileName = path;
                // Define an object to pass to the API for missing parameters
                object missing = Missing.Value;
                doc = word.Documents.Open(ref fileName,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing);

                var read = string.Empty;

                for (int i = 0; i < doc.Paragraphs.Count; i++)
                {
                    progressBar1.Value = (i * 100 / doc.Paragraphs.Count) + 1;

                    doc.Paragraphs[i + 1].Range.Font.Size = 10;
                    var temp = string.IsNullOrEmpty(doc.Paragraphs[i + 1].Range.Text) ? "" : doc.Paragraphs[i + 1].Range.Text;
                    if (temp.Contains(","))
                    {
                        var length = "";
                        var price = "";

                        var items = temp.Split(' ');

                        foreach (var item in items)
                        {
                            if (item.Contains("-"))
                            {
                                length = item;
                            }

                            if (item.Contains(","))
                            {
                                price = GetFinalPrice(item, discount1, discount2);
                            }
                        }

                        doc.Paragraphs[i + 1].Range.Text = string.Concat(length, " ", price, "\r");
                    }
                }

                progressBar1.Hide();
                MessageBox.Show("Arquivo Transformado com sucesso.");
                button2.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro: " + ex.Message);
            }
            finally
            {
                ((_Document)doc).Close();
                ((_Application)word).Quit();
            }
        }

        private string GetFinalPrice(string item, double discount1, double discount2)
        {
            var price = double.Parse(item) * (1 - discount1 / 100);
            price = price * (1 - discount2 / 100);

            var priceMask = string.Concat("1000", price.ToString("N2").Replace(",",""));

            return priceMask;
        }

        #endregion
    }
}
