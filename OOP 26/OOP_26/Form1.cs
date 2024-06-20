using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace OOP_26
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.Add(@"D:\vs\source\OOP_26\OOP_26\НАКАЗ.doc");
            comboBox1.Items.Add(@"D:\vs\source\OOP_26\OOP_26\НАКАЗ_1.doc");
        }
        Word.Application word = new Word.Application();
        Word.Document doc;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Object missingObj = System.Reflection.Missing.Value;
                Object templatePathObj = comboBox1.SelectedItem.ToString();

                doc = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                doc.Activate();

                foreach (Word.FormField f in doc.FormFields)
                {
                    switch (f.Name)
                    {
                        case "Organization":
                            f.Range.Text = textBox1.Text;
                            break;
                        case "Number":
                            f.Range.Text = textBox2.Text;
                            break;
                        case "Date":
                            f.Range.Text = textBox3.Text;
                            break;
                        case "Year":
                            f.Range.Text = textBox4.Text;
                            break;
                        case "PIB":
                            f.Range.Text = textBox5.Text;
                            break;
                        case "TabNum":
                            f.Range.Text = textBox6.Text;
                            break;
                        case "StructPid":
                            f.Range.Text = textBox7.Text;
                            break;
                        case "Posada":
                            f.Range.Text = textBox8.Text;
                            break;
                        case "Hours":
                            f.Range.Text = textBox9.Text;
                            break;
                        case "PIB2":
                            f.Range.Text = textBox10.Text;
                            break;
                        case "Flag1":
                            f.CheckBox.Value = checkBox1.Checked;
                            break;
                        case "Flag2":
                            f.CheckBox.Value = checkBox2.Checked;
                            break;
                        case "Flag3":
                            f.CheckBox.Value = checkBox3.Checked;
                            break;
                        case "Flag4":
                            f.CheckBox.Value = checkBox4.Checked;
                            break;
                        case "Flag5":
                            f.CheckBox.Value = checkBox5.Checked;
                            break;
                        case "Flag6":
                            f.CheckBox.Value = checkBox6.Checked;
                            break;
                        case "Flag7":
                            f.CheckBox.Value = checkBox7.Checked;
                            break;
                    }
                }
                //Збереження по визначеному шляху
                Object savePath = @"D:\Збережений файл.doc";
                doc.SaveAs2(ref savePath);
                //Пошук 
                string findText = textBox11.Text;
                string replaceWith = textBox12.Text;
                bool found = false;

                foreach (Word.Range range in doc.StoryRanges)
                {
                    Word.Find find = range.Find;
                    find.Text = findText;
                    find.Replacement.Text = replaceWith;//заміна тектсу

                    if (find.Execute(Replace: WdReplace.wdReplaceAll))
                    {
                        found = true;
                    }
                }
                
                if (found)
                {
                    MessageBox.Show($"Текст '{findText}' було знайдено та змінено на '{replaceWith}'", "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Текст '{findText}' не було знайдено", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                word.Visible = true;


            }
            catch(Exception ex) 
            {
                if (doc != null)
                {
                    doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                    doc = null;
                }

                if (word != null)
                {
                    word.Quit();
                    word = null;
                }

                MessageBox.Show("Виникла помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (doc != null)
            {
                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                doc = null;
            }

            if (word != null)
            {
                word.Quit();
                word = null;
            }

        }
    }
}
