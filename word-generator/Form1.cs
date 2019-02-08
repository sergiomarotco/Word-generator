using EasyDox;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace word_generator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int parameters_count = 0;
        private void Start()
        {
            if (File.Exists("Parameters.txt"))
            {
                listBox1.Items.Clear(); listBox2.Items.Clear();
                string[] parameters = File.ReadAllLines("Parameters.txt");
                parameters_count = parameters.Length;
                for (int i = 0; i < parameters.Length; i++)
                {
                    string[] parameter = parameters[i].Split('\t');
                    switch (parameter[0])
                    {
                        case "Template_patch":
                            textBox1.Text = parameter[1]; //    Заполняем поле "Папка с шаблонами"
                            try
                            {
                                string[] Templates = new DirectoryInfo(parameter[1]).GetFiles("Template*.docx", SearchOption.TopDirectoryOnly).Select(f => f.Name).ToArray();
                                listBox1.Items.AddRange(Templates); //Заполняем список "Шаблоны в папке"
                                string[] Replasments = new DirectoryInfo(textBox1.Text).GetFiles("Replacement*.txt", SearchOption.TopDirectoryOnly).Select(f => f.Name).ToArray();
                                for (int r = 0; r < Replasments.Length; r++)
                                {
                                    string[] lll = File.ReadAllLines(Replasments[r]);
                                    if (lll[0].Equals("#do not delete this line#"))
                                        listBox2.Items.Add(Replasments[r]);
                                }
                            }
                            catch { }
                            break;
                        case "Template_selected": //    Заполняем поле "Выбранный шаблон документы"
                            if (File.Exists(parameter[1])) textBox2.Text = parameter[1];
                            break;
                        case "Replacement_selected": //  Заполняем поле "Выбранный шаблон замены"
                            if (File.Exists(parameter[1])) textBox3.Text = parameter[1];
                            FilReps();
                            break;
                        default: break;
                    }
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            splitContainer1.SplitterDistance = 419;
            Start();
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = listBox1.SelectedItem.ToString();
        }

        private void ListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                textBox3.Text = listBox2.SelectedItem.ToString();
            }
            catch (Exception ee) { MessageBox.Show(ee.Source + Environment.NewLine + ee.Message); }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            var engine = new Engine();

            Dictionary<string, string> fieldValues = new Dictionary<string, string>(listView1.Items.Count);
            for (int L = 0; L < listView1.Items.Count; L++)            
                fieldValues.Add(listView1.Items[L].SubItems[0].Text, listView1.Items[L].SubItems[1].Text);
    
            string outputPath = textBox1.Text+"\\Output\\" + DateTime.Now.Year.ToString()+"."+ DateTime.Now.Month.ToString() + "."+ DateTime.Now.Day.ToString() + " "+ DateTime.Now.Hour.ToString() + "-"+ DateTime.Now.Minute.ToString() + ".docx";

            var errors = engine.Merge(textBox1.Text+"\\"+textBox2.Text, fieldValues, outputPath);


            foreach (var error in errors)
            {
                Console.WriteLine(error.Accept(new ErrorToRussianString()));
            }

            Process.Start(outputPath);
        }
        private void FilReps()
        {
            try
            {
                listView1.Items.Clear();                
                string[] reps = File.ReadAllLines(textBox1.Text + "\\" + textBox3.Text, Encoding.Default);
                for (int L = 1; L < reps.Length; L++)
                {
                    string[] rep = reps[L].Split('\t');
                    listView1.Items.Add(new ListViewItem(new string[] { rep[0], rep[1] }));
                }
            }
            catch (Exception ee) { }
        }
        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            FilReps();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Save_Parameters();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            folderDlg.SelectedPath = Environment.CurrentDirectory;

            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
            Save_Parameters();
            Start();
          
        }
        private void Save_Parameters()
        {
            string[] content = new string[parameters_count];
            content[0] = "Template_patch\t" + textBox1.Text;
            content[1] = "Template_selected\t" + textBox2.Text;
            content[2] = "Replacement_selected\t" + textBox3.Text;
            File.WriteAllLines(Environment.CurrentDirectory + "\\Parameters.txt", content);
        }
        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Start();
        }
    }
}
