using EasyDox;
using Microsoft.Office.Interop.Word;
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
                                string[] Templates = new DirectoryInfo(parameter[1]).GetFiles("*.doc*", SearchOption.AllDirectories).Select(f => f.FullName).ToArray();
                                listBox1.Items.AddRange(Templates); //Заполняем список "Шаблоны в папке"
                                string[] Replasments = new DirectoryInfo(textBox1.Text).GetFiles("Replacement*.txt", SearchOption.TopDirectoryOnly).Select(f => f.Name).ToArray();
                                for (int r = 0; r < Replasments.Length; r++)
                                {
                                    string[] lll = File.ReadAllLines(textBox1.Text+@"\"+Replasments[r]);
                                    if (lll[0].Equals("#do not delete this line#"))
                                        listBox2.Items.Add(Replasments[r]);
                                }
                            }
                            catch(Exception ee) { MessageBox.Show(ee.Message);}
                            break;
                        case "Template_selected": //    Заполняем поле "Выбранный шаблон документы"
                            if (File.Exists(parameter[1])) textBox2.Text = parameter[1];
                            break;
                        case "Replacement_selected": //  Заполняем поле "Выбранный шаблон замены"
                            if (File.Exists(parameter[1])) textBox3.Text = parameter[1];
                            FilReps();
                            break;                  //  Заполняем поле "Папка с результатами"
                        case "Output_folder":
                            if (Directory.Exists(parameter[1])) textBox4.Text = parameter[1];
                            break;
                        default: break;
                    }
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Icon = Properties.Resources.ico;
            splitContainer1.SplitterDistance = 670;
            Start();
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count > 1)
            { textBox2.Text = ""; }
            else
            { textBox2.Text = listBox1.SelectedItem.ToString(); }
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
            Stopwatch st = new Stopwatch();
            st.Start();
            MessageBox.Show("!После закрытия окна, все процессы winword.exe будут убиты без сохранения!");
            foreach (var proc in Process.GetProcessesByName("winWord"))
            {
                proc.Kill();
            }
            button1.Enabled = false;
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.ShowNewFolderButton = false;
            FBD.Description = "Укажите путь к новой папке для копирования структуры папок";
            if (textBox4.Text.Length == 0)
            {
                if (FBD.ShowDialog() == DialogResult.OK)
                {
                    textBox4.Text = FBD.SelectedPath;
                }
            }
            if (textBox4.Text.Length != 0 && Directory.Exists(textBox4.Text))
            {
                foreach (string dirPath in Directory.GetDirectories(textBox1.Text, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        Directory.CreateDirectory(dirPath.Replace(textBox1.Text, textBox4.Text));//дублируем структуру папок                 
                    }
                    catch { }
                }

                int counter = 0;
                for (int i = 0; i < listBox1.SelectedItems.Count; i++)//Выполняем замену в каждом выделенном файле
                {
                    label6.Text = (i + 1) + " / " + listBox1.SelectedItems.Count + " Working";
                    /* if (FileIsOpen(listBox1.SelectedItems[i].ToString()) == true)// открыт ли уже файл
                     { MessageBox.Show("Необходимо закрыть все процессы Word.exe и повторить задачу");break; }
                     else
                     {*/
                    try
                    {
                        OpenFile(i);//открываем в word документ
                        for (int L = 0; L < listView1.Items.Count; L++)
                        {
                            FindReplace(listView1.Items[L].SubItems[0].Text, listView1.Items[L].SubItems[1].Text);//выполняем в тексте документа замену текста
                        }
                        SaveCloseFile(i);//акрываем открытый в word документ
                    }
                    catch { }
                    //}
                    counter++;

                    label6.Text = counter + " / " + listBox1.SelectedItems.Count + " Done!";
                }
            }
            button1.Enabled = true;
            st.Stop(); label8.Text = st.Elapsed.ToString();
        }
        // глобальные переменные
        public  Microsoft.Office.Interop.Word.Application app;
        Document doc;

        public static Object missing = Type.Missing;

        
        public void OpenFile(int listbox1_id)// Открываем файл .doc 
        {
            app = new Microsoft.Office.Interop.Word.Application();
            string doc_file = listBox1.SelectedItems[listbox1_id].ToString();
            doc = app.Documents.Open(doc_file);          //  app.Documents.Open(textBox1.Text + "\\" + listBox1.SelectedItems[listbox1_id]);
        }

        // Закрытие general и сохранение файла нового файла
        public void SaveCloseFile(int listbox1_id)
        {
            string newfile = listBox1.SelectedItems[listbox1_id] + " "+ DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + " " + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + ".doc"; // новый файл на основе файла-шаблона
            //string[] newfileA = listBox1.SelectedItems[listbox1_id].ToString().Split('.');
            //newfile = textBox4.Text + "\\"+newfileA[newfileA.Length - 1];
            newfile = newfile.Replace(textBox1.Text, textBox4.Text);
            newfile = newfile.Replace("Template_", "");
            app.ActiveDocument.SaveAs(newfile);
            app.ActiveDocument.Close();
            //app.Documents.Close();
            app.Quit();
            app = null;
        }

        // поиск и замена
        public void FindReplace(string str_old, string str_new)
        {   
            object missingObject = null;
            object item = WdGoToItem.wdGoToPage;
            object whichItem = WdGoToDirection.wdGoToFirst;
            object replaceAll = WdReplace.wdReplaceAll;
            object forward = true;
            object matchAllWord = true;
            object matchCase = false;
            object originalText = str_old;
            object replaceText = str_new;

            doc.GoTo(ref item, ref whichItem, ref missingObject, ref missingObject);
            foreach (Range rng in doc.StoryRanges)
            {
                rng.Find.Execute(ref originalText, ref matchCase,
                ref matchAllWord, ref missingObject, ref missingObject, ref missingObject, ref forward,
                ref missingObject, ref missingObject, ref replaceText, ref replaceAll, ref missingObject,
                ref missingObject, ref missingObject, ref missingObject);
            }
            /* Код работающий быстрее но не работающий с колонтитулами:
            Find find = app.Selection.Find;

            find.Text = str_old; // текст поиска
            find.Replacement.Text = str_new; // текст замены
            find.Execute(FindText: str_old, MatchCase: false, MatchWholeWord: true, MatchWildcards: false,
                        MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true, Wrap: WdFindWrap.wdFindContinue,
                        Format: false, ReplaceWith: str_new, Replace: WdReplace.wdReplaceAll);
            object matchCase = true; object matchWholeWord = true;
            object matchWildCards = false; object matchSoundLike = false;
            object nmatchAllForms = false; object forward = true;
            object format = false; object matchKashida = false;
            object matchDiactitics = false; object matchAlefHamza = false;
            object matchControl = false; object read_only = false;
            object visible = true; object replace = 2;
            object wrap = 1;
            foreach (Range range in doc.StoryRanges)
            {
                if (range.StoryType == WdStoryType.wdEvenPagesHeaderStory)
                {
                    find.Execute(FindText: str_new, MatchCase: false, MatchWholeWord: false, MatchWildcards: false,
                    MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true, Wrap: WdFindWrap.wdFindContinue,
                    Format: false, ReplaceWith: str_new, Replace: WdReplace.wdReplaceAll);
                }
            }*/
        }


        private void FilReps()
        {
            try
            {
                listView1.Items.Clear();
                string tt = textBox3.Text;
                string[] reps = File.ReadAllLines(tt, Encoding.Default);
                for (int L = 1; L < reps.Length; L++)
                {
                    string[] rep = reps[L].Split('\t');
                    listView1.Items.Add(new ListViewItem(new string[] { rep[0], rep[1] }));
                }
            }
            catch { }
        }

        public bool FileIsOpen(string path)
        {
            FileStream a = null;

            try
            {
                a = File.Open(path,
                FileMode.Open, FileAccess.Read, FileShare.None);
                return false;
            }
            catch
            {
                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
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
            try
            {
                string[] content = new string[parameters_count];
                content[0] = "Template_patch\t" + textBox1.Text;
                content[1] = "Template_selected\t" + textBox2.Text;
                content[2] = "Replacement_selected\t" + textBox3.Text;
                content[3] = "Output_folder\t" + textBox4.Text; //папка куда выгружать результаты
                File.WriteAllLines(Environment.CurrentDirectory + "\\Parameters.txt", content);
            }
            catch { }
        }
        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Start();
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/sergiomarotco/");
        }

        private void listBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            FilReps();
        }
    }
}
