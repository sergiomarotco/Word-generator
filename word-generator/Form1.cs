using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
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
        /// <summary>
        /// Заполнить параметры в левой части программы из файла Parameters.txt
        /// </summary>
        private void LoadParameters()
        {
            if (File.Exists("Parameters.txt"))
            {
                listView3.Items.Clear(); listView2.Items.Clear();
                string[] parameters = File.ReadAllLines("Parameters.txt");
                parameters_count = parameters.Length;
                for (int i = 0; i < parameters.Length; i++)
                {
                    string[] parameter = parameters[i].Split('\t');
                    switch (parameter[0])
                    {
                        case "Template_patch":
                            if (Directory.Exists(parameter[1]))
                                textBox1.Text = parameter[1]; //    Заполняем поле "Папка с шаблонами"
                            else textBox1.Text = Environment.CurrentDirectory;
                            try
                            {
                                string[] Templates = new DirectoryInfo(textBox1.Text).GetFiles("*.doc*", SearchOption.AllDirectories).Select(f => f.FullName).ToArray();
                                for (int r = 0; r < Templates.Length; r++)
                                {
                                    string[] lll = File.ReadAllLines(Templates[r]);

                                    string[] patches = Templates[r].Split('\\');
                                    listView3.Items.Add(new ListViewItem(new string[] { patches[patches.Length - 1], Templates[r] }));

                                }

                                string[] Replasments = new DirectoryInfo(textBox1.Text).GetFiles("Replacement*.txt", SearchOption.AllDirectories).Select(f => f.FullName).ToArray();
                                for (int r = 0; r < Replasments.Length; r++)
                                {
                                    string[] lll = File.ReadAllLines(Replasments[r]);
                                    if (lll[0].Equals("#do not delete this line#"))
                                    {
                                        string[] patches = Replasments[r].Split('\\');
                                        listView2.Items.Add(new ListViewItem(new string[] { patches[patches.Length-1], Replasments[r] }));
                                    }
                                }
                            }
                            catch(Exception ee) { MessageBox.Show(ee.Message);}
                            break;
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
            LoadParameters();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (listView3.SelectedItems.Count > 0)
            {
                if (listView2.SelectedItems.Count == 1)
                {
                    button1.Enabled = false;
                    if (textBox4.Text.Length == 0)
                    {
                        FolderBrowserDialog FBD = new FolderBrowserDialog();
                        FBD.ShowNewFolderButton = false;
                        FBD.Description = "Укажите путь к новой папке для копирования структуры папок и генерации файлов";
                        FBD.SelectedPath = Environment.CurrentDirectory;
                        if (FBD.ShowDialog() == DialogResult.OK)
                        {
                            textBox4.Text = FBD.SelectedPath;
                        }
                    }
                    MessageBox.Show("!После закрытия окна, все процессы winword.exe будут убиты без сохранения!");
                    Stopwatch st = new Stopwatch();
                    st.Start();
                    foreach (var proc in Process.GetProcessesByName("winWord"))
                    {
                        proc.Kill();
                    }

                    if (!String.IsNullOrEmpty(textBox4.Text))
                    {
                        if (Directory.Exists(textBox4.Text))
                        {
                            foreach (string dirPath in Directory.GetDirectories(textBox1.Text, "*", SearchOption.AllDirectories))
                            {
                                try
                                {
                                    Directory.CreateDirectory(dirPath.Replace(textBox1.Text, textBox4.Text));//дублируем структуру папок                 
                                }
                                catch (Exception ee) { MessageBox.Show(ee.Message.ToString()); }
                            }

                            int counter = 0;
                            for (int i = 0; i < listView3.SelectedItems.Count; i++)//Выполняем замену в каждом выделенном файле
                            {
                                label6.Text = (i + 1) + " / " + listView3.SelectedItems.Count + " Working";
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
                                    SaveCloseFile(i);//закрываем открытый в word документ
                                }
                                catch (Exception ee) { MessageBox.Show(ee.Message.ToString() + "\nВозможно MS Office не установлен"); }
                                //}
                                counter++;

                                label6.Text = counter + " / " + listView3.SelectedItems.Count + " Done!";
                            }
                        }
                        else MessageBox.Show("Папка с результатами не существует");
                    }
                    else MessageBox.Show("Поле 'Папка с результатами' не заполнено");
                    button1.Enabled = true;
                    st.Stop(); label8.Text = st.Elapsed.ToString();
                }
                else MessageBox.Show("В поле 'Шаблоны замены' не выбран ни один файл");
            }
            else MessageBox.Show("В поле 'Шаблоны в папке' не выбран ни один файл");
        }

        // глобальные переменные
        public  Microsoft.Office.Interop.Word.Application app;
        Document doc;
        public static Object missing = Type.Missing;   
        
        public void OpenFile(int listbox1_id)// Открываем файл .doc 
        {
            app = new Microsoft.Office.Interop.Word.Application();
            string doc_file = listView3.SelectedItems[listbox1_id].SubItems[1].Text.ToString();
            doc = app.Documents.Open(doc_file);          //  app.Documents.Open(textBox1.Text + "\\" + listBox1.SelectedItems[listbox1_id]);
        }

        // Закрытие general и сохранение нового файла
        public void SaveCloseFile(int listbox1_id)
        {
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            // ОШИБКА В НАЧАЛЕ ФОРМИРУЕМОГО ФАЙЛА НЕТ УЧЕТА В КАКУЮ ПАПКУ ЕГО ПЕМЕСТИТЬ!!!
            string[] patches = listView3.SelectedItems[listbox1_id].SubItems[1].Text.Split('\\');
            string PatchName = "";
            for (int i = 0; i < patches.Length - 1; i++)
            {
                PatchName += patches[i] + "\\";
            }
            string[] fileTypes = listView3.SelectedItems[listbox1_id].SubItems[0].Text.Split('.');
            string FileName = "";
            for (int i = 0; i < fileTypes.Length - 1; i++)
            {
                if (i != fileTypes.Length - 1)
                    FileName += fileTypes[i] + ".";
                else FileName += fileTypes[i];
            }
            string FileType = fileTypes[fileTypes.Length - 1];
            string newfile = PatchName + FileName + " " + DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + " " + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "." + FileType; // новый файл на основе файла-шаблона
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

        /// <summary>
        /// Подгрузить замены в правое окно
        /// </summary>
        private void FilReps()
        {
            try
            {
                listView1.Items.Clear();
                string tt = listView2.Items[listView2.SelectedIndices[0]].SubItems[1].Text;
                if (File.Exists(tt))
                {
                    string[] reps = File.ReadAllLines(tt, Encoding.Default);
                    for (int L = 1; L < reps.Length; L++)
                    {
                        string[] rep = reps[L].Split('\t');
                        listView1.Items.Add(new ListViewItem(new string[] { rep[0], rep[1] }));
                    }
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
            }
        }
        /// <summary>
        /// Сохранить параметры программы в файл Parameters.txt
        /// </summary>
        private void Save_Parameters()
        {
            try
            {
                if (File.Exists("Parameters.txt"))
                {
                    string[] content = new string[2];
                    content[0] = "Template_patch\t" + textBox1.Text;
                    content[1] = "Output_folder\t" + textBox4.Text; //папка куда выгружать результаты

                    File.WriteAllLines("Parameters.txt", content);
                }
            }
            catch (Exception ee) {
                string ee3 = ee.ToString();
                MessageBox.Show(ee.InnerException.Message.ToString()); }
        }



        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/sergiomarotco/");
        }


        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            folderDlg.SelectedPath = Environment.CurrentDirectory;

            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox4.Text = folderDlg.SelectedPath;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            Save_Parameters();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Save_Parameters();
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            FilReps();
        }
    }
}
