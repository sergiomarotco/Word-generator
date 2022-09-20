using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Word_generator
{
    /// <summary>
    /// Основное окно программы.
    /// </summary>
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Word.Application app;
        private Document doc;

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// Основное окно программы.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Открыть Word файл.
        /// </summary>
        /// <param name="listbox1_id">Номер файла из списка.</param>
        public void OpenFile(int listbox1_id) // Открываем файл .doc
        {
            app = new Microsoft.Office.Interop.Word.Application();
            string doc_file = listView3.SelectedItems[listbox1_id].SubItems[1].Text.ToString();
            doc = app.Documents.Open(doc_file); // app.Documents.Open(textBox1.Text + "\\" + listBox1.SelectedItems[listbox1_id]);
        }

        /// <summary>
        /// Закрытие general и сохранение нового файла.
        /// </summary>
        /// <param name="listbox1_id">Элемент листбокса.</param>
        public void SaveCloseFile(int listbox1_id)
        {
            string[] patches = listView3.SelectedItems[listbox1_id].SubItems[1].Text.Split('\\');
            string patchName = string.Empty;
            for (int i = 0; i < patches.Length - 1; i++)
            {
                patchName += patches[i] + "\\";
            }

            string[] fileTypes = listView3.SelectedItems[listbox1_id].SubItems[0].Text.Split('.');
            string fileName = string.Empty;
            for (int i = 0; i < fileTypes.Length - 1; i++)
            {
                if (i != fileTypes.Length - 1)
                {
                    fileName += fileTypes[i] + ".";
                }
                else
                {
                    fileName += fileTypes[i];
                }
            }

            string fileType = fileTypes[fileTypes.Length - 1];
            string newfile = patchName + fileName + " " + DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + " " + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "." + fileType; // новый файл на основе файла-шаблона
            newfile = newfile.Replace(textBox1.Text, textBox4.Text);
            newfile = newfile.Replace("Template_", string.Empty);
            app.ActiveDocument.SaveAs(newfile);
            app.ActiveDocument.Close();
            app.Quit();
            app = null;
        }

        /// <summary>
        /// Проверка открытия потока чтения файла.
        /// </summary>
        /// <param name="path">Путь к файлу.</param>
        /// <returns>Статус проверки возможности открытия файла на чтение.</returns>
        public bool FileIsOpen(string path)
        {
            FileStream a = null;

            try
            {
                a = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None);
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

        /// <summary>
        /// Функция замены строк в word файле.
        /// </summary>
        /// <param name="str_old">Заменяемая страка.</param>
        /// <param name="str_new">Строка которой заменяют.</param>
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
        /// Заполнить параметры в левой части программы из файла Parameters.txt.
        /// </summary>
        private void LoadParameters()
        {
            if (File.Exists("Parameters.txt"))
            {
                listView3.Items.Clear();
                listView2.Items.Clear();
                string[] parameters = File.ReadAllLines("Parameters.txt"); // загружаем файл в память

                for (int i = 0; i < parameters.Length; i++)
                {
                    string[] parameter = parameters[i].Split('\t');
                    switch (parameter[0])
                    {
                        case "Template_patch":
                            if (Directory.Exists(parameter[1]))
                            {
                                textBox1.Text = parameter[1]; // Заполняем поле "Папка с шаблонами"
                            }
                            else
                            {
                                textBox1.Text = Environment.CurrentDirectory;
                            }

                            try
                            {
                                string[] templates = new DirectoryInfo(textBox1.Text).GetFiles("*.doc*", SearchOption.AllDirectories).Select(f => f.FullName).ToArray();
                                for (int r = 0; r < templates.Length; r++)
                                {
                                    string[] patches = templates[r].Split('\\');
                                    listView3.Items.Add(new ListViewItem(new string[] { patches[patches.Length - 1], templates[r] }));
                                }

                                string[] replasments = new DirectoryInfo(textBox1.Text).GetFiles("Replacement*.txt", SearchOption.AllDirectories).Select(f => f.FullName).ToArray();
                                for (int r = 0; r < replasments.Length; r++)
                                {
                                    string[] lll = File.ReadAllLines(replasments[r]);
                                    if (lll[0].Equals("#do not delete this line#"))
                                    {
                                        string[] patches = replasments[r].Split('\\');
                                        listView2.Items.Add(new ListViewItem(new string[] { patches[patches.Length - 1], replasments[r] }));
                                    }
                                }
                            }
                            catch (Exception ee)
                            {
                                MessageBox.Show(ee.Message);
                            }

                            break;
                        case "Output_folder":
                            if (Directory.Exists(parameter[1]))
                            {
                                textBox4.Text = parameter[1];
                            }
                            else
                            {
                                textBox4.Text = Environment.CurrentDirectory;
                            }

                            break;
                        default: break;
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Icon = Properties.Resources.ico;
            splitContainer1.SplitterDistance = 571;
            LoadParameters();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (listView3.SelectedItems.Count > 0)
            {
                if (listView2.SelectedItems.Count == 1)
                {
                    label6.Text = string.Empty;
                    label8.Text = string.Empty;
                    button1.Enabled = false;
                    if (textBox4.Text.Length == 0)
                    {
                        FolderBrowserDialog fBD = new FolderBrowserDialog
                        {
                            ShowNewFolderButton = false,
                            Description = "Укажите путь к новой папке для копирования структуры папок и генерации файлов",
                            SelectedPath = Environment.CurrentDirectory,
                        };
                        if (fBD.ShowDialog() == DialogResult.OK)
                        {
                            textBox4.Text = fBD.SelectedPath;
                        }
                    }

                    MessageBox.Show("!После закрытия окна, все процессы winword.exe будут убиты без сохранения!");
                    Stopwatch st = new Stopwatch();
                    st.Start();
                    foreach (var proc in Process.GetProcessesByName("winWord"))
                    {
                        proc.Kill();
                    }

                    if (!string.IsNullOrEmpty(textBox4.Text))
                    {// проверка заполнен ли путь с папкой
                        if (Directory.Exists(textBox4.Text))
                        {// проверка создана ли папка куда сохранять
                            foreach (string dirPath in Directory.GetDirectories(textBox1.Text, "*", SearchOption.AllDirectories))
                            {
                                try
                                { // дублируем структуру папок
                                    Directory.CreateDirectory(dirPath.Replace(textBox1.Text, textBox4.Text));
                                }
                                catch (Exception ee)
                                {
                                    MessageBox.Show(ee.Message.ToString());
                                }
                            }

                            int counter = 0;
                            for (int i = 0; i < listView3.SelectedItems.Count; i++)
                            {// Выполняем замену в каждом выделенном файле
                                label6.Text = (i + 1) + " / " + listView3.SelectedItems.Count + " Working";
                                try
                                {
                                    OpenFile(i); // открываем в word документ
                                    for (int l = 0; l < listView1.Items.Count; l++)
                                    {
                                        FindReplace(listView1.Items[l].SubItems[0].Text, listView1.Items[l].SubItems[1].Text); // выполняем в тексте документа замену текста
                                    }

                                    SaveCloseFile(i); // закрываем открытый в word документ
                                }
                                catch (Exception ee)
                                {
                                    MessageBox.Show(ee.Message.ToString() + "\nВозможно MS Office не установлен");
                                }

                                counter++;
                                label6.Text = counter + " / " + listView3.SelectedItems.Count + " Done!";
                            }
                        }
                        else
                        {
                            MessageBox.Show("Папка с результатами не существует");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Поле 'Папка с результатами' не заполнено");
                    }

                    button1.Enabled = true;
                    st.Stop();
                    label8.Text = st.Elapsed.ToString();
                }
                else
                {
                    MessageBox.Show("В поле 'Шаблоны замены' не выбран ни один файл");
                }
            }
            else
            {
                MessageBox.Show("В поле 'Шаблоны в папке' не выбран ни один файл");
            }
        }

        /// <summary>
        /// Подгрузить замены в правое окно.
        /// </summary>
        private void Fill_Replacements()
        {
            try
            {
                listView1.Items.Clear();
                string tt = listView2.Items[listView2.SelectedIndices[0]].SubItems[1].Text;
                if (File.Exists(tt))
                {
                    string[] reps = File.ReadAllLines(tt, Encoding.Default);
                    for (int l = 1; l < reps.Length; l++)
                    {
                        string[] rep = reps[l].Split('\t');
                        listView1.Items.Add(new ListViewItem(new string[] { rep[0], rep[1] }));
                    }
                }
            }
            catch
            {
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Save_Parameters();
        }

        /// <summary>
        /// Выбрать папку с шаблонами.
        /// </summary>
        /// <param name="sender">.</param>
        /// <param name="e">..</param>
        private void Button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog
            {
                ShowNewFolderButton = true,
                SelectedPath = Environment.CurrentDirectory,
            };

            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
            }
        }

        /// <summary>
        /// Сохранить параметры программы в файл Parameters.txt.
        /// </summary>
        private void Save_Parameters()
        {
            try
            {
                if (File.Exists("Parameters.txt"))
                {
                    string[] content = new string[2];
                    content[0] = "Template_patch\t" + textBox1.Text; // папка с шаблонами файлов из которых делаются гтовые
                    content[1] = "Output_folder\t" + textBox4.Text; // папка куда выгружать результаты
                    File.WriteAllLines("Parameters.txt", content);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.InnerException.Message.ToString());
            }
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/sergiomarotco/");
        }

        /// <summary>
        /// Выбрать папку с результатами.
        /// </summary>
        /// <param name="sender">.</param>
        /// <param name="e">..</param>
        private void Button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog
            {
                ShowNewFolderButton = true,
                SelectedPath = Environment.CurrentDirectory,
            };

            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox4.Text = folderDlg.SelectedPath;
            }
        }

        private void TextBox4_TextChanged(object sender, EventArgs e)
        {
            Save_Parameters();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            Save_Parameters();
        }

        private void ListView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Fill_Replacements();
        }

        /// <summary>
        /// Действие нажатия на кнопку открытия папки с результатами.
        /// </summary>
        /// <param name="sender">.</param>
        /// <param name="e">..</param>
        private void Button6_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer", textBox4.Text);
            }
            catch
            {
            }
        }

        /// <summary>
        /// Действие нажатия на кнопку открытия папки с шаблонами.
        /// </summary>
        /// <param name="sender">..</param>
        /// <param name="e">.</param>
        private void Button4_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("explorer", textBox1.Text);
            }
            catch
            {
            }
        }
    }
}
