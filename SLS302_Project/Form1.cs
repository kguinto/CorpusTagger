using System;
using System.Windows.Forms;
using edu.stanford.nlp.ling;
using edu.stanford.nlp.tagger.maxent;

using java.io;
using java.util;
using System.ComponentModel;

using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Data;
using Excel;

namespace SLS302_Project
{
    public partial class Form1 : Form
    {

        /* Single Tagging Page */
        private String inputText = "";
        private String taggedInputText = "";
        private ArrayList wordList = new ArrayList();
        //        private ArrayList lemmaList = new ArrayList();
        private int ttr = 0;
        private int[] wordCount = new int[4];
        private double[] ratios = new double[4];

        DataSet set = new DataSet("AllFiles");
        System.Data.DataTable mainTable = new System.Data.DataTable("student");

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            inputText = textBox1.Text;
            taggedInputText = Functions.runPOS(inputText);
            wordList = Functions.writeWordList(taggedInputText);
            wordCount = Functions.countTags(wordList);
            ratios = Functions.getRatios(wordCount, wordList);
        //    ttr = Functions.getTTR(inputText);

            richTextBox1.Text = taggedInputText;
            richTextBox2.Text = Functions.listVocab(wordList);
            richTextBox3.Text = writeReport();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
  
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
        //      Console.Write("Doing it fam");
        //      text1 = runPOS(textBox1.Text);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        //      richTextBox1.Text = text1;
        //      button1.Text = "Tag";
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            _Workbook oWB;
            _Worksheet oSheet;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (_Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Code";
                oSheet.Cells[1, 2] = "Country";
                oSheet.Cells[1, 3] = "Sex";
                oSheet.Cells[1, 4] = "Age";
                oSheet.Cells[1, 5] = "Grade";
                oSheet.Cells[1, 6] = "Major";

                oSheet.Cells[1, 7] = "VST";
                oSheet.Cells[1, 8] = "CEFR";

                oSheet.Cells[1, 9] = "Noun%";
                oSheet.Cells[1, 10] = "Verb%";
                oSheet.Cells[1, 11] = "Adjective%";
                oSheet.Cells[1, 12] = "Adverb%";
                oSheet.Cells[1, 13] = "TTR";


                //Add data
                oSheet.Cells[2, 9] = "Noun%";
                oSheet.Cells[1, 10] = "Verb%";
                oSheet.Cells[1, 11] = "Adjective%";
                oSheet.Cells[1, 12] = "Adverb%";
                oSheet.Cells[1, 13] = "TTR";


            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private String writeReport()
        {
            String outputString = "";

            outputString =
                     "Number of unique...\n" +
                     "words: " + wordList.size() + "\n" +
                     "nouns: " + wordCount[0] + "\n" +
                     "verbs: " + wordCount[1] + "\n" +
                     "adjectives: " + wordCount[2] + "\n" +
                     "adverbs: " + wordCount[3] + "\n\n" +

                     "Ratio to words...\n" +
                     "nouns: " + ratios[0] + "%\n" +
                     "verbs: " + ratios[1] + "%\n" +
                     "adjectives: " + ratios[2] + "%\n" +
                     "adverbs: " + ratios[3] + "%\n";



            return outputString;
        }

        private void openFileButton_Click(object sender, EventArgs e)
        {
            //            string winDir = System.Environment.GetEnvironmentVariable("windir");
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.

            String[] row = {"", "", "", "", "", "", "", "", "", "", "", "", "", ""};

            if (result == DialogResult.OK) // Test result.
            {
                string text = "";

                string[] files = openFileDialog1.FileNames;

                string[] filePath;
                string[] fileName;

                try
                {

                    int i = 0;
                    foreach (string f in files)
                    {
                        text = System.IO.File.ReadAllText(f);
                        outputBox.Text = "Processing " + f + "...";
                        outputBox.Update();
                        processText(text);

                        filePath = f.Split('\\');
                        fileName = filePath[filePath.Length-1].Split('_');

                        outputBox.Text = fileName[1];
                        outputBox.Update();

                           row[0] = fileName[1]; // Country
                           row[1] = fileName[3]; // Code
                                                 // TTR
                           row[3] = ratios[0].ToString();  // Noun Ratio
                           row[4] = ratios[1].ToString(); // Verb Ratio
                           row[5] = ratios[2].ToString(); // Adj Ratio
                           row[6] = ratios[3].ToString(); //Adv Ratio            
                                                    // VST
                                                 // CEFR

                           dataGridView1.Rows.Add(row);


                        i++;
                    }
                } catch (System.IO.IOException) { }

            }

        }

        private void processText(string text)
        {
            inputText = text;
            taggedInputText = Functions.runPOS(inputText);
            wordList = Functions.writeWordList(taggedInputText);
            wordCount = Functions.countTags(wordList);
            ratios = Functions.getRatios(wordCount, wordList);
        }

        private void addToExcelSheet(Microsoft.Office.Interop.Excel._Worksheet oSheet, int row)
        {

                //Add data
                oSheet.Cells[row, 9] = ratios[0];
                oSheet.Cells[row, 10] = ratios[1];
                oSheet.Cells[row, 11] = ratios[2];
                oSheet.Cells[row, 12] = ratios[3];
               // oSheet.Cells[row, 13] = "TTR";

        }

        private void outputBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void excelSheet_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();

  

            string[] files = openFileDialog1.FileNames;

            FileStream stream = System.IO.File.Open(openFileDialog1.FileNames[0], FileMode.Open, FileAccess.Read);
             //Choose one of either 1 or 2
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet ds = excelReader.AsDataSet();

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();

            System.Data.DataTable dtFinal = new System.Data.DataTable();

            foreach (System.Data.DataTable tmp in ds.Tables)
            {
                dtFinal.Merge(tmp);
            }

            dataGridView2.AutoGenerateColumns = true;
            dataGridView2.DataSource = dtFinal; // dataset


    //        dataGridView1.DataMember = dtFinal; // table name you need to show
        }
    }

    public class Functions
    {
        public static String runPOS(String text)
        {
            var jarRoot = @"C:\Users\Kirk\Documents\Visual Studio 2015\Projects\SLS302_Project\stanford-corenlp-3.6.0-models\edu\stanford\nlp";
            var modelsDirectory = jarRoot + @"\models";

            var text2 = "";

            // Loading POS Tagger
            var tagger = new MaxentTagger(modelsDirectory + @"\english-left3words-distsim.tagger");

            // Text for tagging
            //   var text = "A Part-Of-Speech Tagger (POS Tagger) is a piece of software that reads text in some language "
            //              + "and assigns parts of speech to each word (and other token), such as noun, verb, adjective, etc., although "
            //              + "generally computational applications use more fine-grained POS tags like 'noun-plural'.";

            var sentences = MaxentTagger.tokenizeText(new java.io.StringReader(text)).toArray();
            foreach (ArrayList sentence in sentences)
            {
                var taggedSentence = tagger.tagSentence(sentence);
                text2 += Sentence.listToString(taggedSentence, false);
            }


            return text2;
        }

        public static ArrayList writeWordList(String text)
        {
            string[] words = text.Split(new Char[] { ',', ' ', '.', '\n' });
            string[] splitS;
            ArrayList returnList = new ArrayList();

            foreach (string s in words)
            {
                splitS = s.Split('/');
                if (splitS.Length == 2 && !inList(splitS[0], returnList) && !splitS[0].Equals(""))
                    returnList.add(new VocabWord(splitS[0].ToLower(), splitS[1]));
            }

            return returnList;
        }

        public static Boolean inList(String s, ArrayList al)
        {
            Boolean toReturn = false;
            foreach (VocabWord vw in al)
            {
                if (s.Equals(vw.getWord(), StringComparison.InvariantCultureIgnoreCase))
                    toReturn = true;
            }

            return toReturn;
        }

        public static String listVocab(ArrayList list)
        {
            string output = "";

            foreach (VocabWord vw in list)
            {
                output += vw.getWord() + "\t" + vw.getPOS() + "\n";
            }

            return output;
        }

        public static int[] countTags(ArrayList list)
        {
            int verbCount = 0;
            int nounCount = 0;
            int adjCount = 0;
            int advCount = 0;

            foreach (VocabWord vw in list)
            {
                switch (vw.getPOS())
                {
                    case "JJ":
                        adjCount++;
                        break;
                    case "JJR":
                        adjCount++;
                        break;
                    case "JJS":
                        adjCount++;
                        break;

                    case "NN":
                        nounCount++;
                        break;
                    case "NNP":
                        nounCount++;
                        break;
                    case "NNS":
                        nounCount++;
                        break;
                    case "NNPS":
                        nounCount++;
                        break;
                    case "PRP":
                        nounCount++;
                        break;
                    case "PRP$":
                        nounCount++;
                        break;

                    case "RB":
                        advCount++;
                        break;
                    case "RBR":
                        advCount++;
                        break;
                    case "RBS":
                        advCount++;
                        break;

                    case "VB":
                        verbCount++;
                        break;
                    case "VBD":
                        verbCount++;
                        break;
                    case "VBG":
                        verbCount++;
                        break;
                    case "VBN":
                        verbCount++;
                        break;
                    case "VBP":
                        verbCount++;
                        break;
                    case "VBZ":
                        verbCount++;
                        break;
                }
            }

            int[] toReturn = new int[4];
            toReturn[0] = nounCount;
            toReturn[1] = verbCount;
            toReturn[2] = adjCount;
            toReturn[3] = advCount;

            return toReturn;
        }

        public static double[] getRatios(int[] wc, ArrayList wl)
        {
            double[] toReturn = new double[4];

            toReturn[0] = Math.Round(100.0 * wc[0] / (double)wl.size(), 2);
            toReturn[1] = Math.Round(100.0 * wc[1] / (double)wl.size(), 2);
            toReturn[2] = Math.Round(100.0 * wc[2] / (double)wl.size(), 2);
            toReturn[3] = Math.Round(100.0 * wc[3] / (double)wl.size(), 2);

            return toReturn;
        }

        private class VocabWord
        {
            private String word;
            private String pos;

            public VocabWord(String _word, String _pos)
            {
                word = _word;
                pos = _pos;
            }
            public String getWord() { return word; }
            public String getPOS() { return pos; }
        }
    }
}
