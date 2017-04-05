using System;
using System.Windows.Forms;
using edu.stanford.nlp.ling;
using edu.stanford.nlp.tagger.maxent;

using java.util;
using System.ComponentModel;

using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Data;
using Excel;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;

namespace SLS302_Project
{
    public partial class Form1 : Form
    {

        /* Single Tagging Page */
        private string inputText = "";
        private string taggedInputText = "";
        private ArrayList wordList = new ArrayList();
        private double ttr = -1;
        private int[] wordCount = new int[4];
        private double[] ratios = new double[5];
        private bool debugging = false;         //changing this to true creates a box on the top with useful info
        MaxentTagger tagger;

        DataSet set = new DataSet("AllFiles");
        System.Data.DataTable mainTable = new System.Data.DataTable("student");

        System.Data.DataSet categoryTables = new System.Data.DataSet();


        public Form1()
        {
            InitializeComponent();
           
        }

        /// <summary>
        /// Button no longer in use. Does nothing :-)
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            inputText = textBox1.Text;
            taggedInputText = runPOS(inputText);
            wordList = writeWordList(taggedInputText);
            wordCount = countTags(wordList);
            ratios = getRatios(wordCount, wordList);
            ttr = getTTR(inputText, wordList);

            richTextBox1.Text = taggedInputText;
            richTextBox2.Text = listVocab(wordList);
            richTextBox3.Text = writeReport();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// When this Form loads, 
        /// </summary>
        private void Form1_Shown(object sender, EventArgs e)
        {
            /* If the the debug variable is set to true, a text box is visible so things can be written and shown.*/
            if (debugging)
                debugBox.Visible = true;
            else
                debugBox.Visible = false;


            /* Load tagger file and create the tagger object */
            outputBox.Text = "Loading tagger...";
            outputBox.Update();

            if (debugging) { debugBox.Text = "Current directory: " + Directory.GetCurrentDirectory(); debugBox.Update(); }
            tagger = new MaxentTagger(@"../../resources/english-left3words-distsim.tagger");


            /* Initialize mainTable */
            mainTable.Columns.Add("Country", typeof(string));
            mainTable.Columns.Add("Code", typeof(string));
            mainTable.Columns.Add("Text", typeof(string));
            mainTable.Columns.Add("TTR", typeof(double));
            mainTable.Columns.Add("ContentWordRatio", typeof(double));
            mainTable.Columns.Add("NounRatio", typeof(double));
            mainTable.Columns.Add("VerbRatio", typeof(double));
            mainTable.Columns.Add("AdjRatio", typeof(double));
            mainTable.Columns.Add("AdvRatio", typeof(double));


            UniqueConstraint custUnique = new UniqueConstraint(new DataColumn[] { mainTable.Columns["Code"] });
            mainTable.Constraints.Add(custUnique);

            /* Set source for the Table view to mainTable*/
            dataGridView1.DataSource = mainTable;
            dataGridView1.Columns["Text"].Visible = false;

            /* Set source for the table in Chart view to mainTable*/
            dataGridView2.DataSource = mainTable;
            dataGridView2.Columns["Text"].Visible = false;

            /* Set source for the Chart view to mainTable */
            foreach (DataColumn cl in mainTable.Columns)
            {
                if (!cl.DataType.Equals(typeof(string)))
                {
                    chart1_xBox.Items.Add(cl.ColumnName.ToString());
                    chart1_yBox.Items.Add(cl.ColumnName.ToString());
                }
                else
                {
                    categoryBox.Items.Add(cl.ColumnName.ToString());
                }
            }

            /* Done with everything in loading form */
            outputBox.Text = "Done...";
            outputBox.Update();
        }


        /**********************************************************************************/
        /* These do nothing :-) */
        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e) {}
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {}
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e){}

        /* These used to do stuff, but now also do nothing. :-) */
        private void exportButton_Click(object sender, EventArgs e)
        { /*
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
                string errorMessage;
                errorMessage = "Error: ";
                errorMessage = string.Concat(errorMessage, theException.Message);
                errorMessage = string.Concat(errorMessage, " Line: ");
                errorMessage = string.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            } */
        }
        private string writeReport()
        {
            string outputstring = "";
            /*
            outputstring =
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


            */
            return outputstring; 
        }
        
        /**********************************************************************************/

        /// <summary>
        /// This opens text files, sticks the data into mainTable, and updates the Table view
        /// </summary>
        private void openFileButton_Click(object sender, EventArgs e)
        {
            debugBox.Text = "showing dialog";
            debugBox.Update();

            DialogResult result = textFilesDialog.ShowDialog(); // Show the dialog.

            if (result == DialogResult.OK) // Test result.
            {
                string text = "";

                string[] files = textFilesDialog.FileNames;

                string[] filePath;
                string[] fileName;

                debugBox.Text = "Gonna process files";
                debugBox.Update();

                DataRow row;

                int i = 1;
                foreach (string f in files)
                {
                    row = mainTable.NewRow();
                    text = System.IO.File.ReadAllText(f);

                    outputBox.Text = "Processing file " + i + " of " + (files.Length).ToString() + "...";
                    outputBox.Update();

                    processText(text);

                    filePath = f.Split('\\');
                    fileName = filePath[filePath.Length-1].Split('_');                     

                    row["Country"] = fileName[1]; // Country
                    row["Code"] = "W_" + fileName[1] + "_" + fileName[3]; // Code
                    row["Text"] = inputText;
                    row["TTR"] = ttr.ToString();  // TTR
                    row["NounRatio"] = ratios[0];  // Noun Ratio
                    row["VerbRatio"] = ratios[1]; // Verb Ratio
                    row["AdjRatio"] = ratios[2]; // Adj Ratio
                    row["AdvRatio"] = ratios[3]; //Adv Ratio 
                    row["ContentWordRatio"] = ratios[4];      

                    try
                    {
                        mainTable.Rows.Add(row);
                    } catch(System.Data.ConstraintException) {}                    

                    i++;
                }

                dataGridView1.Update();
                              
                outputBox.Text = "Done.";
                outputBox.Update();
                
            }

        }

        /// <summary>
        /// Calls each function to analyze essay
        /// </summary>
        /// <param name="text"> The essay to be analyzed </param>
        private void processText(string text)
        {
            if (debugging) { debugBox.Text = "Getting input text"; debugBox.Update(); }
            inputText = text;

            if (debugging){debugBox.Text = "Running POS Tagger"; debugBox.Update();}
            taggedInputText = runPOS(inputText);

            if (debugging) { debugBox.Text = "Writing word list"; debugBox.Update(); }
            wordList = writeWordList(taggedInputText);

            if (debugging) { debugBox.Text = "Counting tags"; debugBox.Update(); }
            wordCount = countTags(wordList);

            if (debugging) { debugBox.Text = "Calculating ratios"; debugBox.Update(); }
            ratios = getRatios(wordCount, wordList);

            if (debugging) { debugBox.Text = "Calculating TTR"; debugBox.Update(); }
            ttr = getTTR(inputText, wordList);
            
        }

        /*********************************/
        /* More old functions that do nothing :-) */

        private void addToExcelSheet(Microsoft.Office.Interop.Excel._Worksheet oSheet, int row)
        {
            /*
                //Add data
                oSheet.Cells[row, 9] = ratios[0];
                oSheet.Cells[row, 10] = ratios[1];
                oSheet.Cells[row, 11] = ratios[2];
                oSheet.Cells[row, 12] = ratios[3];
               // oSheet.Cells[row, 13] = "TTR";
               */
        }

        private void outputBox_TextChanged(object sender, EventArgs e){}

        /*********************************/

        /// <summary>
        /// When a cell on table is clicked, the essay text is shown on the side panel.
        /// </summary>
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            if (rowIndex>=0)
               essayBox.Text = mainTable.Rows[rowIndex]["Text"].ToString();
        }

        /// <summary>
        /// Loads additional information about each essay from an excel sheet; most importantly the VST and CEFR scores.
        /// </summary>
        private void excelSheetButton_Click(object sender, EventArgs e)
        {
            DialogResult result = excelFileDialog.ShowDialog();

            outputBox.Text = "Loading data sheet...";
            outputBox.Update();

            if (result == DialogResult.OK)
            {
                string[] files = excelFileDialog.FileNames;

                if (debugging) { debugBox.Text = "Opening excel file"; debugBox.Update(); }

                try
                {
                    FileStream stream = System.IO.File.Open(excelFileDialog.FileNames[0], FileMode.Open, FileAccess.Read);
                    //Choose one of either 1 or 2
                    //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                    //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

                    if (debugging) { debugBox.Text = "opening excelReader"; debugBox.Update(); }
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    if (debugging) { debugBox.Text = "Creating column names"; debugBox.Update(); }
                    //4. DataSet - Create column names from first row
                    excelReader.IsFirstRowAsColumnNames = true;
                    DataSet ds = excelReader.AsDataSet();

                    if (debugging) { debugBox.Text = "Closing excelReader"; debugBox.Update(); }
                    //6. Free resources (IExcelDataReader is IDisposable)
                    excelReader.Close();

                    if (debugging) { debugBox.Text = "Creating dtFinal"; debugBox.Update(); }
                    System.Data.DataTable dtFinal = new System.Data.DataTable();

                    if (debugging) { debugBox.Text = "Merging dtFinal"; debugBox.Update(); }
                    foreach (System.Data.DataTable tmp in ds.Tables)
                    {
                        dtFinal.Merge(tmp);
                    }

                    if (debugging) { debugBox.Text = "Adding columns to mainTable from dtFinal"; debugBox.Update(); }
                    foreach (DataColumn cl in dtFinal.Columns)
                    {
                        if (!mainTable.Columns.Contains(cl.ColumnName))
                        {
                            mainTable.Columns.Add(cl.ColumnName, cl.DataType);
                            if (!cl.DataType.Equals(typeof(string)))
                            {
                                chart1_xBox.Items.Add(cl.ColumnName.ToString());
                                chart1_yBox.Items.Add(cl.ColumnName.ToString());
                            }
                        }
                    }

                    DataRow[] foundRows;

                    int i = 0;
                    foreach (DataRow row in mainTable.Rows)
                    {
                        i++;
                        outputBox.Text = "Importing data for row " + i + " of " + mainTable.Rows.Count;
                        outputBox.Update();

                        if (debugging) { debugBox.Text = "Selecting row in dtFinal containing " + row["Code"]; debugBox.Update(); }
                        foundRows = dtFinal.Select("Code = \'" + row["Code"] + "'"); //Get matching rows from tables

                        if (foundRows.Length > 0)
                        {
                            foreach (DataColumn cl in foundRows[0].Table.Columns)
                            {
                                if (debugging) { debugBox.Text = "Adding column " + cl.ColumnName + " to row " + row["Code"]; debugBox.Update(); }
                                if (row.Table.Columns.Contains(cl.ColumnName))
                                {
                                    row[cl.ColumnName] = foundRows[0][cl.ColumnName];
                                }
                            }

                        }

                        debugBox.Update();
                    }

                    outputBox.Text = "Done.";
                    outputBox.Update();
                }
                catch (Exception ex)
                {
                    outputBox.Text = "!!! Could not load data sheet. It may be open in another program.";
                    outputBox.Update();
                }

            }
        }

        /// <summary>
        /// Runs the POS Tagger on the text
        /// </summary>
        /// <param name="text"></param>
        /// <returns>The tagged version of the text. </returns>
        public string runPOS(string text)
        {
            var text2 = "";

            if (debugging) { debugBox.Text = "Tokenizing text"; debugBox.Update(); }
            var sentences = MaxentTagger.tokenizeText(new java.io.StringReader(text)).toArray();

            if (debugging) { debugBox.Text = "Tagging sentences"; debugBox.Update(); }
            foreach (ArrayList sentence in sentences)
            {
                var taggedSentence = tagger.tagSentence(sentence);
                text2 += Sentence.listToString(taggedSentence, false);
            }

            return text2;
        }

        /// <summary>
        /// Writes out a list of each unique word in the text. This uses a class called VocabWord, which 
        /// contains a word as well as its part of speech.
        /// </summary>
        /// <param name="text">Essay text</param>
        /// <returns>ArrayList of VocabWords containing each word in the text and its POS</returns>
        public ArrayList writeWordList(string text)
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

        /// <summary>
        /// Checks if the word in the text is already in the word list
        /// </summary>
        /// <param name="s">A word in the text</param>
        /// <param name="al">ArrayList of words in the text so far</param>
        /// <returns>True if the word is there already, false if not</returns>
        public Boolean inList(string s, ArrayList al)
        {
            Boolean toReturn = false;
            foreach (VocabWord vw in al)
            {
                if (s.Equals(vw.getWord(), StringComparison.InvariantCultureIgnoreCase))
                    toReturn = true;
            }

            return toReturn;
        }

        /// <summary>
        /// Calculates the Type-Token Ratio of the text
        /// </summary>
        /// <param name="text">Essay text</param>
        /// <param name="wl">list of VocabWords </param>
        /// <returns>The type token ratio of the text</returns>
        public  double getTTR(string text, ArrayList wl)
        {
            double types = wl.size();
            double tokens = (text.Split(new Char[] { ',', ' ', '.', '\n' })).Length;

            return Math.Round(100.0 * types / (double) tokens, 2);
        }

        /* Another no longer used function */
        public  string listVocab(ArrayList list)
        { 
            string output = "";
            /*
            foreach (VocabWord vw in list)
            {
                output += vw.getWord() + "\t" + vw.getPOS() + "\n";
            }
            */
            return output;
        }

        /// <summary>
        /// Iterates through the word list and counts the number of each part of speech
        /// </summary>
        /// <param name="list">list of VocabWords</param>
        /// <returns>An array of integers representing the count of each POS</returns>
        public  int[] countTags(ArrayList list)
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

        /// <summary>
        /// Calculates the ratio of each part of speech to the total size of the vocabulary
        /// </summary>
        /// <param name="wc">An array of integers representing the count of each POS</param>
        /// <param name="wl">The list of VocabWords</param>
        /// <returns>An array of doubles representing the ratio of each part of speech to the total size of the vocabulary</returns>
        public double[] getRatios(int[] wc, ArrayList wl)
        {
            double[] toReturn = new double[5];

            toReturn[0] = Math.Round(100.0 * wc[0] / (double)wl.size(), 2);
            toReturn[1] = Math.Round(100.0 * wc[1] / (double)wl.size(), 2);
            toReturn[2] = Math.Round(100.0 * wc[2] / (double)wl.size(), 2);
            toReturn[3] = Math.Round(100.0 * wc[3] / (double)wl.size(), 2);
            toReturn[4] = Math.Round(100.0 * (wc[0]+wc[1]+wc[2]+wc[3]) / (double)wl.size(), 2);


            return toReturn;
        }

        /// <summary>
        /// Computes the coefficient of two arrays
        /// Credit: http://stackoverflow.com/questions/17447817/correlation-of-two-arrays-in-c-sharp */
        /// </summary>
        /// <param name="values1">The first array</param>
        /// <param name="values2">The second array</param>
        /// <returns></returns>
        public double ComputeCoeff(double[] values1, double[] values2)
        {
            if (values1.Length != values2.Length)
                throw new ArgumentException("values must be the same length");

            var avg1 = values1.Average();
            var avg2 = values2.Average();

            var sum1 = values1.Zip(values2, (x1, y1) => (x1 - avg1) * (y1 - avg2)).Sum();

            var sumSqr1 = values1.Sum(x => Math.Pow((x - avg1), 2.0));
            var sumSqr2 = values2.Sum(y => Math.Pow((y - avg2), 2.0));

            var result = sum1 / Math.Sqrt(sumSqr1 * sumSqr2);

            return result;
        }

        /// <summary>
        /// A class that contains a word and its part of speech, both strings
        /// </summary>
        private class VocabWord
        {
            private string word;
            private string pos;

            public VocabWord(string _word, string _pos)
            {
                word = _word;
                pos = _pos;
            }
            public string getWord() { return word; }
            public string getPOS() { return pos; }
        }

        /* More nothing :-) */
        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e){}

        /* Draws the graph with axes being the two dropdowns */
        private void drawGraphButton_Click(object sender, EventArgs e)
        {
            outputBox.Text = "Drawing graph...";
            outputBox.Update();

            chart1.Series.Clear();

            string xAxis = chart1_xBox.Text;
            string yAxis = chart1_yBox.Text;

            try
            {
                double[] xArray = mainTable.AsEnumerable().Select(r => r.Field<double>(xAxis)).ToArray();
                double[] yArray = mainTable.AsEnumerable().Select(r => r.Field<double>(yAxis)).ToArray(); ;
                double coeff = ComputeCoeff(xArray, yArray);
                coeffBox.Text = "coeff: " + coeff; coeffBox.Update();
            } catch (Exception exc)
            {
                debugBox.Text = "Couldn't get coeff because " + exc.ToString();
            }

         /*   if (!categoryBox.Text.Equals(""))
            {
                DataView view = new DataView(mainTable);
                System.Data.DataTable distinctValues = view.ToTable(true, categoryBox.Text);
                //foreach()
                if (debugging)
                {
                    debugBox.Text = "";
                    foreach (DataRow row in distinctValues.Rows)
                        debugBox.Text += row[categoryBox.Text];
                    debugBox.Update();
                }

                foreach (DataRow row in distinctValues.Rows)
                {
                    string category = "" + row[categoryBox.Text];
                    categoryTables = new DataSet();

                    vfSystem.Data.DataView view = new System.Data.DataView(table);
                    System.Data.DataTable dt = view.ToTable("Selected", false, "col1", "col2", "col6", "col7", "col3");

                    categoryTables.Tables.Add(dt);

                    dt.Columns.Add("" + xAxis);
                    dt.Columns.Add("" + yAxis);

                    DataRow[] foundRows;
                    foundRows = mainTable.query("" + xAxis + ", " + yAxis + "where" + categoryBox.Text + " =  \'" + row[categoryBox.Text] + "\'"); //Get matching rows from tables

                    foreach (DataRow fr in foundRows)
                    {
                        dt.Rows.Add(fr.ItemArray);
                    }

                    chart1.Series.Add(category);
                    chart1.Series[category].ChartType = SeriesChartType.Line;
                    chart1.Series[category].IsXValueIndexed = false;

                    chart1.Series[category].XValueMember = xAxis;
                    chart1.Series[category].YValueMembers = yAxis;
                }
                    chart1.DataSource = categoryTables;
            }
            else
            {*/

            chart1.Series.Add("test");
            chart1.Series["test"].ChartType = SeriesChartType.Point;
            chart1.Series["test"].MarkerStyle = MarkerStyle.Circle;

            chart1.Series["test"].IsXValueIndexed = false;

            chart1.Series["test"].XValueMember = chart1_xBox.Text;
            chart1.Series["test"].YValueMembers = chart1_yBox.Text;

            chart1.DataSource = mainTable;
            

            chart1.DataBind();
            chart1.Series["test"].Sort(PointSortOrder.Ascending, "X");

         //   }
            outputBox.Text = "Done.";
            outputBox.Update();
        }

    }
}
