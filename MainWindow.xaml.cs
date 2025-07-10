/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


using System.Collections.Generic;
using System.Windows;

namespace OfficeTest
{
    /// <summary>
    /// interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// created on: 16.06.25
        /// last edit: 10.07.25
        /// </summary>
        System.Version version = new System.Version("1.0.5");

        Microsoft.Office.Interop.Word.Application wordApp = 
            new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document wordDocument =
            new Microsoft.Office.Interop.Word.Document();
        Microsoft.Office.Interop.Excel.Application excelApp = 
            new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook workbook;

        /// <summary>
        /// standard constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            Closing += Window_Closing;

            Display( "Init ... ok\n" );
            Display("Menüpunkt 'Excel beschicken' erstellt ein Double-Feld und sichert es in Excel.\n");
            Display("Menüpunkt 'Word Clipboard-Daten' nimmt aus der Excelfunktion" +
                " die Clipboard-Daten in ein Word-Dokument.\n");

        }   // end: public MainWindow

        /// <summary>
        /// handler function -> Window_Closing
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            excelApp.Quit();
            
            wordApp.Quit(
                Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges,
                Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument,
                false );

        }   // end: private void Window_Closing

        /// <summary>
        /// handler function -> MenuItem
        /// used for exit routines
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void MenuQuit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();

        }   // end: MenuQuit_Click



        // ---------------------------------------------     helper functions

        /// <summary>
        /// helper function, writing array data into a string
        /// </summary>
        /// <param name="data">2d ragged array </param>
        /// <returns>the data as string</returns>
        public string ArrayToString(double[][] data, bool textWrap = false)
        {
            string text = "";

            foreach (double[] dat in data)
            {
                text += $" [ {string.Join(", ", dat)} ] ";
                if (textWrap)
                    text += "\n";

            }
            text += "\n";
            return (text);

        }   // end: ArrayToString

        /// <summary>
        /// helper function to write the text into the main window
        /// </summary>
        /// <param name="text">input string</param>
        public void Display(string? text)
        {
            if (!string.IsNullOrEmpty(text))
                textBlock.Text += text + "\n";
            textScroll.ScrollToBottom();

        }   // end: Display

        /// <summary>
        /// helper function to write the text into the main window
        /// </summary>
        /// <param name="text">any-object-variant</param>
        private void Display(int obj)
        {
            Display(obj.ToString());

        }   // end: Display

        // ---------------------------------------------------  main functions

        private void MenuExcel_Click(object sender, RoutedEventArgs e)
        {
            // Add a new Excel workbook.
            workbook = excelApp.Workbooks.Add();
            excelApp.Visible = true;
            excelApp.Range["A1"].Value = "Feld X";
            excelApp.Range["B1"].Value = "Feld Y";
            //excelApp.Range["A2"].Select();

            double[,] feldDoubles = new double[2, 2]
                        { { 2.5, 3.6 },
                        { 7.3, 9.6 } };
            excelApp.Range["A2:B3"].Value = feldDoubles;

            excelApp.Columns[1].AutoFit();
            excelApp.Columns[2].AutoFit();
            excelApp.DisplayAlerts = false;

            bool saveOk = workbook.Saved;
            if ( !saveOk)  try
            {
                workbook.SaveAs(
                    "C:\\Users\\marct\\Downloads\\neueExcels", 
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, 
                    Type.Missing,               // password
                    Type.Missing,               // WriteResPassword
                    false,                      // ReadOnlyRecommended
                    false,                      // CreateBackup
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing,               // ConflictResolution
                    Type.Missing,               // AddToMru
                    Type.Missing,               // TextCodepage
                    Type.Missing,               // TextVisualLayout
                    Type.Missing                // Local
                    );

            }
            catch( Exception ex ) 
            {
                Display("MenuExcel_Click: Fehler beim Speichern -> " + ex.Message);

            }



            // Copy the results to the Clipboard.
            excelApp.Range["A1:B3"].Copy();


        }   // end: MenuExcel_Click

        private void MenuWord_Click(object sender, RoutedEventArgs e)
        {

            wordApp.Visible = true;
            wordDocument = wordApp.Documents.Add();
            wordApp.Selection.PasteSpecial(Link: true, DisplayAsIcon: false);
            bool saveOk = wordDocument.Saved;
            if (!saveOk) try
            {
                wordDocument.SaveAs2(
                    "C:\\Users\\marct\\Downloads\\neueWords.docx",
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument,
                    false,      // LockComments
                    "",         // Password
                    true,       // AddToRecentFiles
                    "",         // WritePassword
                    false,      // ReadOnlyRecommended
                    false,      // EmbedTrueTypeFonts
                    true,       // SaveNativePictureFormat
                    false,      // SaveFormsData
                    false,      // SaveAsAOCELetter
                    Microsoft.Office.Core.MsoEncoding.msoEncodingAutoDetect,        // 50001
                    false,      // InsertLineBreaks
                    false,      // AllowSubstitutions
                    Microsoft.Office.Interop.Word.WdLineEndingType.wdCRLF,
                    true        // AddBiDiMarks
                );

            }
            catch (Exception ex)
            {
                Display("MenuWord_Click: Fehler beim Speichern -> " + ex.Message);

            }

        }   // end: MenuWord_Click

    }   // end: class MainWindow

}   // end: namespace OfficeTest
