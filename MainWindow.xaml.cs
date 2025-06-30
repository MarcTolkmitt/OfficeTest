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


// Ignore Spelling: Fwith

using System.Windows;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Access = Microsoft.Office.Interop.Access;

namespace OfficeTest
{
    /// <summary>
    /// interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// created on: 16.06.25
        /// last edit: 16.06.25
        /// </summary>
        Version version = new Version("1.0.1");
        List<Account> bankAccounts = new List<Account>();

        /// <summary>
        /// standard constructor
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            bankAccounts.Add(
                new Account
                {
                    ID = 345,
                    Balance = 541.27
                } );
            bankAccounts.Add(
                new Account
                {
                    ID = 123,
                    Balance = -127.44
                } );

            Display( "Init ... ok" );
            
            DisplayInExcel( bankAccounts, ( account, cell ) =>
            // This multiline lambda expression sets custom processing rules
            // for the bankAccounts.
            {
                cell.Value = account.ID;
                cell.Offset[ 0, 1 ].Value = account.Balance;
                if ( account.Balance < 0 )
                {
                    cell.Interior.Color = 255;
                    cell.Offset[ 0, 1 ].Interior.Color = 255;
                }
            } );

            var wordApp = new Word.Application();
            wordApp.Visible = true;
            wordApp.Documents.Add();
            wordApp.Selection.PasteSpecial( Link: true, DisplayAsIcon: false );

        }   // end: public MainWindow

        /// <summary>
        /// handler function -> Window_Closing
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

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

        void DisplayInExcel( IEnumerable<Account> accounts,
                   Action<Account, Excel.Range> DisplayFunc )
        {
            var excelApp = new Excel.Application();
            // Add a new Excel workbook.
            excelApp.Workbooks.Add();
            excelApp.Visible = true;
            excelApp.Range[ "A1" ].Value = "ID";
            excelApp.Range[ "B1" ].Value = "Balance";
            excelApp.Range[ "A2" ].Select();

            foreach ( var ac in accounts )
            {
                DisplayFunc( ac, excelApp.ActiveCell );
                excelApp.ActiveCell.Offset[ 1, 0 ].Select();
            }
            // Copy the results to the Clipboard.
            excelApp.Range[ "A1:B3" ].Copy();

            excelApp.Columns[ 1 ].AutoFit();
            excelApp.Columns[ 2 ].AutoFit();
        }

    }   // end: class MainWindow

}   // end: namespace OfficeTest
