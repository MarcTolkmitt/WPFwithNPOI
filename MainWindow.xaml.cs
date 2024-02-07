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

using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.SS.UserModel;
using NPOIwrap;

namespace WPFwithNPOI
{
    /// <summary>
    /// interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // variables
        string fileName = @"Excel-Test.xlsx";
        NPOIexcel excel = new NPOIexcel();

        /// <summary>
        /// standardconstructor
        /// </summary>
        public MainWindow( )
        {
            InitializeComponent();

            Display( "Init ... ok\n" );
            string message = "This is the demoprogram for NPOIwrap.\n"
                + "Please try the menupoints in their order.\n"
                + "Every change can be seen in Excel if you are interested.\n"
                + "But you have to close the file in Excel before you can write it here!\n";
            Display( message );

        }   // end: public MainWindow

        /// <summary>
        /// handlerfunction -> Window_Closing
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void Window_Closing( object sender, System.ComponentModel.CancelEventArgs e )
        {

        }   // end: private void Window_Closing

        /// <summary>
        /// handlerfunction -> MenuItem
        /// used for exitroutines
        /// </summary>
        /// <param name="sender">triggering UI-element</param>
        /// <param name="e">send parameter from it</param>
        private void MenuQuit_Click( object sender, RoutedEventArgs e )
        {
            this.Close();

        }   // end: MenuQuit_Click



        // ---------------------------------------------     helperfunctions

        /// <summary>
        /// helperfunction, writing arraydata into a string
        /// </summary>
        /// <param name="data">2d ragged array </param>
        /// <returns>the data as string</returns>
        public string ArrayToString( double[][] data, bool textWrap = false )
        {
            string text = "";

            foreach ( double[] dat in data )
            {
                text += $" [ {string.Join( ", ", dat )} ] ";
                if ( textWrap )
                    text += "\n";

            }
            text += "\n";
            return ( text );

        }   // end: ArrayToString

        /// <summary>
        /// helperfunction to write the text into the mainwindow
        /// </summary>
        /// <param name="text">inputstring</param>
        public void Display( string? text )
        {
            if ( !string.IsNullOrEmpty( text ) )
                textBlock.Text += text + "\n";
            textScroll.ScrollToBottom();

        }   // end: Display

        /// <summary>
        /// helperfunction to write the text into the mainwindow
        /// </summary>
        /// <param name="text">any-objekt-variant</param>
        private void Display( int obj )
        {
            Display( obj.ToString() );

        }   // end: Display

        // -----------------------------------      the rest of the program

        private void MenuCreateHelloWorld_Click( object sender, RoutedEventArgs e )
        {
            excel.CreateHelloWorld();
            Display( $"created the Excel file with name {excel.fileName}.\n" );

        }   // end: MenuCreateHelloWorld_Click

        private void MenuReadHelloWorld_Click( object sender, RoutedEventArgs e )
        {
            excel.ReadHelloWorld();
            string message = "Reading the 'Hello World'-example:\n";
            Display( message + excel.DataListString_ToString() );

        }   // end: MenuReadHelloWorld_Click

        private void MenuCreateDoubleLnoHeader_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListDouble" + myExcel.fileEnding;
            string message = $"creating a new file ( {testName} ) for this.\n" +
                "A number of rows with doubles and no header...\n";
            Display( message );
            myExcel.CreateWorkbook();
            for ( int i = 0; i < 10; i++ )
            {   // add some data to be seen
                ExcelDataRowList newRow = new ExcelDataRowList( CellType.Numeric );
                for ( int j = 0; j < 10; j++ )
                { 
                    double result = i + 0.1 * j;
                    newRow.cellData.Add( result );

                }
                myExcel.dataListDouble.Add( newRow );

            }
            myExcel.CreateSheetFromListDouble( 0, "list of numeric cells" );
            myExcel.SaveWorkbook( testName, true );

        }   // end: MenuCreateDoubleLnoHeader_Click

        private void MenuReadDoubleList_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListDouble" + myExcel.fileEnding;
            string message = "reading the example file without header row:\n";
            myExcel.ReadWorkbook( testName, true );
            myExcel.ReadSheets();
            myExcel.ReadSheetAsListDouble( 0 );
            Display( message + myExcel.DataListDouble_ToString() );

        }   // end: MenuReadDoubleList_Click

        private void MenuReadAndAddHeader_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListDouble" + myExcel.fileEnding;
            string message = "reading the example file without header row ...\n"
                + "and adding an empty header row to it.\n"
                + "this will be table 1\n";
            Display( message );
            myExcel.ReadWorkbook( testName, true );
            myExcel.ReadSheets();
            myExcel.ReadSheetAsListDouble( 0 );
            myExcel.CreateSheetFromListDouble( 1, "double list cells with header", true );
            myExcel.SaveWorkbook( testName, true );

        }   // end: MenuReadAndAddHeader_Click

        private void MenuReadDoubleListHeader_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListDouble" + myExcel.fileEnding;
            string message = "reading the example file with header row:\n";
            myExcel.ReadWorkbook( testName, true );
            myExcel.ReadSheets();
            myExcel.ReadSheetAsListDouble( 1, true );
            Display( message + myExcel.DataListDouble_ToString( 1, true ) );


        }

        private void MenuChangeHeaderDL_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListDouble" + myExcel.fileEnding;
            string message = "changing the header and reloading...\n";
            myExcel.ReadWorkbook( testName, true );
            myExcel.ReadSheets();
            string[] heads = new string[] 
                { "is Mickey Mouse", "looking like", "Elvis",
                    "?", "YEAH", "while golfing..." };
            myExcel.ChangeHeader( 1, heads );
            // reread the file after changing the header
            myExcel.ReadWorkbook( testName, true );
            myExcel.ReadSheets();
            myExcel.ReadSheetAsListDouble( 1, true );
            Display( myExcel.DataListDouble_ToString( 1, true ) );

        }   // end: MenuChangeHeaderDL_Click

        private void MenuCreateMixedList_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListMixed" + myExcel.fileEnding;
            string message = $"creating a new file ( {testName} ) for this.\n" +
                "A number of rows with special data and no header...\n";
            Display( message );
            myExcel.CreateWorkbook();
            for ( int i = 0; i < 10; i++ )
            {   // add some data to be seen
                ExcelDataRow newRow = new ExcelDataRow();
                newRow.exampleIntNumber = i;
                newRow.exampleDoubleNumber = i * 1.1;
                newRow.exampleText = $"example number {i}";
                myExcel.dataListMixed.Add( newRow );

            }
            myExcel.CreateSheetFromListMixed( 0, "list of mixed cells" );
            myExcel.SaveWorkbook( testName, true );

        }   // end: MenuCreateMixedList_Click

        private void MenuReadMixedList_Click( object sender, RoutedEventArgs e )
        {
            NPOIexcel myExcel = new NPOIexcel();
            string testName = @"RowListMixed" + myExcel.fileEnding;
            string message = $"reading the file ( {testName} ) for this.\n" +
                "A number of rows with special data and no header...\n";
            Display( message );
            myExcel.ReadWorkbook( testName, true );
            myExcel.ReadSheets();
            myExcel.ReadSheetAsListMixed( 0 );
            Display( myExcel.DataListMixed_ToString() );

        }   // end: MenuReadMixedList_Click

    }   // end: class MainWindow

}   // end:namespace NPOIwithWPF
