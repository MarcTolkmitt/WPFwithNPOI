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

            Display( "Init ... ok" );

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
            Display( $"created the Excel file with name {excel.fileName}." );

        }   // end: MenuCreateHelloWorld_Click

        private void MenuReadHelloWorld_Click( object sender, RoutedEventArgs e )
        {
            excel.ReadHelloWorld();
            Display( excel.DataListString_ToString() );

        }   // end: MenuReadHelloWorld_Click

    }   // end: class MainWindow

}   // end:namespace NPOIwithWPF
