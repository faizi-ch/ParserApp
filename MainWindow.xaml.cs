using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using iTextSharp.text;
using iTextSharp.text.pdf;
using LumenWorks.Framework.IO.Csv;
using System.Drawing;
using SautinSoft.Document;
using Element = iTextSharp.text.Element;
using Run = System.Windows.Documents.Run;


namespace ParserApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("Work Item ID"), 
                new DataColumn("Issue key"),
                new DataColumn("Summary"),
                new DataColumn("Custom field (Estimated Points)"),
                new DataColumn("Custom field (Actual Points)")
            });

            using (CsvReader csv = new CsvReader(new StreamReader(@"C:\Users\Faizi\Downloads\2018-03-09T14_50_320100.csv"), false))
            {
                int id = 0;
                try
                {
                    while (csv.ReadNextRecord())
                    {

                        //string value = csv[0]; // "9"

                        // do something with value here
                        if (id!=0)
                        {
                            dt.Rows.Add("WI-" + id, csv[0], csv[3], csv[17], csv[18]);

                        }

                        id++;

                    }

                }
                catch (Exception)
                {

                    
                }
            }


            

            DataGrid.DataContext = dt.DefaultView;

            Document document = new Document(PageSize.A4);
            Font NormalFont = FontFactory.GetFont("Arial", 12, Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

            using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
            {


                PdfWriter writer =PdfWriter.GetInstance(document, new FileStream("zzz.pdf", FileMode.Create));

                /*PdfWriter*/ writer = PdfWriter.GetInstance(document, memoryStream);
                Phrase phrase = null;
                PdfPCell cell = null;
                PdfPTable table = null;
                //iTextSharp.text.BaseColor color=null;

                document.Open();

                PdfPTable imageTable=new PdfPTable(1);

                cell = ImageCell(@"logo.png", 30f, PdfPCell.ALIGN_RIGHT);
                imageTable.AddCell(cell);
                document.Add(imageTable);

                table = new PdfPTable(2);
                table.TotalWidth = 510f;//table size
                table.LockedWidth = true;
                table.SpacingBefore = 10f;//both are used to mention the space from heading
                table.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
                table.DefaultCell.Border = PdfPCell.RIGHT_BORDER | PdfPCell.LEFT_BORDER;


                phrase = new Phrase();
                phrase.Add(new Chunk("\nCompany Name\n\n\nAddress:"));
                phrase.Add(new Chunk("\n\nLeistungsnachweis / Proof of Performance\n\n"));
                phrase.Add(new Chunk("\n\n\n\n\nAuftragnehmer / Contractor:\nKunde / Client: \n"));
                phrase.Add(new Chunk("\nLeistungskatalog:"));
                PdfPCell c1 = new PdfPCell(phrase);
                c1.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                c1.BorderColorRight = BaseColor.WHITE;
                c1.BorderColorTop = BaseColor.WHITE;
                c1.BorderColorBottom = BaseColor.WHITE;
                c1.BorderColor = BaseColor.WHITE;
                table.DefaultCell.BorderColor = BaseColor.WHITE;

                

                Phrase phrase22 = new Phrase();
                phrase22.Add(new Chunk("\nDatum:\n\n\nfreigegeben"));
                phrase22.Add(new Chunk(" / approved:\n\n"));
                phrase22.Add(new Chunk("\n\n\n\n\n…………………………………….\nName Kunde\n"));
                phrase22.Add(new Chunk("Name Client\n\n"));
                phrase22.Add(new Chunk("\n\n\n\n\n…………………………………….\nName Kunde\n"));
                phrase22.Add(new Chunk("Name Client"));



                PdfPTable smallTable=new PdfPTable(2);
                Phrase p1 = new Phrase();
                p1.Add(new Chunk("p11:"));
                PdfPCell p1Cell=new PdfPCell(p1);

                Phrase p22 = new Phrase();
                p22.Add(new Chunk("p22:"));
                PdfPCell p2Cell = new PdfPCell(p22);
                smallTable.AddCell(p1Cell);
                smallTable.AddCell(p2Cell);

                PdfPCell cc2 = new PdfPCell(smallTable);
                cc2.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                cc2.BorderColorRight = BaseColor.WHITE;
                cc2.BorderColorTop = BaseColor.WHITE;
                cc2.BorderColorBottom = BaseColor.WHITE;

                table.DefaultCell.BorderColor = BaseColor.WHITE;

                cc2.BorderColor = BaseColor.WHITE;
                table.AddCell(c1);
                table.AddCell(cc2);
                
                table.DefaultCell.BorderColor = BaseColor.WHITE;

                //Header Table
                /*table = new PdfPTable(2);
                table.TotalWidth = 500f;
                table.LockedWidth = true;
                //table.SetWidths(new float[] {0.3f, 0.7f});

                //Company Logo
                cell = ImageCell(@"logo.png", 30f, PdfPCell.ALIGN_RIGHT);
                table.AddCell(cell);


                //Separater Line
                //color = new iTextSharp.text.BaseColor(System.Drawing.ColorTranslator.FromHtml("#A9A9A9"));
                DrawLine(writer, 25f, document.Top - 79f, document.PageSize.Width - 25f, document.Top - 79f, iTextSharp.text.BaseColor.GREEN);
                DrawLine(writer, 25f, document.Top - 80f, document.PageSize.Width - 25f, document.Top - 80f, iTextSharp.text.BaseColor.MAGENTA);
                document.Add(table);

                //table = new PdfPTable(1);
                //table.HorizontalAlignment = Element.ALIGN_LEFT;
                //table.SetWidths(new float[] { 0.3f, 1f });
                //table.SpacingBefore = 20f;

                cell = PhraseCell(new Phrase("Leistungskatalog:\n\n", FontFactory.GetFont("Arial", 14, Font.BOLD, iTextSharp.text.BaseColor.BLACK)), PdfPCell.ALIGN_CENTER);
                //cell.Colspan = 2;
                table.AddCell(cell);
                cell = PhraseCell(new Phrase(), PdfPCell.ALIGN_LEFT);
                //cell.Colspan = 2;
                //cell.PaddingBottom = 50f;
                table.AddCell(cell);

                //Name
                phrase = new Phrase();
                phrase.Add(new Chunk("aafaf", FontFactory.GetFont("Arial", 10, Font.BOLD, iTextSharp.text.BaseColor.BLACK)));
                phrase.Add(new Chunk("xfsfdg", FontFactory.GetFont("Arial", 8, Font.BOLD, iTextSharp.text.BaseColor.BLACK)));
                cell = PhraseCell(phrase, PdfPCell.ALIGN_LEFT);
                cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                table.AddCell(cell);*/
                document.Add(table);

                DrawLine(writer, 160f, 80f, 160f, 690f, iTextSharp.text.BaseColor.BLUE);
                DrawLine(writer, 115f, document.Top - 200f, document.PageSize.Width - 100f, document.Top - 200f, iTextSharp.text.BaseColor.BLACK);

                table = new PdfPTable(2);
                table.SetWidths(new float[] { 0.5f, 2f });
                table.TotalWidth = 340f;
                table.LockedWidth = true;
                table.SpacingBefore = 20f;
                table.HorizontalAlignment = Element.ALIGN_LEFT;


                PdfPTable table2 = new PdfPTable(dt.Columns.Count);
                table2.WidthPercentage = 100;

                //Set columns names in the pdf file
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    PdfPCell cell2 = new PdfPCell(new Phrase(dt.Columns[k].ColumnName));

                    cell2.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell2.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell2.BackgroundColor = new iTextSharp.text.BaseColor(51, 102, 102);

                    table2.AddCell(cell2);
                }

                //Add values of DataTable in pdf file
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        PdfPCell cell3 = new PdfPCell(new Phrase(dt.Rows[i][j].ToString()));

                        //Align the cell in the center
                        cell3.VerticalAlignment = PdfPCell.ALIGN_CENTER;

                        
                        if (j==2)
                        {
                            cell3.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        }
                        else
                        {
                            cell3.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        }
                        

                        table2.AddCell(cell3);
                    }
                }

                iTextSharp.text.Paragraph p;
                p = new iTextSharp.text.Paragraph("Alle oben aufgeführten Liefergegenstände und Dienstleistungen (Stories/Tasks) für diesen Sprint/Monat bzw. Berichtszeitraum wurden übergeben/geliefert. Der Fortschritt und Status der laufenden, noch nicht abgeschlossenen Stories/Tasks entspricht den Vereinbarungen.");
                p.Alignment= Element.ALIGN_JUSTIFIED;

                
                iTextSharp.text.Paragraph p2;
                Font f = new Font(Font.FontFamily.UNDEFINED, p.Font.Size, Font.ITALIC);
                p2 = new iTextSharp.text.Paragraph("\nAll deliverables and services (stories/tasks) for this sprint/month respectively reporting period as listed above have been provided/ delivered. The progress and status of the ongoing, unfinished Stories/Tasks correlates the agreed amount of work.\n",f);
                p2.Alignment = Element.ALIGN_JUSTIFIED;

                DrawLine(writer, 160f, 80f, 160f, 690f, iTextSharp.text.BaseColor.BLACK);
                DrawLine(writer, 115f, document.Top - 200f, document.PageSize.Width - 100f, document.Top - 200f, iTextSharp.text.BaseColor.BLACK);
                document.Add(table);
                document.Add(table2);
                document.Add(p);
                document.Add(p2);

                PdfPTable table4 = new PdfPTable(2);
                table4.TotalWidth = 510f;//table size
                table4.LockedWidth = true;
                table4.SpacingBefore = 10f;//both are used to mention the space from heading
                table4.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
                table4.DefaultCell.Border = PdfPCell.RIGHT_BORDER | PdfPCell.LEFT_BORDER;


                Phrase phrase1 = new Phrase();
                phrase1.Add(new Chunk("\nDatum:\n\n\ngeprüft"));
                phrase1.Add(new Chunk(" / verified:\n\n",f));
                phrase1.Add(new Chunk("\n\n\n\n\n…………………………………….\nName Auftragnehmer\n"));
                phrase1.Add(new Chunk("Name Contractor", f));
                PdfPCell c = new PdfPCell(phrase1);
                c.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                c.BorderColorRight = BaseColor.WHITE;
                c.BorderColorTop = BaseColor.WHITE;
                c.BorderColorBottom = BaseColor.WHITE;
                c.BorderColor= BaseColor.WHITE;
                table4.DefaultCell.BorderColor = BaseColor.WHITE;
                Phrase phrase2 = new Phrase();
                phrase2.Add(new Chunk("\nDatum:\n\n\nfreigegeben"));
                phrase2.Add(new Chunk(" / approved:\n\n", f));
                phrase2.Add(new Chunk("\n\n\n\n\n…………………………………….\nName Kunde\n"));
                phrase2.Add(new Chunk("Name Client\n\n", f));
                phrase2.Add(new Chunk("\n\n\n\n\n…………………………………….\nName Kunde\n"));
                phrase2.Add(new Chunk("Name Client", f));
                PdfPCell cc = new PdfPCell(phrase2);
                cc.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                cc.BorderColorRight=BaseColor.WHITE;
                cc.BorderColorTop= BaseColor.WHITE;
                cc.BorderColorBottom= BaseColor.WHITE;

                table4.DefaultCell.BorderColor = BaseColor.WHITE;

                cc.BorderColor=BaseColor.WHITE;
                table4.AddCell(c);
                table4.AddCell(cc);
                table4.DefaultCell.BorderColor=BaseColor.WHITE;
                document.Add(table4);


                document.Close();
                byte[] bytes = memoryStream.ToArray();
                memoryStream.Close();

                

                writer.Close();



                //Environment.Exit(0);
            }
        }
        private static void DrawLine(PdfWriter writer, float x1, float y1, float x2, float y2, iTextSharp.text.BaseColor color)
        {
            PdfContentByte contentByte = writer.DirectContent;
            contentByte.SetColorStroke(color);
            contentByte.MoveTo(x1, y1);
            contentByte.LineTo(x2, y2);
            contentByte.Stroke();
        }
        private static PdfPCell PhraseCell(Phrase phrase, int align)
        {
            PdfPCell cell = new PdfPCell(phrase);
            cell.BorderColor = iTextSharp.text.BaseColor.WHITE;
            cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
            cell.HorizontalAlignment = align;
            cell.PaddingBottom = 2f;
            cell.PaddingTop = 0f;
            return cell;
        }

        private static PdfPCell ImageCell(string path, float scale, int align)
        {
            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(path); 
            //image.ScalePercent(scale);
            image.ScaleAbsolute(125,40);
            //image.Height = 200;
            PdfPCell cell = new PdfPCell(image);
            cell.BorderColor = iTextSharp.text.BaseColor.WHITE;
            cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
            cell.HorizontalAlignment = align;
            cell.PaddingBottom = 0f;
            cell.PaddingTop = 0f;
            return cell;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
           
            FileInfo pathToDocx = new FileInfo(@"C:\Users\Faizi\Desktop\zzzz.docx");

            // Let's parse docx docuemnt and get all text from it. 
            DocumentCore docx = DocumentCore.Load(pathToDocx.FullName);

            StringBuilder text = new StringBuilder();

            foreach (var par in docx.GetChildElements(true, ElementType.TableRow))
            {

               // MessageBox.Show((par.Content.ToString()));
                Console.WriteLine((par.Content.ToString()));
            }

            // Show extracted text. 
            
            //Console.ReadLine();
        }
    }
}
