using DevExpress.Export;
using DevExpress.Utils.CommonDialogs.Internal;
using DevExpress.XtraPrinting;
using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Levira1
{
    public partial class ProdutosTexteis : System.Web.UI.Page
    {
        readonly string conString = ConfigurationManager.ConnectionStrings
              ["DefaultConnection"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {
          
        }

        protected void ASPxButton1_Click(object sender, EventArgs e)
        {
            ASPxGridViewExporter1.WriteXlsxToResponse(new XlsxExportOptionsEx { ExportType = ExportType.WYSIWYG });
            /*PdfExportOptions options = new PdfExportOptions();
            options.Compressed = false;
            ASPxGridView1.ExportPdfToResponse(options);*/
        }

        public DataTable GetDataTabletFromCSVFile(string csv_file_path)
        {
            DataTable csvData = new DataTable();

            DataColumn code = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "code",
                AutoIncrement = false
            };
            csvData.Columns.Add(code);

            DataColumn description = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "description",
                AutoIncrement = false
            };
            csvData.Columns.Add(description);

            DataColumn Fulldescription = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Fulldescription",
                AutoIncrement = false
            };
            csvData.Columns.Add(Fulldescription);

            DataColumn sku = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "sku",
                AutoIncrement = false
            };
            csvData.Columns.Add(sku);

            DataColumn ean = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "ean",
                AutoIncrement = false
            };
            csvData.Columns.Add(ean);

            DataColumn brand = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "brand",
                AutoIncrement = false
            };
            csvData.Columns.Add(brand);

            DataColumn color = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "color",
                AutoIncrement = false
            };
            csvData.Columns.Add(color);

            DataColumn family = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "family",
                AutoIncrement = false
            };
            csvData.Columns.Add(family);

            DataColumn lenght = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "lenght",
                AutoIncrement = false
            };
            csvData.Columns.Add(lenght);

            DataColumn breadth = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "breadth",
                AutoIncrement = false
            };
            csvData.Columns.Add(breadth);

            DataColumn tall = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tall",
                AutoIncrement = false
            };
            csvData.Columns.Add(tall); 
            
            DataColumn volume = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "volume",
                AutoIncrement = false
            };
            csvData.Columns.Add(volume);

            DataColumn weight = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "weight",
                AutoIncrement = false
            };
            csvData.Columns.Add(weight);

            DataColumn dadson = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "dadson",
                AutoIncrement = false
            };
            csvData.Columns.Add(dadson);

            DataColumn related = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "related",
                AutoIncrement = false
            };
            csvData.Columns.Add(related);

            DataColumn variation1 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "variation1",
                AutoIncrement = false
            };
            csvData.Columns.Add(variation1);

            DataColumn variation2 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "variation2",
                AutoIncrement = false
            };
            csvData.Columns.Add(variation2);

            DataColumn descriptive1 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "descriptive1",
                AutoIncrement = false
            };
            csvData.Columns.Add(descriptive1);

            DataColumn descriptive2 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "descriptive2",
                AutoIncrement = false
            };
            csvData.Columns.Add(descriptive2);

            DataColumn descriptive3 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "descriptive3",
                AutoIncrement = false
            };
            csvData.Columns.Add(descriptive3);

            DataColumn descriptive4 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "descriptive4",
                AutoIncrement = false
            };
            csvData.Columns.Add(descriptive4);

            DataColumn descriptive5 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "descriptive5",
                AutoIncrement = false
            };
            csvData.Columns.Add(descriptive5);

            DataColumn returnpolicy = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "returnpolicy",
                AutoIncrement = false
            };
            csvData.Columns.Add(returnpolicy);

            DataColumn shippingdescription = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "shippingdescription",
                AutoIncrement = false
            };
            csvData.Columns.Add(shippingdescription);           

            DataColumn url1 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "url1",
                AutoIncrement = false
            };
            csvData.Columns.Add(url1);

            DataColumn url8 = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "url8",
                AutoIncrement = false
            };
            csvData.Columns.Add(url8);

            DataColumn size = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "size",
                AutoIncrement = false
            };
            csvData.Columns.Add(size);

            DataColumn category = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "category",
                AutoIncrement = false
            };
            csvData.Columns.Add(category);

            DataColumn tags = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tags",
                AutoIncrement = false
            };
            csvData.Columns.Add(tags);

            DataColumn price = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "price",
                AutoIncrement = false
            };
            csvData.Columns.Add(price);

            DataColumn stock = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "stock",
                AutoIncrement = false
            };
            csvData.Columns.Add(stock);

            /*DataColumn useraction = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "useraction",
                AutoIncrement = false
            };
            csvData.Columns.Add(useraction);

            DataColumn dateauto = new DataColumn
            {
                DataType = Type.GetType("SqlDateTime"),
                ColumnName = "dateauto",
                AutoIncrement = false
            };
            csvData.Columns.Add(dateauto);*/
          
            try
            {  
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    bool firstLine = true;
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    /*string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }*/
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Making empty value as null
                        // get the column headers
                        if (firstLine)
                        {
                            firstLine = false;

                            continue;
                        }
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('" + ex.Message.Replace("\'", " ") + "')</script>");
            }
            return csvData;
        }

        public void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData)
        {
            try
            {
               
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conString))
                {
                    bulkCopy.BulkCopyTimeout = 60;
                    bulkCopy.DestinationTableName = "[dbo].[leviraprodtexteis]";
                    //limite rows a enviar--> bulkCopy.BatchSize = 0;
                    /*bulkCopy.ColumnMappings.Add("code", "code");
                    bulkCopy.ColumnMappings.Add("stock", "stock");*/
                    bulkCopy.WriteToServer(csvFileData);
                    bulkCopy.Close();
                }
                /*using (SqlConnection dbConnection = new SqlConnection(conString))
                {              
                    dbConnection.Open();
                    using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                    {                       
                        s.DestinationTableName = "[dbo].[leviraprodtexteis]";
                        foreach (var column in csvFileData.Columns)
                            s.ColumnMappings.Add(column.ToString(), column.ToString());
                            s.WriteToServer(csvFileData);                  
                    }
                    
                }  */
            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('" + ex.Message.Replace("\'", " ") + "')</script>");                
            }
        }

        /*void DisplayFileContents(HttpPostedFile file)
        {
            System.IO.Stream myStream;
            int fileLen;
            StringBuilder displayString = new StringBuilder();

            // Get the length of the file.
            fileLen = FileUpload1.PostedFile.ContentLength;

            // Display the length of the file in a label.
            LengthLabel.Text = "The length of the file is " +
                               fileLen.ToString() + " bytes.";

            // Create a byte array to hold the contents of the file.
            byte[] Input = new byte[fileLen];

            // Initialize the stream to read the uploaded file.
            myStream = FileUpload1.FileContent;

            // Read the file into the byte array.
            myStream.Read(Input, 0, fileLen);

            // Copy the byte array to a string.
            for (int loop1 = 0; loop1 < fileLen; loop1++)
            {
                displayString.Append(Input[loop1].ToString());
            }

            // Display the contents of the file in a 
            // textbox on the page.
            ContentsLabel.Text = "The contents of the file as bytes:";

            TextBox ContentsTextBox = new TextBox
            {
                TextMode = TextBoxMode.MultiLine,
                Height = Unit.Pixel(300),
                Width = Unit.Pixel(400),
                Text = displayString.ToString()
            };

            // Add the textbox to the Controls collection
            // of the Placeholder control.
            PlaceHolder1.Controls.Add(ContentsTextBox);

        }*/
        
        //public static readonly string[] _csvMimeTypes = new[] { "image/jpeg", "image/png" };
        //public static readonly string[] _csvMimeTypes = new[] { "text/csv" };
        protected void UploadButton_Click(object sender, EventArgs e)
        {
            // Specify the path on the server to
            // save the uploaded file to.
            string savePath = @"C:\testedestino\";
            string csv_file_path = savePath + FileUpload1.FileName;
            // Before attempting to perform operations
            // on the file, verify that the FileUpload 
            // control contains a file.
            //if (FileUpload1.HasFile)
            //if (_imageMimeTypes.Contains(fupFirmLogo.PostedFile.ContentType))
            //testarcsv.IsValidcsv(FileUpload1.PostedFile.ContentType
            if (FileUpload1.HasFile && FileUpload1.FileName != string.Empty && FileUpload1.FileContent.Length > 0 && csv_file_path.IsValidcsv())
            {
                    // Append the name of the file to upload to the path.
                    savePath += FileUpload1.FileName;

                    // Call the SaveAs method to save the 
                    // uploaded file to the specified path.
                    // This example does not perform all
                    // the necessary error checking.               
                    // If a file with the same name
                    // already exists in the specified path,  
                    // the uploaded file overwrites it.
                    FileUpload1.SaveAs(savePath);

                    // Notify the user that the file was uploaded successfully.
                    //UploadStatusLabel.Text = "Your file was uploaded successfully.";

                    // Call a helper routine to display the contents
                    // of the file to upload.
                    //DisplayFileContents(FileUpload1.PostedFile);
                    DataTable csvFileData = GetDataTabletFromCSVFile(csv_file_path);
                    InsertDataIntoSQLServerUsingSQLBulkCopy(csvFileData);
                    Response.Redirect("ProdutosTexteis");                                          
            }
            else
            {
                // Notify the user that a file was not uploaded.
                //UploadStatusLabel.Text = "Not a valid csv file to upload";
                Response.Write("<script>alert('" + "Not a valid csv file to import" + "')</script>");
            }
            //string csv_file_path = @"C:\tester site\Cópia de Lista Calçado Verão Senhora (P_site).xlsx - Inglês.csv";
        }
    }
}