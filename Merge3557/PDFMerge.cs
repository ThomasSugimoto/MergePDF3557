using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Merge3557
{
    public class PDFMerge
    {
        /// <summary>
        ///    This procedure merges the pdf files from the list passed in arguments into on pdf file.
        /// </summary>
        /// <param name="files"></param>
        /// <param name="output"></param>
        public static void Merge(string[] files, string output)
        {
            try
            {
                int f = 0;

                PdfReader reader = new PdfReader(files[f]);

                int n = reader.NumberOfPages;
                Document document = new Document(reader.GetPageSizeWithRotation(1));

                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(output, FileMode.Create));
                document.Open();

                PdfContentByte cb = writer.DirectContent;
                PdfImportedPage page;

                while (f < files.Length)
                {
                    int i = 0;
                    while (i < n)
                    {
                        i++;
                        document.SetPageSize(reader.GetPageSizeWithRotation(i));
                        document.NewPage();
                        page = writer.GetImportedPage(reader, i);
                        int rotation = reader.GetPageRotation(i);
                        if (rotation == 90 || rotation == 270)
                        {
                            cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                        }
                        else
                        {
                            cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                        }
                    }

                    f++;
                    if (f < files.Length)
                    {
                        reader = new PdfReader(files[f]);
                        n = reader.NumberOfPages;
                    }
                }
                document.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format(@"ERROR: {0} - creating merged pdf", ex.Message));
            }
        }
    }
}
