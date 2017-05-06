using Novacode;
using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ConsoleApplication1
{
    class Program
    {

        public static void testc()
        {
            //  Ookii.Dialogs.Wpf.VistaFileDialog dlg = new Ookii.Dialogs.Wpf.VistaFileDialog();
            var dlg = new Ookii.Dialogs.Wpf.VistaOpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "JPEG|*.jpg";
            dlg.FilterIndex = 0;

            dlg.ShowDialog();


            var doc = DocX.Load(System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("ConsoleApplication1.Resources.SOR.docx"));
            var cell = doc.Tables[1].Rows[1].Cells[0];
            var para = cell.Paragraphs[0];
            para.InsertText(@"01all marbles must be finished 45 degrees");
            para.Append("2all marbles must be finished 45 degrees").Append("03all marbles must be finished 45 degrees").Append("all marbles must be finished 45 degrees");
            foreach (string path in dlg.FileNames)
            {
                int orient = 6;
                // var txtx = File.ReadAllLines(path.Replace("jpg", "txt")).ToList();

                //   orient = int.Parse(txtx[0].Trim());


                // para.AppendLine();


                Novacode.Image img = doc.AddImage(path);

                Picture pic1;

                pic1 = img.CreatePicture();
                pic1.Name = Path.GetFileName(path);

                var para1 = cell.InsertParagraph();
                para1.AppendPicture(pic1);
                para1.Alignment = Alignment.center;
                // pic1.SetPictureShape(BasicShapes.rtTriangle);
                try
                {


                    if (orient == 6)
                    {
                        //portrait image
                        pic1.Rotation = 90;
                    }
                    if (orient == 3)
                    {
                        pic1.Rotation = 180;
                    }

                    if (orient == 8)
                    {
                        pic1.Rotation = 270;
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Invalid Picture\r\n" + path);
                    continue;
                }
                //Landscape Image
                //orient = 4;
                if (orient == 3 || orient == 10)
                {
                    pic1.Height = (int)XUnit.FromInch(3.9).Point;
                    pic1.Width = (int)(XUnit.FromInch(6.95).Point);
                }


                //Portrait Image
                else
                {
                    //continue;

                    pic1.Height = (int)XUnit.FromInch(3).Point;
                    pic1.Width = (int)(XUnit.FromInch(5).Point);
                    


                }

                para.AppendLine();
                para1.Alignment = Alignment.center;

                var pr = cell.InsertParagraph().Append("[Image " + (dlg.FileNames.ToList().IndexOf(path) + 1).ToString("00") + "]");
                pr.Alignment = Alignment.center;

            }
            string filename = Path.GetTempFileName() + ".docx";
            if (!Directory.Exists(Path.GetDirectoryName(filename)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filename));
            }
            doc.SaveAs(filename);


            Process.Start(filename);

        }

        static void Main(string[] args)
        {
            testc();
        }
    }
}
