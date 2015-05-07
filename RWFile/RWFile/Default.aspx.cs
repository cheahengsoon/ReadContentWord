using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace RWFile
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            //Uploading file to server
            fuWordFile.SaveAs(Server.MapPath(fuWordFile.FileName));
            //Getting Filename
            object filename = Server.MapPath(fuWordFile.FileName);
            //Initiating Application Class
            Microsoft.Office.Interop.Word.Application AC = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;
            //Reating the word document and adding it to textbox
            try
            {
                doc = AC.Documents.Open(ref filename, ref missing, ref readOnly, 
                    ref missing, ref missing, ref missing, ref missing, ref missing, 
                    ref missing, ref missing, ref missing, ref isVisible);

                String strLine;
                bool bolEOF = false;

                doc.Characters[1].Select();

                int index = 0;
                do
                {
                    object unit = Word.WdUnits.wdLine;
                    object count = 1;
                  //  int countLine=6;
                    AC.Selection.MoveEnd(ref unit, ref count);

                    strLine = AC.Selection.Text;
                    TextBox1.Text += ++index + " - " + strLine + "\r\n"; //for our understanding

                   //Check Line Range
                    //if(index==countLine)
                    //{
                    //    TextBox2.Text += countLine + "-" + strLine;
                    //}
                    if(index>=11 && index <=18)
                    {
                        txtintro.Text += strLine;
                    }

                    if(index>=20 && index<=29)
                    {
                        txtbenefit.Text += strLine;
                    }
                    //if(index==33)
                    //{
                    //    txtcontenttitle.Text += strLine;
                    //}
                    if(index==32||index==33||index==46||index==47||index==64||index==65||index==76
                       ||index==77||index==88||index==99)
                    {
                        txtcontenttitle.Text += strLine;
                    }
                    object direction = Word.WdCollapseDirection.wdCollapseEnd;
                    AC.Selection.Collapse(ref direction);

                    if (AC.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                        bolEOF = true;



                } while (!bolEOF);
               // TextBox1.Text = doc.Content.Text;
                AC.Documents.Close();
            }
            catch (Exception ex)
            {
                TextBox1.Text = ex.ToString();
            }
        }

   

      

      
    }
}