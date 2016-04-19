using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace AlertOutlookAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
        }


        private void Button1_Click(object sender, RibbonControlEventArgs e) 
        {

          Form1 cForm1 = new Form1();
          cForm1.Show();
        
        }




    }
}
