﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace VSTONew
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 fm = new Form1();
            fm.Show();

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
