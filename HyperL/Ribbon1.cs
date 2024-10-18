using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace HyperL
{
    public partial class Ribbon1
    {
        PowerPoint.Application app;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.officeplus.cn/PPT/template/");
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.google.com/imghp");

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.bing.com/images");

        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://image.baidu.com");

        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.pinterest.com");

        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://dribbble.com");

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://huaban.com");

        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://unsplash.com");

        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://pixabay.com");

        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.pexels.com");

        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.flaticon.com");

        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.iconfont.cn");

        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://iconscout.com");

        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://pngtree.com");

        }

        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.stickpng.com");

        }

        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.51yuansu.com");

        }

        private void button17_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://chatgpt.com");

        }

        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://claude.ai");

        }

        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://yiyan.baidu.com");

        }

        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://jingbenzhengyi.com");

        }
    }
}
