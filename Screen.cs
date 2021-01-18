using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InterfaceСППР
{
    public partial class Screen : Form
    {
        public string type = "";
        public Screen()
        {
            InitializeComponent();
        }

        private void buttonOrder_Click(object sender, EventArgs e)
        {
            type = "order";
            Close();
        }

        private void buttonStat_Click(object sender, EventArgs e)
        {
            type = "stat";
            Close();
        }
    }
}
