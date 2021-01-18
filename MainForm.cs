using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace InterfaceСППР
{
    public partial class MainForm : Form
    {
        string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["CoffeeConnectionString"].ConnectionString;
        public MainForm()
        {
            InitializeComponent();
            //заставка
            Screen scr = new Screen();
            scr.ShowDialog();
            if (scr.type == "stat")
            {
                tabControl1.SelectedTab = tabPageStat;
            }

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();

                //заполнение checkbox с основными ингредиентами
                string sql = "SELECT ingr_num, ingr_name FROM ingredient where ingr_type=1 or ingr_type=2";
                SqlCommand command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    checkedListBoxIngrMain.Items.Add(reader[1]);
                }
                checkedListBoxIngrMain.SetItemChecked(checkedListBoxIngrMain.Items.IndexOf("Эспрессо"), true);
                reader.Close();

                //заполнение checkbox с дополнительными ингредиентами (добавками)
                sql = "SELECT ingr_num, ingr_name FROM ingredient where ingr_type=1 or ingr_type=0";
                command = new SqlCommand(sql, conn);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    checkedListBoxIngrAdd.Items.Add(reader[1]);
                }
                reader.Close();

                //установка min и max цены
                sql = "SELECT max(sizeUnit_price), min(sizeUnit_price) FROM coffee_size";
                command = new SqlCommand(sql, conn);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    numericUpDownPriceHigh.Value = Convert.ToInt32(reader[0]);
                    numericUpDownPriceHigh.Minimum = Convert.ToInt32(reader[1]);
                    numericUpDownPriceHigh.Maximum = Convert.ToInt32(reader[0]);
                    numericUpDownPriceLow.Value = Convert.ToInt32(reader[1]);
                    numericUpDownPriceLow.Minimum = Convert.ToInt32(reader[1]);
                    numericUpDownPriceLow.Maximum = Convert.ToInt32(reader[0]);
                }
                reader.Close();
            }

            //стилизация таблиц
            dataGridViewFound.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            dataGridViewFound.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridViewFound.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridViewFound.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridViewFound.BackgroundColor = Color.White;

            dataGridViewFound.EnableHeadersVisualStyles = false;
            dataGridViewFound.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dataGridViewFound.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(20, 25, 72);
            dataGridViewFound.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dataGridViewOrders.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            dataGridViewOrders.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridViewOrders.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dataGridViewOrders.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dataGridViewOrders.BackgroundColor = Color.White;

            dataGridViewOrders.EnableHeadersVisualStyles = false;
            dataGridViewOrders.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dataGridViewOrders.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(20, 25, 72);
            dataGridViewOrders.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            this.Width = 1027;
            this.Height = 613;

            //ограничение даты, установка формата
            book_dateDateTimePicker.CustomFormat = "dd/MM/yyyy HH:mm";
            book_dateDateTimePicker.MinDate = DateTime.Now;

            DateTime date2 = DateTime.Now.Date;
            string date1 = "01." + date2.Month + "." + date2.Year;
            labelReport.Text = "Отчет за период: " + date1 + " - " + date2.ToShortDateString();
        }

        private void ingredientBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.ingredientBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dataSet1);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataSet1.orderUnit". При необходимости она может быть перемещена или удалена.
            this.orderUnitTableAdapter.Fill(this.dataSet1.orderUnit);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataSet1.book". При необходимости она может быть перемещена или удалена.
            this.bookTableAdapter.Fill(this.dataSet1.book);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataSet1.delivery". При необходимости она может быть перемещена или удалена.
            this.deliveryTableAdapter.Fill(this.dataSet1.delivery);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataSet1.size". При необходимости она может быть перемещена или удалена.
            this.sizeTableAdapter.Fill(this.dataSet1.size);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "dataSet1.ingredient". При необходимости она может быть перемещена или удалена.
            this.ingredientTableAdapter.Fill(this.dataSet1.ingredient);
            ChartLoad();
            ChartPriceLoad();
            ChartTypeLoad();
        }

        private void buttonSize_Click(object sender, EventArgs e)
        {
            comboBoxSize.Visible = !comboBoxSize.Visible;
        }

        private void buttonIngr12_Click(object sender, EventArgs e)
        {
            checkedListBoxIngrMain.Visible = !checkedListBoxIngrMain.Visible;
        }

        private void buttonIngr01_Click(object sender, EventArgs e)
        {
            checkedListBoxIngrAdd.Visible = !checkedListBoxIngrAdd.Visible;
        }

        private void buttonPrice_Click(object sender, EventArgs e)
        {
            numericUpDownPriceLow.Visible = !numericUpDownPriceLow.Visible;
            numericUpDownPriceHigh.Visible = !numericUpDownPriceHigh.Visible;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            groupBoxBookDeliv.Visible = !groupBoxBookDeliv.Visible;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            flowLayoutPanelAddress.Visible = !flowLayoutPanelAddress.Visible;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            flowLayoutPanelBook.Visible = !flowLayoutPanelBook.Visible;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            int size = comboBoxSize.SelectedIndex + 1;
            decimal maxPrice = numericUpDownPriceHigh.Value;
            decimal minPrice = numericUpDownPriceLow.Value;
            CheckedListBox.CheckedItemCollection listMainIngr = checkedListBoxIngrMain.CheckedItems;
            List<string[]> data = new List<string[]>();
            List<int> resultCof = new List<int>();

            //строка выбранных добавок
            CheckedListBox.CheckedItemCollection listAddIngr = checkedListBoxIngrAdd.CheckedItems;
            string str = "";
            if (listAddIngr.Count > 0)
            {
                foreach (string ingr in listAddIngr)
                {
                    str += ingr + ", ";
                }
                str = str.Substring(0, str.Length - 2);
            }
            else
                str = "Без добавок";
            decimal add = AddsSum(listAddIngr);

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();

                //получение напитков из бд, удовлетворяющих всем условиям
                string sql = "select coffee.coffee_num FROM coffee INNER JOIN coffee_size ON coffee.coffee_num = coffee_size.coffee_num INNER JOIN size ON coffee_size.size_num = size.size_num where coffee_size.size_num=" + size + " and sizeUnit_price<=" + maxPrice + " and sizeUnit_price>=" + minPrice;
                SqlCommand command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                List<int> coffeeNums = new List<int>();
                int i = 0;  // по массиву основных ингредиентов
                while (reader.Read())
                {
                    coffeeNums.Add(Convert.ToInt32(reader[0]));
                }
                reader.Close();
                int k = 0; // для сравнения количества ингредиентов с выбранными
                for (int j = 0; j < coffeeNums.Count; j++)
                {
                    sql = "select ingr_name from ingredient inner join coffeeUnit on ingredient.ingr_num=coffeeUnit.ingr_num where coffee_num=" + coffeeNums[j];
                    command = new SqlCommand(sql, conn);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        if (i < listMainIngr.Count && reader[0].ToString() == listMainIngr[i].ToString())
                        {
                            k++;
                            i++;
                        }
                    }
                    reader.Close();
                    if (k == listMainIngr.Count)
                        resultCof.Add(coffeeNums[j]);
                    k = 0;
                    i = 0;
                }
                sql = "select coffee.coffee_num, coffee_name, size_name, sizeUnit_price FROM coffee INNER JOIN coffee_size ON coffee.coffee_num = coffee_size.coffee_num INNER JOIN size ON coffee_size.size_num = size.size_num where coffee_size.size_num=" + size + " and sizeUnit_price<=" + maxPrice + " and sizeUnit_price>=" + minPrice;
                command = new SqlCommand(sql, conn);
                reader = command.ExecuteReader();
                k = 0;  // по массиву итоговых кофейных напитков
                while (reader.Read())
                {
                    if (k < resultCof.Count && Convert.ToInt32(reader[0]) == resultCof[k])
                    {
                        data.Add(new string[6]);
                        data[data.Count - 1][0] = reader[1].ToString();
                        data[data.Count - 1][1] = reader[2].ToString();
                        data[data.Count - 1][2] = reader[3].ToString();
                        data[data.Count - 1][3] = str;
                        data[data.Count - 1][4] = AddsSum(listAddIngr).ToString();
                        data[data.Count - 1][5] = "1";
                        k++;
                    }
                }
                reader.Close();
                conn.Close();
            }
            if (data.Count == 0)
            {
                MessageBox.Show("Не найдено!");
                return;
            }

            dataGridViewFound.Rows.Clear();
            foreach (string[] s in data)
                dataGridViewFound.Rows.Add(s);
            dataGridViewFound.Rows[0].Selected = true;
            this.dataGridViewFound.Sort(this.dataGridViewFound.Columns[2], ListSortDirection.Ascending);
        }

        //проверка корректности цен
        private void numericUpDownPriceLow_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownPriceLow.Value > numericUpDownPriceHigh.Value)
            {
                numericUpDownPriceLow.Value = numericUpDownPriceHigh.Value;
            }
        }

        private void numericUpDownPriceHigh_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownPriceLow.Value > numericUpDownPriceHigh.Value)
            {
                numericUpDownPriceHigh.Value = numericUpDownPriceLow.Value;
            }
        }

        //метод, подсчитывающий сумму добавок
        public decimal AddsSum(CheckedListBox.CheckedItemCollection listAddIngr)
        {
            decimal add = 0;
            if (listAddIngr.Count > 0)
            {
                string s = "";
                foreach (string ingr in listAddIngr)
                {
                    s += "'" + ingr + "',";
                }
                s = s.Substring(0, s.Length - 1);
                string connStr = System.Configuration.ConfigurationManager.ConnectionStrings["CoffeeConnectionString"].ConnectionString;
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string sql = "select ingr_price from ingredient where ingr_name in(" + s + ")";
                    SqlCommand command = new SqlCommand(sql, conn);
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        add += Convert.ToDecimal(reader[0]);
                    }
                    reader.Close();
                    conn.Close();
                }
            }
            return add;
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (dataGridViewFound.SelectedCells.Count == 0)
            {
                MessageBox.Show("Не выделены напитки!");
                return;
            }
            decimal sum = 0, add = 0;
            //заполнение таблицы заказов
            foreach (DataGridViewRow r in dataGridViewFound.SelectedRows)
            {
                bool fl = false; //отсутствует строка
                foreach (DataGridViewRow ro in dataGridViewOrders.Rows)
                {
                    if (ro.Cells[0].Value == r.Cells[0].Value &&
                        ro.Cells[1].Value == r.Cells[1].Value &&
                        ro.Cells[3].Value == r.Cells[3].Value)
                    {
                        int oldK = Convert.ToInt32(ro.Cells[5].Value),
                            newK = Convert.ToInt32(r.Cells[5].Value);
                        ro.Cells[5].Value = oldK + newK;
                        sum += Convert.ToDecimal(ro.Cells[2].Value) * (Convert.ToInt32(ro.Cells[5].Value) - oldK);
                        add += Convert.ToDecimal(ro.Cells[4].Value) * (Convert.ToInt32(ro.Cells[5].Value) - oldK);
                        fl = true;
                        break;
                    }
                }
                if (fl) break;
                int index = dataGridViewOrders.Rows.Add(r.Clone() as DataGridViewRow);
                decimal oneSum = 0, oneAdd = 0;
                foreach (DataGridViewCell o in r.Cells)
                {
                    dataGridViewOrders.Rows[index].Cells[o.ColumnIndex].Value = o.Value;
                    if (o.ColumnIndex == 2)
                    {
                        oneSum += Convert.ToDecimal(o.Value); //mainsum
                    }
                    if (o.ColumnIndex == 4)
                    {
                        oneAdd += Convert.ToDecimal(o.Value); //addsum
                    }
                    if (o.ColumnIndex == 5)
                    {
                        sum += oneSum * Convert.ToInt32(o.Value);
                        add += oneAdd * Convert.ToInt32(o.Value);
                    }
                }
            }

            //подсчет суммы
            decimal mainSum;
            Decimal.TryParse(labelMainSum.Text, out mainSum);
            labelMainSum.Text = (mainSum + sum).ToString();

            decimal addSum;
            Decimal.TryParse(labelAddSum.Text, out addSum);
            labelAddSum.Text = (addSum + add).ToString();
            labelSum.Text = (mainSum + addSum + sum + add).ToString();
        }

        private void buttonOrder_Click(object sender, EventArgs e)
        {
            if (dataGridViewOrders.Rows.Count == 0)
            {
                MessageBox.Show("Отсутствует заказ!");
                return;
            }

            int order_num = 0;
            string sql;
            SqlCommand command1;
            SqlDataReader reader;
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                sql = "SELECT TOP 1 * FROM orders ORDER BY order_num DESC";
                command1 = new SqlCommand(sql, conn);
                reader = command1.ExecuteReader();
                while (reader.Read())
                {
                    order_num = Convert.ToInt32(reader[0]);
                }
                reader.Close();
                Dictionary<string, int> orderedCoffeeDict = new Dictionary<string, int>();
                foreach (DataGridViewRow row in dataGridViewOrders.Rows)
                {
                    orderedCoffeeDict[row.Cells[0].Value.ToString()] = Convert.ToInt32(row.Cells[5].Value);
                }
                foreach (string i in orderedCoffeeDict.Keys)
                {
                    sql = "SELECT coffee_num FROM coffee where coffee_name='" + i + "'";
                    command1 = new SqlCommand(sql, conn);
                    reader = command1.ExecuteReader();
                    int coffee_num = 0;
                    while (reader.Read())
                    {
                        coffee_num = Convert.ToInt32(reader[0]);
                    }
                    reader.Close();

                    int orderUnit_kol = 0;

                    sql = "SELECT orderUnit_kol FROM orderUnit where coffee_num=" + coffee_num;
                    command1 = new SqlCommand(sql, conn);
                    reader = command1.ExecuteReader();
                    while (reader.Read())
                    {
                        orderUnit_kol = Convert.ToInt32(reader[0]);
                    }
                    reader.Close();

                    int kol = orderedCoffeeDict[i];

                    if (orderUnit_kol != 0)
                    {
                        string str = string.Format("UPDATE [orderUnit] SET [orderUnit_kol] = {0} WHERE [coffee_num] = {1}", orderUnit_kol + kol, coffee_num);
                        using (SqlCommand command = new SqlCommand(str, conn))
                        {
                            command.Parameters.Add(new SqlParameter("coffee_num", coffee_num));
                            command.Parameters.Add(new SqlParameter("orderUnit_kol", orderUnit_kol));
                            command.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        using (SqlCommand command = new SqlCommand(
                   "insert into orderUnit VALUES(@coffee_num, @orderUnit_kol)", conn))
                        {
                            command.Parameters.Add(new SqlParameter("coffee_num", coffee_num));
                            command.Parameters.Add(new SqlParameter("orderUnit_kol", kol));
                            command.ExecuteNonQuery();
                        }
                    }
                }
                conn.Close();
            }

            //запись в бд "Бронь" или "Доставка" в зависимости от выбранного условия
            if (!radioButtonDeliv.Checked)
            {
                if (book_tableTextBox.Text != "" && book_nameTextBox.Text != "" && deliv_phoneMaskedTextBox.MaskCompleted)
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        DateTime dt = book_dateDateTimePicker.Value;
                        string sql1 = "SELECT order_num FROM book where datepart(yy, book_date)='" + dt.Year + "' and datepart(m, book_date)='" + dt.Month + "' and datepart(d, book_date)='" + dt.Day + "' and datepart(hh, book_date)='" + dt.Hour + "' and datepart(mi, book_date)='" + dt.Minute + "'  and book_table=" + Convert.ToInt32(book_tableTextBox.Text);
                        SqlCommand command2 = new SqlCommand(sql1, conn);
                        reader = command2.ExecuteReader();
                        while (reader.Read())
                        {
                            MessageBox.Show("Невозможно забронировать столик на данное время!\nИзмените время или номер столика.");
                            return;
                        }
                        reader.Close();

                        int order_type = 1;
                        if (radioButtonBuy.Checked) order_type = 3;
                        using (SqlCommand command = new SqlCommand(
                        "INSERT INTO orders VALUES(@order_date, @order_sum, @order_type)", conn))
                        {
                            command.Parameters.Add(new SqlParameter("order_date", DateTime.Today.Date));
                            command.Parameters.Add(new SqlParameter("order_sum", Convert.ToDecimal(labelSum.Text)));
                            command.Parameters.Add(new SqlParameter("order_type", order_type));
                            command.ExecuteNonQuery();
                        }

                        using (SqlCommand command = new SqlCommand(
                            "INSERT INTO book values(@order_num, @book_date, @book_count, @book_table, @book_name, @book_phone, @book_comment)", conn))
                        {
                            command.Parameters.Add(new SqlParameter("order_num", order_num));
                            command.Parameters.Add(new SqlParameter("book_date", book_dateDateTimePicker.Value));
                            command.Parameters.Add(new SqlParameter("book_count", Convert.ToInt32(book_countNumericUpDown.Value)));
                            command.Parameters.Add(new SqlParameter("book_table", Convert.ToInt32(book_tableTextBox.Text)));
                            command.Parameters.Add(new SqlParameter("book_name", book_nameTextBox.Text));
                            command.Parameters.Add(new SqlParameter("book_phone", Convert.ToDecimal(deliv_phoneMaskedTextBox.Text)));
                            command.Parameters.Add(new SqlParameter("book_comment", deliv_commentTextBox.Text));
                            command.ExecuteNonQuery();
                        }
                        MessageBox.Show("Заказ #" + order_num.ToString() + " принят!");
                    }
                }
                else
                {
                    MessageBox.Show("Заполнены не все поля!");
                    return;
                }
            }
            if (radioButtonDeliv.Checked)
            {
                if (deliv_addressTextBox.Text != "" && deliv_phoneMaskedTextBox.MaskCompleted)
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();

                        using (SqlCommand command = new SqlCommand(
                            "INSERT INTO orders VALUES(@order_date, @order_sum, @order_type)", conn))
                        {
                            command.Parameters.Add(new SqlParameter("order_date", DateTime.Today.Date));
                            command.Parameters.Add(new SqlParameter("order_sum", Convert.ToDecimal(labelSum.Text)));
                            command.Parameters.Add(new SqlParameter("order_type", 2));
                            command.ExecuteNonQuery();
                        }

                        using (SqlCommand command = new SqlCommand(
                            "INSERT INTO delivery values(@order_num, @deliv_address, @deliv_phone, @deliv_comment)", conn))
                        {
                            command.Parameters.Add(new SqlParameter("order_num", order_num));
                            command.Parameters.Add(new SqlParameter("deliv_address", deliv_addressTextBox.Text));
                            command.Parameters.Add(new SqlParameter("deliv_phone", Convert.ToDecimal(deliv_phoneMaskedTextBox.Text)));
                            command.Parameters.Add(new SqlParameter("deliv_comment", deliv_commentTextBox.Text));
                            command.ExecuteNonQuery();
                        }
                        MessageBox.Show("Заказ #" + order_num.ToString() + " принят!");
                        ClearSum();
                        buttonClearTextBox_Click(sender, e);
                    }
                }
                else
                {
                    MessageBox.Show("Заполнены не все поля!");
                }
            }
        }

        private void ClearSum()
        {
            dataGridViewOrders.Rows.Clear();
            labelMainSum.Text = "0,00";
            labelAddSum.Text = "0,00";
            labelSum.Text = "0,00";
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Очистить все?", "Удаление", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                ClearSum();
            }
            else if (dialogResult == DialogResult.No)
            {
                if (dataGridViewOrders.Rows.Count > 0)
                {
                    foreach (DataGridViewRow r in dataGridViewOrders.SelectedRows) {
                        int ind = r.Index; //номер строки
                        decimal sum = Convert.ToDecimal(dataGridViewOrders.Rows[ind].Cells[2].Value);
                        decimal add = Convert.ToDecimal(dataGridViewOrders.Rows[ind].Cells[4].Value);

                        labelMainSum.Text = (Convert.ToDecimal(labelMainSum.Text) - sum).ToString();
                        labelAddSum.Text = (Convert.ToDecimal(labelAddSum.Text) - add).ToString();
                        decimal mainSum;
                        Decimal.TryParse(labelMainSum.Text, out mainSum);
                        decimal addSum;
                        Decimal.TryParse(labelAddSum.Text, out addSum);
                        labelSum.Text = (mainSum + addSum).ToString();
                        dataGridViewOrders.Rows.RemoveAt(ind);
                    }
                }
            }
        }

        private void buttonClearTextBox_Click(object sender, EventArgs e)
        {
            book_countNumericUpDown.Value = 1;
            book_tableTextBox.Text = "";
            book_nameTextBox.Text = "";
            deliv_addressTextBox.Text = "";
            deliv_phoneMaskedTextBox.Text = "";
            deliv_commentTextBox.Text = "";
        }

        private void buttonUpdateChart_Click(object sender, EventArgs e)
        {
            ChartLoad();
            ChartPriceLoad();
            ChartTypeLoad();
        }

        //методы для заполнения диаграмм данными
        public void ChartLoad()
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                string sql = "SELECT coffee_name, orderUnit_kol FROM coffee c inner join orderUnit o on c.coffee_num = o.coffee_num where o.coffee_num in (select coffee_num from orderUnit)";
                SqlCommand command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                chartCoffee.Series["Кофейные напитки"].Points.Clear();
                while (reader.Read())
                {
                    chartCoffee.Series["Кофейные напитки"].Points.AddXY(reader[0].ToString(), reader[1].ToString());
                }
                reader.Close();
            }
        }

        public void ChartPriceLoad()
        {
            decimal sum = 0;
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                string sql = "SELECT datepart(d, order_date), datepart(m, order_date), sum(order_sum) from orders where datepart(m, order_date) >=" + DateTime.Now.Month + "group by order_date";
                SqlCommand command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                chartPrice.Series["Цена"].Points.Clear();
                while (reader.Read())
                {
                    string str = reader[0].ToString() + "." + reader[1].ToString();
                    chartPrice.Series["Цена"].Points.AddXY(str, Convert.ToInt32(reader[2]).ToString());
                    sum += Convert.ToDecimal(reader[2]);
                }
                reader.Close();
            }
            chartPrice.ChartAreas[0].AxisY.MajorGrid.Interval = 1000;
            chartPrice.Series[0].BorderWidth = 2;
            labelItog.Text = sum.ToString();
        }

        public void ChartTypeLoad()
        {
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                string sql = "select count(order_type) from orders";
                SqlCommand command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                double all = 0, proc = 0;
                while (reader.Read())
                {
                    all = Convert.ToInt32(reader[0]);
                }
                reader.Close();
                sql = "select count(order_type) from orders where order_type=1";
                command = new SqlCommand(sql, conn);
                reader = command.ExecuteReader();
                chartType.Series["Тип заказа"].Points.Clear();
                while (reader.Read())
                {
                    proc = Math.Round(Convert.ToInt32(reader[0]) * 100 / all, 1);
                    chartType.Series["Тип заказа"].Points.AddXY("Бронь столика (" + proc + "%)", reader[0].ToString());
                }
                reader.Close();
                sql = "select count(order_type) from orders where order_type=2";
                command = new SqlCommand(sql, conn);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    proc = Math.Round(Convert.ToInt32(reader[0]) * 100 / all, 1);
                    chartType.Series["Тип заказа"].Points.AddXY("Доставка (" + proc + "%)", reader[0].ToString());
                }
                reader.Close();
                sql = "select count(order_type) from orders where order_type=3";
                command = new SqlCommand(sql, conn);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    proc = Math.Round(Convert.ToInt32(reader[0]) * 100 / all, 1);
                    chartType.Series["Тип заказа"].Points.AddXY("Покупка (" + proc + "%)", reader[0].ToString());
                }
                reader.Close();
            }
        }

        private void dataGridViewOrders_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int ind = e.RowIndex,
                kol = Convert.ToInt32(dataGridViewOrders.Rows[ind].Cells[e.ColumnIndex].Value);
            if (kol <= 0)
            {
                dataGridViewOrders.Rows[ind].Cells[e.ColumnIndex].Value = oldKol;
                return;
            }

            decimal mainIngr = Convert.ToDecimal(dataGridViewOrders.Rows[ind].Cells[2].Value),
                addIngr = Convert.ToDecimal(dataGridViewOrders.Rows[ind].Cells[4].Value);
            kol -= oldKol;
            decimal mainSum;
            Decimal.TryParse(labelMainSum.Text, out mainSum);
            decimal addSum;
            Decimal.TryParse(labelAddSum.Text, out addSum);
            mainSum = mainSum + mainIngr * kol;
            addSum = addSum + addIngr * kol;
            labelMainSum.Text = (mainSum).ToString();
            labelAddSum.Text = (addSum).ToString();
            labelSum.Text = (mainSum + addSum).ToString();
        }

        int oldKol = 1;
        private void dataGridViewOrders_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            oldKol = Convert.ToInt32(dataGridViewOrders[e.ColumnIndex, e.RowIndex].Value);
        }

        private void dataGridViewFound_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            oldKol = Convert.ToInt32(dataGridViewFound[e.ColumnIndex, e.RowIndex].Value);
        }

        private void dataGridViewFound_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int ind = e.RowIndex,
                kol = Convert.ToInt32(dataGridViewFound.Rows[ind].Cells[e.ColumnIndex].Value);
            if (kol <= 0)
            {
                dataGridViewFound.Rows[ind].Cells[e.ColumnIndex].Value = oldKol;
                return;
            }
        }

        //метода для записи данных в файл с расширением xls
        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet sheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            string date = DateTime.Today.Month + "." + DateTime.Today.Year;
            sheet.Name = "Отчет за период";
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                string sql = "SELECT coffee_name, orderUnit_kol FROM coffee c inner join orderUnit o on c.coffee_num = o.coffee_num where o.coffee_num in (select coffee_num from orderUnit) order by orderUnit_kol desc";
                SqlCommand command = new SqlCommand(sql, conn);
                SqlDataReader reader = command.ExecuteReader();
                sheet.Cells[1, 2] = "Отчет за период: 1." + date + "-" + "11" + "." + date;
                sheet.Cells[2, 1] = "№";
                sheet.Cells[2, 2] = "Название";
                sheet.Cells[2, 3] = "Количество";
                int i = 3;
                while (reader.Read())
                {
                    sheet.Cells[i, 1] = (i - 1).ToString();
                    sheet.Cells[i, 2] = reader[0].ToString();
                    sheet.Cells[i, 3] = reader[1].ToString();
                    i++;
                }
                i--;
                reader.Close();
                sheet.Columns[2].ColumnWidth = 20;
                sheet.Columns[3].ColumnWidth = 10;
                sheet.Cells[i + 2, 2] = "Общая выручка:";
                sheet.Cells[i + 2, 3] = labelItog.Text;

                Excel.Range range1 = sheet.Range[sheet.Cells[1, 1], sheet.Cells[i + 2, 3]];
                range1.Cells.Font.Name = "Times New Roman";
                range1.Cells.Font.Size = 10;
                Excel.Range range2 = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]];
                range2.Cells.Font.Bold = true;
                Excel.Range range3 = sheet.Range[sheet.Cells[1, 1], sheet.Cells[i, 1]];
                range3.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range3 = sheet.Range[sheet.Cells[1, 3], sheet.Cells[i, 3]];
                range3.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                Excel.Range range4 = sheet.Range[sheet.Cells[1, 2], sheet.Cells[1, 3]];
                range4.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                AllBorders(range1.Borders);
                for (int j = 2; j <= i; j++)
                {
                    if (j % 2 == 1)
                    {
                        for (int k = 1; k <= 3; k++)
                        {
                            sheet.Cells[j, k].Interior.Color = Color.AliceBlue;
                        }
                    }
                }

                string filename = "Отчет. " + DateTime.Today.Day + "." + date + ".xls";
                try
                {

                    xlWorkBook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                catch
                {
                    MessageBox.Show("Не удалось сохранить файл!");
                    return;
                }
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(sheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Данные успешно выгружены!");

                string filePath = @"C:\Users\Aigul\Document\" + filename;
                FileInfo fi = new FileInfo(filePath);
                if (fi.Exists)
                {
                    System.Diagnostics.Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("Файл не найден!");
                }
            }
        }

        //стилизация к виду таблицы
        private void AllBorders(Excel.Borders _borders)
        {
            _borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders.Color = Color.Black;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Завершить программу?", "Выход", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void поБазеДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"..\..\help\БД.pdf");
        }

        private void поБазеМоделиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"..\..\help\БМ.pdf");
        }
    }
}