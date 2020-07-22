using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Color = System.Drawing.Color;

namespace ClientsWPF
{
    public partial class Form1 : MaterialForm
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Clients2";
        static readonly string SpreadsheetId = "1zg-2dB6oeelWHYM7oHj-hfMg4BEDcNmtO4oaK8FFSZQ";
        static SheetsService service;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForInternetConnection();
            if (CheckForInternetConnection().Equals(true))
            {
                GoogleCredential credential;
                using (var stream = new FileStream("My First Project-ac5d44b52132.json", FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(Scopes);
                }

                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                ReadEntries();
            }
            else
                MessageBox.Show("Отсутствует подключение к интернету");
        }

        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }

        private void ReadEntries()
        {
            var range = "A2:K";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                //ListViewItem item = new ListViewItem();
                foreach (var row in values)
                {
                    dataGridView1.Rows.Add(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10]);
                }
            }
            else
            {
                MessageBox.Show("Ошибка чтения данных");
            }
        }

        private void findBtn_Click(object sender, EventArgs e)
        {
            string searchValue = searchTextBox.Text;
            bool valueResult = false;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                            if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(searchValue.ToLower()) && searchValue.Length != 0)
                            {
                                dataGridView1.Rows[i].Selected = true;
                                valueResult = true;
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                                break;
                            }
                }
                if (!valueResult)
                {
                    MessageBox.Show("Запись " + "'" + searchTextBox.Text + "'" + " не найдена", "Не найдено");
                    return;
                }
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int idc, idr;
            char sheetY = ' ';
            idc = Convert.ToInt32(dataGridView1.CurrentCell.ColumnIndex.ToString());
            idr = Convert.ToInt32(dataGridView1.CurrentCell.RowIndex.ToString()) + 2;
            string s = dataGridView1[idc, idr - 2].Value.ToString();
            switch (idc)
            {
                case 0:
                    sheetY = 'A';
                    break;
                case 1:
                    sheetY = 'B';
                    break;
                case 2:
                    sheetY = 'C';
                    break;
                case 3:
                    sheetY = 'D';
                    break;
                case 4:
                    sheetY = 'E';
                    break;
                case 5:
                    sheetY = 'F';
                    break;
                case 6:
                    sheetY = 'G';
                    break;
                case 7:
                    sheetY = 'H';
                    break;
                case 8:
                    sheetY = 'I';
                    break;
                case 9:
                    sheetY = 'J';
                    break;
                case 10:
                    sheetY = 'K';
                    break;
            }

            //MessageBox.Show(sheetY.ToString() + idr.ToString());
            var range = sheetY.ToString() + idr.ToString();
            var valueRange = new ValueRange();

            var oblist = new List<object>() { s };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            int idr;
            idr = Convert.ToInt32(dataGridView1.CurrentCell.RowIndex.ToString()) + 2;
            //MessageBox.Show($"A{ idr }:K{ idr }");
            var range = $"A{ idr }:K{ idr }";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var deleteReponse = deleteRequest.Execute();
        }

        private void btnAddZ_Click(object sender, EventArgs e)
        {
            if (tbClient.Text != null && tbPassport.Text != null && tbAddress.Text != null &&
                tbPhone.Text != null && cbProduct.SelectedItem != null && cbColor.SelectedItem != null &&
                tbProdCount.Text != null && tbPodCount.Text != null && tbZakaz.Text != null && tbDopInfo.Text != null)
            {
                var range = "A:K";
                var valueRange = new ValueRange();

                var oblist = new List<object>() { DateTime.Now.ToString("dd-MM-yy"), tbClient.Text.ToString()
                , tbPassport.Text.ToString(), tbAddress.Text.ToString(), tbPhone.Text.ToString()
                , cbProduct.SelectedItem.ToString(), cbColor.SelectedItem.ToString() + " " + cbType.SelectedItem.ToString()
                , tbProdCount.Text.ToString(), tbPodCount.Text.ToString(), tbZakaz.Text.ToString(), tbDopInfo.Text.ToString() };
                
                valueRange.Values = new List<IList<object>> { oblist };

                var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendReponse = appendRequest.Execute();

                tabControl1.SelectedIndex = 0;
                dataGridView1.Rows.Clear();
                ReadEntries();
            }
            else
                MessageBox.Show("Не все поля заполнены!");
        }

        private void tbPodCount_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void tbPodCount_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 44) // цифры, клавиша BackSpace и запятая
            {
                e.Handled = true;
            }
        }

        private void tbProdCount_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 44) // цифры, клавиша BackSpace и запятая
            {
                e.Handled = true;
            }
        }

        private void searchTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (searchTextBox != null)
                {
                    findBtn.PerformClick();
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8 && number != 44) // цифры, клавиша BackSpace и запятая
            {
                e.Handled = true;
            }
        }

        public void Stoimost()
        {
            if (textBox1.Text.Length != 0)
            {
                double price = Convert.ToDouble(textBox1.Text.Replace(',', '.'));
                double priceMinusNDS;
                double pricePlusNDS;
                double podCount = Convert.ToDouble(tbPodCount.Text.Replace(',', '.'));
                if (checkBox2.Checked == true)
                {
                    priceMinusNDS = Math.Round((price * Convert.ToDouble(tbProdCount.Text.Replace(',', '.'))) + (podCount * 8.75), 2);
                    pricePlusNDS = Math.Round(priceMinusNDS + (priceMinusNDS * 0.2), 2);
                    textBox3.Text = Convert.ToString(priceMinusNDS);
                    textBox2.Text = Convert.ToString(pricePlusNDS);
                    return;
                }
                else
                {
                    priceMinusNDS = Math.Round(price * Convert.ToDouble(tbProdCount.Text.Replace(',', '.')), 2);
                    pricePlusNDS = Math.Round(priceMinusNDS + (priceMinusNDS * 0.2), 2);
                    textBox3.Text = Convert.ToString(priceMinusNDS);
                    textBox2.Text = Convert.ToString(pricePlusNDS);
                    return;
                }
            }
            else
                return;
        }

        public void PodCount()
        {
            if (tbPodCount.Text.Length != 0)
            {
                int podCount = Convert.ToInt32(tbPodCount.Text.ToString());
                double prodCount;
                if (checkBox1.Checked == true)
                {
                    string product = cbProduct.SelectedItem.ToString();
                    if (product.Equals("СОЛПоРУГОб"))
                    {
                        prodCount = podCount * 288;
                        tbProdCount.Text = Convert.ToString(prodCount);
                        Stoimost();
                        return;
                    }

                    else if (product.Equals("СУЛПу(3)") || product.Equals("СУЛПу(14)") || product.Equals("СУРПу(14)"))
                    {
                        prodCount = podCount * 324.96;
                        tbProdCount.Text = Convert.ToString(prodCount);
                        Stoimost();
                        return;
                    }

                    else if (product.Equals("СОЛПо") || product.Equals("СОРПо"))
                    {
                        prodCount = podCount * 336;
                        tbProdCount.Text = Convert.ToString(prodCount);
                        Stoimost();
                        return;
                    }

                    else
                        MessageBox.Show("Заполните поле 'Продукция'!");
                }
                else
                    tbPodCount.Text.ToString();
            }
            else
                return;
        }

        private void cbProduct_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cbProduct.Text.Length > 0)
            {
                PodCount();
                Stoimost();
            }
        }

        private void tbPodCount_TextChanged_1(object sender, EventArgs e)
        {
            PodCount();
            Stoimost();
        }

        private void tbProdCount_TextChanged(object sender, EventArgs e)
        {
            PodCount();
            Stoimost();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            PodCount();
            Stoimost();
        }

        private void checkBox2_CheckStateChanged(object sender, EventArgs e)
        {
            PodCount();
            Stoimost();
        }
    }
}
