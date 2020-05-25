using Microsoft.WindowsAPICodePack.Dialogs;
using MySql.Data.MySqlClient;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Security.AccessControl;

namespace WindowsFormsApp6
{
    public partial class Form1 : Form
    {
        private readonly object mis = Type.Missing;
        private readonly System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
        private readonly string pathToTXT_Categorii;
        private readonly string pathToTXT_Produse;
        private readonly string pathToTXT_Vinzari;
        private readonly DataGridView DW1 = new DataGridView();
        private readonly DataGridView DW2 = new DataGridView();
        private readonly DataGridView DW3 = new DataGridView();
        private readonly string pathToFolder = "";

        // DB Connection Variables
        string userid;
        string password;
        string con2;
        MySqlConnection con = null;
        string dbn;



        public Form1()
        {
            InitializeComponent();
            bool ok = false;


            while (!ok)
            {
                using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
                {
                    dialog.IsFolderPicker = true;
                    dialog.Title = "Выберите папку где находятся файлы: Categorii.txt, Produse.txt, Vinzari.txt";

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        string[] files = Directory.GetFiles(dialog.FileName);
                        int count = 0;
                        foreach (string item in files)
                        {
                            if (item.Contains("Categorii.txt") || item.Contains("Produse.txt") || item.Contains("Vinzari.txt"))
                            {
                                count++;
                            }
                        }
                        if (count == 3)
                        {
                            MessageBox.Show("Файлы найденны успешно!!! ", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ok = true;
                            pathToFolder = dialog.FileName;
                        }
                        else
                        {
                            DialogResult res = MessageBox.Show("Ошибка, файлы не найденны, попробовать еще?", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (res == DialogResult.Yes)
                            {
                                ok = false;
                            }
                            if (res == DialogResult.No)
                            {
                                MessageBox.Show("До свидания!!!");

                                Environment.Exit(1);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("До свидания!!!");
                        Environment.Exit(1);
                    }
                }
                pathToTXT_Categorii = $@"{pathToFolder}\Categorii.txt";
                pathToTXT_Produse = $@"{pathToFolder}\Produse.txt";
                pathToTXT_Vinzari = $@"{pathToFolder}\Vinzari.txt";

            }

        }




        private void readFromTXT_toListBox()
        {
            string line = "";
            listBox1.Items.Clear();
            //listBox2.Items.Clear();
            //listBox3.Items.Clear();

            System.IO.StreamReader fileCategorii = new System.IO.StreamReader(pathToTXT_Categorii);
            while ((line = fileCategorii.ReadLine()) != null)
            {
                listBox1.Items.Add(line);
            }
            fileCategorii.Close();

            System.IO.StreamReader fileProduse = new System.IO.StreamReader(pathToTXT_Produse);
            while ((line = fileProduse.ReadLine()) != null)
            {
                listBox2.Items.Add(line);
            }
            fileProduse.Close();

            System.IO.StreamReader fileVinzari = new System.IO.StreamReader(pathToTXT_Vinzari);
            while ((line = fileVinzari.ReadLine()) != null)
            {
                listBox3.Items.Add(line);
            }
            fileVinzari.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            readFromTXT_toListBox();
        }

        private bool createDataGridView(DataGridView DW, ListBox listBox, int x, int y)
        {
            DW.EnableHeadersVisualStyles = false;
            DW.ColumnHeadersDefaultCellStyle.BackColor = Color.AliceBlue;
            DW.ColumnHeadersHeight = 60;
            DW.Font = new System.Drawing.Font("Lobster", 12F);
            try
            {
                DW.ColumnCount = listBox.Items[0].ToString().Split('\t').Length;
                for (int i = 0; i < DW.ColumnCount; i++)
                {
                    DW.Columns[i].Name = listBox.Items[0].ToString().Split('\t')[i];
                }

                DW.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                DW.Location = new System.Drawing.Point(x, y);
                DW.Name = "Products";
                DW.Size = new System.Drawing.Size(400, 222);
                Controls.Add(DW);

                int nr = listBox.Items.Count;
                DW.Rows.Clear();
                string[] ss;
                for (int i = 0; i < nr; i++)
                {
                    if (i != 0)
                    {
                        ss = listBox.Items[i].ToString().Split('\t');
                        int nc = ss.Length;
                        DW.Rows.Add();
                        for (int j = 0; j < nc; j++)
                        {
                            DW.Rows[i - 1].Cells[j].Value = ss[j].ToString();
                        }
                    }
                }
                return true;
            }
            catch (System.ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!createDataGridView(DW1, listBox1, 15, 300))
            {
                MessageBox.Show(
                        "2. Нажмите кнопку ProductToListBox",
                        "Сообщение",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly
                    );
            }
            createDataGridView(DW2, listBox2, 420, 300);
            createDataGridView(DW3, listBox3, 825, 300);
        }

        private int gridToMySql(string table)
        {
            if(userid == null || password == null)
            {
                return 2;
            }

            // Connection
            string con2 = $"server='localhost';user id='{userid}'; password='{password}';";
            MySqlConnection con;
            try
            {
                con = new MySqlConnection(con2);
                con.Open();
            }
            catch (Exception)
            {
                return 1;
            }
            var cmd = new MySqlCommand();


            string myquery = "";
            string rows = "";
            string result = "";
            DataGridView DW = null;
            switch (table)
            {
                case "categorii":
                    rows = "cod, catn, descr";
                    DW = DW1;
                    break;
                case "vinzari":
                    rows = "PrId, Date1, QuantS";
                    DW = DW3;
                    break;
                case "produse":
                    rows = "PrId, PrN, Price, UnM, Cat";
                    DW = DW2;
                    break;
                default:
                    throw new Exception();
            }
            myquery = $"INSERT INTO {dbn}.{table} ({rows}) VALUES (";
            if (DW.Rows.Count <= 0)
                return 1;
            // DW1 -- categorii
            for (int i = 0; i < DW.Rows.Count - 1; i++)
            {
                myquery = $"INSERT INTO {dbn}.{table} ({rows}) VALUES (";

                for (int j = 0; j < DW.Columns.Count; j++)
                {
                    myquery += $"'{DW.Rows[i].Cells[j].Value.ToString()}',";
                }

                myquery = myquery.Remove(myquery.Length - 1);
                myquery += ");";

                // Insert line into DB
                cmd = new MySqlCommand(myquery, con);
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    result += "\n" + ex.ToString();
                    return 1;
                }

            }
            return 0;
        }

        private void createExcel(DataGridView DW1, DataGridView DW2, DataGridView DW3)
        {
            Excel.Application app = new Excel.Application();
            if (app == null)
            {
                MessageBox.Show("Excel не установлен на вашем устройстве", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            app.DisplayAlerts = false;
            Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Excel._Worksheet worksheet1 = null;
            app.Visible = true;
            try
            {
                worksheet1 = workbook.Sheets["Sheet1"];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                worksheet1 = workbook.Sheets["Лист1"];
            }
            Excel._Worksheet worksheet2 = null;
            Excel._Worksheet worksheet3 = null;
            worksheet2 = workbook.Sheets.Add(After: worksheet1);
            worksheet3 = workbook.Sheets.Add(After: worksheet2);
            worksheet1.Delete();
            worksheet1 = workbook.Sheets.Add(Before: worksheet2);
            worksheet1 = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet1.Name = "Categorii";
            worksheet2.Name = "Produse";
            worksheet3.Name = "Vinzari";

            //_________________ DW1 __________________
            // storing header part in Excel  
            for (int i = 1; i < DW1.Columns.Count + 1; i++)
            {
                worksheet1.Cells[1, i] = DW1.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < DW1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < DW1.Columns.Count; j++)
                {
                    worksheet1.Cells[i + 2, j + 1] = DW1.Rows[i].Cells[j].Value.ToString();
                }
            }

            //_________________ DW2 __________________
            for (int i = 1; i < DW2.Columns.Count + 1; i++)
            {
                worksheet2.Cells[1, i] = DW2.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < DW2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < DW2.Columns.Count; j++)
                {
                    worksheet2.Cells[i + 2, j + 1] = DW2.Rows[i].Cells[j].Value.ToString();
                }
            }

            //_________________ DW3 __________________
            for (int i = 1; i < DW3.Columns.Count + 1; i++)
            {
                worksheet3.Cells[1, i] = DW3.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < DW3.Rows.Count - 1; i++)
            {
                for (int j = 0; j < DW3.Columns.Count; j++)
                {
                    worksheet3.Cells[i + 2, j + 1] = DW3.Rows[i].Cells[j].Value.ToString();
                }
            }
            // Summ Formule 
            worksheet3.Cells[1,5] = "=SUM(C2:C200)";

            var cellValue = (double)(worksheet3.Cells[1, 5] as Excel.Range).Value;

            textBox1.Text = "Sum = " + cellValue.ToString();

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = "output",
                DefaultExt = ".xlsx"
            };


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            // Exit from the application  
            try
            {
                app.Quit();

            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show(
                        "Не удалось сохранить файл",
                        "Ошибка!",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly
                    );
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
             if(DW1.Rows.Count == 0)
            {
                MessageBox.Show(
                        "Нажмите кнопку 3. ListBoxToGrid",
                        "Сообщение",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly
                    );
                return;
            }

            createExcel(DW1, DW2, DW3);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // **************** создание базы данных MySql ****************
            // Connection
            bool ok = false;
            //string userid;
            //string password;
            //string con2;
            //MySqlConnection con = null;

            // Пока не удалось подключиться к mysql
            while (!ok)
            {
                userid = Interaction.InputBox("Введите user id пользователя MySql", "MySql User ID", "root");
                password = Interaction.InputBox("Введите password пользователя MySql", "MySql Password", "root");
                con2 = $"server='localhost';user id='{userid}'; password='{password}';";
                con = new MySqlConnection(con2);

                try
                {
                    con.Open();
                    ok = true;
                }
                catch
                {
                    MessageBox.Show($"Ошибка!!! Имя пользователя или пароль неверны", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ok = false;
                }

            }



            // Data Base Creation 
            
            dbn = Interaction.InputBox("Введите название базы данных: ", "Название БД", "individual_work");
            string sql = " CREATE DATABASE IF NOT EXISTS " + dbn + ";";
            var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            MessageBox.Show("data base " + dbn + " was succcesfully created!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information); 

            // Delete Tables if Exists
            string sql1 = "DROP TABLE if exists `" + dbn + "`.`categorii`;";
            var cmd1 = new MySql.Data.MySqlClient.MySqlCommand(sql1, con);
            cmd1.ExecuteNonQuery();

            sql1 = "DROP TABLE if exists `" + dbn + "`.`produse`;";
            cmd1 = new MySql.Data.MySqlClient.MySqlCommand(sql1, con);
            cmd1.ExecuteNonQuery();

            sql1 = "DROP TABLE if exists `" + dbn + "`.`vinzari`;";
            cmd1 = new MySql.Data.MySqlClient.MySqlCommand(sql1, con);
            cmd1.ExecuteNonQuery();

            // **************** создание таблиц Categorii, Produse, Vinzari ****************
            // categorii
            sql = "CREATE TABLE `" + dbn + "`.`categorii` ( " +
            " `cod` INT NULL, `catn` VARCHAR(45) NULL, `descr` VARCHAR(45) NULL) ; ";
            cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            //vinzari
            sql = "CREATE TABLE `" + dbn + "`.`vinzari` ( " +
           " `PrId` INT NULL, `Date1` DATE NULL, `QuantS` DOUBLE NULL) ; ";
            cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, con);
            cmd.ExecuteNonQuery();

            //produse
            sql = "CREATE TABLE `" + dbn + "`.`produse` ( " +
           " `PrId` INT NULL, `PrN` VARCHAR(20) NULL, `Price` DOUBLE NULL, `UnM` VARCHAR(20) NULL, `Cat` INT NULL) ; ";
            cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, con);
            cmd.ExecuteNonQuery();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(textBox1.TextLength == 0)
            {

                MessageBox.Show(
                        "Нажмите кнопку 4. Grid To Excel",
                        "Сообщение",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly
                    );
                return;
            }
            textBox1.Visible = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            bool err1 = false;
            bool err2 = false;

            string[] list = { "categorii", "vinzari", "produse" };
            foreach (string i in list)
            {
                switch (gridToMySql(i))
                {
                    case 0:
                        break;
                    case 1:
                        err1 = true;
                        break;
                    case 2:
                        err2 = true;
                        break;
                }
            }

            if (err2)
            {
                MessageBox.Show("Нажмите на кнопку 1. Create DB MySql!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (err1)
            {
                MessageBox.Show("Insert error!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("Insert Success!!!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
