using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.IO;

namespace DAI
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        private OleDbConnection connection;
        private const string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ДДМА\3 Курс\Курсовая ОБДЗ\DAI\DAI\DAI.accdb";
        public Form1()
        {
            InitializeComponent();
            connection = new OleDbConnection(connectionString);
            FillInspectorComboBox();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
        }
        public void LoadData()
        {
            try
            {
                using (OleDbDataAdapter lostItemsAdapter = new OleDbDataAdapter("SELECT * FROM [Подія]", connection))
                {
                    DataTable lostItemsTable = new DataTable();
                    lostItemsAdapter.Fill(lostItemsTable);
                    dataGridView1.DataSource = lostItemsTable;
                }

                using (OleDbDataAdapter documentsAdapter = new OleDbDataAdapter("SELECT * FROM [Відомості про учасників ДТП]", connection))
                {
                    DataTable documentsTable = new DataTable();
                    documentsAdapter.Fill(documentsTable);
                    dataGridView2.DataSource = documentsTable;
                }

                using (OleDbDataAdapter clientsAdapter = new OleDbDataAdapter("SELECT * FROM [Інформація про інспекторів]", connection))
                {
                    DataTable clientsTable = new DataTable();
                    clientsAdapter.Fill(clientsTable);
                    dataGridView3.DataSource = clientsTable;
                }
                // Встановлення автоматичного розширення стовпців для всіх DataGridView
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(this);
            form2.ShowDialog();
        }
        private void FillInspectorComboBox()
        {
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("SELECT [ID інспектора] FROM [Інформація про інспекторів]", connection);
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox3.Items.Add(reader["ID інспектора"].ToString());
                    comboBox2.Items.Add(reader["ID інспектора"].ToString());
                }
                reader.Close();
                // Встановлення ValueMember
                comboBox2.ValueMember = "ID інспектора що вносив дані";
                comboBox3.ValueMember = "ID інспектора що вносив дані";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Створюємо список фільтрів
            List<string> filters = new List<string>();

            // Перевіряємо кожне текстове поле та комбіноване поле
            if (!string.IsNullOrEmpty(textBox1.Text.Trim()))
                filters.Add($"[ID запису] = {textBox1.Text.Trim()}");
            if (!string.IsNullOrEmpty(textBox17.Text.Trim()))
                filters.Add($"[ID ДТП] = {textBox17.Text.Trim()}");
            if (comboBox1.SelectedItem != null && !string.IsNullOrEmpty(comboBox1.SelectedItem.ToString()))
                filters.Add($"[Тип події] LIKE '*{comboBox1.SelectedItem.ToString()}*'");
            if (comboBox2.SelectedItem != null && !string.IsNullOrEmpty(comboBox2.SelectedItem.ToString()))
                filters.Add($"[ID інспектора що вносив дані] = '{comboBox2.SelectedItem.ToString()}'");
            if (!string.IsNullOrEmpty(textBox2.Text.Trim()))
                filters.Add($"[Дата події] = '{textBox2.Text.Trim()}'");
            if (!string.IsNullOrEmpty(textBox3.Text.Trim()))
                filters.Add($"[Локація події] LIKE '*{textBox3.Text.Trim()}*'");
            if (!string.IsNullOrEmpty(textBox5.Text.Trim()))
                filters.Add($"[Причина аварії] LIKE '*{textBox5.Text.Trim()}*'");

            // Об'єднуємо всі фільтри за допомогою оператора AND
            string filterExpression = string.Join(" AND ", filters);

            // Виконуємо фільтрацію та відображення результату у dataGridView1
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = filterExpression;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Отримання значення ID запису для редагування
                int recordId = string.IsNullOrEmpty(textBox1.Text) ? -1 : int.Parse(textBox1.Text);

                // Перевірка, чи вказано значення ID запису для редагування
                if (recordId == -1)
                {
                    MessageBox.Show("Вкажіть ID запису для редагування");
                    return;
                }

                // Формування SQL-запиту для редагування інформації
                string updateQuery = "UPDATE [Подія] SET ";
                List<string> updateFields = new List<string>();

                if (!string.IsNullOrEmpty(textBox17.Text))
                    updateFields.Add($"[ID ДТП] = {textBox17.Text}");
                if (comboBox1.SelectedItem != null && !string.IsNullOrEmpty(comboBox1.SelectedItem.ToString()))
                    updateFields.Add($"[Тип події] = '{comboBox1.SelectedItem.ToString()}'");
                if (!string.IsNullOrEmpty(textBox2.Text))
                    updateFields.Add($"[Дата події] = '{textBox2.Text}'");
                if (!string.IsNullOrEmpty(textBox3.Text))
                    updateFields.Add($"[Локація події] = '{textBox3.Text}'");
                if (!string.IsNullOrEmpty(textBox5.Text))
                    updateFields.Add($"[Причина аварії] = '{textBox5.Text}'");

                // Складання полів для оновлення
                updateQuery += string.Join(", ", updateFields);

                // Додавання WHERE умови для визначення запису, який треба відредагувати
                updateQuery += $" WHERE [ID запису] = {recordId}";

                // Виконання SQL-запиту на оновлення інформації
                OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection);
                updateCommand.ExecuteNonQuery();
                LoadData();
                MessageBox.Show("Інформація успішно відредагована");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
                try
                {
                    connection.Open();

                    // Отримання значення ID запису для видалення
                    int recordId = string.IsNullOrEmpty(textBox1.Text) ? -1 : int.Parse(textBox1.Text);

                    // Перевірка, чи вказано значення ID запису для видалення
                    if (recordId == -1)
                    {
                        MessageBox.Show("Вкажіть ID запису для видалення");
                        return;
                    }

                    // Отримання ID ДТП для даного запису з таблиці "Подія"
                    string getIdDtpQuery = $"SELECT [ID ДТП] FROM [Подія] WHERE [ID запису] = {recordId}";
                    OleDbCommand getIdDtpCommand = new OleDbCommand(getIdDtpQuery, connection);
                    object result = getIdDtpCommand.ExecuteScalar();

                    // Перевірка, чи існує запис з таким ID
                    if (result == null)
                    {
                        MessageBox.Show("Запис з таким ID не знайдено");
                        return;
                    }

                    int idDtp = (int)result;

                    // Формування SQL-запиту для видалення записів з таблиці "Подія" з тим самим ID ДТП
                    string deleteFromPodiaQuery = $"DELETE FROM [Подія] WHERE [ID ДТП] = {idDtp}";
                    // Логування запиту
                    Console.WriteLine(deleteFromPodiaQuery);

                    // Виконання SQL-запиту на видалення записів з таблиці "Подія"
                    OleDbCommand deleteFromPodiaCommand = new OleDbCommand(deleteFromPodiaQuery, connection);
                    deleteFromPodiaCommand.ExecuteNonQuery();

                    // Формування SQL-запиту для видалення записів з таблиці "Відомості про учасників ДТП" з тим самим ID ДТП
                    string deleteFromUchasnykyQuery = $"DELETE FROM [Відомості про учасників ДТП] WHERE [ID ДТП] = {idDtp}";
                    // Логування запиту
                    Console.WriteLine(deleteFromUchasnykyQuery);

                    // Виконання SQL-запиту на видалення записів з таблиці "Відомості про учасників ДТП"
                    OleDbCommand deleteFromUchasnykyCommand = new OleDbCommand(deleteFromUchasnykyQuery, connection);
                    deleteFromUchasnykyCommand.ExecuteNonQuery();

                    LoadData();
                    MessageBox.Show("Записи успішно видалені");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Помилка: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                }
            }
        //////////////////////////////////////////////////////////////////////////////////////
        private void button6_Click(object sender, EventArgs e)
        {
            // Створюємо список фільтрів
            List<string> filters = new List<string>();

            // Перевіряємо кожне текстове поле та комбіноване поле
            if (!string.IsNullOrEmpty(textBox10.Text.Trim()))
                filters.Add($"[ID запису] = {textBox10.Text.Trim()}");
            if (!string.IsNullOrEmpty(textBox8.Text.Trim()))
                filters.Add($"[Прізвище] LIKE '*{textBox8.Text.Trim()}*'");
            if (comboBox1.SelectedItem != null && !string.IsNullOrEmpty(comboBox1.SelectedItem.ToString()))
                filters.Add($"[Тип події] LIKE '*{comboBox1.SelectedItem.ToString()}*'");
            if (!string.IsNullOrEmpty(textBox11.Text.Trim()))
                filters.Add($"[Ім'я] LIKE '*{textBox11.Text.Trim()}*'");
            if (!string.IsNullOrEmpty(textBox6.Text.Trim()))
                filters.Add($"[По батькові] LIKE '*{textBox6.Text.Trim()}*'");
            if (!string.IsNullOrEmpty(textBox7.Text.Trim()))
                filters.Add($"[Вік] = {textBox7.Text.Trim()}");
            if (!string.IsNullOrEmpty(textBox12.Text.Trim()))
                filters.Add($"[Номерний знак] LIKE '*{textBox12.Text.Trim()}*'");
            if (comboBox3.SelectedItem != null && !string.IsNullOrEmpty(comboBox3.SelectedItem.ToString()))
                filters.Add($"[ID інспектора що вносив дані] = '{comboBox3.SelectedItem.ToString()}'");
            if (!string.IsNullOrEmpty(textBox14.Text.Trim()))
                filters.Add($"[Модель машини] LIKE '*{textBox14.Text.Trim()}*'");
            if (!string.IsNullOrEmpty(textBox9.Text.Trim()))
                filters.Add($"[Номер водійського посвідчення] LIKE '*{textBox9.Text.Trim()}*'");
            if (!string.IsNullOrEmpty(textBox13.Text.Trim()))
                filters.Add($"[Номер телефону] LIKE '*{textBox13.Text.Trim()}*'");
            if (!string.IsNullOrEmpty(textBox15.Text.Trim()))
                filters.Add($"[ID ДТП] = {textBox15.Text.Trim()}");
            if (!string.IsNullOrEmpty(textBox16.Text.Trim()))
                filters.Add($"[Адреса] LIKE '*{textBox16.Text.Trim()}*'");

            // Об'єднуємо всі фільтри за допомогою оператора AND
            string filterExpression = string.Join(" AND ", filters);

            // Виконуємо фільтрацію та відображення результату у dataGridView2
            (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = filterExpression;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Отримання значення ID запису для редагування
                int recordId = string.IsNullOrEmpty(textBox10.Text) ? -1 : int.Parse(textBox10.Text);

                // Перевірка, чи вказано значення ID запису для редагування
                if (recordId == -1)
                {
                    MessageBox.Show("Вкажіть ID запису для редагування");
                    return;
                }

                // Перевірка правильності введення номера телефону
                if (!string.IsNullOrEmpty(textBox13.Text))
                {
                    string phoneNumber = textBox13.Text;
                    string phonePattern = @"^(?:\+380|380|0)\d{9}$";
                    if (!System.Text.RegularExpressions.Regex.IsMatch(phoneNumber, phonePattern))
                    {
                        MessageBox.Show("Номер телефону введено невірно. Номер має бути в форматі +380XXXXXXXXX, 380XXXXXXXXX або 0XXXXXXXXX.");
                        return;
                    }
                }

                // Перевірка правильності введення номера водійського посвідчення
                if (!string.IsNullOrEmpty(textBox9.Text))
                {
                    string licenseNumber = textBox9.Text;
                    string licensePattern = @"^[A-Z]{3}\d{6}$";
                    if (!System.Text.RegularExpressions.Regex.IsMatch(licenseNumber, licensePattern))
                    {
                        MessageBox.Show("Номер водійського посвідчення введено невірно. Він має бути у форматі 3 великі букви та 6 цифр.");
                        return;
                    }
                }

                // Формування SQL-запиту для редагування інформації
                string updateQuery = "UPDATE [Відомості про учасників ДТП] SET ";
                List<string> updateFields = new List<string>();

                if (!string.IsNullOrEmpty(textBox8.Text))
                    updateFields.Add($"[Прізвище] = '{textBox8.Text}'");
                if (comboBox1.SelectedItem != null && !string.IsNullOrEmpty(comboBox1.SelectedItem.ToString()))
                    updateFields.Add($"[Тип події] = '{comboBox1.SelectedItem.ToString()}'");
                if (!string.IsNullOrEmpty(textBox11.Text))
                    updateFields.Add($"[Ім'я] = '{textBox11.Text}'");
                if (!string.IsNullOrEmpty(textBox6.Text))
                    updateFields.Add($"[По батькові] = '{textBox6.Text}'");
                if (!string.IsNullOrEmpty(textBox7.Text))
                    updateFields.Add($"[Вік] = {textBox7.Text}");
                if (!string.IsNullOrEmpty(textBox12.Text))
                    updateFields.Add($"[Номерний знак] = '{textBox12.Text}'");
                if (!string.IsNullOrEmpty(textBox14.Text))
                    updateFields.Add($"[Модель машини] = '{textBox14.Text}'");
                if (!string.IsNullOrEmpty(textBox9.Text))
                    updateFields.Add($"[Номер водійського посвідчення] = '{textBox9.Text}'");
                if (!string.IsNullOrEmpty(textBox13.Text))
                    updateFields.Add($"[Номер телефону] = '{textBox13.Text}'");
                if (!string.IsNullOrEmpty(textBox15.Text))
                    updateFields.Add($"[ID ДТП] = {textBox15.Text}");
                if (!string.IsNullOrEmpty(textBox16.Text))
                    updateFields.Add($"[Адреса] = '{textBox16.Text}'");
                if (comboBox3.SelectedItem != null && !string.IsNullOrEmpty(comboBox3.SelectedItem.ToString()))
                    updateFields.Add($"[ID інспектора що вносив дані] = '{comboBox3.SelectedItem.ToString()}'");

                // Складання полів для оновлення
                updateQuery += string.Join(", ", updateFields);

                // Додавання WHERE умови для визначення запису, який треба відредагувати
                updateQuery += $" WHERE [ID запису] = {recordId}";

                // Виконання SQL-запиту на оновлення інформації
                OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection);
                updateCommand.ExecuteNonQuery();

                LoadData();
                MessageBox.Show("Інформація успішно відредагована");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Отримання значення ID запису для видалення
                int recordId = string.IsNullOrEmpty(textBox10.Text) ? -1 : int.Parse(textBox10.Text);

                // Перевірка, чи вказано значення ID запису для видалення
                if (recordId == -1)
                {
                    MessageBox.Show("Вкажіть ID запису для видалення");
                    return;
                }

                // Отримання ID ДТП для даного запису з таблиці "Подія"
                string getIdDtpQuery = $"SELECT [ID ДТП] FROM [Подія] WHERE [ID запису] = {recordId}";
                OleDbCommand getIdDtpCommand = new OleDbCommand(getIdDtpQuery, connection);
                object result = getIdDtpCommand.ExecuteScalar();

                // Перевірка, чи існує запис з таким ID
                if (result == null)
                {
                    MessageBox.Show("Запис з таким ID не знайдено");
                    return;
                }

                int idDtp = (int)result;

                // Формування SQL-запиту для видалення записів з таблиці "Подія" з тим самим ID ДТП
                string deleteFromPodiaQuery = $"DELETE FROM [Подія] WHERE [ID ДТП] = {idDtp}";
                // Логування запиту
                Console.WriteLine(deleteFromPodiaQuery);

                // Виконання SQL-запиту на видалення записів з таблиці "Подія"
                OleDbCommand deleteFromPodiaCommand = new OleDbCommand(deleteFromPodiaQuery, connection);
                deleteFromPodiaCommand.ExecuteNonQuery();

                // Формування SQL-запиту для видалення записів з таблиці "Відомості про учасників ДТП" з тим самим ID ДТП
                string deleteFromUchasnykyQuery = $"DELETE FROM [Відомості про учасників ДТП] WHERE [ID ДТП] = {idDtp}";
                // Логування запиту
                Console.WriteLine(deleteFromUchasnykyQuery);

                // Виконання SQL-запиту на видалення записів з таблиці "Відомості про учасників ДТП"
                OleDbCommand deleteFromUchasnykyCommand = new OleDbCommand(deleteFromUchasnykyQuery, connection);
                deleteFromUchasnykyCommand.ExecuteNonQuery();

                LoadData();
                MessageBox.Show("Записи успішно видалені");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> searchConditions = new List<string>();

                if (!string.IsNullOrEmpty(textBox23.Text))
                    searchConditions.Add($"[ID інспектора] = '{textBox23.Text}'");
                if (!string.IsNullOrEmpty(textBox20.Text))
                    searchConditions.Add($"[Прізвище] LIKE '%{textBox20.Text}%'");
                if (!string.IsNullOrEmpty(textBox19.Text))
                    searchConditions.Add($"[Ім'я] LIKE '%{textBox19.Text}%'");
                if (!string.IsNullOrEmpty(textBox18.Text))
                    searchConditions.Add($"[По батькові] LIKE '%{textBox18.Text}%'");
                if (!string.IsNullOrEmpty(textBox21.Text))
                    searchConditions.Add($"[Звання] LIKE '%{textBox21.Text}%'");
                if (!string.IsNullOrEmpty(textBox22.Text))
                    searchConditions.Add($"[Відділ] LIKE '%{textBox22.Text}%'");

                string filterExpression = string.Join(" AND ", searchConditions);

                (dataGridView3.DataSource as DataTable).DefaultView.RowFilter = filterExpression;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Отримання значення ID інспектора для додавання
                if (string.IsNullOrEmpty(textBox23.Text))
                {
                    MessageBox.Show("Вкажіть ID інспектора для додавання");
                    return;
                }

                string inspectorId = textBox23.Text;

                // Перевірка правильності введення ID інспектора
                string idPattern = @"^[A-Z]{2}\d{4}$";
                if (!System.Text.RegularExpressions.Regex.IsMatch(inspectorId, idPattern))
                {
                    MessageBox.Show("ID інспектора введено невірно. Він має бути у форматі 2 великі букви та 4 цифри.");
                    return;
                }

                // Формування SQL-запиту для додавання нового інспектора
                string insertQuery = "INSERT INTO [Інформація про інспекторів] ([ID інспектора], [Прізвище], " +
                    "[Ім'я], [По батькові], [Звання], [Відділ]) VALUES ";
                List<string> insertFields = new List<string>
        {
            $"'{inspectorId}'",
            !string.IsNullOrEmpty(textBox20.Text) ? $"'{textBox20.Text}'" : "NULL",
            !string.IsNullOrEmpty(textBox19.Text) ? $"'{textBox19.Text}'" : "NULL",
            !string.IsNullOrEmpty(textBox18.Text) ? $"'{textBox18.Text}'" : "NULL",
            !string.IsNullOrEmpty(textBox21.Text) ? $"'{textBox21.Text}'" : "NULL",
            !string.IsNullOrEmpty(textBox22.Text) ? $"'{textBox22.Text}'" : "NULL"
        };

                insertQuery += $"({string.Join(", ", insertFields)})";

                // Виконання SQL-запиту на додавання нового інспектора
                OleDbCommand insertCommand = new OleDbCommand(insertQuery, connection);
                insertCommand.ExecuteNonQuery();

                LoadData();
                MessageBox.Show("Новий інспектор успішно доданий");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Отримання значення ID інспектора для видалення
                if (string.IsNullOrEmpty(textBox23.Text))
                {
                    MessageBox.Show("Вкажіть ID інспектора для видалення");
                    return;
                }

                string inspectorId = textBox23.Text;

                // Перевірка правильності введення ID інспектора
                string idPattern = @"^[A-Z]{2}\d{4}$";
                if (!System.Text.RegularExpressions.Regex.IsMatch(inspectorId, idPattern))
                {
                    MessageBox.Show("ID інспектора введено невірно. Він має бути у форматі 2 великі букви та 4 цифри.");
                    return;
                }

                // Формування SQL-запиту для видалення інспектора
                string deleteQuery = $"DELETE FROM [Інформація про інспекторів] WHERE [ID інспектора] = '{inspectorId}'";

                // Виконання SQL-запиту на видалення інспектора
                OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection);
                deleteCommand.ExecuteNonQuery();

                LoadData();
                MessageBox.Show("Інспектор успішно видалений");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            LoadData();
        }
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DAI.accdb";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    string query = $"SELECT * FROM [Подія] INNER JOIN [Відомості про учасників ДТП] ON " +
                        $"[Подія].[ID ДТП] = [Відомості про учасників ДТП].[ID ДТП] WHERE [Подія].[ID ДТП] = @ID_DTP";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@ID_DTP", textBox25.Text);
                        OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);

                        string report = "Звіт:\n\n";

                        foreach (System.Data.DataRow row in dataTable.Rows)
                        {
                            foreach (System.Data.DataColumn col in dataTable.Columns)
                            {
                                string value = row[col].ToString();
                                if (col.DataType == typeof(DateTime))
                                {
                                    DateTime dateTimeValue = DateTime.Parse(value);
                                    value = dateTimeValue.ToShortDateString();
                                }
                                report += $"{col.ColumnName}: {value}\n";
                            }
                            report += "\n";
                        }

                        MessageBox.Show(report, "Звіт");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }
    }
 }