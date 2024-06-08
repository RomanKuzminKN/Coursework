using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace DAI
{
    public partial class Form2 : Form
    {
        private OleDbConnection connection;
        private const string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\ДДМА\3 Курс\Курсовая ОБДЗ\DAI\DAI\DAI.accdb";
        private Form1 form1;
        public Form2(Form1 form1)
        {
            InitializeComponent();
            connection = new OleDbConnection(connectionString);
            FillInspectorComboBox();
            this.form1 = form1;
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
                    comboBox2.Items.Add(reader["ID інспектора"].ToString());
                }
                reader.Close();
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
        private bool IsValidPhoneNumber(string phoneNumber)
        {
            // Перевірка номеру телефону на відповідність одному з форматів: +380XXXXXXXXX, 380XXXXXXXXX, 0XXXXXXXXX
            return System.Text.RegularExpressions.Regex.IsMatch(phoneNumber, @"^\+?380\d{9}$|^0\d{9}$");
        }

        private bool IsValidDriverLicenseNumber(string licenseNumber)
        {
            // Перевірка номера водійського посвідчення на формат: 3 великі літери та 6 цифр
            return System.Text.RegularExpressions.Regex.IsMatch(licenseNumber, @"^[A-Z]{3}\d{6}$");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Збираємо дані з форми
                string inspectorId = comboBox2.SelectedItem?.ToString();
                string[] licensePlates = textBox5.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string eventType = comboBox1.SelectedItem?.ToString();
                DateTime eventDate = dateTimePicker1.Value.Date;
                string location = textBox2.Text;
                string reason = textBox4.Text;
                string[] participants = textBox1.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string[] carModels = textBox3.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string[] ages = textBox6.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string[] addresses = textBox7.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string[] phoneNumbers = textBox8.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string[] driverLicenses = textBox9.Text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                if (inspectorId == null || eventType == null)
                {
                    MessageBox.Show("Будь ласка, оберіть значення в комбо-боксах.");
                    return;
                }

                // Генеруємо ID для нового ДТП
                OleDbCommand idCommand = new OleDbCommand("SELECT MAX([ID ДТП]) FROM [Подія]", connection);
                object result = idCommand.ExecuteScalar();
                int accidentId = (result != DBNull.Value) ? Convert.ToInt32(result) + 1 : 1;

                // Вставка даних в таблицю "Подія"
                foreach (var carModel in carModels)
                {
                    OleDbCommand eventCommand = new OleDbCommand("INSERT INTO [Подія] ([ID ДТП], [Тип події], [Дата події], [Локація події], [Причина аварії], [ID інспектора що вносив дані]) VALUES (?, ?, ?, ?, ?, ?)", connection);
                    eventCommand.Parameters.AddWithValue("@ID ДТП", accidentId);
                    eventCommand.Parameters.AddWithValue("@Тип події", eventType);
                    eventCommand.Parameters.AddWithValue("@Дата події", eventDate);
                    eventCommand.Parameters.AddWithValue("@Локація події", location);
                    eventCommand.Parameters.AddWithValue("@Причина аварії", reason);
                    eventCommand.Parameters.AddWithValue("@ID інспектора що вносив дані", inspectorId);
                    eventCommand.ExecuteNonQuery();
                }

                // Вставка даних в таблицю "Відомості про учасників ДТП"
                int licensePlateIndex = 0;
                foreach (var participant in participants)
                {
                    var names = participant.Trim().Split(' ');
                    string licensePlate = licensePlateIndex < licensePlates.Length ? licensePlates[licensePlateIndex].Trim() : string.Empty;
                    string carModel = licensePlateIndex < carModels.Length ? carModels[licensePlateIndex].Trim() : string.Empty;
                    string age = licensePlateIndex < ages.Length ? ages[licensePlateIndex].Trim() : string.Empty;
                    string address = licensePlateIndex < addresses.Length ? addresses[licensePlateIndex].Trim() : string.Empty;
                    string phoneNumber = licensePlateIndex < phoneNumbers.Length ? phoneNumbers[licensePlateIndex].Trim() : string.Empty;
                    string driverLicense = licensePlateIndex < driverLicenses.Length ? driverLicenses[licensePlateIndex].Trim() : string.Empty;

                    if (!string.IsNullOrEmpty(phoneNumber) && !IsValidPhoneNumber(phoneNumber))
                    {
                        MessageBox.Show($"Невірний формат номеру телефону: {phoneNumber}");
                        return;
                    }

                    if (!string.IsNullOrEmpty(driverLicense) && !IsValidDriverLicenseNumber(driverLicense))
                    {
                        MessageBox.Show($"Невірний формат номеру водійського посвідчення: {driverLicense}");
                        return;
                    }

                    OleDbCommand participantCommand = new OleDbCommand("INSERT INTO [Відомості про учасників ДТП] ([ID ДТП], [Прізвище], [Ім'я], [По батькові], " +
                        "[Номерний знак], [Модель машини], [ID інспектора що вносив дані], [Вік], [Адреса], [Номер телефону], [Номер водійського посвідчення]) " +
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", connection);
                    participantCommand.Parameters.AddWithValue("@ID ДТП", accidentId);
                    participantCommand.Parameters.AddWithValue("@Прізвище", names.Length > 0 ? names[0] : string.Empty);
                    participantCommand.Parameters.AddWithValue("@Ім'я", names.Length > 1 ? names[1] : string.Empty);
                    participantCommand.Parameters.AddWithValue("@По батькові", names.Length > 2 ? names[2] : string.Empty);
                    participantCommand.Parameters.AddWithValue("@Номерний знак", licensePlate);
                    participantCommand.Parameters.AddWithValue("@Модель машини", carModel);
                    participantCommand.Parameters.AddWithValue("@ID інспектора що вносив дані", inspectorId);
                    participantCommand.Parameters.AddWithValue("@Вік", age);
                    participantCommand.Parameters.AddWithValue("@Адреса", address);
                    participantCommand.Parameters.AddWithValue("@Номер телефону", phoneNumber);
                    participantCommand.Parameters.AddWithValue("@Номер водійського посвідчення", driverLicense);
                    participantCommand.ExecuteNonQuery();

                    licensePlateIndex++;
                }

                MessageBox.Show("Дані успішно додано до бази даних!");
                form1.LoadData();
                this.Close();
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
    }
}