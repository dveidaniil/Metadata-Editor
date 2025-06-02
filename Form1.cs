using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Security.Cryptography;



namespace Metadata_Editor
{
    public partial class Form1 : Form
    {

        string filename;
        string filepath;
        string userName = Environment.UserName;

        DateTime thisDay = DateTime.Today;

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.Title = "Открыть файл";
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            filename = openFileDialog1.FileName;
            filepath = openFileDialog1.FileName;


        }

        

        

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = Path.GetFileName(openFileDialog1.FileName);
                filepath = openFileDialog1.FileName;
                // например, вывести путь на форму, если хотите
                // label1.Text = filename;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true))
            {
                var packageProps = doc.PackageProperties;
                var extendedProps = doc.ExtendedFilePropertiesPart.Properties;

                if (packageProps != null)
                {
                    // Генерируем случайные временные метки
                    var random = new Random();
                    DateTime now = DateTime.Now;

                    // @ERROR НЕ РАБОТАЕТ ВАЩЕ
                    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    // Случайное время редактирования (9-25 минут)
                    // @ERROR НЕ РАБОТАЕТ ВАЩЕ
                    int editMinutes = random.Next(9, 26);
                    int editSeconds = random.Next(0, 60);
                    TimeSpan editTime = new TimeSpan(0, editMinutes, editSeconds);

                    // Дата создания - случайное время в прошлом (до 30 дней)
                    DateTime created = now.AddDays(-random.Next(1, 13))
                                       .AddHours(-random.Next(0, 24))
                                       .AddMinutes(-random.Next(0, 60));

                    // Дата изменения = дата создания + время редактирования
                    DateTime modified = created.Add(editTime);

                    // Дата последнего открытия - между созданием и текущим временем
                    DateTime lastPrinted = created.AddDays(random.Next(0, (now - created).Days));

                    // Устанавливаем свойства
                    packageProps.Creator = userName; // Удаляем автора
                    packageProps.LastModifiedBy = userName; // Удаляем "Кем сохранен"
                    packageProps.Created = created;
                    packageProps.Modified = modified;
                    packageProps.LastPrinted = lastPrinted;
                    


                    // Дополнительные свойства для правдоподобности
                    packageProps.Revision = "1";
                    packageProps.Version = "1";

                    int totalTimeMinutes = random.Next(23, 64); // Случайное время в минутах (от 1 до 180)
                    extendedProps.TotalTime.Text = Convert.ToString(totalTimeMinutes); // например, 123 минуты
                    extendedProps.Save();
                    
                    MessageBox.Show("Метаданные успешно изменены!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
