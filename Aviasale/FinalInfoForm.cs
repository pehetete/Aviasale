using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Aviasale
{
    public partial class FinalInfoForm : Form
    {
        public FinalInfoForm() => InitializeComponent();

        private void FinalInfoForm_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < PersonalInfoForm.Passenger.Count; i++)
            {
                ListViewItem item = new ListViewItem(PersonalInfoForm.Passenger[i][0].ToString());
                item.SubItems.Add(PersonalInfoForm.Passenger[i][1].ToString());
                item.SubItems.Add(PersonalInfoForm.Passenger[i][2].ToString());
                item.SubItems.Add(PersonalInfoForm.Passenger[i][3].ToString());
                item.SubItems.Add(PersonalInfoForm.Passenger[i][4].ToString());
                item.SubItems.Add(PersonalInfoForm.Passenger[i][5].ToString());

                listView1.Items.Add(item);
            }
            
        }

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < PersonalInfoForm.Passenger.Count; i++)
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var wordDoc = wordApp.Documents.Add(System.Windows.Forms.Application.StartupPath + "\\Sig.dotx");

                string FullName = $"{PersonalInfoForm.Passenger[i][0].ToString()} {PersonalInfoForm.Passenger[i][1].ToString()} {PersonalInfoForm.Passenger[i][2].ToString()}";
                int Age = int.Parse(PersonalInfoForm.Passenger[i][3].ToString());
                ReplaceStub("#Город", GetInfoForm.ToWhere, wordDoc);
                ReplaceStub("#стоимость", PersonalInfoForm.Passenger[i][4].ToString(), wordDoc);
                ReplaceStub("#ФИО", FullName, wordDoc);
                ReplaceStub("#паспорт", PersonalInfoForm.Passenger[i][6].ToString(), wordDoc);
                ReplaceStub("#дата", DateTime.Today.ToString(), wordDoc);
                if (PersonalInfoForm.Passenger[i][4].ToString() == "0")
                    ReplaceStub("#вид", "Бесплатный", wordDoc);
                else if (Age >= 2 && Age <= 12)
                    ReplaceStub("#вид", "Уцененный", wordDoc);
                else ReplaceStub("#вид", "Платный", wordDoc);
                ReplaceStub("#возраст", Age.ToString(), wordDoc);

                wordDoc.SaveAs($"{System.Windows.Forms.Application.StartupPath}\\Билет №{i + 1} {FullName}.docx");
                wordDoc.Close();
                wordApp.Quit();
            }
        }

        private void ReplaceStub(string stubToReplace, string text, Document worldDocument)
        {
            var range = worldDocument.Content;
            range.Find.ClearFormatting();
            object wdReplaceAll = WdReplace.wdReplaceAll;
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: wdReplaceAll);
        }
    }
}