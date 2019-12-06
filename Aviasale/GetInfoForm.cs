using System;
using System.Windows.Forms;

namespace Aviasale
{
    public partial class GetInfoForm : Form
    {
        public GetInfoForm() => InitializeComponent();

        internal static int Adults { get; set; }
        internal static int Kids { get; set; }
        internal static string ToWhere { get; set; }
        internal static double Price { get; set; }


        private void btnNext_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtAdult.Text) && !string.IsNullOrEmpty(txtKids.Text) && 
               (Convert.ToInt32(txtAdult.Text) + Convert.ToInt32(txtKids.Text)) != 0)
            {
                Adults = int.Parse(txtAdult.Text);
                Kids = int.Parse(txtKids.Text);
                ToWhere = coBoxTo.Text;

                Hide();
                PersonalInfoForm personalInfoForm = new PersonalInfoForm();
                personalInfoForm.Show();
            }
            else MessageBox.Show("Заполните все поля!","Ошибка", MessageBoxButtons.OK,MessageBoxIcon.Error);
        }

        private void txtValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
                e.Handled = true;
        }

        private void coBoxTo_TextChanged(object sender, EventArgs e)
        {
            if (coBoxTo.SelectedIndex == 0)
                Price = 1000;
            else if (coBoxTo.SelectedIndex == 1)
                Price = 1200;
            else if (coBoxTo.SelectedIndex == 2)
                Price = 1400;
            else if (coBoxTo.SelectedIndex == 3)
                Price = 1600;
        }
    }
}