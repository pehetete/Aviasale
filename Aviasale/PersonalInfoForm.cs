using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Aviasale
{
    public partial class PersonalInfoForm : Form
    {
        public PersonalInfoForm() => InitializeComponent();

        internal static List<List<object>> Passenger = new List<List<object>>();
        int Passangers = GetInfoForm.Adults + GetInfoForm.Kids,
            PassangerCounter = 1;

        private void btnNext_Click(object sender, EventArgs e)
        {
            int Age = DateTime.Now.Year - Convert.ToDateTime(txtBirthDate.Text).Year;
            double Price = 0;
            bool IsAdult = false;

            if (Age < 2)
                Price = 0;
            else if (Age >= 2 && Age <= 12)
                Price = GetInfoForm.Price * 0.5;
            else if (Age > 12)
            {
                Price = GetInfoForm.Price;
                IsAdult = true;
            }

            Passenger.Add(new List<object> { txtSecond.Text, txtFirst.Text, txtMiddle.Text, Age, Price, IsAdult, txtPassport.Text });

            if (PassangerCounter < Passangers)
            {
                PassangerCounter++;
                Text = "Пассажир №" + (PassangerCounter);
            }
            else
            {
                Hide();
                FinalInfoForm finalInfoForm = new FinalInfoForm();
                finalInfoForm.Show();
            }
        }

        private void PersonalInfoForm_Load(object sender, EventArgs e)
        {
            Text = "Пассажир №1";
        }
    }
}
