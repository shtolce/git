using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DivGameApi.Models;
namespace DivGameApi.Services
{
    public class DialogService
    {
        public Dialog1 dialogText { get; set; }
        public DialogService()
        {
            Dialog1 dialog = new Dialog1();
            dialog.PersonId = "1";
            dialog.Text = "fuck off";
            dialogText = dialog;
        }


    }
}
