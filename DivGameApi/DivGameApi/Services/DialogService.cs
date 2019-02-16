using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DivGameApi.Models;
namespace DivGameApi.Services
{
    public class DialogService
    {
        public Dialog1 dialogText { get; set; }
        public async void Upd()
        {
                await Task.Run(() =>
                {
                    int i = 0;
                    while (true)
                    {
                        i++;
                        dialogText.Text = "fuck off" +i.ToString();
                        Thread.Sleep(3000);
                        if (i > 10000) i = 0;
                    }
                });
        }
        public DialogService()
        {
            Dialog1 dialog = new Dialog1();
            dialog.PersonId = "1";
            dialog.Text = "fuck off";
            dialogText = dialog;
            Upd();   


        }


    }
}
