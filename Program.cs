using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.ReplyMarkups;

namespace Telegram_NCD
{
    class ExcelSearchResult
    {
        public bool Found { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static TelegramBotClient botClient;
        static string excelFilePath = @"NCD_Accounts.xlsx"; // Update this with the path to your Excel file

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Or LicenseContext.Commercial if applicable
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            //CheckExcelData();

            try
            {
                botClient = new TelegramBotClient("YOUR_Telegram_Bot_Token");

            }
            catch (Exception e)
            {

                Console.Write(e.Message);
            }




            botClient.OnMessage += Bot_OnMessage;
            Console.Write("Bot is working");
            botClient.StartReceiving();
            Console.ReadLine();
        }

        private static ExcelSearchResult SearchExcelForContact(string phoneNumber)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assumes the data is in the first worksheet
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Text == phoneNumber) // Assumes phone numbers are in the first column
                    {
                        return new ExcelSearchResult
                        {
                            Found = true,
                            UserName = worksheet.Cells[row, 2].Text, // Assumes user names are in the second column
                            Password = worksheet.Cells[row, 3].Text // Assumes passwords are in the third column
                        };
                    }
                }
            }
            return new ExcelSearchResult { Found = false };
        }

        static async void Bot_OnMessage(object sender, MessageEventArgs e)
        {
            var contact = e.Message.Contact;
            var senderId = e.Message.From.Id;


            if (e.Message.Contact != null)
            {
                if (contact.UserId == senderId)
                {
                    string phoneNumber = e.Message.Contact.PhoneNumber;
                    if (phoneNumber.Length >= 9)
                    {
                        phoneNumber = phoneNumber.Substring(phoneNumber.Length - 9);
                    }
                    var result = SearchExcelForContact(phoneNumber);
                    string response = result.Found ? $"Your user name : {result.UserName}, password : {result.Password}" : "Sorry, not found";

                    await botClient.SendTextMessageAsync(
                        chatId: e.Message.Chat,
                        text: response
                    );
                }
                else
                {
                    string response = "You are trying to steal someone else's Account, do not use bot for cheating be honest";

                    await botClient.SendTextMessageAsync(
                        chatId: e.Message.Chat,
                        text: response
                    );

                }

            }
            else
            {
                // Create a single button
                var requestContactButton = new KeyboardButton("Send contact") { RequestContact = true };

                // Wrap the button in an array and create a keyboard
                var keyboard = new ReplyKeyboardMarkup(new[] { new[] { requestContactButton } }, resizeKeyboard: true, oneTimeKeyboard: true);

                await botClient.SendTextMessageAsync(
                    chatId: e.Message.Chat,
                    text: "Please send your contact",
                    replyMarkup: keyboard
                );
            }
        }

        //private static void CheckExcelData()
        //{
        //    using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        //    {
        //        var worksheet = package.Workbook.Worksheets[0]; // Assumes the data is in the first worksheet
        //        int totalRows = worksheet.Dimension.End.Row;

        //        for (int row = 2; row <= totalRows; row++)
        //        {
        //            string phoneNumber = worksheet.Cells[row, 1].Text; // Assumes phone numbers are in the first column
        //            string userName = worksheet.Cells[row, 2].Text;    // Assumes user names are in the second column
        //            string password = worksheet.Cells[row, 3].Text;    // Assumes passwords are in the third column

        //            Console.WriteLine($"Row {row}: Phone = {phoneNumber}, User Name = {userName}, Password = {password}");
        //        }
        //    }
        //}


    }
}


