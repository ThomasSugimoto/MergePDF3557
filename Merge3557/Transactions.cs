using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Merge3557
{
    class Transactions
    {
        public string ProcessDate = string.Empty;
        public Int32 BatchNo;
        public string SeqNo = string.Empty;
        public string MODE = string.Empty;
        public string BankAccount = string.Empty;
        public string CheckNo = string.Empty;
        public double AmountInCents;
        public string PageType = string.Empty;
        public string FileName = string.Empty;

        /// <summary>
        ///     Creates the class Transactions for Lists
        /// </summary>
        /// <param name="ProcessDate"></param>
        /// <param name="BatchNo"></param>
        /// <param name="SeqNo"></param>
        /// <param name="MODE"></param>
        /// <param name="BankAccount"></param>
        /// <param name="CheckNo"></param>
        /// <param name="AmountInCents"></param>
        /// <param name="PageType"></param>
        /// <param name="FileName"></param>
        public Transactions(string ProcessDate, Int32 BatchNo, string SeqNo, string MODE, string BankAccount, string CheckNo, double AmountInCents, string PageType, string FileName)
        {
            try
            {
                this.ProcessDate = ProcessDate;
                this.BatchNo = BatchNo;
                this.SeqNo = SeqNo;
                this.MODE = MODE;
                this.BankAccount = BankAccount;
                this.CheckNo = CheckNo;
                this.AmountInCents = AmountInCents;
                this.PageType = PageType;
                this.FileName = FileName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
        }
        public Transactions() { }
    }
}
