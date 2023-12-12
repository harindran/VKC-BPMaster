using System;
using System.Collections.Generic;
using System.Text;

namespace Project_1
{
  
    static class SubMain
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        //[STAThread]
        static void Main()
        {
            try
            {
                InitialSettings AddOn = InitialSettings.Instance;
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message); }
        }
    }
}
 
