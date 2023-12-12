using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;

namespace Project_1
{

    class General
    {
        public struct PairValue
        {
            public string value;
            public string desc;
        }


        #region DisableControl
        public void DisableControl(SAPbouiCOM.Form frm, string ControlName)
        {
            try
            {
                SAPbouiCOM.Item txtDisable = (SAPbouiCOM.Item)frm.Items.Item(ControlName);
                txtDisable.Enabled = false;
            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "", "", "");
            }
        }
        #endregion

        #region clear combo
        //public void ClearCombo(SAPbouiCOM.ComboBox ddlClear, bool blnBlankNeeded)
        //{
        //    try
        //    {
        //        for (int i = ddlClear.ValidValues.Count - 1; i >= 0; i--)
        //        {
        //            ddlClear.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
        //        }
        //        if (blnBlankNeeded)
        //        {
        //            ddlClear.ValidValues.Add("-1", "  ");
        //            ddlClear.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
        //    }
        //}
        #endregion

      

        public void UpdateDB()
        {
            UserFields UserFields = new UserFields();
            
        }


        


    }
}
