using System;
using System.Collections.Generic;
using System.Text;

namespace Project_1
{
    class Vitemmaster
    {
        General gen = new General();
        bool ActForm = false;
        #region Instance creation of Vitemmaster class


        private static Vitemmaster instance;
        public static Vitemmaster Instance
        {
            get
            {
                if (instance == null)
                    instance = new Vitemmaster();

                return instance;
            }
        }
        #endregion
        #region SapApplication_FormDataEvent
        internal void SapApplication_FormDataEvent(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                if (BusinessObjectInfo.FormTypeEx == "150" & BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction)
                {

                    SAPbouiCOM.Form Form = Global.SapApplication.Forms.Item(BusinessObjectInfo.FormUID);

                }


            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        #endregion
        #region SapApplication_MenuEvent
        internal void SapApplication_MenuEvent(SAPbouiCOM.MenuEvent pVal)
        {


            try
            {
                //SAPbouiCOM.Form Form = Global.SapApplication.Forms.ActiveForm;
                if (pVal.MenuUID == "150" & pVal.BeforeAction == false)
                {
                }

                #region Ok Mode
                if ((pVal.MenuUID == "1288" | pVal.MenuUID == "1289" | pVal.MenuUID == "1290" | pVal.MenuUID == "1291") & pVal.BeforeAction == false)
                {


                }
                #endregion
            }

            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "ok", "", "");
            }
        }
        #endregion
        #region SapApplication_ItemEvent
        internal bool SapApplication_ItemEvent(SAPbouiCOM.ItemEvent pVal)
        {
            //bool bubval = true;
            try
            {

                if (pVal.FormTypeEx == "150")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "OK", "", "");
            }

        }
        #endregion 
    }
}
