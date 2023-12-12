using System;
using System.Collections.Generic;
using System.Text;

namespace Project_1
{
    class Vitem
    {

        General gen = new General();
        bool ActForm = false;
        #region Instance creation of Vitem class
       

        private static Vitem instance;
        public static Vitem Instance
        {
            get
            {
                if (instance == null)
                    instance = new Vitem();

                return instance;
            }
        }
        #endregion
        #region SapApplication_FormDataEvent
        internal void SapApplication_FormDataEvent(SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                if (BusinessObjectInfo.FormTypeEx == "134" & BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction )
                {

                    SAPbouiCOM.Form Form = Global.SapApplication.Forms.Item(BusinessObjectInfo.FormUID);
                    //Mitem.Instance.AddButton1(BusinessObjectInfo);
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
               SAPbouiCOM.Form Form = Global.SapApplication.Forms.ActiveForm;
                if (pVal.MenuUID == "134" & pVal.BeforeAction == false)
                {
                }
                #region Ok Mode
                if ((pVal.MenuUID == "1288" | pVal.MenuUID == "1289" | pVal.MenuUID == "1290" | pVal.MenuUID == "1291") & pVal.BeforeAction == false)
                {
                    Mitem.Instance.ClearMatrixUnit(Form);
                    Mitem.Instance.Fillseries(Form);
                    Mitem.Instance.FillseriesBrand(Form);


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
            bool bubval = true;
            try
            {
                if (pVal.FormTypeEx == "-62" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                {
                   // Mitem.Instance.WareHouseUnitFill(pVal);
                }
                if (pVal.FormTypeEx == "-134")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                    {
                         
                    }

                   
                    if (pVal.ItemUID == "U_Country" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS & pVal.BeforeAction == false & pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE )
                    {
                        Mitem.Instance.BPCodeGeneration(pVal);
                    }
                    if (pVal.ItemUID == "U_State" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN& pVal.BeforeAction == false & pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE )
                    {
                        Mitem.Instance.BPCodeGeneration(pVal);
                    }

                }
                if (pVal.FormTypeEx == "134")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                    {

                        Mitem.Instance.ClickOnItm(pVal);

                    }

                    if (pVal.ItemUID == "40" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false & pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        Mitem.Instance.BPCodeGeneration(pVal);
                    }

                    if (pVal.ItemUID == "us" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.Before_Action == true)
                    {
                      //  Mitem.Instance.AddButton(pVal);
                        Global.SapApplication.Forms.Item(pVal.FormUID).PaneLevel = 40;
                        Mitem.Instance.ChangePane(pVal);

                    }
                    if (pVal.ItemUID == "brand" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.Before_Action == true)
                    {
                       // Mitem.Instance.AddButtonBrand(pVal);
                        Global.SapApplication.Forms.Item(pVal.FormUID).PaneLevel = 45;
                        Mitem.Instance.ChangePaneBrand(pVal);

                    }
                   
                    if ((pVal.ItemUID == "us"))
                    {
                        SAPbouiCOM.Form frm = Global.SapApplication.Forms.Item(pVal.FormUID);
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                        {
                        Mitem.Instance.Fillseries(frm); ///modified on 12-07-2012
                           // Mitem.Instance.ItemDisable(pVal);
                            

                        }

                    }
                    if ((pVal.ItemUID == "brand"))
                    {
                        SAPbouiCOM.Form frm = Global.SapApplication.Forms.Item(pVal.FormUID);
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                        {
                            Mitem.Instance.FillseriesBrand(frm);

                        }
                    }
                    if (pVal.ItemUID == "ivctrl")
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                           
                        }
                    }
                    if (pVal.ItemUID == "prchs")
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                           
                        }
                    }
                    if (pVal.ItemUID == "sals")
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                         
                        }
                    }
                                             
                  
                    if ((pVal.ItemUID == "1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) & (pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)  & (pVal.BeforeAction == true))
                    {
                       Mitem.Instance.BPCodeGeneration(pVal);
                       Mitem.Instance.addtotable(pVal);
                       Mitem.Instance.addtotableBrand(pVal);
                        

                    }
                    if ((pVal.ItemUID == "1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) & (pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) & (pVal.BeforeAction == true))
                    {
                        Mitem.Instance.UpdateTable(pVal);
                        Mitem.Instance.UpdateTableBrand(pVal);

                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false & ((pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)))
                    {
                       // Mitem.Instance.AddButton(pVal);

                    }
                }
            }
            catch (Exception e)
            {
                Global.SapApplication.MessageBox(e.Message, 1, "OK", "", "");
            }
            return bubval;
        }
        #endregion 
    }
}