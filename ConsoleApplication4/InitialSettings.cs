using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace Project_1
{
    class InitialSettings
    {
        General gen = new General();

        #region Singleton

        private static InitialSettings instance;

        public static InitialSettings Instance
        {
            get
            {
                if (instance == null) instance = new InitialSettings();

                return instance;
            }
        }

        #endregion
        #region InitialSettings
        public InitialSettings()
        {
            try
            {
                ApplicationSetUp();
                if (Global.SapApplication.Language == SAPbouiCOM.BoLanguages.ln_English | Global.SapApplication.Language == SAPbouiCOM.BoLanguages.ln_English_Cy | Global.SapApplication.Language == SAPbouiCOM.BoLanguages.ln_English_Gb | Global.SapApplication.Language == SAPbouiCOM.BoLanguages.ln_English_Sg)
                {
                    Boolean oFlag = false;
                    oFlag = IsValid();
                    if (oFlag == true)
                    {
                        ConnectCompany();
                    }
                    else if (oFlag == false)
                    {
                        Global.SapApplication.MessageBox("Installing Add-On failed due to License mismatch", 1, "Ok", "", "");
                    }
                    //ConnectCompany();
                   UserFields.Instance.UpdateDatabase();
                    SetMenuItems();

                }

               Global.SapApplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SapApplication_MenuEvent);
               Global.SapApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SapApplication_AppEvent);
               Global.SapApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SapApplication_ItemEvent);
               Global.SapApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SapApplication_FormDataEvent);
            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
  #endregion
        #region Intialize Application Instance
        private void ApplicationSetUp()
        {
            SAPbouiCOM.SboGuiApi GuiApi = new SAPbouiCOM.SboGuiApi();
            try
            {
                string ConnString = null;
                ConnString = Environment.GetCommandLineArgs().GetValue(1).ToString();
                GuiApi.Connect(ConnString);
                Global.SapApplication = GuiApi.GetApplication(-1);

            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("The Sap Buisness One Application could not be found");
                System.Environment.Exit(0);
            }

        }
         #endregion
        #region Connect to the company
        private void ConnectCompany()
        {
            try
            {
                string sErrorMsg;
                string cookie;
                string connStr;
                Global.SapCompany = new SAPbobsCOM.Company();
                cookie = Global.SapCompany.GetContextCookie();
                connStr = Global.SapApplication.Company.GetConnectionContext(cookie);
                Global.SapCompany.SetSboLoginContext(connStr);
                Global.SapCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                Global.SapCompany.Connect();
                sErrorMsg = Global.SapCompany.GetLastErrorDescription();
                Global.SapApplication.StatusBar.SetText("Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch
            {
                Global.SapApplication.MessageBox(Global.SapCompany.GetLastErrorDescription().ToString(), 1, "Ok", "", "");
            }
        }
        #endregion
        #region sap application menu event 
        private void SapApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            string _str_Objtype = "";
            try
            {

                Global.bubblevalue = true;
                SAPbouiCOM.Form form = Global.SapApplication.Forms.ActiveForm;
                if (form.TypeEx == "134")
                {
                    Vitem.Instance.SapApplication_MenuEvent(pVal);

                }
               
            }
            catch (Exception ex)
            {
                Global.SapApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                Global.bubblevalue = false;
            }
            BubbleEvent = Global.bubblevalue;
        }
        #endregion 
        #region Item Event

        private void SapApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            bool bubbleval = true;
            try
            {
                if (pVal.FormTypeEx == "134" | pVal.FormTypeEx == "-62" | pVal.FormTypeEx == "-134")
                    bubbleval = Vitem.Instance.SapApplication_ItemEvent(pVal);

                BubbleEvent = bubbleval;
            }
            catch (Exception ex)
            {
            }
        }

        #endregion Item Event
        #region DataEvent
        void SapApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
           
            {
                string _str_Objtype = "";

                try
                {
                    
                    if (BusinessObjectInfo.FormTypeEx == "134")
                    {
                        Vitem.Instance.SapApplication_FormDataEvent(BusinessObjectInfo);
                    }
                }
                catch (Exception ex)
                {
                    Global.SapApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    Global.bubblevalue = false;
                }
            }
            BubbleEvent = Global.bubblevalue;
        }
         #endregion
        #region AppEvent
        private void SapApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Global.SapApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    System.Environment.Exit(0);
                    break;

            }
        }
        #endregion
        private void CreateMenuItem(SAPbouiCOM.BoMenuType mType, string uniqueID, string desc, int position, string menuItemId)
        {
            SAPbouiCOM.Menus Menu = null;
            SAPbouiCOM.MenuItem MenuItem = null;
            Menu = Global.SapApplication.Menus;
            string rootPath = System.Windows.Forms.Application.StartupPath;
            rootPath = rootPath.Remove(rootPath.Length - 9, 9);

            SAPbouiCOM.MenuCreationParams CreationPara = null;
            CreationPara = (SAPbouiCOM.MenuCreationParams)(Global.SapApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams));
            MenuItem = Global.SapApplication.Menus.Item(menuItemId);

            try
            {
                Menu = MenuItem.SubMenus;
                CreationPara.Type = mType;
                CreationPara.UniqueID = uniqueID;
                CreationPara.String = desc;
                CreationPara.Position = position;
                CreationPara.Enabled = true;
                Menu.AddEx(CreationPara);
            }
            catch (Exception ex)
            {
                //Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }


        internal void Enable(SAPbouiCOM.Form form, bool Enable)
        {
            try
            {
                form.Freeze(true);
            }
            catch
            {
            }
        }
        public void SetMenuItems()
        {

            //CreateMenuItem(SAPbouiCOM.BoMenuType.mt_STRING, "Coordnator", "CoordnatorZone ", 1, "Coordnator");

            //CreateMenuItem(SAPbouiCOM.BoMenuType.mt_POPUP, "Document", "Document", 20, "43520");
            //CreateMenuItem(SAPbouiCOM.BoMenuType.mt_STRING, "DocumentAdd", "DocumentAdd ", 1, "Document");






        }


        #region
        public Boolean IsValid()
        {
            try
            {
                
                Global.SapApplication.ActivateMenuItem("257");   
                SAPbouiCOM.Form oform = Global.SapApplication.Forms.ActiveForm;
                SAPbouiCOM.EditText oHWKEY = (SAPbouiCOM.EditText)oform.Items.Item("79").Specific;
                Global.HWKEY = new string[] { "N2092941383", "F1534030594", "Y1334940735", "D1206114874", "Q0021813522", "A0335651095", "A0061802481", "E0649908341", "K0718181110" };
                String CRRHWKEY = oHWKEY.Value.ToString();
                Global.SapApplication.Forms.ActiveForm.Close();     
                for (int i = 0; i <= Global.HWKEY.Length - 1; i++)
                {
                    if (CRRHWKEY == Global.HWKEY[i])
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                return false;
            }
        }

        #endregion
        }

    }

