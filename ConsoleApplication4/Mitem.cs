using System;
using System.Collections.Generic;
using System.Text;

namespace Project_1
{
    class Mitem
    {
        General gen = new General();
        #region Instance for Sub class
        private static Mitem instance;
        public string _str_Grpnum = "";
        public static Mitem Instance
        {
            get
            {
                if (instance == null)
                    instance = new Mitem();
                return instance;
            }
        }
        #endregion
        #region Create Instance For Vitem
        public Mitem()
        {
            Vitem vw = Vitem.Instance;
        }
        #endregion

        #region click on item
        public void ClickOnItm(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.Item oItem = null;
            try
            {
                SAPbouiCOM.Form NorForm = Global.SapApplication.Forms.Item(pVal.FormUID);
                addProducttab(NorForm);
                AddButton(pVal);
                AddButtonBrand(pVal);
                oItem = NorForm.Items.Item("7");
                oItem.Click(SAPbouiCOM.BoCellClickType.ct_Double);
                NorForm.PaneLevel = 1;

            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        #endregion

        #region add product tab
        public void addProducttab(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Folder oFolder;
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.Item oItem1;
            try
            {
                oForm.Freeze(true);
                oItem = oForm.Items.Add("us", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oItem.AffectsFormMode = false;
                oFolder.Caption = "UNIT SELECTION";
                oFolder.GroupWith("9");
                oItem.Width = 125;
                oItem1 = oForm.Items.Item("9");
                oItem.Left = oItem1.Left + oItem1.Width;
                oItem.Enabled = true;
                oItem.Visible = true;
                oForm.PaneLevel = 1;
               // int k = oItem1.FromPane;
                oItem = oForm.Items.Add("brand", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oItem.AffectsFormMode = false;
                oFolder.Caption = "BRAND SELECTION";
                oFolder.GroupWith("9");
                oItem.Width = 125;
                oItem1 = oForm.Items.Item("9");
                oItem.Left = oItem1.Left + oItem1.Width;
                oItem.Enabled = true;
                oItem.Visible = true;
                oForm.PaneLevel = 1;
                int k = oItem1.FromPane;
                oForm.Freeze(false);

                oForm.Freeze(false);
            }
            catch (Exception e)
            {
                Global.SapApplication.MessageBox(e.Message, 1, "Ok", "", "");
            }
        }
        #endregion
        #region chagepane
        public void ChangePane(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbobsCOM.Recordset oSelectOFPR = null, oSelectCINF = null, oUPDCINF = null;
            try
            {

                Global.SapApplication.Forms.ActiveForm.Freeze(true);

                Global.SapApplication.Forms.ActiveForm.PaneLevel = 40;
                Global.SapApplication.Forms.ActiveForm.Freeze(false);

            }
            catch (Exception e)
            {
                Global.SapApplication.MessageBox(e.Message, 1, "OK", "", "");
            }
        }
        #endregion
        //--------------------------------------------------------------
        #region chagepaneBrand
        public void ChangePaneBrand(SAPbouiCOM.ItemEvent pVal)
        {
            SAPbobsCOM.Recordset oSelectOFPR = null, oSelectCINF = null, oUPDCINF = null;
            try
            {

                Global.SapApplication.Forms.ActiveForm.Freeze(true);

                Global.SapApplication.Forms.ActiveForm.PaneLevel = 45;
                Global.SapApplication.Forms.ActiveForm.Freeze(false);

            }
            catch (Exception e)
            {
                Global.SapApplication.MessageBox(e.Message, 1, "OK", "", "");
            }
        }
        #endregion
        //--------------------------------------------------------------
        #region addbutton
        internal void AddButton(SAPbouiCOM.ItemEvent val)
        {
            try
            {
                SAPbouiCOM.Form form = Global.SapApplication.Forms.Item(val.FormUID);
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.Item _itm_Ref1;
               
                  SAPbouiCOM.Item itm_mat;
                SAPbouiCOM.Columns oColumns;
                SAPbouiCOM.Column oColumn;
                UserFields Ousr = new UserFields();
                Ousr.CreateTables();
                if (val.FormTypeEx == "134")
                {
                  
                    SAPbouiCOM.Matrix MAT_Unit;
                    _itm_Ref1 = form.Items.Item("21");
                    itm_mat = form.Items.Add("mtx_Unit", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    itm_mat.Left = 20;
                    itm_mat.Top = _itm_Ref1.Top + 40;
                    itm_mat.Width = 240;
                    itm_mat.Height = 180;
                    itm_mat.FromPane = 40;
                    itm_mat.ToPane = 40;
                    MAT_Unit = (SAPbouiCOM.Matrix)itm_mat.Specific;
                    oColumns = MAT_Unit.Columns;
                    oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Width = 20;
                    oColumn.Editable = false;
                    SAPbouiCOM.Matrix _matUserFieldsP = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Unit").Specific;
                    oColumns = MAT_Unit.Columns;
                    oColumn = oColumns.Add("col_Unit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtUom = (SAPbouiCOM.Column)_matUserFieldsP.Columns.Item("col_Unit");
                    form.DataSources.UserDataSources.Add("Uom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtUom.DataBind.SetBound(true, "", "Uom");
                    oColumn.Description = "Unit";
                    oColumn.TitleObject.Caption = "UNIT";
                    oColumn.Width = 120; 
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;
                    oColumns = MAT_Unit.Columns;
                    oColumn = oColumns.Add("col_Select", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    SAPbouiCOM.Column _edtslct = (SAPbouiCOM.Column)_matUserFieldsP.Columns.Item("col_Select");
                    form.DataSources.UserDataSources.Add("Selct", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtslct.DataBind.SetBound(true, "", "Selct");
                    oColumn .Description ="col_Select";
                    oColumn.TitleObject.Caption = "SELECT";
                    oColumn.Width = 120;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;
                    oColumns = MAT_Unit.Columns;
                    oColumn = oColumns.Add("col_Credit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtCredit = (SAPbouiCOM.Column)_matUserFieldsP.Columns.Item("col_Credit");
                    form.DataSources.UserDataSources.Add("Credit", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtCredit.DataBind.SetBound(true, "", "Credit");
                    oColumn.Description = "col_Credit";
                    oColumn.TitleObject.Caption = "Credit Limit";
                    oColumn.Width = 120;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;
                    //-------New On 12-07-2012-----------------------//
                    oColumns = MAT_Unit.Columns;
                    oColumn = oColumns.Add("col_Deflt", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    SAPbouiCOM.Column _edtdefault = (SAPbouiCOM.Column)_matUserFieldsP.Columns.Item("col_Deflt");
                    form.DataSources.UserDataSources.Add("Default", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtdefault.DataBind.SetBound(true, "", "Default");
                    oColumn.Description = "col_Deflt";
                    oColumn.TitleObject.Caption = "Is Default";
                    oColumn.Width = 80;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;

                   
                    
                }
                form.Freeze(false);
            }
            catch (Exception ex)
            {
                
            }
        }
        #endregion

        #region addbuttonBrand
        internal void AddButtonBrand(SAPbouiCOM.ItemEvent val)
        {
            try
            {
                SAPbouiCOM.Form form = Global.SapApplication.Forms.Item(val.FormUID);
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.Item _itm_Ref1;

                SAPbouiCOM.Item itm_mat;
                SAPbouiCOM.Columns oColumns;
                SAPbouiCOM.Column oColumn;
                if (val.FormTypeEx == "134")
                {                  
                    //-----------------------For Brand on 04-06-2012

                    SAPbouiCOM.Matrix MAT_Brand;
                    _itm_Ref1 = form.Items.Item("75");
                    itm_mat = form.Items.Add("mtx_Brand", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    itm_mat.Left = 20;
                    itm_mat.Top = _itm_Ref1.Top + 40;
                    itm_mat.Width = 540;
                    itm_mat.Height = 180;
                    itm_mat.FromPane = 45;
                    itm_mat.ToPane = 45;
                    MAT_Brand = (SAPbouiCOM.Matrix)itm_mat.Specific;
                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Width = 20;
                    oColumn.Editable = false;
                    SAPbouiCOM.Matrix _matUserFieldsPBrand = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Brand").Specific;

                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("col_Unit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtUomBrand = (SAPbouiCOM.Column)_matUserFieldsPBrand.Columns.Item("col_Unit");
                    form.DataSources.UserDataSources.Add("BUnit", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtUomBrand.DataBind.SetBound(true, "", "BUnit");
                    oColumn.Description = "Unit";
                    oColumn.TitleObject.Caption = "UNIT";
                    oColumn.Width = 120;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;

                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("col_Select", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    SAPbouiCOM.Column _edtslctBrand = (SAPbouiCOM.Column)_matUserFieldsPBrand.Columns.Item("col_Select");
                    form.DataSources.UserDataSources.Add("BSelect", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtslctBrand.DataBind.SetBound(true, "", "BSelect");
                    oColumn.Description = "col_Select";
                    oColumn.TitleObject.Caption = "SELECT";
                    oColumn.Width = 80;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;

                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("col_Brand", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtBrand = (SAPbouiCOM.Column)_matUserFieldsPBrand.Columns.Item("col_Brand");
                    form.DataSources.UserDataSources.Add("Brand", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtBrand.DataBind.SetBound(true, "", "Brand");
                    oColumn.Description = "col_Brand";
                    oColumn.TitleObject.Caption = "BRAND";
                    oColumn.Width = 200;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;

                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("col_BrandN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtBrandName = (SAPbouiCOM.Column)_matUserFieldsPBrand.Columns.Item("col_BrandN");
                    form.DataSources.UserDataSources.Add("BrandName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    _edtBrandName.DataBind.SetBound(true, "", "BrandName");
                    oColumn.Description = "col_BrandN";
                    oColumn.TitleObject.Caption = "BRAND Name";
                    oColumn.Width = 200;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;

                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("col_TPairs", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtTarPairs = (SAPbouiCOM.Column)_matUserFieldsPBrand.Columns.Item("col_TPairs");
                    form.DataSources.UserDataSources.Add("TPairs", SAPbouiCOM.BoDataType.dt_PRICE, 30);
                    _edtTarPairs.DataBind.SetBound(true, "", "TPairs");
                    oColumn.Description = "col_TPairs";
                    oColumn.TitleObject.Caption = "Target Pairs";
                    oColumn.Width = 100;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;

                    oColumns = MAT_Brand.Columns;
                    oColumn = oColumns.Add("col_TValue", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    SAPbouiCOM.Column _edtTarValue= (SAPbouiCOM.Column)_matUserFieldsPBrand.Columns.Item("col_TValue");
                    form.DataSources.UserDataSources.Add("TValue", SAPbouiCOM.BoDataType.dt_PRICE, 30);
                    _edtTarValue.DataBind.SetBound(true, "", "TValue");
                    oColumn.Description = "col_TValue";
                    oColumn.TitleObject.Caption = "Target Value";
                    oColumn.Width = 100;
                    oColumn.Editable = true;
                    oColumn.DisplayDesc = true;


                }
                form.Freeze(false);
            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region fillcombo
        internal void Fillseries(SAPbouiCOM.Form form)
        {
          //  SAPbouiCOM.Form form = Global.SapApplication.Forms.Item(pVal.FormUID);
           
            try
            {

                int ini = 0;
      
                form.Freeze(true);
                SAPbouiCOM.Matrix matrx = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Unit").Specific;
                SAPbouiCOM.EditText _edtCode = (SAPbouiCOM.EditText)form.Items.Item("5").Specific;
                SAPbobsCOM.Recordset oRecordSet1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (matrx.RowCount == 0)
                {
                  

                  // matrx.AddRow(1, 0);

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string _str_Query = "select Code   from [@NOR_UNITMASTER] ";
                    oRecordSet.DoQuery(_str_Query);
                    while (!oRecordSet.EoF)
                    {
                        matrx.AddRow(1, 0);
                        matrx.GetLineData(matrx.RowCount);
                        form.DataSources.UserDataSources.Item("Uom").Value = oRecordSet.Fields.Item(0).Value.ToString();
                        form.DataSources.UserDataSources.Item("Selct").Value = "N";
                        form.DataSources.UserDataSources.Item("Default").Value = "N";
                        form.DataSources.UserDataSources.Item("Credit").Value = "0.00";
                        matrx.SetLineData(matrx.RowCount);

                        oRecordSet.MoveNext();
                    }
                }
                    if (_edtCode.Value.ToString().Trim() != "")
                    {
                        string _str_Query1 = "SELECT A.Code,B.U_status,B.U_CrediLmt ,Isnull(C.U_IsDefault,'N')U_IsDefault FROm [@NOR_UNITMASTER] A Left join [@NOR_UNITALLOC] B on  A.Code=B.U_unitcode Left join [@NOR_DELRCONSON] C on B.U_unitcode = C.U_UnitCode and B.U_CusCode = C.U_DealerCode WHERE B.U_cuscode='" + _edtCode.Value.ToString().Trim() + "'";
                        oRecordSet1.DoQuery(_str_Query1);
                        if (oRecordSet1.RecordCount > 0)
                        {
                            while (!oRecordSet1.EoF)
                            {
                                string Unit = oRecordSet1.Fields.Item("Code").Value.ToString();
                                for (int j = 1; j <= matrx.RowCount; j++)
                                {
                                    matrx.GetLineData(j);
                                    string Unit1 = form.DataSources.UserDataSources.Item("Uom").Value;
                                    if (Unit == Unit1)
                                    {
                                        form.DataSources.UserDataSources.Item("Credit").Value = oRecordSet1.Fields.Item("U_CrediLmt").Value.ToString();
                                       
                                        string strStatus = oRecordSet1.Fields.Item("U_status").Value.ToString();
                                        string strDefault = oRecordSet1.Fields.Item("U_IsDefault").Value.ToString();
                                        if (strStatus == "Y")
                                        {
                                            form.DataSources.UserDataSources.Item("Selct").Value = "Y";

                                        }
                                        else
                                        {
                                            form.DataSources.UserDataSources.Item("Selct").Value = "N";
                                        }
                                        if (strDefault == "Y")
                                        {
                                            form.DataSources.UserDataSources.Item("Default").Value = "Y";
                                        }
                                        else
                                        {
                                            form.DataSources.UserDataSources.Item("Default").Value = "N";
                                        }
                                        matrx.SetLineData(j);
                                    }
                                   
                                }
                                oRecordSet1.MoveNext();
                            }
                        }
                    }
                    form.Freeze(false);
            
            }
            catch (Exception ex)
            {
                form.Freeze(false);
                Global.SapApplication.MessageBox(ex.Message, 1, "ok", "", "");
            }
        }

        
        #endregion
        #region Delete Matrix
        internal void ClearMatrix(SAPbouiCOM.Form  form)
        {
          //  SAPbouiCOM.Form form = Global.SapApplication.Forms.ActiveForm;
            SAPbouiCOM.Matrix matrx = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Brand").Specific;
            while( matrx.RowCount>0)
            {
                matrx.DeleteRow(matrx.RowCount);
            }

        }
        #endregion
        #region Clear Matrix Unit
        internal void ClearMatrixUnit(SAPbouiCOM.Form form)
        {
            //SAPbouiCOM.Form form = Global.SapApplication.Forms.ActiveForm;
            SAPbouiCOM.Matrix matrx = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Unit").Specific;
            while (matrx.RowCount > 0)
            {
                matrx.DeleteRow(matrx.RowCount);
            }

        }
         #endregion

        #region fillBrand
        internal void FillseriesBrand( SAPbouiCOM.Form form)
        {
           // SAPbouiCOM.Form form = Global.SapApplication.Forms.ActiveForm;
            try
            {

                int ini = 0;
                form.Freeze(true);
                SAPbouiCOM.Matrix matrx = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Brand").Specific;
                SAPbouiCOM.Matrix matrxUnit = (SAPbouiCOM.Matrix)form.Items.Item("mtx_Unit").Specific;
                string _units = "";
                int k = matrxUnit.RowCount;
                while (k > 0)
                {
                    matrxUnit.GetLineData(k);

                    if (form.DataSources.UserDataSources.Item("Selct").Value == "Y")
                    {
                        if (_units == "")
                        {
                            _units = "'" + form.DataSources.UserDataSources.Item("Uom").Value.ToString() + "'";
                        }
                        else
                        {
                            _units = _units + "," + "'" + form.DataSources.UserDataSources.Item("Uom").Value.ToString() + "'";
                        }
                    }
                    k--;
                }
                SAPbouiCOM.EditText _edtCode = (SAPbouiCOM.EditText)form.Items.Item("5").Specific;
                ClearMatrix(form);
                if(_units != "")
                {
                    if (matrx.RowCount == 0)
                    {
                      //  matrx.AddRow(1, 0);

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset oRecordSet1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        //matrx.GetLineData(matrx.RowCount);
                        string _str_Query = "select U_Unit ,b.Code ,b.Name from [@NOR_UNIT_BRAND] a inner join [@BRAND] b on a.U_Brand = b.Code where U_Unit in (" + _units + " ) group by U_Unit,b.Name,b.Code";

                        oRecordSet.DoQuery(_str_Query);
                        while (!oRecordSet.EoF)
                        {
                            matrx.AddRow(1, 0);
                            matrx.GetLineData(matrx.RowCount);
                            form.DataSources.UserDataSources.Item("BUnit").Value = oRecordSet.Fields.Item(0).Value.ToString();
                            form.DataSources.UserDataSources.Item("Brand").Value = oRecordSet.Fields.Item(1).Value.ToString();
                            form.DataSources.UserDataSources.Item("BrandName").Value = oRecordSet.Fields.Item(2).Value.ToString();
                            form.DataSources.UserDataSources.Item("BSelect").Value = "N";
                            form.DataSources.UserDataSources.Item("TPairs").Value = "0.00";
                            form.DataSources.UserDataSources.Item("TValue").Value = "0.00";
                            matrx.SetLineData(matrx.RowCount);                    
                            oRecordSet.MoveNext();
                        }
                        if (_edtCode.Value.ToString().Trim() != "")
                        {
                            string _str_Query1 = "SELECT A.U_Unit, A.U_Brand,B.U_Status ,B.U_TarPairs,B.U_TarValue FROm [@NOR_UNIT_BRAND] A Left join [@NOR_BP_BRAND] B on  A.U_Unit=B.U_Unit and A.U_Brand = B.U_Brand WHERE B.U_CustCode ='" + _edtCode.Value.ToString().Trim() + "'";
                            oRecordSet1.DoQuery(_str_Query1);
                            if (oRecordSet1.RecordCount > 0)
                            {
                                while (!oRecordSet1.EoF)
                                {
                                    string Unit = oRecordSet1.Fields.Item("U_Unit").Value.ToString();
                                    string Brand = oRecordSet1.Fields.Item("U_Brand").Value.ToString();
                                    string TarPairs = oRecordSet1.Fields.Item("U_TarPairs").Value.ToString();
                                    string TarValue = oRecordSet1.Fields.Item("U_TarValue").Value.ToString();
                                    for (int j = 1; j <= matrx.RowCount; j++)
                                    {
                                        matrx.GetLineData(j);
                                        string Unit1 = form.DataSources.UserDataSources.Item("BUnit").Value.ToString();
                                        string Brand1 = form.DataSources.UserDataSources.Item("Brand").Value.ToString();
                                       
                                        if (Unit == Unit1 && Brand == Brand1)
                                        {
                                            string i = oRecordSet1.Fields.Item("U_Status").Value.ToString();
                                            if (i == "Y")
                                            {
                                                form.DataSources.UserDataSources.Item("BSelect").Value = "Y";
                                                form.DataSources.UserDataSources.Item("TPairs").Value = TarPairs;
                                                form.DataSources.UserDataSources.Item("TValue").Value = TarValue;
                                            }
                                            else
                                            {
                                                form.DataSources.UserDataSources.Item("BSelect").Value = "N";
                                            }
                                        }
                                        matrx.SetLineData(j);
                                    }
                                    oRecordSet1.MoveNext();
                                }
                            }
                        }

                    }
                }
                form.Freeze(false);
            }
            catch (Exception ex)
            {
                form.Freeze(false);
                Global.SapApplication.MessageBox(ex.Message, 1, "ok", "", "");
            }
        }
        #endregion
   
       
       
        #region Checkbox Click
        //internal void PurchsCheckClick(SAPbouiCOM.Form frmpurchse)
        //{
        //    SAPbouiCOM.DBDataSource dsRec = null;
        //    string elemPC = "";
        //    try
        //    {
        //        dsRec = frmpurchse.DataSources.DBDataSources.Add("@NOR_ITB1");
        //        dsRec = frmpurchse.DataSources.DBDataSources.Item("@NOR_ITB1");
        //        elemPC = dsRec.GetValue("", 0).Trim();

        //        if (elemPC.Equals("Y"))
        //        {

        //            dsRec.SetValue("1", 0, "P");

        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        Global.SapApplication.MessageBox(e.Message, 1, "OK", "", "");
        //    }
        //}
        
        #endregion

        #region ItemDisable
        //internal void ItemDisable(SAPbouiCOM.ItemEvent val)
        //{
        //    SAPbouiCOM.Form frm = Global.SapApplication.Forms.Item(val.FormUID);
           
        //    try
        //    {
        //        frm.Items.Item("prchs").Enabled = false;
        //        frm.Items.Item("sals").Enabled = false;
        //        frm.Items.Item("phctbl").Enabled = false;
        //        frm.Items.Item("saltbl").Enabled = false;
        //        frm.Items.Item("cmbuom").Enabled = false;
        //    }
        //    catch { }
        //}
        #endregion

        #region CHKInventryControll
        //internal void CHKInventryControll(SAPbouiCOM.ItemEvent val)
        //{
        //    SAPbouiCOM.Form frm = Global.SapApplication.Forms.Item(val.FormUID);
        //    SAPbouiCOM.CheckBox chkIvctr = (SAPbouiCOM.CheckBox)frm.Items.Item("ivctrl").Specific;
        //    SAPbouiCOM.ComboBox combo = (SAPbouiCOM.ComboBox)frm.Items.Item("cmbuom").Specific;
        //    SAPbouiCOM.Matrix _mtrp = (SAPbouiCOM.Matrix)frm.Items.Item("phctbl").Specific;
        //    SAPbouiCOM.Matrix _mtrxsal = (SAPbouiCOM.Matrix)frm.Items.Item("saltbl").Specific;

        //    try
        //    {
        //        if (frm.DataSources.UserDataSources.Item("ivctrl").Value == "Y")
        //        {

        //            if (frm.DataSources.UserDataSources.Item("chkPurc").Value == "Y")
        //            {
        //                frm.Items.Item("phctbl").Enabled = true;
                        
        //            }
        //            else if (frm.DataSources.UserDataSources.Item("chkPurc").Value == "N")
        //            {
        //                frm.Items.Item("phctbl").Enabled = false;
                       


        //            }
        //            if (frm.DataSources.UserDataSources.Item("chkSales").Value == "Y")
        //            {
        //                frm.Items.Item("saltbl").Enabled = true;
        //            }
        //            else if (frm.DataSources.UserDataSources.Item("chkSales").Value == "N")
        //            {
        //               frm.Items.Item("saltbl").Enabled = false;
        //              //_mtrxsal.Clear();
        //              // Fillseries();

        //            }

        //            frm.Items.Item("prchs").Enabled = true;
        //            frm.Items.Item("sals").Enabled = true;
        //            frm.Items.Item("cmbuom").Enabled = true;

        //           // frm.Items.Item("phctbl").Enabled = true;


        //           // frm.Items.Item("saltbl").Enabled = true;


        //        }
        //        else
        //        {
        //            frm.Items.Item("prchs").Enabled = false;
        //            frm.Items.Item("sals").Enabled = false;
        //            frm.Items.Item("cmbuom").Enabled = false;
        //            frm.Items.Item("phctbl").Enabled = false;
        //            frm.Items.Item("saltbl").Enabled = false;
                    
        //           // ClearCombo(combo,true);
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        //  Global.SapApplication.MessageBox(ex.Message, 1, "", "", "");
        //    }
        //}
        #endregion

        #region BP Code Generation
        public void BPCodeGeneration(SAPbouiCOM.ItemEvent pVal)
        {
                SAPbouiCOM.Form frm = Global.SapApplication.Forms.Item(pVal.FormUID);
             SAPbouiCOM.DBDataSource _dbds_ocrdcode = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
             string BPCode = "", BPId = "", Country = "", State = "",CustGrp="";
             BPId = _dbds_ocrdcode.GetValue("CardType", 0).ToString().Trim();

             if (BPId == "S" || BPId == "C")
             {
                 if (BPId == "S")
                 {
                     BPId = "V";
                 }
                 Country = _dbds_ocrdcode.GetValue("U_Country", 0).ToString().Trim();
                 State = _dbds_ocrdcode.GetValue("U_State", 0).ToString().Trim();
                 CustGrp = _dbds_ocrdcode.GetValue("GroupCode", 0).ToString().Trim();
                 if (BPId != "" && Country != "")
                 {

                     if (Country == "IN" & State != "")
                     {
                         BPCode = BPId + "-" + State + "-";
                     }
                     else if (Country != "IN")
                     {
                         BPCode = BPId + "-" + Country + "-";
                     }
                 }
                 string nextBPCode = GetNextItemCode(BPCode, Country,CustGrp);
                 if (nextBPCode != "")
                 {
                     if (nextBPCode == "0000")
                     {
                         nextBPCode = "0001";
                     }
                     BPCode = BPCode + nextBPCode;
                     // SAPbouiCOM.Form frmBP = Global.SapApplication.Forms.ActiveForm;
                     if (pVal.FormTypeEx == "134")
                     {
                         SAPbouiCOM.EditText txtCardCode = (SAPbouiCOM.EditText)frm.Items.Item("5").Specific;
                         txtCardCode.Value = BPCode;
                     }
                     else
                     {
                         //string frmUID = frm.DataSources.UserDataSources.Item("FormID").ValueEx;
                         SAPbouiCOM.Form frmBP = Global.SapApplication.Forms.GetFormByTypeAndCount(134, 1);
                         SAPbouiCOM.EditText txtCardCode = (SAPbouiCOM.EditText)frmBP.Items.Item("5").Specific;
                         txtCardCode.Value = BPCode;
                     }


                 }
                 else
                 {
                     SAPbouiCOM.Form frmBP = Global.SapApplication.Forms.GetFormByTypeAndCount(134, 1);
                     SAPbouiCOM.EditText txtCardCode = (SAPbouiCOM.EditText)frmBP.Items.Item("5").Specific;
                     txtCardCode.Value = "";
                 }
             }

        }
        #endregion
#region GetNextCode
        internal string GetNextItemCode(string strCode , string Country,string CustGroup)
        {
            try
            {
                if (strCode != "" & Country != "")
                {
                    string strQry = "SELECT  ISNULL(MAX(CAST (SUBSTRING(CardCode,LEN('" + strCode + "')+1,4) AS NUMERIC)),0)+1 FROM OCRD WHERE  SUBSTRING(OCRD.CardCode,0,LEN('" + strCode + "')+1)= '" + strCode + "'";
                    SAPbobsCOM.Recordset rsNextCode = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsNextCode.DoQuery(strQry);
                    int a = rsNextCode.RecordCount;
                    if (Country == "IN")
                    {
                        if (rsNextCode.EoF || rsNextCode.RecordCount == 0)
                            return "0001";
                        else
                            return rsNextCode.Fields.Item(0).Value.ToString().PadLeft(4, '0');
                    }
                    else
                    {
                        if (CustGroup == "131" &  strCode.Substring(0,1)=="C")
                        {

                            string _strCardType = strCode.Substring(0, 1);
                            if (_strCardType == "V")
                            {
                                _strCardType = "S";
                            }
                            strQry = "SELECT ISNULL( MAX(CAST(RIGHT(Isnull(cardcode,0),3) as NUMERIC )),0) +1 from OCRD where CardType ='" + _strCardType + "' and U_Country != 'IN'";
                            rsNextCode = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsNextCode.DoQuery(strQry);

                            if (rsNextCode.EoF || rsNextCode.RecordCount == 0)
                                return "0001";
                            else
                            {
                                string str = rsNextCode.Fields.Item(0).Value.ToString().PadLeft(3, '0');
                                return rsNextCode.Fields.Item(0).Value.ToString().PadLeft(3, '0');
                            }


                        }
                        else
                        {


                            string _strCardType = strCode.Substring(0, 1);
                            if (_strCardType == "V")
                            {
                                _strCardType = "S";
                            }
                            strQry = "SELECT ISNULL( MAX(CAST(RIGHT(Isnull(cardcode,0),4) as NUMERIC )),0) +1 from OCRD where CardType ='" + _strCardType + "' and U_Country != 'IN'";
                            rsNextCode = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsNextCode.DoQuery(strQry);

                            if (rsNextCode.EoF || rsNextCode.RecordCount == 0)
                                return "0001";
                            else
                            {
                                string str = rsNextCode.Fields.Item(0).Value.ToString().PadLeft(4, '0');
                                return rsNextCode.Fields.Item(0).Value.ToString().PadLeft(4, '0');
                            }
                        }
                       
                    }
                }
                else
                {
                    return "";
                }

            }
            catch (Exception ex)
            {
                Global.SapApplication.MessageBox(ex.Message, 1, "", "", "");
                return "";
            }
        }
#endregion
    
        
       
       

        #region addt to table Brand
        public void addtotableBrand(SAPbouiCOM.ItemEvent val)
        {
            SAPbouiCOM.Form frm = Global.SapApplication.Forms.ActiveForm;
            string code = "";
            string unit = "";
            string name = "";
            string status = "";
            SAPbouiCOM.DBDataSource _dbds_ocrdcode = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
            string _str_itemcode = _dbds_ocrdcode.GetValue("CardCode", 0).ToString().Trim();
            //SAPbouiCOM.DBDataSource _dbds_ocrdname = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
            string _str_itemname = _dbds_ocrdcode.GetValue("CardName", 0).ToString().Trim();

            SAPbobsCOM.Recordset rsDocEntry = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            SAPbouiCOM.Matrix _matUserField = (SAPbouiCOM.Matrix)frm.Items.Item("mtx_Brand").Specific;

            string col1 = "";
            string col2 = "";
            string col3 = "";
            string col4 = "";
            string col5 = "";


            for (int i = 1; i <= _matUserField.RowCount; i++)
            {
                _matUserField.GetLineData(i);
                SAPbouiCOM.CheckBox _chkSelect = (SAPbouiCOM.CheckBox)_matUserField.Columns.Item("col_Select").Cells.Item(i).Specific;

                if (frm.DataSources.UserDataSources.Item("Unit").Value != "")
                {
                    string strQryDoc = "Select MAX(isnull(cast(Code as float),0))+1 as Code from [@NOR_BP_BRAND]";
                    rsDocEntry.DoQuery(strQryDoc);
                    string strDoc = rsDocEntry.Fields.Item("Code").Value.ToString();
                    col1 = frm.DataSources.UserDataSources.Item("Unit").Value;
                    col2 = frm.DataSources.UserDataSources.Item("Selct").Value;
                    col3 = frm.DataSources.UserDataSources.Item("Brand").Value;
                    col4 = frm.DataSources.UserDataSources.Item("TPairs").Value;
                    col5 = frm.DataSources.UserDataSources.Item("TValue").Value;
                    //Credit
                    SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string _str_sql1 = @"insert into [@NOR_BP_BRAND] (Code,Name,[U_Unit],[U_Brand],[U_CustCode],[U_TarPairs],[U_TarValue]) values ('" + strDoc + "','" + strDoc + "','" + col1 + "','" + col3 + "', '" + _str_itemcode + "','" + col4 + "','" + col5 + "')";
                    rs_insert1.DoQuery(_str_sql1);
                }

            }
        }

        #endregion 
        #region addt to table
        public void addtotable(SAPbouiCOM.ItemEvent val)
        {
            try
            {
                SAPbouiCOM.Form frm = Global.SapApplication.Forms.ActiveForm;
                string code = "";
                string unit = "";
                string name = "";
                string status = "";
                SAPbouiCOM.DBDataSource _dbds_ocrdcode = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
                string _str_itemcode = _dbds_ocrdcode.GetValue("CardCode", 0).ToString().Trim();
                //SAPbouiCOM.DBDataSource _dbds_ocrdname = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
                string _str_itemname = _dbds_ocrdcode.GetValue("CardName", 0).ToString().Trim();

                SAPbobsCOM.Recordset rsDocEntry = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                SAPbouiCOM.Matrix _matUserField = (SAPbouiCOM.Matrix)frm.Items.Item("mtx_Unit").Specific;

                string col1 = "";
                string col2 = "";
                string col3 = "";
                string colDeflt = "";


                for (int i = 1; i <= _matUserField.RowCount; i++)
                {
                    _matUserField.GetLineData(i);
                   // SAPbouiCOM.CheckBox _chkSelect = (SAPbouiCOM.CheckBox)_matUserField.Columns.Item("col_Select").Cells.Item(i).Specific;

                    if (frm.DataSources.UserDataSources.Item("Uom").Value != "")
                    {
                        string strQryDoc = "Select MAX(isnull(cast(Code as float),0))+1 as Code from [@NOR_UNITALLOC]";
                        rsDocEntry.DoQuery(strQryDoc);
                        string strDoc = rsDocEntry.Fields.Item("Code").Value.ToString();
                        col1 = frm.DataSources.UserDataSources.Item("Uom").Value;
                        col2 = frm.DataSources.UserDataSources.Item("Selct").Value;
                        col3 = frm.DataSources.UserDataSources.Item("Credit").Value;
                        colDeflt = frm.DataSources.UserDataSources.Item("Default").Value;
                        //Credit
                        SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string _str_sql1 = @"insert into [@NOR_UNITALLOC] (Code,Name,[U_cuscode],[U_unitcode],[U_status],[U_CrediLmt]) values ('" + strDoc + "','" + strDoc + "','" + _str_itemcode + "','" + col1 + "','" + col2 + "','" + col3 + "')";
                        rs_insert1.DoQuery(_str_sql1);

                        SAPbobsCOM.Recordset rs_IsnChkDealer = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string strIns = "SELECT * FROM [@NOR_DELRCONSON] WHERE U_UnitCode='" + col1 + "' and U_DealerCode='" + _str_itemcode + "'";
                        rs_IsnChkDealer.DoQuery(strIns);
                        if (rs_IsnChkDealer.RecordCount > 0)
                        {
                            rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            _str_sql1 = @"Update [@NOR_DELRCONSON] SET  U_CreditLimit='" + col3 + "',U_IsDefault = '" + colDeflt + "'  WHERE U_UnitCode='" + col1 + "' and U_DealerCode='" + _str_itemcode + "'";
                            rs_IsnChkDealer.DoQuery(_str_sql1);

                        }
                        else
                        {
                            rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            _str_sql1 = @"insert into [@NOR_DELRCONSON] (Code,Name,[U_DealerCode],[U_DealerName],[U_UnitCode],[U_UnitName],[U_CreditLimit],[U_IsDefault]) values ((select MAX( CONVERT(INT, code))+1  from [@NOR_DELRCONSON] ),(select MAX( CONVERT(INT, code))+1  from [@NOR_DELRCONSON] ),'" + _str_itemcode + "',(select cardName from OCRD where cardcode ='" + _str_itemcode + "'),'" + col1 + "', (select name   from [@NOR_UNITMASTER] where code ='" + col1 + "'),'" + col3 + "','" + colDeflt + "')";
                            rs_IsnChkDealer.DoQuery(_str_sql1);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
            }
        }

        #endregion 
      
        #region UpdateTableBrand

        public void UpdateTableBrand(SAPbouiCOM.ItemEvent val)
        {
            SAPbouiCOM.Form frm = Global.SapApplication.Forms.ActiveForm;
            string code = "";
            string unit = "";
            string name = "";
            string status = "";
            SAPbouiCOM.DBDataSource _dbds_ocrdcode = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
            string _str_itemcode = _dbds_ocrdcode.GetValue("CardCode", 0).ToString().Trim();
            string _str_itemname = _dbds_ocrdcode.GetValue("CardName", 0).ToString().Trim();

            SAPbobsCOM.Recordset rsDocEntry = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset rsExist = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            SAPbouiCOM.Matrix _matUserField = (SAPbouiCOM.Matrix)frm.Items.Item("mtx_Brand").Specific;

            string col1 = "";
            string col2 = "";
            string col3 = "";
            string col4 = "";
            string col5 = "";


            for (int i = 1; i <= _matUserField.RowCount; i++)
            {
                _matUserField.GetLineData(i);
                SAPbouiCOM.CheckBox _chkSelect = (SAPbouiCOM.CheckBox)_matUserField.Columns.Item("col_Select").Cells.Item(i).Specific;
                if (_chkSelect.Checked == false)
                {
                    col1 = frm.DataSources.UserDataSources.Item("BUnit").Value;
                    col2 = frm.DataSources.UserDataSources.Item("Brand").Value;
                    col3 = frm.DataSources.UserDataSources.Item("BSelect").Value;
                    col4 = frm.DataSources.UserDataSources.Item("TPairs").Value;
                    col5 = frm.DataSources.UserDataSources.Item("TValue").Value;
                    string strExist = "SELECT U_Unit, U_Status,U_Brand FROM [@NOR_BP_BRAND] WHERE  U_status='Y' and U_Custcode='" + _str_itemcode + "' and U_Unit='" + col1 + "' and U_Brand='" + col2 + "'";
                    rsExist.DoQuery(strExist);
                    if (rsExist.RecordCount > 0)
                    {
                        SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string _str_sql1 = @"Update [@NOR_BP_BRAND] SET U_Status='" + col3 + "'  WHERE U_Unit='" + col1 + "' and U_Brand='" + col2 + "' and U_Custcode='" + _str_itemcode + "'";
                        rs_insert1.DoQuery(_str_sql1);

                    }
                }
                else if (_chkSelect.Checked == true)
                {
                    if (frm.DataSources.UserDataSources.Item("BUnit").Value != "")
                    {
                        string strQryDoc = "Select MAX(isnull(cast(Code as float),0))+1 as Code from [@NOR_BP_BRAND]";
                        rsDocEntry.DoQuery(strQryDoc);
                        string strDoc = rsDocEntry.Fields.Item("Code").Value.ToString();
                        col1 = frm.DataSources.UserDataSources.Item("BUnit").Value;
                        col2 = frm.DataSources.UserDataSources.Item("BSelect").Value;
                        col3 = frm.DataSources.UserDataSources.Item("Brand").Value;
                        col4 = frm.DataSources.UserDataSources.Item("TPairs").Value;
                        col5 = frm.DataSources.UserDataSources.Item("TValue").Value;
                        //Credit
                        SAPbobsCOM.Recordset rs_IsnChk = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string strIns = "SELECT * FROM [@NOR_BP_BRAND] WHERE U_Unit='" + col1 + "' and U_Brand='" + col3 +"' and U_Custcode='" + _str_itemcode + "'";
                        rs_IsnChk.DoQuery(strIns);
                        if (rs_IsnChk.RecordCount > 0)
                        {
                            SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string _str_sql1 = @"Update [@NOR_BP_BRAND] SET U_Status='" + col2 + "'  WHERE U_Unit='" + col1 + "' and U_Brand='" + col3 + "' and U_Custcode='" + _str_itemcode + "'";
                            rs_insert1.DoQuery(_str_sql1);
                        }
                        else
                        {
                            SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string _str_sql1 = @"insert into [@NOR_BP_BRAND] (Code,Name,[U_Custcode],[U_Unit],[U_Status],[U_Brand],[U_TarPairs],[U_TarValue]) values ('" + strDoc + "','" + strDoc + "','" + _str_itemcode + "','" + col1 + "','" + col2 + "','" + col3 + "','" + col4 + "','" + col5 + "')";
                            rs_insert1.DoQuery(_str_sql1);
                        }
                    }
                }
            }
        }


        #endregion
        #region UpdateTable
        
        public void UpdateTable(SAPbouiCOM.ItemEvent val)
        {
            try
            {
            SAPbouiCOM.Form frm = Global.SapApplication.Forms.ActiveForm;
            string code = "";
            string unit = "";
            string name = "";
            string status = "";
            SAPbouiCOM.DBDataSource _dbds_ocrdcode = (SAPbouiCOM.DBDataSource)frm.DataSources.DBDataSources.Item("OCRD");
            string _str_itemcode = _dbds_ocrdcode.GetValue("CardCode", 0).ToString().Trim();
            string _str_itemname = _dbds_ocrdcode.GetValue("CardName", 0).ToString().Trim();

            SAPbobsCOM.Recordset rsDocEntry = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset rsExist = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            SAPbouiCOM.Matrix _matUserField = (SAPbouiCOM.Matrix)frm.Items.Item("mtx_Unit").Specific;

            string col1 = "";
            string col2 = "";
            string col3 = "";
            string colDeflt = "";


            for (int i = 1; i <= _matUserField.RowCount; i++)
            {
                _matUserField.GetLineData(i);
              //  SAPbouiCOM.CheckBox _chkSelect = (SAPbouiCOM.CheckBox)_matUserField.Columns.Item("col_Select").Cells.Item(i).Specific;
                if (frm.DataSources.UserDataSources.Item("Selct").Value == "N")
                {
                    SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    col1 = frm.DataSources.UserDataSources.Item("Uom").Value;
                    col3 = frm.DataSources.UserDataSources.Item("Credit").Value;
                    string strExist = "SELECT U_unitcode, U_status FROM [@NOR_UNITALLOC] WHERE  U_status='Y' and U_cuscode='" + _str_itemcode + "' and U_unitcode='" + col1 + "'";
                    rsExist.DoQuery(strExist);
                    if (rsExist.RecordCount > 0)
                    {
                        col2 = "N";
                        string _str_sql1 = @"Update [@NOR_UNITALLOC] SET U_status='" + col2 + "' , U_CrediLmt='" + col3 + "' WHERE U_unitcode='" + col1 + "' and U_cuscode='" + _str_itemcode + "'";
                        rs_insert1.DoQuery(_str_sql1);

                    }
                    string _str_sql2 = @"Update [@NOR_UNITALLOC] SET U_CrediLmt='" + col3 + "' WHERE U_unitcode='" + col1 + "' and U_cuscode='" + _str_itemcode + "'";
                    rs_insert1.DoQuery(_str_sql2);
                }
                else if (frm.DataSources.UserDataSources.Item("Selct").Value == "Y")
                {
                    if (frm.DataSources.UserDataSources.Item("Uom").Value != "")
                    {
                        string strQryDoc = "Select MAX(isnull(cast(Code as float),0))+1 as Code from [@NOR_UNITALLOC]";
                        rsDocEntry.DoQuery(strQryDoc);
                        string strDoc = rsDocEntry.Fields.Item("Code").Value.ToString();
                        col1 = frm.DataSources.UserDataSources.Item("Uom").Value;
                        col2 = frm.DataSources.UserDataSources.Item("Selct").Value;
                        col3 = frm.DataSources.UserDataSources.Item("Credit").Value;
                        if (col3 == "")
                        {
                            col3 = "0";
                        }
                        colDeflt = frm.DataSources.UserDataSources.Item("Default").Value;
                        //Credit
                        SAPbobsCOM.Recordset rs_IsnChk = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset rs_IsnChkDealer = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                       

                        string strIns = "SELECT * FROM [@NOR_UNITALLOC] WHERE U_unitcode='" + col1 + "' and U_cuscode='" + _str_itemcode + "'";
                        rs_IsnChk.DoQuery(strIns);
                        if (rs_IsnChk.RecordCount > 0)
                        {
                            SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string _str_sql1 = @"Update [@NOR_UNITALLOC] SET U_status='" + col2 + "' , U_CrediLmt='" + col3 + "' WHERE U_unitcode='" + col1 + "' and U_cuscode='" + _str_itemcode + "'";
                            rs_insert1.DoQuery(_str_sql1);


                        }
                        else
                        {
                            SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string _str_sql1 = @"insert into [@NOR_UNITALLOC] (Code,Name,[U_cuscode],[U_unitcode],[U_status],[U_CrediLmt]) values ('" + strDoc + "','" + strDoc + "','" + _str_itemcode + "','" + col1 + "','" + col2 + "','" + col3 + "')";
                            rs_insert1.DoQuery(_str_sql1);
                        }
                        strIns = "SELECT * FROM [@NOR_DELRCONSON] WHERE U_UnitCode='" + col1 + "' and U_DealerCode='" + _str_itemcode + "'";
                        rs_IsnChkDealer.DoQuery(strIns);
                        if (rs_IsnChkDealer.RecordCount > 0)
                        {
                            SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string _str_sql1 = @"Update [@NOR_DELRCONSON] SET  U_CreditLimit='" + col3 + "',U_IsDefault = '" + colDeflt + "'  WHERE U_UnitCode='" + col1 + "' and U_DealerCode='" + _str_itemcode + "'";
                            rs_IsnChkDealer.DoQuery(_str_sql1);

                        }
                        else
                        {
                            SAPbobsCOM.Recordset rs_insert1 = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //string _str_sql1 = @"insert into [@NOR_DELRCONSON] (Code,Name,[U_DealerCode],[U_UnitCode],[U_CreditLimit],[U_IsDefault]) values ((select MAX( CONVERT(INT, code))+1  from [@NOR_DELRCONSON] ),(select MAX( CONVERT(INT, name))+1  from [@NOR_DELRCONSON] ),'" + _str_itemcode + "','" + col1 + "','" + col3 + "','"+ colDeflt +"')";
                            string _str_sql1 = @"insert into [@NOR_DELRCONSON] (Code,Name,[U_DealerCode],[U_DealerName],[U_UnitCode],[U_UnitName],[U_CreditLimit],[U_IsDefault]) values (isnull((select MAX( CONVERT(INT, code))+1  from [@NOR_DELRCONSON] ),0),isnull((select MAX( CONVERT(INT, code))+1  from [@NOR_DELRCONSON] ),0),'" + _str_itemcode + "',(select cardName from OCRD where cardcode ='" + _str_itemcode + "'),'" + col1 + "', (select name   from [@NOR_UNITMASTER] where code ='" + col1 + "'),'" + col3 + "','" + colDeflt + "')";

                            rs_IsnChkDealer.DoQuery(_str_sql1);
                        }
                    }
                }
            }
        }
            catch(Exception ex)
            {
            }
        }

        
        #endregion

     

        #region Combofill
        #region FillCombo
        #region UnitCombofill
        public void FillCombo(SAPbouiCOM.Form frmFill, SAPbouiCOM.ComboBox ddlFill, string tableName, string value, string description, string strWhere, bool DefineNew, bool Blank)
        {
            SAPbobsCOM.Recordset rsFill;
            string strQry;
            frmFill.Freeze(true);
            try
            {
                ClearCombo(ddlFill, false);
                strQry = "SELECT DISTINCT " + value + ", " + description + " FROM " + tableName + " " + strWhere;
                rsFill = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsFill.DoQuery(strQry);
                if (Blank)
                {
                    ddlFill.ValidValues.Add("-1", "   ");
                }
                while (!rsFill.EoF)
                {
                    ddlFill.ValidValues.Add(rsFill.Fields.Item(value).Value.ToString(), rsFill.Fields.Item(description).Value.ToString());
                    rsFill.MoveNext();
                }

                if (DefineNew)
                {
                    ddlFill.ValidValues.Add("-999", "Define New");
                }
            }
            catch (Exception ex)
            {
                frmFill.Freeze(false);
                Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
            frmFill.Freeze(false);
        }


         #endregion

         #endregion


        #region ClearCombo
        public void ClearCombo(SAPbouiCOM.ComboBox ddlClear, bool blnBlankNeeded)
        {
            try
            {
                for (int i = ddlClear.ValidValues.Count - 1; i >= 0; i--)
                {
                    ddlClear.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                if (blnBlankNeeded)
                {
                    ddlClear.ValidValues.Add("-1", "All");
                    ddlClear.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                //Global.SapApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }


        #endregion
        #endregion

        public void UpdateDealerConsolidation()
        {

        }
    } 
}
