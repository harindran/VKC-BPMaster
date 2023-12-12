using System;
using System.Collections.Generic;
using System.Text;

namespace Project_1
{
    public class UserFields
    {

        #region Singleton

        private static UserFields instance;

        public static UserFields Instance
        {
            get
            {
                if (instance == null) instance = new UserFields();

                return instance;
            }
        }

        #endregion

        #region Update

        internal void UpdateDatabase()
        {
            try
            {
                CreateTables();
               // CreateUDO();
               // UpdateProcedure();
            }
            catch { }
        }

        #endregion

        #region User Tables and Fields

        internal void CreateTables()
        {
            try
            {
                

                //CreateUserFields("OCRD", "UnitCode", "purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "");
                //CreateUserFields("OCRD", "Status", "inventory controlled", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, "");
                //CreateUserFields("OITB", "ItemId", "Item ID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 4, "");


                CreateTable("NOR_UNITALLOC", "Unit Allocation");
                CreateUserFields("NOR_UNITALLOC", "cuscode", "customer code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "");
                CreateUserFields("NOR_UNITALLOC", "unitcode", "unit of measurement", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "");
                CreateUserFields("NOR_UNITALLOC", "status", "selected status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, "");
                CreateUserFields("NOR_UNITALLOC", "CrediLmt", "Credit Limit", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, "");
                //CreateUserFields("NOR_ITB1", "cofact", "conversion factor", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 20, "");
                //CreateUserFields("NOR_ITB1", "type", "type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1,"P|PURCHASE&S|SALES","");
                CreateTable("NOR_UNIT_BRAND", "Brand Unitwise");
                CreateUserFields("NOR_UNIT_BRAND", "Unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "NOR_UNITMASTER", "");
                CreateUserFields("NOR_UNIT_BRAND", "Brand", "Brand", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "BRAND", "");

                CreateTable("NOR_BP_BRAND", "Business Partner Brands");
                CreateUserFields("NOR_BP_BRAND", "Unit", "Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "NOR_UNITMASTER", "");
                CreateUserFields("NOR_BP_BRAND", "Brand", "Brand", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "BRAND", "");
                CreateUserFields("NOR_BP_BRAND", "CustCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "");
                CreateUserFields("NOR_BP_BRAND", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
              // Added on 21-06-2012----------------------------------------//

               // CreateUserFields("OCRD", "ID", "Business Partner ID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, "", "C|Customer & V|Vendor");
                CreateUserFields("OCRD", "Country", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, "");
                CreateUserFields("OCRD", "State", "State", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 3, "");

            }
            catch { }
        }

        #endregion



        internal void CreateTable(string tablename, string desc)
        {
            int lRetCode = 0;
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            oUserTablesMD.TableName = tablename;
            oUserTablesMD.TableDescription = desc;
            oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_NoObject;
            lRetCode = oUserTablesMD.Add();
            string err = Global.SapCompany.GetLastErrorDescription();
           // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);

        }

        internal void CreateMasterTable(string tablename, string desc)
        {
            int lRetCode = 0;
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            oUserTablesMD.TableName = tablename;
            oUserTablesMD.TableDescription = desc;
            oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData;
            lRetCode = oUserTablesMD.Add();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);

        }
        #region
        internal void CreateUDODoc(string code, string name, string table, string tableLine)
        {
            string err;
            int lRetCode = 0;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;

            oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

            oUserObjectMD.Code = code;

            oUserObjectMD.Name = name;

            oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
            oUserObjectMD.TableName = table;

            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;

            oUserObjectMD.FindColumns.ColumnAlias = "DocEntry";
            oUserObjectMD.FindColumns.Add();
            oUserObjectMD.FindColumns.ColumnAlias = "DocNum";
            oUserObjectMD.FindColumns.Add();

            if (tableLine != "")
            {
                oUserObjectMD.ChildTables.TableName = tableLine;

                oUserObjectMD.ChildTables.Add();
            }
            lRetCode = oUserObjectMD.Add();
            err = Global.SapCompany.GetLastErrorDescription();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
        }
        internal void CreateUDOMaster(string code, string name, string table, string tableLine)
        {
            int lRetCode = 0;
            string err = "";
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;

            oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

            oUserObjectMD.Code = code;

            oUserObjectMD.Name = name;

            oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
            oUserObjectMD.TableName = table;

            //oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

            //oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;

            //oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;

            //oUserObjectMD.FindColumns.ColumnAlias = "Code";
            //oUserObjectMD.FindColumns.Add();
            //oUserObjectMD.FindColumns.ColumnAlias = "Name";
            //oUserObjectMD.FindColumns.Add();

            if (tableLine != "")
            {

                oUserObjectMD.ChildTables.TableName = tableLine;
                oUserObjectMD.ChildTables.Add();
            }


            lRetCode = oUserObjectMD.Add();
            err = Global.SapCompany.GetLastErrorDescription();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);

        }

        internal void CreateUDO(string Code, string Name, bool Doc, string HTable, bool LogTable, string LTable)
        {
            string err, _str_Field1 = "", _str_Field2 = "";
            int lRetCode = 0;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;

            oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

            oUserObjectMD.Code = Code;
            oUserObjectMD.Name = Name;

            if (Doc)
            {
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
                _str_Field1 = "DocEntry";
                _str_Field2 = "DocNum";
            }
            else
            {
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                _str_Field1 = "Code";
                _str_Field2 = "Name";
            }
            oUserObjectMD.TableName = HTable;

            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
            if (Doc)
            {
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
            }
            else
            {
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
            }
            if (LogTable)
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;


            oUserObjectMD.FindColumns.ColumnAlias = _str_Field1;
            oUserObjectMD.FindColumns.Add();
            oUserObjectMD.FindColumns.ColumnAlias = _str_Field2;
            oUserObjectMD.FindColumns.Add();

            string[] _str_arr = LTable.Split('|');
            if (_str_arr[0] != "")
            {
                for (int i = 0; i < _str_arr.Length; i++)
                {
                    oUserObjectMD.ChildTables.TableName = _str_arr[i];
                    oUserObjectMD.ChildTables.Add();
                }
            }
            lRetCode = oUserObjectMD.Add();
            err = Global.SapCompany.GetLastErrorDescription();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
        }
        #endregion
        #region
        internal void CreateMasterLines(string tablename, string desc)
        {
            int lRetCode = 0;
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            oUserTablesMD.TableName = tablename;
            oUserTablesMD.TableDescription = desc;
            oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterDataLines;
            lRetCode = oUserTablesMD.Add();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);

        }
        internal void CreateDocumentTable(string tablename, string desc)
        {
            int lRetCode = 0;
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            oUserTablesMD.TableName = tablename;
            oUserTablesMD.TableDescription = desc;
            oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_Document;
            lRetCode = oUserTablesMD.Add();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);

        }
        internal void CreateDocumentLines(string tablename, string desc)
        {
            try
            {
                int lRetCode = 0;
                SAPbobsCOM.UserTablesMD oUserTablesMD = null;

                oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

                oUserTablesMD.TableName = tablename;
                oUserTablesMD.TableDescription = desc;
                oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_DocumentLines;
                lRetCode = oUserTablesMD.Add();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
            }
            catch (Exception ex)
            {

            }
        }
        internal void CreateObjectTable(string tablename, string desc)
        {
            int lRetCode = 0;
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            oUserTablesMD.TableName = tablename;
            oUserTablesMD.TableDescription = desc;
            oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_NoObject;
            lRetCode = oUserTablesMD.Add();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);

        }
        #endregion
        #region
        internal void AddQCategory(string QCategory)
        {
            SAPbobsCOM.QueryCategories QC = ((SAPbobsCOM.QueryCategories)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories)));
            QC.Name = QCategory;
            QC.Permissions = "YYYYYYYYYYYYYYY";
            QC.Add();
        }
        internal void AddQuery(string QCategory, string QName, string Query)
        {
            SAPbobsCOM.UserQueries UQ = ((SAPbobsCOM.UserQueries)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)));
            SAPbobsCOM.Recordset _rs = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            _rs.DoQuery("select CategoryId from OQCN where CatName='" + QCategory + "'");
            UQ.QueryCategory = Convert.ToInt32(_rs.Fields.Item(0).Value.ToString()); ;
            UQ.QueryDescription = QName;
            UQ.Query = Query;
            UQ.Add();
        }
        internal void SetFormattedSearch(string FormType, string ItemUID, string ColUID, string QCategory, string QueryName, bool Refresh, string FieldID, bool FrcRefresh, bool ByField)
        {
            SAPbobsCOM.FormattedSearches FS = ((SAPbobsCOM.FormattedSearches)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches)));
            SAPbobsCOM.Recordset _rs = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            bool _bool = false;
            if (ColUID == "")
                ColUID = "-1";

            string _str_Select = "select IndexID from CSHS where FormID='" + FormType + "' and ItemId='" + ItemUID + "' and ColID='" + ColUID + "'";
            _rs.DoQuery("select IndexID from CSHS where FormID='" + FormType + "' and ItemId='" + ItemUID + "' and ColID='" + ColUID + "'");
            if (Convert.ToInt32(_rs.Fields.Item(0).Value.ToString()) > 0)
            {
                _bool = true;
                FS.GetByKey(Convert.ToInt32(_rs.Fields.Item(0).Value.ToString()));
            }
            _rs.DoQuery("select IntrnalKey from OUQR INNER JOIN OQCN ON OUQR.QCategory=OQCN.CategoryId WHERE QName='" + QueryName + "' AND CatName='" + QCategory + "'");
            FS.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
            FS.FormID = FormType;
            FS.ItemID = ItemUID;
            if (ColUID == "")
                ColUID = "-1";
            FS.ColumnID = ColUID;
            FS.QueryID = Convert.ToInt32(_rs.Fields.Item(0).Value.ToString()); ;

            if (Refresh)
                FS.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
            else
                FS.Refresh = SAPbobsCOM.BoYesNoEnum.tNO;

            FS.FieldID = FieldID;

            if (FrcRefresh)
                FS.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tYES;
            else
                FS.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO;

            if (ByField)
                FS.ByField = SAPbobsCOM.BoYesNoEnum.tYES;
            else
                FS.ByField = SAPbobsCOM.BoYesNoEnum.tNO;

            int lRetCode;
            if (_bool)
                lRetCode = FS.Update();
            else
                lRetCode = FS.Add();
        }
        #endregion
        internal void DeleteUserFields(string Table, string Field)
        {
            SAPbobsCOM.Recordset oSelect = (SAPbobsCOM.Recordset)Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oSelect.DoQuery("Select FieldID from CUFD where TableID='" + Table + "' and AliasID='" + Field + "'");

            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
            oUserFieldsMD.GetByKey(Table, Convert.ToInt32(oSelect.Fields.Item(0).Value.ToString()));
            oUserFieldsMD.Remove();
        }
        internal void CreateUserFields(string Table, string Title, string Description, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes Structure, int Size, string LinkTable)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));

            int lRetCode = 0;
            oUserFieldsMD.TableName = Table;
            oUserFieldsMD.Name = Title;
            oUserFieldsMD.Description = Description;
            oUserFieldsMD.Type = Type;
            if (Type != SAPbobsCOM.BoFieldTypes.db_Numeric)
                oUserFieldsMD.SubType = Structure;
            oUserFieldsMD.EditSize = Size;

            if (LinkTable != "")
                oUserFieldsMD.LinkedTable = LinkTable;

            lRetCode = oUserFieldsMD.Add();
            string err = Global.SapCompany.GetLastErrorDescription();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
        }

        internal void CreateUserFields(string Table, string Title, string Description, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes Structure, int Size, string LinkTable, string ValidValues)
        {
            try
            {
                SAPbobsCOM.UserFieldsMD oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Global.SapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                int lRetCode = 0;




                oUserFieldsMD.TableName = Table;
                oUserFieldsMD.Name = Title;
                oUserFieldsMD.Description = Description;
                oUserFieldsMD.Type = Type;
                if (Type != SAPbobsCOM.BoFieldTypes.db_Numeric)
                    oUserFieldsMD.SubType = Structure;
                oUserFieldsMD.EditSize = Size;
                if (LinkTable != "")
                    oUserFieldsMD.LinkedTable = LinkTable;

                if (ValidValues != "")
                {
                    string[] _str_Values = new string[2];
                    string[] _arr_ValidValues = ValidValues.Split('&');
                    for (int i = 0; i < _arr_ValidValues.Length; i++)
                    {
                        _str_Values = _arr_ValidValues[i].Split('|');

                        oUserFieldsMD.ValidValues.Value = _str_Values[0];
                        oUserFieldsMD.ValidValues.Description = _str_Values[1];
                        oUserFieldsMD.ValidValues.Add();

                    }
                }

                lRetCode = oUserFieldsMD.Add();
                string err = Global.SapCompany.GetLastErrorDescription();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            }
            catch
            { }
        }
    }
}