using CRM.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hi_Life_API_FixedAssets
{
    class Model
    {
        private IOrganizationService service = EnvironmentSetting.Service;

        List<Entity> fixedAssetsList;
        List<Entity> storeInformationList;

        List<OptionMetadata> new_29;
        List<OptionMetadata> new_15;
        List<OptionMetadata> new_2;
        List<OptionMetadata> new_13;
        List<OptionMetadata> new_9;
        List<OptionMetadata> new_11;
        List<OptionMetadata> new_7;

        String[] fieldArray = {"new_name","new_vsde","new_ewfew","new_29","new_fwegtd",
                                "new_6","new_17","new_vwse",null,"new_30",
                                "new_0","new_22","new_34","new_ewet","new_ewe",
                                "new_31","new_15","new_32","new_35","new_2",
                                null,null,"new_210",null,"new_ewfsa",
                                "new_13","new_9","new_cost","new_5","new_11",
                                "new_salvage","new_depreciation","new_7","new_depreciationaccumulation","new_presentvalue",
                                "new_18","new_19","new_3","new_21",null,
                                null,"overriddencreatedon",null,"new_14","new_totalremainingcost",
                                "new_originalstoreid","new_20","new_1","new_28","new_27",
                                "new_4","new_depreciationbase","new_25","new_16","new_yy",
                                "new_mm","new_dd","new_23","new_26","new_24",
                                "new_33"
                                };

        internal void doModel(List<ICell> list)
        {
            //EntityLogicalName
            fixedAssetsList = Lookup.RetrieveEntityAllRecord("new_fixedassets", new String[] { "new_name" });
            storeInformationList = Lookup.RetrieveEntityAllRecord("new_storeinformation", new String[] { "new_name" });

            new_7 = OptionSet.GetoptionsetText("new_fixedassets", "new_7");
            new_9 = OptionSet.GetoptionsetText("new_fixedassets", "new_9");
            new_11 = OptionSet.GetoptionsetText("new_fixedassets", "new_11");
            new_13 = OptionSet.GetoptionsetText("new_fixedassets", "new_13");
            new_2 = OptionSet.GetoptionsetText("new_fixedassets", "new_2");
            new_15 = OptionSet.GetoptionsetText("new_fixedassets", "new_15");
            new_29 = OptionSet.GetoptionsetText("new_fixedassets", "new_29");
        }

        internal Guid IsExist(String key)
        {
            try
            {
                var list = from e in fixedAssetsList where e["new_name"].ToString() == key select e.Id;
                if (list.Any())
                    return list.First();
                else
                    return Guid.Empty;
                //return Lookup.RetrieveEntityGuid("new_fixedassets", key, "new_name");
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "檢查資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
                return Guid.Empty;
            }
        }

        public TransactionStatus CreateForCRM(IRow row)
        {
            return AddAttributeForRecord(row, Guid.Empty);
        }
        public TransactionStatus UpdateForCRM(IRow row, Guid productId)
        {
            return AddAttributeForRecord(row, productId);
        }
        private TransactionStatus AddAttributeForRecord(IRow row, Guid entityId)
        {
            try
            {
                Entity entity = new Entity("new_fixedassets");

                ICell cell;
                Guid recordGuid;
                Int32 optionValue;
                for (int i = 0; i < 55; i++)
                {
                    cell = row.GetCell(i);
                    if (cell == null && fieldArray[i] != null)
                        entity[fieldArray[i]] = null;
                    else
                    {
                        if (i == 7)
                        {
                            /// CRM欄位名稱     縣市      new_vwse
                            /// CRM關聯實體     縣市      new_storeinformation
                            /// CRM關聯欄位     名稱      new_name
                            /// 
                            var temp = from e in storeInformationList where e["new_name"].ToString() == cell.ToString() select e.Id;
                            if (!temp.Any())
                            {
                                try
                                {
                                    Entity storeinformation = new Entity("new_storeinformation");
                                    storeinformation["new_name"] = cell.ToString();
                                    storeinformation["new_storename"] = row.GetCell(i + 1).ToString();
                                    recordGuid = service.Create(storeinformation);
                                    storeInformationList = Lookup.RetrieveEntityAllRecord("new_storeinformation", new String[] { "new_name" });
                                }
                                catch
                                {
                                    EnvironmentSetting.ErrorMsg = "CRM 建立門市資訊失敗";
                                    Console.WriteLine(EnvironmentSetting.ErrorMsg);
                                    return TransactionStatus.Fail;
                                }
                            }
                            else
                                recordGuid = temp.First();
                            entity[fieldArray[i]] = new EntityReference("new_storeinformation", recordGuid);
                        }
                        else if (i == 45)
                        {
                            ///// CRM欄位名稱     門管部      new_originalstoreid
                            ///// CRM關聯實體     門管別      new_storeinformation
                            ///// CRM關聯欄位     門管部      new_name
                            ///// 
                            //var temp = from e in businessUnitList where e["new_name"].ToString() == cell.NumericCellValue.ToString() select e.Id;
                            //recordGuid = temp.First();
                            //if (recordGuid == Guid.Empty)
                            //{
                            //    EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                            //    EnvironmentSetting.ErrorMsg += "\tCRM實體 : new_storeinformation\n";
                            //    Console.WriteLine(EnvironmentSetting.ErrorMsg);
                            //    return TransactionStatus.Fail;
                            //}
                            //entity[fieldArray[i]] = new EntityReference("new_storeinformation", recordGuid);
                        }
                        else if (i == 8 || i == 20 || i == 21 || i == 23 || i == 39 || i == 40 || i == 42)
                        {

                        }
                        else if (i == 6 || i == 24 || i == 41 || i == 43 || i == 57 || i == 59 || i == 60)
                        {
                            //if (cell.NumericCellValue == 0)
                            //    entity[fieldArray[i]] = null;
                            //else
                            //    entity[fieldArray[i]] = cell.DateCellValue;
                        }
                        else if (i == 3)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_29, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 16)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_15, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 19)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_2, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 25)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_13, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 26)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_9, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 29)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_11, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 32)
                        {
                            optionValue = OptionSet.getOptionSetValue(new_7, cell.ToString());
                            if (optionValue == -1)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = new OptionSetValue(optionValue);
                        }
                        else if (i == 4 || i == 27 || i == 30 || i == 31 || i == 33 || i == 44 || i == 51)
                        {
                            entity[fieldArray[i]] = new Money(Convert.ToDecimal(cell.ToString()));
                        }
                        else if (i == 18 || i == 28)
                        {
                            entity[fieldArray[i]] = Convert.ToDecimal(cell.ToString());
                        }
                        else if (i == 10 || i == 35 || i == 36 || i == 54 || i == 55 || i == 56)
                        {
                            entity[fieldArray[i]] = Convert.ToInt32(cell.ToString());
                        }
                        else
                        {
                            entity[fieldArray[i]] = cell.ToString();
                        }
                    }
                }
                try
                {
                    if (entityId == Guid.Empty)
                        service.Create(entity);
                    else
                    {
                        entity["new_fixedassetsid"] = entityId;
                        service.Update(entity);
                    }
                    return TransactionStatus.Success;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("資產編號 : " + row.GetCell(0).ToString());
                    Console.WriteLine(ex.Message);
                    EnvironmentSetting.ErrorMsg = ex.Message;
                    return TransactionStatus.Fail;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(row.GetCell(0).ToString());
                Console.WriteLine("欄位讀取錯誤");
                Console.WriteLine(ex.Message);
                EnvironmentSetting.ErrorMsg = "欄位讀取錯誤\n" + ex.Message;
                return TransactionStatus.Fail;
            }
        }
    }
}
