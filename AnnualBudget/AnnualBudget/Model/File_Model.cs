using AnnualBudget.BOs;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnnualBudget.Model
{
    class File_Model
    {
        private static log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public File_Model() { }

        public static void WriteToExcel(string departNo, string departName, string userDeptNo, string year, string templateExcelName, System.Windows.Forms.DataGridView[] dgvs, Dictionary<string, Object> keyValues, List<object> summaryList) 
        {                        
            using (FileStream fsIn = new FileStream("Templates\\" + templateExcelName, FileMode.Open))
            {
                _log.Debug("檔案路徑：" + "Templates\\" + templateExcelName);
                //開啟Excel
                XSSFWorkbook workbook = new XSSFWorkbook(fsIn);

                fsIn.Close();

                //取得工作表：人力需求預算表
                _log.Debug("匯出人力需求預算表...");
                ISheet sheet0 = workbook.GetSheetAt(0);
                sheet0 = WriteANBTN_ToExcel(departNo, departName, year, sheet0, dgvs[0]);

                //取得工作表：人力需求預算表
                _log.Debug("匯出人力需求預算表...");
                ISheet sheet1 = workbook.GetSheetAt(1);
                sheet1 = WriteHumanBudgetToExcel(departNo, departName, year, sheet1, dgvs[1]);

                //取得工作表：教育訓練計畫表
                _log.Debug("匯出教育訓練計畫表...");
                ISheet sheet2 = workbook.GetSheetAt(2);
                sheet2 = WriteEmployeeTrainingToExcel(departNo, departName, year, sheet2, dgvs[2]);

                //取得工作表：出差計劃表
                _log.Debug("匯出出差計劃表...");
                ISheet sheet3 = workbook.GetSheetAt(3);
                sheet3 = WriteBizTripToExcel(departNo, departName, year, sheet3, dgvs[3]);

                //取得工作表：資本支出預算表
                _log.Debug("匯出資本支出預算表...");
                ISheet sheet4 = workbook.GetSheetAt(4);
                sheet4 = WriteCapExToExcel(departNo, departName, year, sheet4, dgvs[4]);

                if ("610".Equals(departNo) || "900".Equals(departNo))
                {
                    //取得工作表：研發專案彙總表、資本支出預算表-模具(折舊計算用免印)
                    _log.Debug("匯出研發專案彙總表、資本支出預算表-模具(折舊計算用免印)...");
                    ISheet sheet5 = workbook.GetSheetAt(5); // 研發專案彙總表
                    ISheet sheet6 = workbook.GetSheetAt(6); // 資本支出預算表-模具(折舊計算用免印)
                    sheet5 = WriteRD_PRJ_ToExcel(departNo, departName, year, sheet5, sheet6, keyValues);


                    _log.Debug("匯出預算總表...");
                    ISheet sheet7 = workbook.GetSheetAt(7); // 預算總表
                    sheet7 = WriteSummaryToExcel(departNo, departName, userDeptNo, year, sheet7, dgvs[7], summaryList);
                }
                else
                {
                    _log.Debug("匯出預算總表...");
                    ISheet sheet5 = workbook.GetSheetAt(5); // 預算總表
                    sheet5 = WriteSummaryToExcel(departNo, departName, userDeptNo, year, sheet5, dgvs[7], summaryList);
                }




                SaveFileDialog dlg = new SaveFileDialog();
                dlg.FileName = year + "年部門預算表-(" + departNo + ").xlsx"; // Default file name                
                dlg.Filter = "Excel活頁簿 | *.xlsx"; // Filter files by extension
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    //另存為Result.xls
                    using (FileStream fsOut = new FileStream(dlg.FileName, FileMode.Create))
                    {
                        workbook.Write(fsOut);
                        fsOut.Close();
                    }
                }
            }
        }

        private static ISheet WriteANBTN_ToExcel(string departNo, string departName, string year, ISheet sheet, System.Windows.Forms.DataGridView dgv)
        {
            int index = 3;  // 可填寫內容的起始列數
            try
            {
                
                sheet.GetRow(0).GetCell(1).SetCellValue(year + "年部門人力編制");       // Sheet名稱                
                sheet.GetRow(0).GetCell(5).SetCellValue(departNo);   // 部門代號
                sheet.GetRow(0).GetCell(6).SetCellValue(departName); // 部門名稱
                sheet.GetRow(0).GetCell(8).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (dgv.Rows.Count >= 2)
                {
                    for (int i = 0; i < dgv.Rows.Count - 1; i++)
                    {                           
                        if (dgv.Rows[i].Cells["dgv_Org_cln_Dept"].Value != null)                        
                            sheet.GetRow(index).GetCell(0).SetCellValue(dgv.Rows[i].Cells["dgv_Org_cln_Dept"].Value.ToString());        // 單位

                        if (dgv.Rows[i].Cells["dgv_Org_cln_JobTitle"].Value != null)                        
                            sheet.GetRow(index).GetCell(1).SetCellValue(dgv.Rows[i].Cells["dgv_Org_cln_JobTitle"].Value.ToString());    // 職位/職稱

                        if (dgv.Rows[i].Cells["dgv_Org_cln_Rank"].Value != null)                        
                            sheet.GetRow(index).GetCell(2).SetCellValue(dgv.Rows[i].Cells["dgv_Org_cln_Rank"].Value.ToString());        // 職等/職級
                        
                        sheet.GetRow(index).GetCell(3).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_Org_cln_ActNum"].Value));    // 實際人數


                        if (dgv.Rows[i].Cells["dgv_Org_cln_OnJob"].Value != null)
                            sheet.GetRow(index).GetCell(4).SetCellValue(dgv.Rows[i].Cells["dgv_Org_cln_OnJob"].Value.ToString());      // 現職者

                        sheet.GetRow(index).GetCell(5).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_Org_cln_EstNum"].Value));      // 計劃人數
                        sheet.GetRow(index).GetCell(6).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_Org_cln_DiffNum"].Value));      // 計畫與實際差額
                        sheet.GetRow(index).GetCell(7).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_Org_cbc_DiffMonth"].Value));    // 增減月份

                        if (dgv.Rows[i].Cells["dgv_Org_cln_Reason"].Value != null)
                            sheet.GetRow(index).GetCell(8).SetCellValue(dgv.Rows[i].Cells["dgv_Org_cln_Reason"].Value.ToString());   // 增減人力原因
                        
                        
                        index++;
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e)
            {
                _log.Debug("匯出組織編制表時發生錯誤：" + e.Message);
            }
            return sheet;
        }


        private static ISheet WriteHumanBudgetToExcel(string departNo, string departName, string year, ISheet sheet, System.Windows.Forms.DataGridView dgv) 
        {            
            int index = 4;  // 可填寫內容的起始列數
            try
            {
                sheet.GetRow(1).GetCell(0).SetCellValue(year + "年度人力需求預算表");     // Sheet名稱
                sheet.GetRow(2).GetCell(1).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(2).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(2).GetCell(10).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (dgv.Rows.Count >= 2)
                {
                    for (int i = 0; i < dgv.Rows.Count - 1; i++)
                    {
                        sheet.GetRow(index).GetCell(0).SetCellValue(i + 1);     // 序號
                        if (dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value != null)
                            sheet.GetRow(index).GetCell(1).SetCellValue(dgv.Rows[i].Cells["dgv_HBT_cln_Role"].Value.ToString());         // 直接/間接人員
                        if (dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].Value != null)
                            sheet.GetRow(index).GetCell(2).SetCellValue(dgv.Rows[i].Cells["dgv_HBT_cln_JobTitle"].Value.ToString());         // 職稱
                        if (dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].Value != null)
                            sheet.GetRow(index).GetCell(3).SetCellValue(dgv.Rows[i].Cells["dgv_HBT_cln_Rank"].Value.ToString());         // 職等/職級
                        if (dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value != null)
                            sheet.GetRow(index).GetCell(4).SetCellValue(dgv.Rows[i].Cells["dgv_HBT_cln_JobContent"].Value.ToString());         // 工作職掌
                        sheet.GetRow(index).GetCell(7).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_HBT_cln_ActNum"].Value));      // 實際人數
                        sheet.GetRow(index).GetCell(8).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_HBT_cln_EstNum"].Value));      // 計劃人數
                        sheet.GetRow(index).GetCell(9).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_HBT_cln_StartMonth"].Value));      // 增減起始月份
                        sheet.GetRow(index).GetCell(10).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_HBT_cln_DiffNum"].Value));     // 增減人數
                        sheet.GetRow(index).GetCell(11).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_HBT_cln_Salary"].Value));     // 增減之人員每月薪資
                        sheet.GetRow(index).GetCell(12).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells["dgv_HBT_cln_TotalChangeSalary"].Value));     // 每月增減薪資
                        if (dgv.Rows[i].Cells[10].Value != null)
                            sheet.GetRow(index).GetCell(13).SetCellValue(dgv.Rows[i].Cells["dgv_HBT_cln_Reason"].Value.ToString());       //增減人力原因                    
                        index++;
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出人力需求預算表時發生錯誤：" + e.Message);
            }
            return sheet;
        }

        private static ISheet WriteEmployeeTrainingToExcel(string departNo, string departName, string year, ISheet sheet, System.Windows.Forms.DataGridView dgv) {
            
            int index = 5;  // 可填寫內容的起始列數

            try
            {
                sheet.GetRow(1).GetCell(0).SetCellValue(year + "年度教育訓練計劃表");     // Sheet名稱
                sheet.GetRow(2).GetCell(1).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(2).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(2).GetCell(8).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (dgv.Rows.Count >= 2)
                {
                    for (int i = 0; i < dgv.Rows.Count - 1; i++)
                    {
                        //sheet.GetRow(index).GetCell(0).SetCellValue(i + 1);     // 序號
                        if (dgv.Rows[i].Cells[0].Value != null)
                            sheet.GetRow(index).GetCell(0).SetCellValue(dgv.Rows[i].Cells[0].Value.ToString());         // 職務能力

                        if (dgv.Rows[i].Cells[1].Value != null)
                            sheet.GetRow(index).GetCell(1).SetCellValue(dgv.Rows[i].Cells[1].Value.ToString());         // 訓練重點

                        if (dgv.Rows[i].Cells[2].Value != null)
                            sheet.GetRow(index).GetCell(3).SetCellValue(dgv.Rows[i].Cells[2].Value.ToString());         // 對象


                        sheet.GetRow(index).GetCell(4).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[3].Value));         // 人數

                        if (dgv.Rows[i].Cells[4].Value != null)
                            sheet.GetRow(index).GetCell(5).SetCellValue(dgv.Rows[i].Cells[4].Value.ToString());      // 內/外訓
                        sheet.GetRow(index).GetCell(6).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[5].Value));      // 起始月
                        sheet.GetRow(index).GetCell(7).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[6].Value));      // 結束月
                        sheet.GetRow(index).GetCell(8).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[7].Value));     // 時數
                        sheet.GetRow(index).GetCell(9).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[8].Value));     // 每月費用
                        sheet.GetRow(index).GetCell(10).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[9].Value));     // 該計劃費用總額

                        if (dgv.Rows[i].Cells[10].Value != null)
                            sheet.GetRow(index).GetCell(11).SetCellValue(dgv.Rows[i].Cells[10].Value.ToString());       // 備註

                        index++;
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出教育訓練計劃表時發生錯誤：" + e.Message);
                MessageBox.Show("匯出教育訓練計劃表時發生錯誤！");
            }
            return sheet;
        }

        private static ISheet WriteBizTripToExcel(string departNo, string departName, string year, ISheet sheet, System.Windows.Forms.DataGridView dgv) {
            int Month_1_index = 7;  // 一月的起始行數
            int multiple = 12;      // (二月以後)每個月之間的間隔行數
            int index = 1;      // 可填寫內容的起始列數
            int month = 0;      // 月份
            int[] monthCount = new int[12];  // 儲存各個月份的項目個數

            try
            {

                sheet.GetRow(1).GetCell(0).SetCellValue(year + "年度出差計劃表");         // Sheet名稱
                sheet.GetRow(3).GetCell(1).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(3).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(3).GetCell(13).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (dgv.Rows.Count >= 2)
                {
                    for (int i = 0; i < dgv.Rows.Count - 1; i++)
                    {
                        if (Convert.ToInt32(dgv.Rows[i].Cells[0].Value) == 1)
                        {
                            index = Month_1_index - 1;
                            month = Convert.ToInt32(dgv.Rows[i].Cells[0].Value);
                            index = index + monthCount[month - 1];
                        }
                        else if (Convert.ToInt32(dgv.Rows[i].Cells[0].Value) > 1 && Convert.ToInt32(dgv.Rows[i].Cells[0].Value) <= 12)
                        {
                            index = (((Convert.ToInt32(dgv.Rows[i].Cells[0].Value) - 1) * multiple) + Month_1_index) - 1;
                            month = Convert.ToInt32(dgv.Rows[i].Cells[0].Value);
                            index = index + monthCount[month - 1];
                        }

                        if (dgv.Rows[i].Cells[1].Value != null)
                            sheet.GetRow(index).GetCell(1).SetCellValue(dgv.Rows[i].Cells[1].Value.ToString());         // 出差者

                        if (dgv.Rows[i].Cells[2].Value != null)
                            sheet.GetRow(index).GetCell(2).SetCellValue(dgv.Rows[i].Cells[2].Value.ToString());         // 出差地區


                        sheet.GetRow(index).GetCell(3).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[3].Value));      // 天數
                        sheet.GetRow(index).GetCell(4).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[4].Value));      // 機票款
                        sheet.GetRow(index).GetCell(5).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[5].Value));      // 住宿費
                        sheet.GetRow(index).GetCell(6).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[6].Value));      // 交通費
                        sheet.GetRow(index).GetCell(7).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[7].Value));      // 雜費
                        sheet.GetRow(index).GetCell(8).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[8].Value));      // 日支費
                        sheet.GetRow(index).GetCell(9).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[9].Value));      // 餐費
                        sheet.GetRow(index).GetCell(10).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[10].Value));    // 旅費小計
                        sheet.GetRow(index).GetCell(11).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[11].Value));    // 交際費
                        sheet.GetRow(index).GetCell(12).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[12].Value));    // 旅平險


                        if (dgv.Rows[i].Cells[13].Value != null)
                            sheet.GetRow(index).GetCell(13).SetCellValue(dgv.Rows[i].Cells[13].Value.ToString());       // 出差目的說明

                        monthCount[month - 1] = monthCount[month - 1] + 1;

                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出出差計劃表時發生錯誤：" + e.Message);
                MessageBox.Show("匯出出差計劃表時發生錯誤！");
            }
            return sheet;
        }

        private static ISheet WriteCapExToExcel(string departNo, string departName, string year, ISheet sheet, System.Windows.Forms.DataGridView dgv) {
            int index = 0;
            int A_Row_index = 7;
            int B_Row_index = 7;
            int column_index = 0;
            int A_column_start_index = 0;
            int B_column_start_index = 14;
            int RowCount = 0;
            int A_RowCount = 0;
            int B_RowCount = 0;

            try
            {
                sheet.GetRow(1).GetCell(0).SetCellValue(year + "年度資本支出預算表");     // Sheet名稱
                sheet.GetRow(3).GetCell(1).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(3).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(3).GetCell(25).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (dgv.Rows.Count >= 2)
                {
                    for (int i = 0; i < dgv.Rows.Count - 1; i++)
                    {
                        if (dgv.Rows[i].Cells[0].Value != null) { 
                            if ("固定資產".Equals(dgv.Rows[i].Cells[0].Value.ToString().Trim()))
                            {
                                index = A_Row_index;
                                A_Row_index++;

                                column_index = A_column_start_index;

                                A_RowCount++;
                                RowCount = A_RowCount;
                            }

                            else if ("電腦軟體與專利權等無形資產".Equals(dgv.Rows[i].Cells[0].Value.ToString().Trim()))
                            {
                                index = B_Row_index;
                                B_Row_index++;
                                column_index = B_column_start_index;

                                B_RowCount++;
                                RowCount = B_RowCount;
                            }


                            sheet.GetRow(index).GetCell(0 + column_index).SetCellValue(RowCount);     // 項次

                            if (dgv.Rows[i].Cells[1].Value != null)
                                sheet.GetRow(index).GetCell(1 + column_index).SetCellValue(dgv.Rows[i].Cells[1].Value.ToString());         // 設備名稱

                            if (dgv.Rows[i].Cells[2].Value != null)
                                sheet.GetRow(index).GetCell(2 + column_index).SetCellValue(dgv.Rows[i].Cells[2].Value.ToString());         // 規格

                            sheet.GetRow(index).GetCell(3 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[3].Value));      // 數量
                            sheet.GetRow(index).GetCell(4 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[4].Value));      // 單價
                            sheet.GetRow(index).GetCell(5 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[5].Value));      // 總價金額
                            sheet.GetRow(index).GetCell(6 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[6].Value));      // 耐用年限
                            sheet.GetRow(index).GetCell(7 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[7].Value));      // 預定取得日(年)
                            sheet.GetRow(index).GetCell(8 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[8].Value));      // 預定取得日(月)
                            sheet.GetRow(index).GetCell(9 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[9].Value));      // 預定取得日(日)
                            sheet.GetRow(index).GetCell(10 + column_index).SetCellValue(Convert.ToDouble(dgv.Rows[i].Cells[10].Value));    // 年月折舊

                            if (dgv.Rows[i].Cells[11].Value != null)
                                sheet.GetRow(index).GetCell(11 + column_index).SetCellValue(dgv.Rows[i].Cells[11].Value.ToString());       // 增設(改善)目的

                            if (dgv.Rows[i].Cells[12].Value != null)
                                sheet.GetRow(index).GetCell(12 + column_index).SetCellValue(dgv.Rows[i].Cells[12].Value.ToString());       // 預期效益評估

                            if (dgv.Rows[i].Cells[13].Value != null)
                                sheet.GetRow(index).GetCell(13 + column_index).SetCellValue(dgv.Rows[i].Cells[13].Value.ToString());       // 交易對象
                        }
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出資本支出預算表時發生錯誤：" + e.Message);
                MessageBox.Show("匯出資本支出預算表時發生錯誤！");
            }
            return sheet;
        }

        private static ISheet WriteRD_PRJ_ToExcel(string departNo, string departName, string year, ISheet sheet, ISheet sheet_CapEx, Dictionary<string, Object> keyValues)
        {
            int item_IncreNum = 10;     // 每一個一樣名稱的項目，各間隔幾個row            
            int Start_Row_index = 6;    // 資料從第幾個ROW開始

            try
            {

                sheet.GetRow(1).GetCell(0).SetCellValue(year + "年度研發專案彙總表");     // Sheet名稱
                sheet.GetRow(3).GetCell(1).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(3).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(3).GetCell(17).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (keyValues != null && keyValues.Count >= 1)
                {
                    for (int i = 0; i < keyValues.Count; i++)
                    {
                        Object obj = keyValues.ElementAt(i).Value;
                        ANBTG anbtg = (ANBTG)obj;
                        //RDP rdp = (RDP)obj;
                        //rdp_id = rdp.RDP_Id;
                        //rdp = (RDP)keyValues[rdp_id];

                        sheet.GetRow(Start_Row_index).GetCell(2).SetCellValue(anbtg.Tg005);         // 機型
                        sheet.GetRow(Start_Row_index + 1).GetCell(2).SetCellValue(anbtg.Tg006);     // 客戶
                        sheet.GetRow(Start_Row_index + 2).GetCell(2).SetCellValue(anbtg.Tg007);     // 預計開案
                        sheet.GetRow(Start_Row_index + 3).GetCell(2).SetCellValue(anbtg.Tg008);     // TR預排日
                        sheet.GetRow(Start_Row_index + 4).GetCell(2).SetCellValue(anbtg.Tg009);     // 預計結案
                        sheet.GetRow(Start_Row_index + 5).GetCell(2).SetCellValue(anbtg.Tg010);     // 量產地
                        sheet.GetRow(Start_Row_index + 6).GetCell(2).SetCellValue(anbtg.Tg011);     // 類型
                        sheet.GetRow(Start_Row_index + 7).GetCell(2).SetCellValue(Convert.ToDouble(anbtg.Tg012));     // 耐用年限


                        for (int x = 0; x < anbtg.ANBTH_List.Count; x++) // 一筆單頭裡面的單身資料筆數，正常來說應該會有9筆
                        {
                            ANBTH anbth = (ANBTH)anbtg.ANBTH_List[x];

                            for (int y = 0; y < 12; y++)    // 1-12月份
                            {
                                sheet.GetRow(x + Start_Row_index).GetCell(y + 5).SetCellValue(Convert.ToDouble(anbth.MonthlyData[y]));
                            }


                            if ("8".Equals(anbth.Th005))    // 第8個項目，填入Excel頁籤表『資本支出預算表-模具(折舊計算用免印)』
                            {
                                WriteCapEx_Mould_ToExcel(departNo, departName, year, sheet_CapEx, anbth, i);
                            }
                        }
                        Start_Row_index += item_IncreNum;
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出研發專案彙總表時發生錯誤：" + e.Message);
                MessageBox.Show("匯出研發專案彙總表時發生錯誤！");
            }
            return sheet;
        }


        private static void WriteCapEx_Mould_ToExcel(string departNo, string departName, string year, ISheet sheet, ANBTH anbth, int dataNum)
        {
            
            int item_IncreColumnNum = 3;    // 每一個一樣名稱的項目，各間隔幾個column            
            int Start_Row_index = 6;        // 資料從第幾個ROW開始
            int Start_Column_index = 1;     // 資料從第幾個Column開始

            try
            {
                sheet.GetRow(1).GetCell(0).SetCellValue(year + "年度資本支出計畫表-模具");     // Sheet名稱
                sheet.GetRow(3).GetCell(1).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(3).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(3).GetCell(33).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (anbth != null)
                {
                    Start_Row_index += dataNum;     // 準備寫入的RowIndex要加上目前資料的筆數，做為每一筆row的間隔

                    for (int i = 0; i < 12; i++)    // 月份
                    {
                        sheet.GetRow(Start_Row_index).GetCell(Start_Column_index).SetCellValue(Convert.ToDouble(anbth.MonthlyData[i]));    // 每個專案的[金額]資料
                        sheet.GetRow(Start_Row_index).GetCell(Start_Column_index + 2).SetCellValue(Convert.ToDouble(anbth.Th024));     // 每個專案的[每月折舊]資料
                        Start_Column_index += item_IncreColumnNum;  // 每填一月份的資料要跳3個column
                    }

                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出「資本支出計畫表-模具」時發生錯誤：" + e.Message);
                MessageBox.Show("匯出「資本支出計畫表-模具」時發生錯誤！");
            }
            //return sheet;
        }


        public static ISheet WriteSummaryToExcel(string departNo, string departName, string userDeptNo, string year, ISheet sheet, System.Windows.Forms.DataGridView dgv, List<Object> myList)
        {            
            int key_column_1 = 2;   // 比對欄位_1的欄位位置
            int key_column_2 = 3;   // 比對欄位_2的欄位位置
            int key_column_3 = 4;   // 比對欄位_3的欄位位置

            string key_1 = "";      // 比對欄位_1的欄位內容
            string key_2 = "";      // 比對欄位_2的欄位內容
            string key_3 = "";      // 比對欄位_3的欄位內容

            int startColumn = 5;    // 開始塞值的欄位位置

            try
            {

                sheet.GetRow(0).GetCell(5).SetCellValue("明躍健康科技股份有限公司" + year + "年費用預算表");     // Sheet名稱
                sheet.GetRow(0).GetCell(2).SetCellValue(departNo);                      // 部門別
                sheet.GetRow(1).GetCell(2).SetCellValue(userDeptNo);                    // 最後課別
                                                                                        //sheet.GetRow(3).GetCell(2).SetCellValue(departName);                    // 部門名稱
                sheet.GetRow(1).GetCell(18).SetCellValue("填表日期：　" + DateTime.Today.ToString("yyyy/MM/dd"));      // 填表日期

                if (myList != null && myList.Count > 0)
                {
                    for (int i = 0; i < myList.Count - 1; i++)
                    {
                        ANBTI obj = (ANBTI)myList[i];
                        for (int j = 3; j < sheet.LastRowNum; j++)
                        {
                            key_1 = sheet.GetRow(j).GetCell(key_column_1).ToString();
                            key_2 = sheet.GetRow(j).GetCell(key_column_2).ToString();
                            key_3 = sheet.GetRow(j).GetCell(key_column_3).ToString();

                            
                            // 完全比對成功，塞值
                            if (obj.Ti007.Equals(key_1) && obj.Ti008.Equals(key_2) && obj.Ti009.Equals(key_3))
                            {
                                //if (obj.Ti007.Equals("6111"))
                                //    MessageBox.Show("Hold");
                                sheet.GetRow(j).GetCell(startColumn + 0).SetCellValue(Convert.ToDouble(obj.Ti010));  // 1月
                                sheet.GetRow(j).GetCell(startColumn + 1).SetCellValue(Convert.ToDouble(obj.Ti011));  // 2月
                                sheet.GetRow(j).GetCell(startColumn + 2).SetCellValue(Convert.ToDouble(obj.Ti012));  // 3月
                                sheet.GetRow(j).GetCell(startColumn + 3).SetCellValue(Convert.ToDouble(obj.Ti013));  // 4月
                                sheet.GetRow(j).GetCell(startColumn + 4).SetCellValue(Convert.ToDouble(obj.Ti014));  // 5月
                                sheet.GetRow(j).GetCell(startColumn + 5).SetCellValue(Convert.ToDouble(obj.Ti015));  // 6月
                                sheet.GetRow(j).GetCell(startColumn + 6).SetCellValue(Convert.ToDouble(obj.Ti016));  // 7月
                                sheet.GetRow(j).GetCell(startColumn + 7).SetCellValue(Convert.ToDouble(obj.Ti017));  // 8月
                                sheet.GetRow(j).GetCell(startColumn + 8).SetCellValue(Convert.ToDouble(obj.Ti018));  // 9月
                                sheet.GetRow(j).GetCell(startColumn + 9).SetCellValue(Convert.ToDouble(obj.Ti019));  // 10月
                                sheet.GetRow(j).GetCell(startColumn + 10).SetCellValue(Convert.ToDouble(obj.Ti020)); // 11月
                                sheet.GetRow(j).GetCell(startColumn + 11).SetCellValue(Convert.ToDouble(obj.Ti021)); // 12月
                                break;  // 塞值完後直接跳出迴圈，不再比對
                            }
                        }
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e) {
                _log.Debug("匯出「預算總表」時發生錯誤：" + e.Message);
                MessageBox.Show("匯出「預算總表」時發生錯誤！");
            }
            return sheet;
        }

        public static ISheet WriteSumByAccountsToExcel(string year, ISheet sheet, DataTable dt)
        {
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (sheet.GetRow(i + 1) == null)
                            sheet.CreateRow(i + 1);

                        sheet.GetRow(i + 1).CreateCell(0).SetCellValue(dt.Rows[i]["TI007"].ToString());
                        sheet.GetRow(i + 1).CreateCell(1).SetCellValue(dt.Rows[i]["TI008"].ToString());
                        sheet.GetRow(i + 1).CreateCell(2).SetCellValue(dt.Rows[i]["TI009"].ToString());
                        sheet.GetRow(i + 1).CreateCell(3).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI010"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(4).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI011"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(5).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI012"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(6).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI013"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(7).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI014"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(8).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI015"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(9).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI016"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(10).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI017"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(11).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI018"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(12).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI019"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(13).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI020"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(14).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI021"].ToString()));
                        sheet.GetRow(i + 1).CreateCell(15).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI022"].ToString()));
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e)
            {
                _log.Debug("執行「WriteSumByAccountsToExcel」方法時發生錯誤：" + e.Message);
                MessageBox.Show("執行「WriteSumByAccountsToExcel」方法時發生錯誤！");
            }
            return sheet;
        }

        public static ISheet WriteSumByDeptsToExcel(string year, ISheet sheet, DataTable dt)
        {
            string deptName = "";
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    DataTable deptsTable = Tmpl_Model.GetDept_Table();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (sheet.GetRow(i + 1) == null)
                            sheet.CreateRow(i + 1);

                        string select = "部門代號 = " + dt.Rows[i]["TI002"].ToString();
                        DataRow[] row = deptsTable.Select(select);

                        if (row != null && row.Length > 0)
                            deptName = row[0]["部門名稱"].ToString();


                        sheet.GetRow(i + 1).CreateCell(0).SetCellValue(dt.Rows[i]["TI002"].ToString());
                        sheet.GetRow(i + 1).CreateCell(1).SetCellValue(deptName);


                        sheet.GetRow(i + 1).CreateCell(2).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI010"]));
                        sheet.GetRow(i + 1).CreateCell(3).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI011"]));
                        sheet.GetRow(i + 1).CreateCell(4).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI012"]));
                        sheet.GetRow(i + 1).CreateCell(5).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI013"]));
                        sheet.GetRow(i + 1).CreateCell(6).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI014"]));
                        sheet.GetRow(i + 1).CreateCell(7).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI015"]));
                        sheet.GetRow(i + 1).CreateCell(8).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI016"]));
                        sheet.GetRow(i + 1).CreateCell(9).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI017"]));
                        sheet.GetRow(i + 1).CreateCell(10).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI018"]));
                        sheet.GetRow(i + 1).CreateCell(11).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI019"]));
                        sheet.GetRow(i + 1).CreateCell(12).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI020"]));
                        sheet.GetRow(i + 1).CreateCell(13).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI021"]));
                        sheet.GetRow(i + 1).CreateCell(14).SetCellValue(Convert.ToDouble(dt.Rows[i]["TI022"]));
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
            }
            catch (Exception e)
            {
                _log.Debug("執行「WriteSumByDeptsToExcel」方法時發生錯誤：" + e.Message);
                MessageBox.Show("執行「WriteSumByDeptsToExcel」方法時發生錯誤！");
            }
            return sheet;
        }


        private static void InsertRow(ISheet sheet, int insertRow)
        {
            sheet.CreateRow(insertRow);
            sheet.ShiftRows(insertRow, sheet.LastRowNum, 1);
            // 如果要插入的列數是0，複制下方列，反之，複制上方列
            sheet.CopyRow((insertRow == 0) ? insertRow + 1 : insertRow - 1, insertRow);
            // 清空插入列的值
            var row = sheet.GetRow(insertRow);
            for (int i = 0; i < row.LastCellNum; i++)
            {
                var cell = row.GetCell(i);
                if (cell != null)
                    cell.SetCellValue("");
            }
        }




        public static void WriteToExcel(List<Object> bom_List) {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook(); //建立活頁簿
            ISheet sheet = xssfworkbook.CreateSheet("sheet1"); //建立sheet

            ICellStyle headerStyle = xssfworkbook.CreateCellStyle();
            IFont headerfont = xssfworkbook.CreateFont();
            headerStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平置中
            headerStyle.VerticalAlignment = VerticalAlignment.Center; //垂直置中
            headerfont.FontName = "微軟正黑體";
            headerfont.FontHeightInPoints = 20;
            headerfont.IsBold = true;
            headerStyle.SetFont(headerfont);


            //新增標題列
            sheet.CreateRow(0); //需先用CreateRow建立,才可通过GetRow取得該欄位
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 2)); //合併1~2列及A~C欄儲存格
            sheet.GetRow(0).CreateCell(0).SetCellValue("昕力大學");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle; //套用樣式
            sheet.CreateRow(2).CreateCell(0).SetCellValue("學生編號");
            sheet.GetRow(2).CreateCell(1).SetCellValue("學生姓名");
            sheet.GetRow(2).CreateCell(2).SetCellValue("就讀科系");

            //填入資料
            int rowIndex = 3;
            for (int row = 0; row < bom_List.Count(); row++)
            {
                /*
                Material material = (Material)bom_List[row];
                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue(material.Id);
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue(material.Itm_name);
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(material.Itm_desc);
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(material.TreeLevel);
                rowIndex++;*/

            }

            FileStream myFile = new FileStream(@"D:\myTestExcel.xlsx", FileMode.Create);

            //var excelDatas = new MemoryStream();
            xssfworkbook.Write(myFile);
            
            myFile.Close();
            xssfworkbook.Close();
            //return File(excelDatas.ToArray(), "application/vnd.ms-excel", string.Format($"MyData.xlsx"));
        }

        /// <summary>
        /// 連線共享資料夾
        /// </summary>
        /// <param name="path">共享路徑</param>
        /// <param name="user">使用者名稱</param>
        /// <param name="pass">密碼</param>
        /// <returns></returns>
        public static void LinkFile(string path, string user, string pwd)
        {
            string cLinkUrl = @"Net Use " + path + " " + pwd + " /user:" + user;
            CallCmd(cLinkUrl);
        }

        /// <summary>
        /// 關閉所有共享連線
        /// </summary>
        public static void KillAllLink()
        {
            string cKillCmd = @"Net Use /delete * /yes";
            CallCmd(cKillCmd);
        }

        /// <summary>
        /// 關閉指定連線
        /// </summary>
        /// <param name="path">共享路徑</param>
        public static void KillLink(string path)
        {
            string cKillCmd = @"Net Use " + path + " /delete /yes";
            CallCmd(cKillCmd);
        }

        /// <summary> 
        /// 呼叫Cmd命令 
        /// </summary> 
        /// <param name="strCmd">命令列引數</param> 
        private static void CallCmd(string strCmd)
        {
            //呼叫cmd命令 
            Process myProcess = new Process();
            try
            {
                myProcess.StartInfo.FileName = "cmd.exe";
                myProcess.StartInfo.Arguments = "/c " + strCmd;
                myProcess.StartInfo.UseShellExecute = false;    //關閉Shell的使用 
                myProcess.StartInfo.RedirectStandardInput = true;  //重定向標準輸入 
                myProcess.StartInfo.RedirectStandardOutput = true; //重定向標準輸出 
                myProcess.StartInfo.RedirectStandardError = true;  //重定向錯誤輸出 
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.Start();
            }
            catch (Exception ex) {
                _log.Debug("登入PDM資料夾時發生錯誤：" + ex.Message);
            }   
            finally
            {
                myProcess.WaitForExit();
                if (myProcess != null)
                {
                    myProcess.Close();
                }
            }
        }

        public static IntPtr LogDictionary() {
            //File_Model.LinkFile(@"\\192.168.123.22\HPDM_Vault$\", "itadmin", "god357/");
            string MachineName = "192.168.123.22";
            string UserName = "itadmin";
            string Pw = "god357/";
            string IPath = String.Format(@"\\{0}\HPDM_Vault$", MachineName);
            const int LOGON32_PROVIDER_DEFAULT = 0;
            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
            IntPtr tokenHandle = new IntPtr(0);
            tokenHandle = IntPtr.Zero;

            bool returnValue = LogonUser(UserName, MachineName, Pw,
            LOGON32_LOGON_NEW_CREDENTIALS,
            LOGON32_PROVIDER_DEFAULT,
            ref tokenHandle);

            WindowsIdentity w = new WindowsIdentity(tokenHandle);
            w.Impersonate();
            /*if (false == returnValue)
            {
                return;
            }*/

            return tokenHandle;
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
            string lpszUsername,
            string lpszDomain,
            string lpszPassword,
            int dwLogonType,
            int dwLogonProvider,
            ref IntPtr phToken
        );

        //登出
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);
    }

    


    public class Win32Helper
    {
        internal const string IShellItem2Guid = "7E9FB0D3-919F-4307-AB2E-9B1860310C93";

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string path,
            // The following parameter is not used - binding context.
            IntPtr pbc,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItem shellItem);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DeleteObject(IntPtr hObject);

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        internal interface IShellItem
        {
            void BindToHandler(IntPtr pbc,
                [MarshalAs(UnmanagedType.LPStruct)]Guid bhid,
                [MarshalAs(UnmanagedType.LPStruct)]Guid riid,
                out IntPtr ppv);

            void GetParent(out IShellItem ppsi);
            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        };
        public enum ThumbnailOptions : uint
        {
            None = 0,
            ReturnOnlyIfCached = 1,
            ResizeThumbnail = 2,
            UseCurrentScale = 4
        }
        internal enum SIGDN : uint
        {
            NORMALDISPLAY = 0,
            PARENTRELATIVEPARSING = 0x80018001,
            PARENTRELATIVEFORADDRESSBAR = 0x8001c001,
            DESKTOPABSOLUTEPARSING = 0x80028000,
            PARENTRELATIVEEDITING = 0x80031001,
            DESKTOPABSOLUTEEDITING = 0x8004c000,
            FILESYSPATH = 0x80058000,
            URL = 0x80068000
        }

        internal enum HResult
        {
            Ok = 0x0000,
            False = 0x0001,
            InvalidArguments = unchecked((int)0x80070057),
            OutOfMemory = unchecked((int)0x8007000E),
            NoInterface = unchecked((int)0x80004002),
            Fail = unchecked((int)0x80004005),
            ElementNotFound = unchecked((int)0x80070490),
            TypeElementNotFound = unchecked((int)0x8002802B),
            NoObject = unchecked((int)0x800401E5),
            Win32ErrorCanceled = 1223,
            Canceled = unchecked((int)0x800704C7),
            ResourceInUse = unchecked((int)0x800700AA),
            AccessDenied = unchecked((int)0x80030005)
        }

        [ComImportAttribute()]
        [GuidAttribute("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellItemImageFactory
        {
            [PreserveSig]
            HResult GetImage(
            [In, MarshalAs(UnmanagedType.Struct)] NativeSize size,
            [In] ThumbnailOptions flags,
            [Out] out IntPtr phbm);
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct NativeSize
        {
            private int width;
            private int height;

            public int Width { set { width = value; } }
            public int Height { set { height = value; } }
        };

        [StructLayout(LayoutKind.Sequential)]
        public struct RGBQUAD
        {
            public byte rgbBlue;
            public byte rgbGreen;
            public byte rgbRed;
            public byte rgbReserved;
        }
    }
    public class ThumbnailHelper
    {

        #region instance
        private static object ooo = new object();
        private static ThumbnailHelper _ThumbnailHelper;
        string originalFileName = "";
        private ThumbnailHelper() { }
        public static ThumbnailHelper GetInstance()
        {
            if (_ThumbnailHelper == null)
            {
                lock (ooo)
                {
                    if (_ThumbnailHelper == null)
                        _ThumbnailHelper = new ThumbnailHelper();
                }
            }
            return _ThumbnailHelper;
        }
        #endregion

        #region public methods

        public string GetJPGThumbnail(string filename, int width = 500, int height = 500, Win32Helper.ThumbnailOptions options = Win32Helper.ThumbnailOptions.None)
        {
            if (!File.Exists(filename))
                return string.Empty;
            Bitmap bit = GetBitmapThumbnail(filename, width, height, options);
            originalFileName = Path.GetFileName(filename);
            if (bit == null)
                return string.Empty;
            return GetThumbnailFilePath(bit);
        }
        #endregion

        #region private methods
        private Bitmap GetBitmapThumbnail(string fileName, int width, int height, Win32Helper.ThumbnailOptions options = Win32Helper.ThumbnailOptions.None)
        {
            IntPtr hBitmap = GetHBitmap(System.IO.Path.GetFullPath(fileName), width, height, options);

            try
            {
                Bitmap bmp = Bitmap.FromHbitmap(hBitmap);

                if (Bitmap.GetPixelFormatSize(bmp.PixelFormat) < 32)
                    return bmp;

                return CreateAlphaBitmap(bmp, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            }
            finally
            {
                // delete HBitmap to avoid memory leaks
                Win32Helper.DeleteObject(hBitmap);
            }
        }
        private Bitmap CreateAlphaBitmap(Bitmap srcBitmap, System.Drawing.Imaging.PixelFormat targetPixelFormat)
        {
            Bitmap result = new Bitmap(srcBitmap.Width, srcBitmap.Height, targetPixelFormat);

            System.Drawing.Rectangle bmpBounds = new System.Drawing.Rectangle(0, 0, srcBitmap.Width, srcBitmap.Height);

            BitmapData srcData = srcBitmap.LockBits(bmpBounds, ImageLockMode.ReadOnly, srcBitmap.PixelFormat);

            bool isAlplaBitmap = false;

            try
            {
                for (int y = 0; y <= srcData.Height - 1; y++)
                {
                    for (int x = 0; x <= srcData.Width - 1; x++)
                    {
                        System.Drawing.Color pixelColor = System.Drawing.Color.FromArgb(
                            Marshal.ReadInt32(srcData.Scan0, (srcData.Stride * y) + (4 * x)));

                        if (pixelColor.A > 0 & pixelColor.A < 255)
                        {
                            isAlplaBitmap = true;
                        }

                        result.SetPixel(x, y, pixelColor);
                    }
                }
            }
            finally
            {
                srcBitmap.UnlockBits(srcData);
            }

            if (isAlplaBitmap)
            {
                return result;
            }
            else
            {
                return srcBitmap;
            }
        }

        private IntPtr GetHBitmap(string fileName, int width, int height, Win32Helper.ThumbnailOptions options)
        {
            Win32Helper.IShellItem nativeShellItem;
            Guid shellItem2Guid = new Guid(Win32Helper.IShellItem2Guid);
            int retCode = Win32Helper.SHCreateItemFromParsingName(fileName, IntPtr.Zero, ref shellItem2Guid, out nativeShellItem);

            if (retCode != 0)
                throw Marshal.GetExceptionForHR(retCode);

            Win32Helper.NativeSize nativeSize = new Win32Helper.NativeSize();
            nativeSize.Width = width;
            nativeSize.Height = height;

            IntPtr hBitmap;
            Win32Helper.HResult hr = ((Win32Helper.IShellItemImageFactory)nativeShellItem).GetImage(nativeSize, options, out hBitmap);

            Marshal.ReleaseComObject(nativeShellItem);

            if (hr == Win32Helper.HResult.Ok) return hBitmap;

            throw Marshal.GetExceptionForHR((int)hr);
        }

        /// <summary>
        /// 获取临时文件
        /// </summary>
        /// <returns></returns>
        private FileStream GetTemporaryFilePath(ref string filePath)
        {
            //string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName());
            string path = System.IO.Path.Combine("D:\\test\\", originalFileName);

            var index = path.IndexOf('.');
            string temp = path.Substring(0, index) + ".png";

            var fileStream = File.Create(temp);
            filePath = temp;
            return fileStream;
        }
        /// <summary>
        /// 参数
        /// </summary>
        /// <param name="mimeType"></param>
        /// <returns></returns>
        private ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }
        const int ExpectHeight = 200;
        const int ExpectWidth = 200;
        /// <summary>
        /// 得到图片缩放后的大小  图片大小过小不考虑缩放了
        /// </summary>
        /// <param name="originSize">原始大小</param>
        /// <returns>改变后大小</returns>    
        private System.Drawing.Size ScalePhoto(System.Drawing.Size originSize, bool can)
        {
            if (originSize.Height * originSize.Width < ExpectHeight * ExpectWidth)
            {
                can = false;
            }
            if (can)
            {
                bool isportrait = false;

                if (originSize.Width <= originSize.Height)
                {
                    isportrait = true;
                }

                if (isportrait)
                {
                    double ratio = (double)ExpectHeight / (double)originSize.Height;
                    return new System.Drawing.Size((int)(originSize.Width * ratio), (int)(originSize.Height * ratio));
                }
                else
                {
                    double ratio = (double)ExpectWidth / (double)originSize.Width;
                    return new System.Drawing.Size((int)(originSize.Width * ratio), (int)(originSize.Height * ratio));
                }
            }
            else
            {
                return new System.Drawing.Size(originSize.Width, originSize.Height);
            }

        }
        /// <summary>
        /// 获取缩略图文件
        /// </summary>
        /// <param name="BitmapIcon">缩略图</param>
        /// <returns>路径</returns>
        private string GetThumbnailFilePath(Bitmap bitmap)
        {
            var filePath = "";
            var fileStream = GetTemporaryFilePath(ref filePath);

            //bitmap.Save(filePath);

            ImageCodecInfo myImageCodecInfo = GetEncoderInfo("image/png"); //image code info 图形图像解码压缩
                                                                           //ImageCodecInfo myImageCodecInfo = GetEncoderInfo("image/jpeg"); //image code info 图形图像解码压缩
            System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
            EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 100L);
            EncoderParameters encoderParameters = new EncoderParameters { Param = new EncoderParameter[] { myEncoderParameter } };
            var sizeScaled = ScalePhoto(bitmap.Size, true);
            //去黑色背景
            var finalBitmap = ClearBlackBackground(bitmap);
            finalBitmap.Save(fileStream, myImageCodecInfo, encoderParameters);
            fileStream.Dispose();
            return filePath;
        }

        private Bitmap ClearBlackBackground(Bitmap originBitmap)
        {
            using (var tempBitmap = new Bitmap(originBitmap.Width, originBitmap.Height))
            {
                tempBitmap.SetResolution(originBitmap.HorizontalResolution, originBitmap.VerticalResolution);

                using (var g = Graphics.FromImage(tempBitmap))
                {
                    g.Clear(System.Drawing.Color.White);
                    g.DrawImageUnscaled(originBitmap, 0, 0);
                }
                return new Bitmap(tempBitmap);
            }
        }

        #endregion
    }
}
