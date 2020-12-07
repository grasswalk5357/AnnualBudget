using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTG : ERP_Common
    {
        //private string rdp_id = "";
        public ANBTG(string rdp_id) {
            this.Tg004 = rdp_id;
            CreateBodyItem(rdp_id);
        }
        public ANBTG() {

            //CreateBodyItem(rdp_id);
        }

        

        private decimal tg000 = 0;  // 流水號
        private string tg001 = "";  // 單別代號
        private string tg002 = "";  // 部門代號
        private string tg003 = "";  // 年度

        private string tg004 = "";  // 專案序號
        private string tg005 = "";  // 機型
        private string tg006 = "";  // 客戶
        private string tg007 = "";  // 預計開案
        private string tg008 = "";  // TR預排日
        private string tg009 = "";  // 預計結案
        private string tg010 = "";  // 量產地
        private string tg011 = "";  // 類型
        private decimal tg012 = 0;  // 耐用年限
        private string tg013 = "N";  // 是否標記刪除
        private List<Object> ANBTH_list = null;  // 儲存單身的List

        public decimal Tg000 { get => tg000; set => tg000 = value; }
        public string Tg001 { get => tg001; set => tg001 = value; }
        public string Tg002 { get => tg002; set => tg002 = value; }
        public string Tg003 { get => tg003; set => tg003 = value; }
        public string Tg004 { get => tg004; set => tg004 = value; }
        public string Tg005 { get => tg005; set => tg005 = value; }
        public string Tg006 { get => tg006; set => tg006 = value; }
        public string Tg007 { get => tg007; set => tg007 = value; }
        public string Tg008 { get => tg008; set => tg008 = value; }
        public string Tg009 { get => tg009; set => tg009 = value; }
        public string Tg010 { get => tg010; set => tg010 = value; }
        public string Tg011 { get => tg011; set => tg011 = value; }
        public decimal Tg012 { get => tg012; set => tg012 = value; }
        public string Tg013 { get => tg013; set => tg013 = value; }        
        public List<object> ANBTH_List { get => ANBTH_list; set => ANBTH_list = value; }

        private void CreateBodyItem(string rdp_id) {
            ANBTH rdp_c;
            List<Object> list = new List<Object>();
            for (int i = 0; i < 9; i++)
            {
                rdp_c = new ANBTH(rdp_id);
                list.Add(rdp_c);
            }
            this.ANBTH_list = list;
        }
    }
}
