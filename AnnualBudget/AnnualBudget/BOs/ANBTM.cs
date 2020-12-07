using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnnualBudget.BOs
{
    class ANBTM : ERP_Common
    {
        private decimal tm000 = 0;  //流水號
        private string tm001 = "";  //部門代號
        private string tm002 = "";  //部門名稱
        private decimal tm003 = 0;  //預算比率
        private decimal tm004 = 0;  //1月
        private decimal tm005 = 0;  //2月
        private decimal tm006 = 0;  //3月
        private decimal tm007 = 0;  //4月
        private decimal tm008 = 0;  //5月
        private decimal tm009 = 0;  //6月
        private decimal tm010 = 0;  //7月
        private decimal tm011 = 0;  //8月
        private decimal tm012 = 0;  //9月
        private decimal tm013 = 0;  //10月
        private decimal tm014 = 0;  //11月
        private decimal tm015 = 0;  //12月
        private decimal tm016 = 0;  //合計
        private string tm017 = "";  //費用類別代號
        private string tm018 = "";  //費用類別名稱
        

        public decimal Tm000 { get => tm000; set => tm000 = value; }
        public string Tm001 { get => tm001; set => tm001 = value; }
        public string Tm002 { get => tm002; set => tm002 = value; }
        public decimal Tm003 { get => tm003; set => tm003 = value; }
        public decimal Tm004 { get => tm004; set => tm004 = value; }
        public decimal Tm005 { get => tm005; set => tm005 = value; }
        public decimal Tm006 { get => tm006; set => tm006 = value; }
        public decimal Tm007 { get => tm007; set => tm007 = value; }
        public decimal Tm008 { get => tm008; set => tm008 = value; }
        public decimal Tm009 { get => tm009; set => tm009 = value; }
        public decimal Tm010 { get => tm010; set => tm010 = value; }
        public decimal Tm011 { get => tm011; set => tm011 = value; }
        public decimal Tm012 { get => tm012; set => tm012 = value; }
        public decimal Tm013 { get => tm013; set => tm013 = value; }
        public decimal Tm014 { get => tm014; set => tm014 = value; }
        public decimal Tm015 { get => tm015; set => tm015 = value; }
        public decimal Tm016 { get => tm016; set => tm016 = value; }
        public string Tm017 { get => tm017; set => tm017 = value; }
        public string Tm018 { get => tm018; set => tm018 = value; }
    }
}
