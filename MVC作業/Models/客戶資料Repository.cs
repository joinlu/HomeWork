using System;
using System.Linq;
using System.Collections.Generic;
	
namespace MVC作業.Models
{   
	public  class 客戶資料Repository : EFRepository<客戶資料>, I客戶資料Repository
	{
        public override IQueryable<客戶資料> All()
        {
            return base.All().Where(x=>x.IsDeleted==false);
        }

        public IQueryable<客戶資料> fn取得所有資料()
        {
            return this.All();
        }

        public IQueryable<客戶資料> fn客戶篩選資料(string 客戶分類)
        {
            return this.All().Where(x => x.客戶分類.Equals(客戶分類));
        }

        public IQueryable<客戶資料> fn客戶名稱搜尋(string 姓名)
        {
            return this.All().Where(x => x.客戶名稱.Contains(姓名));
        }

        public List<String> fn分類清單()
        {
            List<String> item客戶分類 = new List<string>();
            item客戶分類.Add("全部分類");
            var 分類清單 = this.All().Select(x=>x.客戶分類).Distinct().ToList();
            foreach (var item in 分類清單)
            {
                if (item != null)
                {
                    item客戶分類.Add(item);
                }
            }
            return item客戶分類;
        }

        public 客戶資料 fn單筆搜尋(int id)
        {
            return this.All().FirstOrDefault(x=>x.Id==id);
        }

        public IQueryable<客戶資料> fn重新排序(string 排序)
        {
            var 客戶資料 = this.All();
            switch (排序)
            {
                case "Name desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.客戶名稱);
                    break;
                case "Name":
                    客戶資料 = 客戶資料.OrderBy(s => s.客戶名稱);
                    break;
                case "Uc desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.客戶分類);
                    break;
                case "Uc":
                    客戶資料 = 客戶資料.OrderBy(s => s.客戶分類);
                    break;
                case "Num desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.統一編號);
                    break;
                case "Num":
                    客戶資料 = 客戶資料.OrderBy(s => s.統一編號);
                    break;
                case "Tel desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.電話);
                    break;
                case "Tel":
                    客戶資料 = 客戶資料.OrderBy(s => s.電話);
                    break;
                case "Fax desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.傳真);
                    break;
                case "Fax":
                    客戶資料 = 客戶資料.OrderBy(s => s.傳真);
                    break;
                case "Add desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.地址);
                    break;
                case "Add":
                    客戶資料 = 客戶資料.OrderBy(s => s.地址);
                    break;
                case "Email desc":
                    客戶資料 = 客戶資料.OrderByDescending(s => s.Email);
                    break;
                case "Email":
                    客戶資料 = 客戶資料.OrderBy(s => s.Email);
                    break;

                default:
                    break;
            }

            return 客戶資料;
        }
    }

	public  interface I客戶資料Repository : IRepository<客戶資料>
	{

	}
}