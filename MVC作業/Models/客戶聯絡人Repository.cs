using System;
using System.Linq;
using System.Collections.Generic;
	
namespace MVC作業.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        public override IQueryable<客戶聯絡人> All()
        {
            return base.All().Where(x => x.IsDeleted == false);
        }

        public IQueryable<客戶聯絡人> fn取得所有資料()
        {
            return this.All();
        }
        public IQueryable<客戶聯絡人> fn客戶篩選資料(string 職稱)
        {
            return this.All().Where(x => x.職稱.Equals(職稱));
        }

        public IQueryable<客戶聯絡人> fn客戶名稱搜尋(string 姓名)
        {
            return this.All().Where(x => x.姓名.Contains(姓名));
        }


        public List<String> fn職稱清單()
        {
            List<String> item職稱分類 = new List<string>();
            item職稱分類.Add("全部職稱");
            var 分類清單 = this.All().Select(x => x.職稱).Distinct().ToList();
            foreach (var item in 分類清單)
            {
                if (item != null)
                {
                    item職稱分類.Add(item);
                }
            }
            return item職稱分類;
        }

        public IQueryable<客戶聯絡人> fn重新排序(string 排序)
        {
            var 客戶聯絡人 = this.All();
            switch (排序)
            {
                case "Name desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.姓名);
                    break;
                case "Name":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.姓名);
                    break;
                case "Email desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.Email);
                    break;
                case "Email":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.Email);
                    break;
                case "Phone desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.手機);
                    break;
                case "Phone":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.手機);
                    break;
                case "Tel desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.電話);
                    break;
                case "Tel":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.電話);
                    break;
                case "User desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.客戶Id);
                    break;
                case "User":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.客戶Id);
                    break;
                case "Job desc":
                    客戶聯絡人 = 客戶聯絡人.OrderByDescending(s => s.職稱);
                    break;
                case "Job":
                    客戶聯絡人 = 客戶聯絡人.OrderBy(s => s.職稱);
                    break;

                default:
                    break;
            }

            return 客戶聯絡人;
        }

        public 客戶聯絡人 fn單筆搜尋(int id)
        {
            return this.All().FirstOrDefault(x => x.Id == id);
        }

        public 客戶聯絡人 fnEail是否已存在(string Email)
        {
          return this.All().Where(x => x.Email.Equals(Email)).FirstOrDefault();
        }
    }

    public  interface I客戶聯絡人Repository : IRepository<客戶聯絡人>
	{

	}
}