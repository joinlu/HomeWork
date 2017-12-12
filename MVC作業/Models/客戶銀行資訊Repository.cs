using System;
using System.Linq;
using System.Collections.Generic;
	
namespace MVC作業.Models
{   
	public  class 客戶銀行資訊Repository : EFRepository<客戶銀行資訊>, I客戶銀行資訊Repository
	{
        public override IQueryable<客戶銀行資訊> All()
        {
            return base.All().Where(x => x.IsDeleted == false);
        }

        public IQueryable<客戶銀行資訊> fn取得所有資料()
        {
            return this.All();
        }

        public IQueryable<客戶銀行資訊> fn銀行名稱搜尋(string 銀行名稱)
        {
            return this.All().Where(x => x.銀行名稱.Contains(銀行名稱));
        }

        public 客戶銀行資訊 fn單筆搜尋(int id)
        {
            return this.All().FirstOrDefault(x => x.Id == id);
        }
        public IQueryable<客戶銀行資訊> fn單筆搜尋編輯用(int id)
        {
            return this.All().Where(x => x.Id == id);
        }
    }

    public  interface I客戶銀行資訊Repository : IRepository<客戶銀行資訊>
	{

	}
}