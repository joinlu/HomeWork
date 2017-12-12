using System;
using System.Linq;
using System.Collections.Generic;
	
namespace MVC作業.Models
{   
	public  class UserListViewRepository : EFRepository<UserListView>, IUserListViewRepository
	{

	}

	public  interface IUserListViewRepository : IRepository<UserListView>
	{

	}
}