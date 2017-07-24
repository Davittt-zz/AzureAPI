using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessServices.Base
{
	public interface ICrudClass<T>
	{
		T Create(T Item);
		T Read(int ID);
		List<T> Read(object filter);
		bool Update(T Item);
		bool Delete(int ID);
	}
}
