using HAA.BusinessEntities;

namespace HAA.BusinessServices.Base
{
    public interface IUserServices
	{
		int Authenticate(string userName, string password);

		UserEntity GetUserById(int elementId);
	}
}
