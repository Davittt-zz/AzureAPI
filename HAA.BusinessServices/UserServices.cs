using AutoMapper;
using HAA.BusinessEntities;
using HAA.BusinessServices.Base;
using HAA.DataModel;
using HAA.DataModel.UnitOfWork;

namespace HAA.BusinessServices
{
    /// <summary>
    /// Offers services for user specific operations
    /// </summary>
    public class UserServices : IUserServices
	{
		private readonly UnitOfWork _unitOfWork;

		/// <summary>
		/// Public constructor.
		/// </summary>
		public UserServices(UnitOfWork unitOfWork)
		{
			_unitOfWork = unitOfWork;
		}

		/// <summary>
		/// Public method to authenticate user by user name and password.
		/// </summary>
		/// <param name="userName"></param>
		/// <param name="password"></param>
		/// <returns></returns>
		public int Authenticate(string userName, string password)
		{
			var user = _unitOfWork.UserRepository.Get(u => u.UserID == userName && u.Password == password);
			if (user != null)
			{
				return user.Index;
			}
			return 0;
		}

		/// <summary>
		/// Fetches element details by id
		/// </summary>
		/// <param name="elementId"></param>
		/// <returns></returns>
		public UserEntity GetUserById(int userId)
		{
			var user = _unitOfWork.UserRepository.GetByID(userId);
			if (user != null)
			{
				var config = new MapperConfiguration(cfg =>
				{
					cfg.CreateMap<User, UserEntity>().ForMember(x => x.Password, opt => opt.Ignore());
				});

				IMapper mapper = config.CreateMapper();
				var usersModel = mapper.Map<User, UserEntity>(user);
				return usersModel;
			}
			return null;
		}

	}
}
