using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Data.Entity.Validation;
using HAA.DataModel.Repositories;

namespace HAA.DataModel.UnitOfWork
{
	/// <summary>
	/// Unit of Work class responsible for DB transactions
	/// </summary>
	public class UnitOfWork : IDisposable, IUnitOfWork
	{
		private readonly WebApiDbEntities _context = null;
		private GenericRepository<User> _userRepository;
		private GenericRepository<UserActivity> _userActivityRepository;
		private GenericRepository<Element> _elementRepository;
        private SpeakerRepository _speakerConfigRepository;
        private GenericRepository<Token> _tokenRepository;

		public UnitOfWork()
		{
			_context = new WebApiDbEntities();
		}

		#region Public Repository Creation properties...

		/// <summary>
		/// Get/Set Property for user repository.
		/// </summary>
		public GenericRepository<User> UserRepository
		{
			get
			{
				if (this._userRepository == null)
					this._userRepository = new GenericRepository<User>(_context);
				return _userRepository;
			}
		}

		/// <summary>
		/// Get/Set Property for User Activity.
		/// </summary>
		public GenericRepository<UserActivity> UserActivityRepository
		{
			get
			{
				if (this._userActivityRepository == null)
					this._userActivityRepository = new GenericRepository<UserActivity>(_context);
				return _userActivityRepository;
			}
		}

		/// <summary>
		/// Get/Set Property for token repository.
		/// </summary>
		public GenericRepository<Element> ElementRepository
		{
			get
			{
				if (this._elementRepository == null)
					this._elementRepository = new GenericRepository<Element>(_context);
				return _elementRepository;
			}
		}

        public SpeakerRepository SpeakerConfigRepository
        {
            get
            {
                if (this._speakerConfigRepository == null)
                    this._speakerConfigRepository = new SpeakerRepository(_context);
                return _speakerConfigRepository;
            }
        }

        /// <summary>
        /// Get/Set Property for token repository.
        /// </summary>
        public GenericRepository<Token> TokenRepository
		{
			get
			{
				if (this._tokenRepository == null)
					this._tokenRepository = new GenericRepository<Token>(_context);
				return _tokenRepository;
			}
		}
		#endregion

		#region Public member methods...
		/// <summary>
		/// Save method.
		/// </summary>
		public void Save()
		{
			try
			{
				_context.SaveChanges();
			}
			catch (DbEntityValidationException e)
			{
				var outputLines = new List<string>();
				foreach (var eve in e.EntityValidationErrors)
				{
					outputLines.Add(string.Format(
						"{0}: Entity of type \"{1}\" in state \"{2}\" has the following validation errors:", DateTime.Now,
						eve.Entry.Entity.GetType().Name, eve.Entry.State));
					foreach (var ve in eve.ValidationErrors)
					{
						outputLines.Add(string.Format("- Property: \"{0}\", Error: \"{1}\"", ve.PropertyName, ve.ErrorMessage));
					}
				}
				System.IO.File.AppendAllLines(@"C:\errors.txt", outputLines);

				throw e;
			}
		}
		#endregion

		#region Implementing IDisposable...

		#region private dispose variable declaration...
		private bool disposed = false;
		#endregion

		/// <summary>
		/// Protected Virtual Dispose method
		/// </summary>
		/// <param name="disposing"></param>
		protected virtual void Dispose(bool disposing)
		{
			if (!this.disposed)
			{
				if (disposing)
				{
					Debug.WriteLine("UnitOfWork is being disposed");
					_context.Dispose();
				}
			}
			this.disposed = true;
		}

		/// <summary>
		/// Dispose method
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}