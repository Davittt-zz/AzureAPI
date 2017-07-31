using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using AzureAPI.Models;
using BusinessEntities;
using BusinessServices.Base;
using AzureAPI.Filters;
using AzureAPI.ActionFilters;
using Microsoft.AspNet.Identity.Owin;

namespace AzureAPI.Controllers
{
	[AuthorizationRequired]
	public class UserEntitiesController : Controller
	{
		private readonly IUserServices _userServices;

		#region Public Constructor

		/// <summary>
		/// Public constructor to initialize element service instance with DI
		/// 
		/// </summary>
		public UserEntitiesController(IUserServices userServices)
		{
			_userServices = userServices;
		}

		#endregion
		// GET: UserEntities
		// public ActionResult Index()
		// {
		//	var elements = _userServices.GetAllUsers();
		//	return View(elements);
		// }

		// GET: UserEntities/Details/
		public ActionResult Details()
		{
			int id = 1;
			if (id > 1)
			{
				return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
			}
			UserEntity userEntity = _userServices.GetUserById(id);
			if (userEntity == null)
			{
				return HttpNotFound();
			}
			return View(userEntity);
		}

		// GET: UserEntities/Create
		// public ActionResult Create()
		// {
		//     return View();
		// }

		// POST: UserEntities/Create
		// Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
		// más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
		//[HttpPost]
		//[ValidateAntiForgeryToken]
		//public ActionResult Create([Bind(Include = "Id,FirstName,LastName,Phone,UserType,Password,UserLevel,Username,Email")] UserEntity userEntity)
		//{
		//    if (ModelState.IsValid)
		//    {
		//        db.UserEntities.Add(userEntity);
		//        db.SaveChanges();
		//        return RedirectToAction("Index");
		//    }
		//
		//    return View(userEntity);
		//}

		// GET: UserEntities/Edit/5
		public ActionResult Edit(int id)
		{
			if (id > 0)
			{
				return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
			}
			UserEntity userEntity = _userServices.GetUserById(id);
			if (userEntity == null)
			{
				return HttpNotFound();
			}
			return View(userEntity);
		}

		// POST: UserEntities/Edit/5
		// Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
		// más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Edit(int id, [Bind(Include = "Id,FirstName,LastName,Phone,UserType,Password,UserLevel,Username,Email")] UserEntity userEntity)
		{
			if (ModelState.IsValid)
			{
				//_userServices.UpdateUser(id);
				return RedirectToAction("Index");
			}
			return View(userEntity);
		}

		// GET: UserEntities/Delete/5
		//public ActionResult Delete(int? id)
		//{
		//    if (id == null)
		//    {
		//        return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
		//    }
		//    UserEntity userEntity = db.UserEntities.Find(id);
		//    if (userEntity == null)
		//    {
		//        return HttpNotFound();
		//    }
		//    return View(userEntity);
		//}

		// POST: UserEntities/Delete/5
		//[HttpPost, ActionName("Delete")]
		//[ValidateAntiForgeryToken]
		//public ActionResult DeleteConfirmed(int id)
		//{
		//    UserEntity userEntity = db.UserEntities.Find(id);
		//    db.UserEntities.Remove(userEntity);
		//    db.SaveChanges();
		//    return RedirectToAction("Index");
		//}

	}
}
