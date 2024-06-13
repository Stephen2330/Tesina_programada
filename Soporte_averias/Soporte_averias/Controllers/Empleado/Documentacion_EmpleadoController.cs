﻿using Soporte_averias.Models;
using Soporte_averias.Permissions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Soporte_averias.Controllers.Empleado
{

	[ValidarSesionAttribute]
	[PermisosRol(Rol.Empleado)]
	public class Documentacion_EmpleadoController : Controller
    {
        // GET: Documentacion_Empleado
        public ActionResult Index()
        {
            return View();
        }
		public FileResult DescargaManualTecnico()
		{

			string archivoPDF = "~/PDF/Manual_tecnico.pdf";


			return File(archivoPDF, "application/pdf", "Manual técnico del sistema.pdf");
		}

		public FileResult DescargaManualUsuario()
		{

			string archivoPDF = "~/PDF/Manual_usuario.pdf";
			string nombrePersonalizado = "Manual de usuario del sistema.pdf";

			Response.AddHeader("Content-Disposition", "attachment; filename= " + nombrePersonalizado);

			return File(archivoPDF, "application/pdf");
		}
	}
}