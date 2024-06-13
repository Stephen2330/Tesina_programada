using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using iTextSharp.text.pdf;
using iTextSharp.text;
using OfficeOpenXml.Style;
using PagedList;
using Soporte_averias.Models;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Soporte_averias.Permissions;

namespace Soporte_averias.Controllers
{
	[ValidarSesionAttribute]
	[PermisosRol(Rol.Administrador)]
	public class CasoController : Controller
    {
        private SOPORTEEntities db = new SOPORTEEntities();

		public ActionResult Index(int? page, int? selectEstado)
		{
			ViewBag.Title = "Lista de casos";

			// Obtener la lista de estados de caso
			var estados = db.TBL_EstadoCaso.OrderBy(e => e.TC_Nombre).ToList();
			var estadosSelectList = new SelectList(estados, "TN_IdEstadoCaso", "TC_Nombre");

			ViewBag.Estados = estadosSelectList;

			int pageSize = 10;
			int pageNumber = (page ?? 1);
			ViewBag.PageNumber = pageNumber;

			IQueryable<TBL_Caso> tBL_CasoQuery = db.TBL_Caso
				.Include(t => t.TBL_Cliente)
				.Include(t => t.TBL_DescripcionCaso)
				.Include(t => t.TBL_Empleado)
				.Include(t => t.TBL_EstadoCaso)
				.Include(t => t.TBL_FechaCierreCaso)
				.Include(t => t.TBL_FechaCreacionCaso)
				.Include(t => t.TBL_Municipalidad)
				.Include(t => t.TBL_PeriodoGarantia)
				.Include(t => t.TBL_PrioridadCaso)
				.Include(t => t.TBL_Producto)
				.Include(t => t.TBL_Usuario);

			if (selectEstado.HasValue)
			{
				tBL_CasoQuery = tBL_CasoQuery.Where(c => c.TBL_EstadoCaso.TN_IdEstadoCaso == selectEstado);
			}

			int totalItems = tBL_CasoQuery.Count(); // Cant. elementos totales
			int totalPages = (int)Math.Ceiling((double)totalItems / pageSize); // Cant. total de páginas
			ViewBag.TotalPages = totalPages;

			// Obtener datos de los empleados con nombre y primer apellido
			var empleados = db.TBL_Empleado.OrderBy(e => e.TC_Nombre).ToList()
							.Select(e => new SelectListItem
							{
								Value = e.TN_IdEmpleado.ToString(),
								Text = $"{e.TC_Nombre} {e.TC_PrimerApellido}" // Concatenar nombre y primer apellido
							}).ToList();

			// Agregar opción por defecto para empleado
			empleados.Insert(0, new SelectListItem { Value = "", Text = "Seleccione empleado" });

			// Asignar lista de empleados al ViewBag
			ViewBag.TN_IdEmpleadoList = new SelectList(empleados, "Value", "Text");

			var casosPaginas = tBL_CasoQuery.OrderBy(c => c.TBL_Cliente.TC_Nombre)
											.ToPagedList(pageNumber, pageSize);

			return View(casosPaginas);
		}

		// GET: Caso/Details/5
		public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TBL_Caso tBL_Caso = db.TBL_Caso.Find(id);
            if (tBL_Caso == null)
            {
                return HttpNotFound();
            }
            return View(tBL_Caso);
        }

		// GET: Caso/Create
		public ActionResult Create()
		{
			// Obtener fechas de creación formateadas
			var fechasCreacion = db.TBL_FechaCreacionCaso
									.OrderBy(f => f.TD_FechaCreacionCaso)
									.ToList()
									.Select(f => new SelectListItem
									{
										Value = f.TN_IdFechaCreacionCaso.ToString(),
										Text = f.TD_FechaCreacionCaso?.ToString("dd-MM-yyyy")
									}).ToList();

			// Agregar opción por defecto para la fecha de creación
			fechasCreacion.Insert(0, new SelectListItem { Value = "", Text = "Seleccione fecha" });

			var fechasCierre = db.TBL_FechaCierreCaso
								.OrderBy(f => f.TD_FechaCierreCaso)
								.ToList()
								.Select(f => new SelectListItem
								{
									Value = f.TN_IdFechaCierreCaso.ToString(),
									Text = f.TD_FechaCierreCaso?.ToString("dd-MM-yyyy")
								}).ToList();

			// Agregar opción por defecto para la fecha de cierre
			fechasCierre.Insert(0, new SelectListItem { Value = "", Text = "Seleccione fecha" });


			// Obtener datos de los empleados con nombre y primer apellido
			var empleados = db.TBL_Empleado.OrderBy(e => e.TC_Nombre).ToList()
							.Select(e => new SelectListItem
							{
								Value = e.TN_IdEmpleado.ToString(),
								Text = $"{e.TC_Nombre} {e.TC_PrimerApellido}" // Concatenar nombre y primer apellido
							}).ToList();


			// Agregar opción por defecto para empleado
			empleados.Insert(0, new SelectListItem { Value = "", Text = "Seleccione empleado" });

			// Asignar lista de empleados al ViewBag
			ViewBag.TN_IdEmpleadoList = new SelectList(empleados, "Value", "Text");

			var usuarios = db.TBL_Usuario.OrderBy(u => u.TC_Nombre).ToList()
							.Select(u => new SelectListItem
							{
								Value = u.TN_IdUsuario.ToString(),
								Text = $"{u.TC_Nombre} {u.TC_PrimerApellido}" // Concatenar nombre y primer apellido
							}).ToList();
			usuarios.Insert(0, new SelectListItem { Value = "", Text = "Seleccione usuario" });
			ViewBag.TN_IdUsuarioList = new SelectList(usuarios, "Value", "Text");

			// Definir la opción por defecto para todos los dropdown lists
			var selectCliente = new SelectListItem { Value = "", Text = "Seleccione cliente" };
			var selectDescripcion = new SelectListItem { Value = "", Text = "Seleccione descripción" };
			var selectEmpleado = new SelectListItem { Value = "", Text = "Seleccione empleado" };
			var selectEstado = new SelectListItem { Value = "", Text = "Seleccione estado" };
			var selectMunicipalidad = new SelectListItem { Value = "", Text = "Seleccione municipalidad" };
			var selectPeriodoGarantia = new SelectListItem { Value = "", Text = "Seleccione periodo de garantía" };
			var selectPrioridadCaso = new SelectListItem { Value = "", Text = "Seleccione prioridad" };
			var selectProducto = new SelectListItem { Value = "", Text = "Seleccione producto" };
			var selectUsuario = new SelectListItem { Value = "", Text = "Seleccione usuario" };

			// Agregar las opciones por defecto al principio de cada dropdown list
			ViewBag.TN_IdCliente = new SelectList((new SelectListItem[] { selectCliente }).Concat(db.TBL_Cliente.OrderBy(c => c.TC_Nombre).Select(c => new SelectListItem { Value = c.TN_IdCliente.ToString(), Text = c.TC_Nombre }).ToList()), "Value", "Text");
			ViewBag.TN_IdDescripcionCaso = new SelectList((new SelectListItem[] { selectDescripcion }).Concat(db.TBL_DescripcionCaso.OrderBy(d => d.TC_Descripcion).Select(d => new SelectListItem { Value = d.TN_IdDescripcionCaso.ToString(), Text = d.TC_Descripcion }).ToList()), "Value", "Text");
			ViewBag.TN_IdEmpleado = new SelectList((new SelectListItem[] { selectEmpleado }).Concat(db.TBL_Empleado.OrderBy(e => e.TC_Nombre).Select(e => new SelectListItem { Value = e.TN_IdEmpleado.ToString(), Text = e.TC_Nombre }).ToList()), "Value", "Text");
			ViewBag.TN_IdEstadoCaso = new SelectList((new SelectListItem[] { selectEstado }).Concat(db.TBL_EstadoCaso.OrderBy(e => e.TC_Nombre).Select(e => new SelectListItem { Value = e.TN_IdEstadoCaso.ToString(), Text = e.TC_Nombre }).ToList()), "Value", "Text");
			ViewBag.TN_IdMunicipalidad = new SelectList((new SelectListItem[] { selectMunicipalidad }).Concat(db.TBL_Municipalidad.OrderBy(m => m.TC_Nombre).Select(m => new SelectListItem { Value = m.TN_IdMunicipalidad.ToString(), Text = m.TC_Nombre }).ToList()), "Value", "Text");
			ViewBag.TN_IdPeriodoGarantia = new SelectList((new SelectListItem[] { selectPeriodoGarantia }).Concat(db.TBL_PeriodoGarantia.OrderBy(p => p.TC_PeriodoGarantia).Select(p => new SelectListItem { Value = p.TN_IdPeriodoGarantia.ToString(), Text = p.TC_PeriodoGarantia }).ToList()), "Value", "Text");
			ViewBag.TN_IdPrioridadCaso = new SelectList((new SelectListItem[] { selectPrioridadCaso }).Concat(db.TBL_PrioridadCaso.OrderBy(p => p.TC_Nombre).Select(p => new SelectListItem { Value = p.TN_IdPrioridadCaso.ToString(), Text = p.TC_Nombre }).ToList()), "Value", "Text");
			ViewBag.TN_IdProducto = new SelectList((new SelectListItem[] { selectProducto }).Concat(db.TBL_Producto.OrderBy(p => p.TC_Nombre).Select(p => new SelectListItem { Value = p.TN_IdProducto.ToString(), Text = p.TC_Nombre }).ToList()), "Value", "Text");
			ViewBag.TN_IdUsuario = new SelectList((new SelectListItem[] { selectUsuario }).Concat(db.TBL_Usuario.OrderBy(u => u.TC_Nombre).Select(u => new SelectListItem { Value = u.TN_IdUsuario.ToString(), Text = u.TC_Nombre }).ToList()), "Value", "Text");

			ViewBag.TN_IdFechaCierreCaso = new SelectList(fechasCierre, "Value", "Text");
			ViewBag.TN_IdFechaCreacionCaso = new SelectList(fechasCreacion, "Value", "Text");

			return View();
		}
		// POST: Caso/Create
		// To protect from overposting attacks, enable the specific properties you want to bind to, for 
		// more details see https://go.microsoft.com/fwlink/?LinkId=317598.
		[HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "TN_IdCaso,TN_IdPeriodoGarantia,TN_IdFechaCreacionCaso,TN_IdEmpleado,TN_IdMunicipalidad,TN_IdCliente,TN_IdProducto,TN_IdDescripcionCaso,TN_IdPrioridadCaso,TN_IdEstadoCaso,TN_IdFechaCierreCaso,TN_IdUsuario")] TBL_Caso tBL_Caso)
        {
            if (ModelState.IsValid)
            {
                db.TBL_Caso.Add(tBL_Caso);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.TN_IdCliente = new SelectList(db.TBL_Cliente, "TN_IdCliente", "TC_Nombre", tBL_Caso.TN_IdCliente);
            ViewBag.TN_IdDescripcionCaso = new SelectList(db.TBL_DescripcionCaso, "TN_IdDescripcionCaso", "TC_Descripcion", tBL_Caso.TN_IdDescripcionCaso);
            ViewBag.TN_IdEmpleado = new SelectList(db.TBL_Empleado, "TN_IdEmpleado", "TC_Nombre", tBL_Caso.TN_IdEmpleado);
            ViewBag.TN_IdEstadoCaso = new SelectList(db.TBL_EstadoCaso, "TN_IdEstadoCaso", "TC_Nombre", tBL_Caso.TN_IdEstadoCaso);
            ViewBag.TN_IdFechaCierreCaso = new SelectList(db.TBL_FechaCierreCaso, "TN_IdFechaCierreCaso", "TC_Descripcion", tBL_Caso.TN_IdFechaCierreCaso);
            ViewBag.TN_IdFechaCreacionCaso = new SelectList(db.TBL_FechaCreacionCaso, "TN_IdFechaCreacionCaso", "TC_Descripcion", tBL_Caso.TN_IdFechaCreacionCaso);
            ViewBag.TN_IdMunicipalidad = new SelectList(db.TBL_Municipalidad, "TN_IdMunicipalidad", "TC_Nombre", tBL_Caso.TN_IdMunicipalidad);
            ViewBag.TN_IdPeriodoGarantia = new SelectList(db.TBL_PeriodoGarantia, "TN_IdPeriodoGarantia", "TC_PeriodoGarantia", tBL_Caso.TN_IdPeriodoGarantia);
            ViewBag.TN_IdPrioridadCaso = new SelectList(db.TBL_PrioridadCaso, "TN_IdPrioridadCaso", "TC_Nombre", tBL_Caso.TN_IdPrioridadCaso);
            ViewBag.TN_IdProducto = new SelectList(db.TBL_Producto, "TN_IdProducto", "TC_Nombre", tBL_Caso.TN_IdProducto);
            ViewBag.TN_IdUsuario = new SelectList(db.TBL_Usuario, "TN_IdUsuario", "TC_Nombre", tBL_Caso.TN_IdUsuario);
            return View(tBL_Caso);
        }

		// GET: Caso/Edit/5
		public ActionResult Edit(int? id)
		{
			if (id == null)
			{
				return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
			}
			TBL_Caso tBL_Caso = db.TBL_Caso.Find(id);
			if (tBL_Caso == null)
			{
				return HttpNotFound();
			}

			// Obtener datos de los empleados con nombre y primer apellido
			var empleados = db.TBL_Empleado.OrderBy(e => e.TC_Nombre).ToList()
							.Select(e => new SelectListItem
							{
								Value = e.TN_IdEmpleado.ToString(),
								Text = $"{e.TC_Nombre} {e.TC_PrimerApellido}" // Concatenar nombre y primer apellido
							}).ToList();

			// Agregar opción por defecto para empleado
			empleados.Insert(0, new SelectListItem { Value = "", Text = "Seleccione empleado" });

			// Asignar lista de empleados al ViewBag
			ViewBag.TN_IdEmpleadoList = empleados;



			// Obtener lista de usuarios y seleccionar el usuario asociado al caso
			var usuarios = db.TBL_Usuario.OrderBy(u => u.TC_Nombre).ToList()
							.Select(u => new SelectListItem
							{
								Value = u.TN_IdUsuario.ToString(),
								Text = $"{u.TC_Nombre} {u.TC_PrimerApellido}"
							}).ToList();

			// Agregar opción "Seleccione usuario" al principio de la lista
			usuarios.Insert(0, new SelectListItem { Value = "", Text = "Seleccione usuario" });

			var usuarioAsoc = tBL_Caso.TN_IdUsuario.ToString();
			// Establecer ViewBag para la lista de usuarios
			ViewBag.TN_IdUsuarioList = new SelectList(usuarios, "Value", "Text", usuarioAsoc);



			// Obtener las demás opciones para los dropdownlist
			var periodosGarantia = db.TBL_PeriodoGarantia.OrderBy(p => p.TC_PeriodoGarantia).ToList();
			var municipalidades = db.TBL_Municipalidad.OrderBy(m => m.TC_Nombre).ToList();
			var clientes = db.TBL_Cliente.OrderBy(c => c.TC_Nombre).ToList();
			var productos = db.TBL_Producto.OrderBy(p => p.TC_Nombre).ToList();
			var descripcionesCaso = db.TBL_DescripcionCaso.OrderBy(d => d.TC_Descripcion).ToList();
			var prioridadesCaso = db.TBL_PrioridadCaso.OrderBy(p => p.TC_Nombre).ToList();
			var estadosCaso = db.TBL_EstadoCaso.OrderBy(e => e.TC_Nombre).ToList();

			// Insertar opción "Seleccione" al principio de cada lista
			periodosGarantia.Insert(0, new TBL_PeriodoGarantia { TN_IdPeriodoGarantia = 0, TC_PeriodoGarantia = "Seleccione periodo" });
			municipalidades.Insert(0, new TBL_Municipalidad { TN_IdMunicipalidad = 0, TC_Nombre = "Seleccione municipalidad" });
			clientes.Insert(0, new TBL_Cliente { TN_IdCliente = 0, TC_Nombre = "Seleccione cliente" });
			productos.Insert(0, new TBL_Producto { TN_IdProducto = 0, TC_Nombre = "Seleccione producto" });
			descripcionesCaso.Insert(0, new TBL_DescripcionCaso { TN_IdDescripcionCaso = 0, TC_Descripcion = "Seleccione descripción" });
			prioridadesCaso.Insert(0, new TBL_PrioridadCaso { TN_IdPrioridadCaso = 0, TC_Nombre = "Seleccione prioridad" });
			estadosCaso.Insert(0, new TBL_EstadoCaso { TN_IdEstadoCaso = 0, TC_Nombre = "Seleccione estado" });

			// Establecer ViewBag para cada dropdownlist
			ViewBag.TN_IdPeriodoGarantia = new SelectList(periodosGarantia, "TN_IdPeriodoGarantia", "TC_PeriodoGarantia", tBL_Caso.TN_IdPeriodoGarantia);
			ViewBag.TN_IdEmpleado = new SelectList(empleados, "Value", "Text", tBL_Caso.TN_IdEmpleado);
			ViewBag.TN_IdMunicipalidad = new SelectList(municipalidades, "TN_IdMunicipalidad", "TC_Nombre", tBL_Caso.TN_IdMunicipalidad);
			ViewBag.TN_IdCliente = new SelectList(clientes, "TN_IdCliente", "TC_Nombre", tBL_Caso.TN_IdCliente);
			ViewBag.TN_IdProducto = new SelectList(productos, "TN_IdProducto", "TC_Nombre", tBL_Caso.TN_IdProducto);
			ViewBag.TN_IdDescripcionCaso = new SelectList(descripcionesCaso, "TN_IdDescripcionCaso", "TC_Descripcion", tBL_Caso.TN_IdDescripcionCaso);
			ViewBag.TN_IdPrioridadCaso = new SelectList(prioridadesCaso, "TN_IdPrioridadCaso", "TC_Nombre", tBL_Caso.TN_IdPrioridadCaso);
			ViewBag.TN_IdEstadoCaso = new SelectList(estadosCaso, "TN_IdEstadoCaso", "TC_Nombre", tBL_Caso.TN_IdEstadoCaso);

			// Obtener fechas de creación y cierre
			var fechasCreacion = db.TBL_FechaCreacionCaso
				.OrderBy(f => f.TD_FechaCreacionCaso)
				.ToList()
				.Select(f => new SelectListItem
				{
					Value = f.TN_IdFechaCreacionCaso.ToString(),
					Text = f.TD_FechaCreacionCaso.HasValue ? f.TD_FechaCreacionCaso.Value.ToString("dd-MM-yyyy") : ""
				})
				.ToList();

			var fechasCierre = db.TBL_FechaCierreCaso
				.OrderBy(f => f.TD_FechaCierreCaso)
				.ToList()
				.Select(f => new SelectListItem
				{
					Value = f.TN_IdFechaCierreCaso.ToString(),
					Text = f.TD_FechaCierreCaso.HasValue ? f.TD_FechaCierreCaso.Value.ToString("dd-MM-yyyy") : ""
				})
				.ToList();

			// Insertar opción "Seleccione" al principio de cada lista
			fechasCreacion.Insert(0, new SelectListItem { Value = "", Text = "Seleccione fecha" });
			fechasCierre.Insert(0, new SelectListItem { Value = "", Text = "Seleccione fecha" });

			// Establecer ViewBag para fechas de creación y cierre
			ViewBag.TN_IdFechaCreacionCaso = new SelectList(fechasCreacion, "Value", "Text", tBL_Caso.TN_IdFechaCreacionCaso);
			ViewBag.TN_IdFechaCierreCaso = new SelectList(fechasCierre, "Value", "Text", tBL_Caso.TN_IdFechaCierreCaso);

			return View(tBL_Caso);
		}
		// POST: Caso/Edit/5
		// To protect from overposting attacks, enable the specific properties you want to bind to, for 
		// more details see https://go.microsoft.com/fwlink/?LinkId=317598.
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Edit([Bind(Include = "TN_IdCaso,TN_IdPeriodoGarantia,TN_IdFechaCreacionCaso,TN_IdEmpleado,TN_IdMunicipalidad,TN_IdCliente,TN_IdProducto,TN_IdDescripcionCaso,TN_IdPrioridadCaso,TN_IdEstadoCaso,TN_IdFechaCierreCaso,TN_IdUsuario")] TBL_Caso tBL_Caso)
		{
			if (ModelState.IsValid)
			{
				
				db.Entry(tBL_Caso).State = EntityState.Modified;
				db.SaveChanges();
				return RedirectToAction("Index");
			}
			ViewBag.TN_IdCliente = new SelectList(db.TBL_Cliente, "TN_IdCliente", "TC_Nombre", tBL_Caso.TN_IdCliente);
			ViewBag.TN_IdDescripcionCaso = new SelectList(db.TBL_DescripcionCaso, "TN_IdDescripcionCaso", "TC_Descripcion", tBL_Caso.TN_IdDescripcionCaso);
			ViewBag.TN_IdEmpleado = new SelectList(db.TBL_Empleado, "TN_IdEmpleado", "TC_Nombre", tBL_Caso.TN_IdEmpleado);
			ViewBag.TN_IdEstadoCaso = new SelectList(db.TBL_EstadoCaso, "TN_IdEstadoCaso", "TC_Nombre", tBL_Caso.TN_IdEstadoCaso);
			ViewBag.TN_IdFechaCierreCaso = new SelectList(db.TBL_FechaCierreCaso, "TN_IdFechaCierreCaso", "TC_Descripcion", tBL_Caso.TN_IdFechaCierreCaso);
			ViewBag.TN_IdFechaCreacionCaso = new SelectList(db.TBL_FechaCreacionCaso, "TN_IdFechaCreacionCaso", "TC_Descripcion", tBL_Caso.TN_IdFechaCreacionCaso);
			ViewBag.TN_IdMunicipalidad = new SelectList(db.TBL_Municipalidad, "TN_IdMunicipalidad", "TC_Nombre", tBL_Caso.TN_IdMunicipalidad);
			ViewBag.TN_IdPeriodoGarantia = new SelectList(db.TBL_PeriodoGarantia, "TN_IdPeriodoGarantia", "TC_PeriodoGarantia", tBL_Caso.TN_IdPeriodoGarantia);
			ViewBag.TN_IdPrioridadCaso = new SelectList(db.TBL_PrioridadCaso, "TN_IdPrioridadCaso", "TC_Nombre", tBL_Caso.TN_IdPrioridadCaso);
			ViewBag.TN_IdProducto = new SelectList(db.TBL_Producto, "TN_IdProducto", "TC_Nombre", tBL_Caso.TN_IdProducto);
			ViewBag.TN_IdUsuario = new SelectList(db.TBL_Usuario, "TN_IdUsuario", "TC_Nombre", tBL_Caso.TN_IdUsuario);
			return View(tBL_Caso);
		}
		
		
		// GET: Caso/Delete/5
		public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TBL_Caso tBL_Caso = db.TBL_Caso.Find(id);
            if (tBL_Caso == null)
            {
                return HttpNotFound();
            }
            return View(tBL_Caso);
        }

        // POST: Caso/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            TBL_Caso tBL_Caso = db.TBL_Caso.Find(id);
            db.TBL_Caso.Remove(tBL_Caso);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

		public ActionResult ExportToPdf(int? selectEstado, int? page)
		{
			int pageNumber = page ?? 1;

			var actividad = db.TBL_Caso.AsQueryable();

			if (selectEstado.HasValue)
			{
				actividad = actividad.Where(m => m.TBL_EstadoCaso.TN_IdEstadoCaso.Equals(selectEstado));
			}
			actividad = actividad.OrderBy(m => m.TN_IdEstadoCaso);
			var pagedActividad = actividad.ToList();

			// Crear el documento PDF
			Document pdfDoc = new Document(PageSize.A4.Rotate(), 10f, 10f, 10f, 10f);
			MemoryStream memoryStream = new MemoryStream();
			PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
			pdfDoc.Open();

			// Estampar la fecha y hora en el pie de página
			string fechaHoraDescarga = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
			pdfDoc.Add(new Paragraph($"Informe generado\n{fechaHoraDescarga}", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL)));

			//// Crear una celda divisora con fondo gris y borde inferior
			//PdfPCell dividerCell = new PdfPCell(new Phrase(" "));
			//dividerCell.Colspan = 6; // Establece el número de columnas que la celda ocupará
			//dividerCell.BackgroundColor = new BaseColor(192, 192, 192); // Gris claro
			//dividerCell.Border = PdfPCell.BOTTOM_BORDER; // Agrega un borde inferior a la celda

			// Crear tabla en PDF
			PdfPTable pdfTable = new PdfPTable(11);
			pdfTable.WidthPercentage = 100;


			// Crear estilo para los encabezados (negrita y fondo gris)
			Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD);
			BaseColor headerBackgroundColor = new BaseColor(192, 192, 192); // Gris claro

			// Añadir encabezados a la tabla con el estilo
			PdfPCell headerCell;

			headerCell = new PdfPCell(new Phrase("Período de garantía", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Fecha creación caso", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Empleado asignado", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Municipalidad", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Cliente", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Producto", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Descripción caso", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Prioridad caso", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Estado caso", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Quien reporta caso", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Fecha cierre caso", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			foreach (var item in pagedActividad)
			{
				pdfTable.AddCell(item.TBL_PeriodoGarantia.TC_PeriodoGarantia.ToString());
				pdfTable.AddCell(item.TBL_FechaCreacionCaso.TD_FechaCreacionCaso.ToString());
				pdfTable.AddCell(item.TBL_Empleado.TC_Nombre.ToString() + " " + item.TBL_Empleado.TC_PrimerApellido.ToString());
				pdfTable.AddCell(item.TBL_Municipalidad.TC_Nombre.ToString());
				pdfTable.AddCell(item.TBL_Cliente.TC_Nombre.ToString());
				pdfTable.AddCell(item.TBL_Producto.TC_Nombre.ToString());
				pdfTable.AddCell(item.TBL_DescripcionCaso.TC_Descripcion.ToString());
				pdfTable.AddCell(item.TBL_PrioridadCaso.TC_Nombre.ToString());
				pdfTable.AddCell(item.TBL_EstadoCaso.TC_Nombre.ToString());
				pdfTable.AddCell(item.TBL_Usuario.TC_Nombre.ToString() + " " + item.TBL_Usuario.TC_PrimerApellido.ToString());
				pdfTable.AddCell(item.TBL_FechaCierreCaso.TD_FechaCierreCaso.ToString());


			}



			pdfDoc.Add(pdfTable);
			pdfDoc.Close();

			byte[] bytes = memoryStream.ToArray();
			memoryStream.Close();
			return File(bytes, "application/pdf", "Datos_casos.pdf");
		}

		//Crear ExportToExcel
		public ActionResult ExportToExcel(int? selectEstado)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var actividad = db.TBL_Caso.AsQueryable();

			if (selectEstado.HasValue)
			{
				actividad = actividad.Where(m => m.TBL_EstadoCaso.TN_IdEstadoCaso.Equals(selectEstado));
			}
				
			else
			{
				ViewData["Mensaje"] = "*No se encontraron datos*";
			}

			// Ordenar los datos por FECHA
			actividad = actividad.OrderBy(m => m.TN_IdEstadoCaso);
			var data = actividad.ToList();

			// Crear el archivo Excel utilizando EPPlus
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Casos");

				// Encabezados
				worksheet.Cells[1, 1].Value = "Período de garantía";
				worksheet.Cells[1, 2].Value = "Fecha creación caso";
				worksheet.Cells[1, 3].Value = "Empleado asignado";
				worksheet.Cells[1, 4].Value = "Municipalidad";
				worksheet.Cells[1, 5].Value = "Cliente";
				worksheet.Cells[1, 6].Value = "Producto";
				worksheet.Cells[1, 7].Value = "Descripción caso";
				worksheet.Cells[1, 8].Value = "Prioridad caso";
				worksheet.Cells[1, 9].Value = "Estado caso";
				worksheet.Cells[1, 10].Value = "Quien reporta caso";
				worksheet.Cells[1, 11].Value = "Fecha cierre caso";


				// Aplicar formato a los encabezados
				using (var range = worksheet.Cells[1, 1, 1, 11])
				{
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
				}

				// Llenar el contenido de la tabla
				for (int i = 0; i < data.Count; i++)
				{
					worksheet.Cells[i + 2, 1].Value = data[i].TBL_PeriodoGarantia.TC_PeriodoGarantia.ToString();
					worksheet.Cells[i + 2, 2].Value = data[i].TBL_FechaCreacionCaso.TD_FechaCreacionCaso.ToString();
					worksheet.Cells[i + 2, 3].Value = data[i].TBL_Empleado.TC_Nombre.ToString() + " " + data[i].TBL_Empleado.TC_PrimerApellido.ToString();
					worksheet.Cells[i + 2, 4].Value = data[i].TBL_Municipalidad.TC_Nombre.ToString();
					worksheet.Cells[i + 2, 5].Value = data[i].TBL_Cliente.TC_Nombre.ToString();
					worksheet.Cells[i + 2, 6].Value = data[i].TBL_Producto.TC_Nombre.ToString();
					worksheet.Cells[i + 2, 7].Value = data[i].TBL_DescripcionCaso.TC_Descripcion.ToString();
					worksheet.Cells[i + 2, 8].Value = data[i].TBL_PrioridadCaso.TC_Nombre.ToString();
					worksheet.Cells[i + 2, 9].Value = data[i].TBL_EstadoCaso.TC_Nombre.ToString();
					worksheet.Cells[i + 2, 10].Value = data[i].TBL_Usuario.TC_Nombre.ToString() + " " + data[i].TBL_Usuario.TC_PrimerApellido.ToString();
					worksheet.Cells[i + 2, 11].Value = data[i].TBL_FechaCierreCaso.TD_FechaCierreCaso.ToString();

				}

				// Configurar el ancho de las columnas automáticamente
				worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

				// Convertir el archivo Excel en un arreglo de bytes
				byte[] excelBytes = package.GetAsByteArray();

				// Descargar el archivo
				return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Datos_casos.xlsx");
			}
		}


		protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
