﻿using System;
using System.Collections.Generic;
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
using Soporte_averias.Models;
using OfficeOpenXml;
using Soporte_averias.Permissions;

namespace Soporte_averias.Controllers
{


	[ValidarSesionAttribute]
	[PermisosRol(Rol.Administrador)]
	public class PrioridadCasoController : Controller
    {
        private SOPORTEEntities db = new SOPORTEEntities();

		// GET: PrioridadCaso
		[HttpGet]
		public ActionResult Index(int? page, string searchText)
		{

			ViewBag.Title = "Lista de prioridades de caso";


			int pageSize = 10;
			int pageNumber = (page ?? 1);
			ViewBag.PageNumber = pageNumber;
			IEnumerable<TBL_PrioridadCaso> prioridad;
			prioridad = db.TBL_PrioridadCaso.AsQueryable();

			if (!string.IsNullOrEmpty(searchText))
			{

				prioridad = prioridad.Where(m => m.TC_Nombre.Contains(searchText));
			}

			int totalItems = prioridad.Count(); //Cant. elementos totales
			int totalPages = (int)Math.Ceiling((double)totalItems / pageSize); //Cant. total de páginas
			ViewBag.totalPages = totalPages;
			ViewBag.CurrentFilter = searchText;
			var prioridadesOrdenados = prioridad.OrderBy(m => m.TC_Nombre);
			var prioridadesPaginas = prioridadesOrdenados.Skip((pageNumber - 1) * pageSize).Take(pageSize);

			return View(prioridadesPaginas);

		}

		// GET: PrioridadCaso/Details/5
		public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TBL_PrioridadCaso tBL_PrioridadCaso = db.TBL_PrioridadCaso.Find(id);
            if (tBL_PrioridadCaso == null)
            {
                return HttpNotFound();
            }
            return View(tBL_PrioridadCaso);
        }

        // GET: PrioridadCaso/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: PrioridadCaso/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "TN_IdPrioridadCaso,TC_Nombre,TC_Descripcion")] TBL_PrioridadCaso tBL_PrioridadCaso)
        {
            if (ModelState.IsValid)
            {
                db.TBL_PrioridadCaso.Add(tBL_PrioridadCaso);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(tBL_PrioridadCaso);
        }

        // GET: PrioridadCaso/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TBL_PrioridadCaso tBL_PrioridadCaso = db.TBL_PrioridadCaso.Find(id);
            if (tBL_PrioridadCaso == null)
            {
                return HttpNotFound();
            }
            return View(tBL_PrioridadCaso);
        }

        // POST: PrioridadCaso/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "TN_IdPrioridadCaso,TC_Nombre,TC_Descripcion")] TBL_PrioridadCaso tBL_PrioridadCaso)
        {
            if (ModelState.IsValid)
            {
                db.Entry(tBL_PrioridadCaso).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(tBL_PrioridadCaso);
        }

		public ActionResult ExportToPdf(string searchText, int? page)
		{
			int pageNumber = page ?? 1;

			var actividad = db.TBL_PrioridadCaso.AsQueryable();

			if (!string.IsNullOrEmpty(searchText))
			{
				actividad = actividad.Where(m => m.TC_Nombre.Contains(searchText));
			}
			actividad = actividad.OrderBy(m => m.TC_Nombre);
			var pagedActividad = actividad.ToList();

			// Crear el documento PDF
			Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
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
			PdfPTable pdfTable = new PdfPTable(2);
			pdfTable.WidthPercentage = 100;


			// Crear estilo para los encabezados (negrita y fondo gris)
			Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD);
			BaseColor headerBackgroundColor = new BaseColor(192, 192, 192); // Gris claro

			// Añadir encabezados a la tabla con el estilo
			PdfPCell headerCell;

			headerCell = new PdfPCell(new Phrase("Nombre", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);

			headerCell = new PdfPCell(new Phrase("Descripción", headerFont));
			headerCell.BackgroundColor = headerBackgroundColor;
			headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
			pdfTable.AddCell(headerCell);



			foreach (var item in pagedActividad)
			{
				pdfTable.AddCell(item.TC_Nombre.ToString());
				pdfTable.AddCell(item.TC_Descripcion.ToString());

			}



			pdfDoc.Add(pdfTable);
			pdfDoc.Close();

			byte[] bytes = memoryStream.ToArray();
			memoryStream.Close();
			return File(bytes, "application/pdf", "Datos_prioridades.pdf");
		}

		//Crear ExportToExcel
		public ActionResult ExportToExcel(string searchText)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			var actividad = db.TBL_PrioridadCaso.AsQueryable();

			if (!string.IsNullOrEmpty(searchText))
			{
				actividad = actividad.Where(m => m.TC_Nombre.ToString().Contains(searchText));
			}
			else
			{
				ViewData["Mensaje"] = "*No se encontraron datos*";
			}

			// Ordenar los datos por FECHA
			actividad = actividad.OrderBy(m => m.TC_Nombre);

			var data = actividad.ToList();

			// Crear el archivo Excel utilizando EPPlus
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Prioridades casos");

				// Encabezados
				worksheet.Cells[1, 1].Value = "Nombre";
				worksheet.Cells[1, 2].Value = "Descripción";


				// Aplicar formato a los encabezados
				using (var range = worksheet.Cells[1, 1, 1, 10])
				{
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
				}

				// Llenar el contenido de la tabla
				for (int i = 0; i < data.Count; i++)
				{
					worksheet.Cells[i + 2, 1].Value = data[i].TC_Nombre.ToString();
					worksheet.Cells[i + 2, 2].Value = data[i].TC_Descripcion.ToString();

				}

				// Configurar el ancho de las columnas automáticamente
				worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

				// Convertir el archivo Excel en un arreglo de bytes
				byte[] excelBytes = package.GetAsByteArray();

				// Descargar el archivo
				return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Datos_prioridades.xlsx");
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
