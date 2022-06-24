using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Microsoft.Office.Interop.Excel;
using ReporteExcel02.Models;

namespace ReporteExcel02.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Peru()
        {
            PeruCrea("C:\\temp\\peru.xlsx");
            return File("C:\\temp\\peru.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Peru.xlsx");
        }

        public virtual void PeruCrea(string doc_excel)
        {
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = wb.Sheets.Add();
            //--------------------------------------------------------------------------

            //DaoPeru daoPeru = new DaoPeru();
            //List<departamentos> list = daoPeru.peruDepaProv();

            //int i = 1;

            //foreach(departamentos depa in list)
            //{
            //    string d = depa.departamento;

            //    sheet.Cells[i, 1] = d;
            //    i++;
            //    foreach (provincias prov in depa.provincias)
            //    {

            //        sheet.Cells[i, 2] = prov.provincia;

            //        d = "";
            //        i++;
            //    }
            //}

            //sheet.Columns.AutoFit();

            DaoPeru daoPeru = new DaoPeru();
            List<departamentos> list = daoPeru.peruDepaProvDist();

            int i = 1;

            foreach (departamentos depa in list)
            {
                string d = depa.departamento;

                sheet.Cells[i, 1] = d;
                i++;

                
                foreach (provincias prov in depa.provincias)
                {
                    string p = prov.provincia;
                    sheet.Cells[i, 1] = " ";
                    sheet.Cells[i, 2] = prov.provincia;
                    i++;

                    foreach (distritos dist in prov.distritos)
                    {
                        sheet.Cells[i, 2] = " ";
                        sheet.Cells[i, 3] = dist.distrito;
                        d = "";
                        p = "";
                        i++;
                    }
                }
            }

            sheet.Cells[i, 1] = "=COUNTA(A1:A" + (i - 2) + ")";
            sheet.Cells[i, 2] = "=COUNTA(B1:B" + (i - 2) + ")";
            sheet.Cells[i, 3] = "=COUNTA(C1:C" + (i - 2) + ")";

            sheet.Columns.AutoFit();

            //--------------------------------------------------------------------------
            wb.SaveAs(doc_excel,
                 XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close();
            excelApp.Quit();
        }
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}