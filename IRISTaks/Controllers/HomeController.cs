using IRISTaks.Models;
using System;
using System.Web.Mvc;

namespace IRISTaks.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var responseModel = new Response();
            return View(responseModel);
        }

        [HttpPost]
        public ActionResult ExtractExcelHeader(ExcelForm modal)
        {
            var _response = new Response();
            if (modal != null)
            {
                if (string.IsNullOrEmpty(modal.Request))
                {
                    _response.Error = "Invalid request";
                }
                else
                {
                    if (modal.Action == ActionType.LTC)
                    {
                        _response.Result = GetExcelColumnNumber(modal.Request.ToUpper());
                    }
                    else if (modal.Action == ActionType.CTL)
                    {
                        _response.Result = GetExcelColumnName(int.Parse(modal.Request));
                    }
                }
            }
            else
            {
                _response.Error = "Invalid Request!";
            }
            return View("Index", _response);
        }

        private string GetExcelColumnNumber(string columnName)
        {
            int result = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                result *= 26;
                result += columnName[i] - 'A' + 1;
            }
            return result.ToString();
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

    }
}