using System;
using System.Collections.Generic;
using HostMgd.ApplicationServices;
using Multicad.Runtime;
using Multicad.DatabaseServices;
using DocumentFormat.OpenXml.Spreadsheet;
using NCadCustom.Code;

using App = HostMgd.ApplicationServices;
using Db = Teigha.DatabaseServices;
using Ed = HostMgd.EditorInput;
using System.Runtime.Intrinsics.Arm;
using Rtm = Teigha.Runtime;
using System.Windows.Forms;

namespace NCadCustom
{
    public class Commands : IExtensionApplication
    {
        public void Initialize()
        {
            App.DocumentCollection dm = App.Application.DocumentManager;
            Ed.Editor ed = dm.MdiActiveDocument.Editor;
            string msg = "PRPR_objxldata - импорт данных из файла внешней таблицы Params.xlsx";
            ed.WriteMessage(msg);
        }
        public void Terminate()
        {
        }

        static readonly string[] excelExtentions = { ".xlsx", ".xls", ".xlsb", ".xlsm" };

        /// <summary>
        /// Создание СПДС объектов на чертеже по таблице из эксель
        /// Параметры и ID объектов находятся в эксель
        /// </summary>
        /// 


        
        [Rtm.CommandMethod("PRPR_objxldata", Rtm.CommandFlags.Session)]
        public static void MainCreateObjBySpreadSheet()
        {
            InputJig jig = new InputJig();
            HostMgd.EditorInput.Editor ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            App.Document doc = App.Application.DocumentManager.MdiActiveDocument;



            //string paramsFilePath = jig.GetText("Укажите полный путь до файла параметров(Excel)", false);
            //paramsFilePath = paramsFilePath.Trim('"').ToLower();

            string dwgName = doc.Name; // метод получения полного пути и имени текущего dwg-файла
            int pos = dwgName.LastIndexOf("\\"); // позиция последнего слеша в полном пути до файла
            string dwgPath = dwgName.Remove(pos, dwgName.Length - pos); // Путь до dwg файла (без имени файла)
            string paramsFilePath = dwgPath + "\\Params.xlsx"; // путь до файла параметров (Excel)

            if (!File.Exists(paramsFilePath))
            {
                ed.WriteMessage("Выбран не существующий путь! Программа завершена.");
                return;
            }

            if (!excelExtentions.Contains(Path.GetExtension(paramsFilePath)))
            {
                ed.WriteMessage("Выбран не Excel файл! Программа завершена.");
                return;
            }

            try
            {
                ShWorker shWorker = new ShWorker(paramsFilePath);
                List<Row> dataRows = shWorker.dataRows;

                List<string> headers = shWorker.GetHeaders(shWorker.sst, dataRows.ElementAt(0));
                List<ObjectForInsert> objs = new List<ObjectForInsert>();

                // за исключением шапки - остальное строки с данными о деталях.
                Row rw = new Row();
                for (int iRow = 1; iRow < dataRows.Count; iRow++)
                {
                    rw = dataRows[iRow];
                    ObjectForInsert oneObject = new ObjectForInsert(shWorker.sst, headers, rw);
                    objs.Add(oneObject);
                }

                foreach (ObjectForInsert obj in objs)
                {
                    obj.PlaceToModelSpace();
                }

                ed.WriteMessage("Обработано!");
            }
            catch (Exception e)
            {
                ed.WriteMessage($"Ошибка : {e}");
            }
        }
    }
}
