﻿using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Multicad.Geometry;
using Multicad;
using Multicad.Objects;

namespace NCadCustom.Code
{
    internal class ObjectForInsert
    {
        /// <summary>
        /// Парсинг строки эксель в объект который в дальнейшем помещается на чертеж.
        /// Завязано на номера столбцов
        /// </summary>
        /// <param name="sst"></param>
        /// <param name="headers"> названия шапок из эксель</param>
        /// <param name="objectData">строка эксель</param>
        internal ObjectForInsert(SharedStringTable sst, List<string> headers, Row objectData)
        {
            CreateObject(sst, headers, objectData);
        }
        internal string dataBaseID { get; set; } //hexadecimal string 16
        internal double x_coord { get; set; }
        internal double y_coord { get; set; }
        internal Dictionary<string, string> objParams { get; set; } // Словарем определен тип данных пары значений - Строка, Строка
        private void CreateObject(SharedStringTable sst, List<string> headers, Row objectData)
        {
            IEnumerable<Cell> cells = objectData.Elements<Cell>();
            objParams = new Dictionary<string, string>();

            Cell cl = cells.ElementAt(0);
            if (cl.DataType != null && cl.DataType == CellValues.SharedString)
            {
                int ssid = int.Parse(cl.CellValue.Text);
                string cellValue = sst.ChildElements[ssid].InnerText;
                dataBaseID = cellValue;
            }

            // X,Y coords parse
            double doubleParsedValue = 0;
            if (double.TryParse(cells.ElementAt(1).CellValue.Text, out doubleParsedValue))
            {
                x_coord = doubleParsedValue;
            }
            if (double.TryParse(cells.ElementAt(2).CellValue.Text, out doubleParsedValue))
            {
                y_coord = doubleParsedValue;
            }

            // начинаются параметры для объектов БД
            for (int iCol = 3; iCol < cells.Count(); iCol++)
            {

                    if (cells.ElementAt(iCol).DataType != null && cells.ElementAt(iCol).DataType == CellValues.SharedString)
                    {
                        int ssid = int.Parse(cells.ElementAt(iCol).CellValue.Text);
                        string cellValue = sst.ChildElements[ssid].InnerText; // Извлечение текстового значения ячейки тблицы
                        
                        objParams.Add(headers[iCol], cellValue);

                    }

            }
        }

        /// <summary>
        /// Помещение объекта в пространство модели
        /// </summary>
        internal void PlaceToModelSpace()
        {
            McParametricObject parObj = new McParametricObject(false);
            parObj.DbEntity.AddToCurrentDocument();

            long id = Convert.ToInt64(dataBaseID, 16);
            parObj.Initialize(id);
            parObj.SetImplementationAndProcess("Implementation1");
            parObj.ViewInXY = McParametricObject.ViewType.Front; // Вид спереди

            List<ExValue> paramsToChange = new List<ExValue>();
            foreach (KeyValuePair<string, string> paramsPair in objParams)
            {
                paramsToChange.Add(new ExValue(paramsPair.Key, paramsPair.Value)); // Вставка пары строковых значений "ИмяПараметра-ЗначениеПараметра"
            }

            parObj.Change(paramsToChange, true);

            Matrix3d tfm = Matrix3d.MakeTranslation(x_coord, y_coord, 0);
            parObj.DbEntity.Transform(tfm);
        }

    }
}
