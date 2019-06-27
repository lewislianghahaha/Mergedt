using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeDt.DB
{
    public class DtList
    {
        /// <summary>
        /// 表头
        /// </summary>
        /// <returns></returns>
        public DataTable Get_MeasureMentColordt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 6; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "BMMEASUREMENTID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "COLORCODE";
                        dc.DataType = Type.GetType("System.String"); 
                        break;
                    case 2:
                        dc.ColumnName = "FORMULAVERSIONDATE";
                        dc.DataType = Type.GetType("System.DateTime"); 
                        break;
                    case 3:
                        dc.ColumnName = "DIFFUSECOARSENESS";
                        dc.DataType = Type.GetType("System.Decimal");  
                        break;
                    case 4:
                        dc.ColumnName = "CREATEDDATE";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                    case 5:
                        dc.ColumnName = "MEASUREMENTTIME";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 表体
        /// </summary>
        /// <returns></returns>
        public DataTable Get_MeasureMentEmptydt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 39; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "BMAID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "ANGLE";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 2:
                        dc.ColumnName = "BMMEASUREMENTID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 3:
                        dc.ColumnName = "L";
                        dc.DataType = Type.GetType("System.Decimal"); 
                        break;
                    case 4:
                        dc.ColumnName = "A";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 5:
                        dc.ColumnName = "B";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 6:
                        dc.ColumnName = "C";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 7:
                        dc.ColumnName = "H";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 8:
                        dc.ColumnName = "BMA_R400";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 9:
                        dc.ColumnName = "BMA_R410";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 10:
                        dc.ColumnName = "BMA_R420";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 11:
                        dc.ColumnName = "BMA_R430";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 12:
                        dc.ColumnName = "BMA_R440";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 13:
                        dc.ColumnName = "BMA_R450";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 14:
                        dc.ColumnName = "BMA_R460";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 15:
                        dc.ColumnName = "BMA_R470";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 16:
                        dc.ColumnName = "BMA_R480";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 17:
                        dc.ColumnName = "BMA_R490";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 18:
                        dc.ColumnName = "BMA_R500";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 19:
                        dc.ColumnName = "BMA_R510";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 20:
                        dc.ColumnName = "BMA_R520";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 21:
                        dc.ColumnName = "BMA_R530";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 22:
                        dc.ColumnName = "BMA_R540";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 23:
                        dc.ColumnName = "BMA_R550";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 24:
                        dc.ColumnName = "BMA_R560";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 25:
                        dc.ColumnName = "BMA_R570";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 26:
                        dc.ColumnName = "BMA_R580";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 27:
                        dc.ColumnName = "BMA_R590";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 28:
                        dc.ColumnName = "BMA_R600";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 29:
                        dc.ColumnName = "BMA_R610";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 30:
                        dc.ColumnName = "BMA_R620";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 31:
                        dc.ColumnName = "BMA_R630";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 32:
                        dc.ColumnName = "BMA_R640";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 33:
                        dc.ColumnName = "BMA_R650";
                        dc.DataType=Type.GetType("System.Decimal");
                        break;
                    case 34:
                        dc.ColumnName = "BMA_R660";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 35:
                        dc.ColumnName = "BMA_R670";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 36:
                        dc.ColumnName = "BMA_R680";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 37:
                        dc.ColumnName = "BMA_R690";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 38:
                        dc.ColumnName = "BMA_R700";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }
    }
}
