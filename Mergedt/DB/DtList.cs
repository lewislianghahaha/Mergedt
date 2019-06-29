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
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 4:
                        dc.ColumnName = "CREATEDDATE";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                    case 5:
                        dc.ColumnName = "MEASUREMENTTIME";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                        #region ColumnName
                        //case 6:
                        //    dc.ColumnName = "L";
                        //    dc.DataType = Type.GetType("System.Double");
                        //    break;
                        //case 7:
                        //    dc.ColumnName = "A";
                        //    dc.DataType = Type.GetType("System.Double");
                        //    break;
                        //case 8:
                        //    dc.ColumnName = "B";
                        //    dc.DataType = Type.GetType("System.Double");
                        //    break;
                        //case 9:
                        //    dc.ColumnName = "C";
                        //    dc.DataType = Type.GetType("System.Double");
                        //    break;
                        //case 10:
                        //    dc.ColumnName = "H";
                        //    dc.DataType = Type.GetType("System.Double");
                        //    break;
                        #endregion
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
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 4:
                        dc.ColumnName = "A";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 5:
                        dc.ColumnName = "B";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 6:
                        dc.ColumnName = "C";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 7:
                        dc.ColumnName = "H";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 8:
                        dc.ColumnName = "BMA_R400";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 9:
                        dc.ColumnName = "BMA_R410";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 10:
                        dc.ColumnName = "BMA_R420";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 11:
                        dc.ColumnName = "BMA_R430";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 12:
                        dc.ColumnName = "BMA_R440";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 13:
                        dc.ColumnName = "BMA_R450";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 14:
                        dc.ColumnName = "BMA_R460";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 15:
                        dc.ColumnName = "BMA_R470";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 16:
                        dc.ColumnName = "BMA_R480";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 17:
                        dc.ColumnName = "BMA_R490";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 18:
                        dc.ColumnName = "BMA_R500";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 19:
                        dc.ColumnName = "BMA_R510";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 20:
                        dc.ColumnName = "BMA_R520";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 21:
                        dc.ColumnName = "BMA_R530";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 22:
                        dc.ColumnName = "BMA_R540";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 23:
                        dc.ColumnName = "BMA_R550";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 24:
                        dc.ColumnName = "BMA_R560";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 25:
                        dc.ColumnName = "BMA_R570";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 26:
                        dc.ColumnName = "BMA_R580";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 27:
                        dc.ColumnName = "BMA_R590";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 28:
                        dc.ColumnName = "BMA_R600";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 29:
                        dc.ColumnName = "BMA_R610";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 30:
                        dc.ColumnName = "BMA_R620";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 31:
                        dc.ColumnName = "BMA_R630";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 32:
                        dc.ColumnName = "BMA_R640";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 33:
                        dc.ColumnName = "BMA_R650";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 34:
                        dc.ColumnName = "BMA_R660";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 35:
                        dc.ColumnName = "BMA_R670";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 36:
                        dc.ColumnName = "BMA_R680";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 37:
                        dc.ColumnName = "BMA_R690";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 38:
                        dc.ColumnName = "BMA_R700";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 获取导入模板临时表(导入时使用)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_Importdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 113; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "ID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "bhcode";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "rank";
                        dc.DataType = Type.GetType("System.Double");  
                        break;
                    case 3:
                        dc.ColumnName = "L_D65_25";
                        dc.DataType = Type.GetType("System.Double"); 
                        break;
                    case 4:
                        dc.ColumnName = "a_D65_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 5:
                        dc.ColumnName = "b_D65_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 6:
                        dc.ColumnName = "C_D65_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 7:
                        dc.ColumnName = "h_D65_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 8:
                        dc.ColumnName = "w400nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 9:
                        dc.ColumnName = "w410nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 10:
                        dc.ColumnName = "w420nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 11:
                        dc.ColumnName = "w430nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 12:
                        dc.ColumnName = "w440nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 13:
                        dc.ColumnName = "w450nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 14:
                        dc.ColumnName = "w460nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 15:
                        dc.ColumnName = "w470nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 16:
                        dc.ColumnName = "w480nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 17:
                        dc.ColumnName = "w490nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 18:
                        dc.ColumnName = "w500nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 19:
                        dc.ColumnName = "w510nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 20:
                        dc.ColumnName = "w520nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 21:
                        dc.ColumnName = "w530nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 22:
                        dc.ColumnName = "w540nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 23:
                        dc.ColumnName = "w550nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 24:
                        dc.ColumnName = "w560nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 25:
                        dc.ColumnName = "w570nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 26:
                        dc.ColumnName = "w580nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 27:
                        dc.ColumnName = "w590nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 28:
                        dc.ColumnName = "w600nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 29:
                        dc.ColumnName = "w610nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 30:
                        dc.ColumnName = "w620nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 31:
                        dc.ColumnName = "w630nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 32:
                        dc.ColumnName = "w640nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 33:
                        dc.ColumnName = "w650nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 34:
                        dc.ColumnName = "w660nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 35:
                        dc.ColumnName = "w670nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 36:
                        dc.ColumnName = "w680nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 37:
                        dc.ColumnName = "w690nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 38:
                        dc.ColumnName = "w700nm_25";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 39:
                        dc.ColumnName = "L_D65_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 40:
                        dc.ColumnName = "a_D65_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 41:
                        dc.ColumnName = "b_D65_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 42:
                        dc.ColumnName = "C_D65_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 43:
                        dc.ColumnName = "h_D65_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 44:
                        dc.ColumnName = "w400nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 45:
                        dc.ColumnName = "w410nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 46:
                        dc.ColumnName = "w420nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 47:
                        dc.ColumnName = "w430nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 48:
                        dc.ColumnName = "w440nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 49:
                        dc.ColumnName = "w450nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 50:
                        dc.ColumnName = "w460nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 51:
                        dc.ColumnName = "w470nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 52:
                        dc.ColumnName = "w480nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 53:
                        dc.ColumnName = "w490nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 54:
                        dc.ColumnName = "w500nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 55:
                        dc.ColumnName = "w510nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 56:
                        dc.ColumnName = "w520nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 57:
                        dc.ColumnName = "w530nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 58:
                        dc.ColumnName = "w540nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 59:
                        dc.ColumnName = "w550nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 60:
                        dc.ColumnName = "w560nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 61:
                        dc.ColumnName = "w570nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 62:
                        dc.ColumnName = "w580nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 63:
                        dc.ColumnName = "w590nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 64:
                        dc.ColumnName = "w600nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 65:
                        dc.ColumnName = "w610nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 66:
                        dc.ColumnName = "w620nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 67:
                        dc.ColumnName = "w630nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 68:
                        dc.ColumnName = "w640nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 69:
                        dc.ColumnName = "w650nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 70:
                        dc.ColumnName = "w660nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 71:
                        dc.ColumnName = "w670nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 72:
                        dc.ColumnName = "w680nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 73:
                        dc.ColumnName = "w690nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 74:
                        dc.ColumnName = "w700nm_45";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 75:
                        dc.ColumnName = "L_D65_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 76:
                        dc.ColumnName = "a_D65_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 77:
                        dc.ColumnName = "b_D65_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 78:
                        dc.ColumnName = "C_D65_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 79:
                        dc.ColumnName = "h_D65_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 80:
                        dc.ColumnName = "w400nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 81:
                        dc.ColumnName = "w410nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 82:
                        dc.ColumnName = "w420nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 83:
                        dc.ColumnName = "w430nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 84:
                        dc.ColumnName = "w440nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 85:
                        dc.ColumnName = "w450nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 86:
                        dc.ColumnName = "w460nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 87:
                        dc.ColumnName = "w470nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 88:
                        dc.ColumnName = "w480nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 89:
                        dc.ColumnName = "w490nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 90:
                        dc.ColumnName = "w500nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 91:
                        dc.ColumnName = "w510nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 92:
                        dc.ColumnName = "w520nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 93:
                        dc.ColumnName = "w530nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 94:
                        dc.ColumnName = "w540nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 95:
                        dc.ColumnName = "w550nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 96:
                        dc.ColumnName = "w560nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 97:
                        dc.ColumnName = "w570nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 98:
                        dc.ColumnName = "w580nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 99:
                        dc.ColumnName = "w590nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 100:
                        dc.ColumnName = "w600nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 101:
                        dc.ColumnName = "w610nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 102:
                        dc.ColumnName = "w620nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 103:
                        dc.ColumnName = "w630nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 104:
                        dc.ColumnName = "w640nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 105:
                        dc.ColumnName = "w650nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 106:
                        dc.ColumnName = "w660nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 107:
                        dc.ColumnName = "w670nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 108:
                        dc.ColumnName = "w680nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 109:
                        dc.ColumnName = "w690nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 110:
                        dc.ColumnName = "w700nm_110";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 111:
                        dc.ColumnName = "datatime";
                        dc.DataType = Type.GetType("System.DateTime"); 
                        break;
                    case 112:
                        dc.ColumnName = "pfversion";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }
    }
}
