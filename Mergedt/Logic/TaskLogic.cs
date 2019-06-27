using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mergedt.Logic;

namespace MergeDt.Logic
{
    public class TaskLogic
    {
        Import import=new Import();
        Generate generate=new Generate();
        Export export=new Export();

        private int _taskid;
        private string _fileAddress;        //文件地址
        private DataTable _resultTable;     //记录集DATATABLE（从EXCEL导入的DataTable）



        #region Set

            /// <summary>
            /// 中转ID
            /// </summary>
            public int TaskId { set { _taskid = value; } }

            /// <summary>
            /// //接收文件地址信息
            /// </summary>
            public string FileAddress { set { _fileAddress = value; } }

        #endregion

        #region Get
            /// <summary>
            ///返回DataTable至主窗体
            /// </summary>
            public DataTable RestulTable => _resultTable;

            

        #endregion


        public void StartTask()
        {
            Thread.Sleep(1000);

            switch (_taskid)
            {
                //打开EXCEL并导入至DataTable功能
                case 1:
                    OpenExcelImporttoDt(_fileAddress);
                    break;
                //运算
                case 2:
                    GenerateRecord();
                    break;
                //导出
                case 3:
                    ExportDtToExcel();
                    break;
            }
        }

        /// <summary>
        /// 打开EXCEL并导入至DataTable功能
        /// </summary>
        private void OpenExcelImporttoDt(string fileAddress)
        {
            
        }

        /// <summary>
        /// 运算
        /// </summary>
        private void GenerateRecord()
        {
            
        }

        /// <summary>
        /// 导出
        /// </summary>
        private void ExportDtToExcel()
        {
            
        }

    }
}
