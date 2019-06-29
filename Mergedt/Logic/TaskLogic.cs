using System.Data;
using System.Threading;
using Mergedt.Logic;

namespace MergeDt.Logic
{
    public class TaskLogic
    {
        Import import=new Import();
        Generate generate=new Generate();
        Export export=new Export();

        private int _taskid;
        private string _fileAddress;       //文件地址
        private DataTable _dt;             //获取dt(从EXCEL获取的DT)

        private DataTable _resultTable;   //返回DT
        private bool _resultMark;        //返回是否成功标记

        private DataTable _tempdt;         //返回运算成功的表头DT(导出时使用)
        private DataTable _tempdtldt;    //返回运算成功的表体DT(导出时使用)

        #region Set

            /// <summary>
            /// 中转ID
            /// </summary>
            public int TaskId { set { _taskid = value; } }

            /// <summary>
            /// //接收文件地址信息
            /// </summary>
            public string FileAddress { set { _fileAddress = value; } }

            /// <summary>
            /// 获取dt(从EXCEL获取的DT)
            /// </summary>
            public DataTable Data { set { _dt = value; } }

        #endregion

        #region Get

        /// <summary>
        ///返回DataTable至主窗体
        /// </summary>
        public DataTable RestulTable => _resultTable;

        /// <summary>
        ///  返回是否成功标记
        /// </summary>
        public bool ResultMark => _resultMark;

        /// <summary>
        /// 返回运算成功的表头DT(导出时使用)
        /// </summary>
        public DataTable Tempdt => _tempdt;

        /// <summary>
        /// 返回运算成功的表体DT(导出时使用)
        /// </summary>
        public DataTable Tempdtldt => _tempdtldt;

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
                    GenerateRecord(_dt);
                    break;
                //导出
                case 3:
                    ExportDtToExcel(_fileAddress,_tempdt,_tempdtldt);
                    break;
            }
        }

        /// <summary>
        /// 打开EXCEL并导入至DataTable功能
        /// </summary>
        private void OpenExcelImporttoDt(string fileAddress)
        {
            _resultTable = import.OpenExcelImporttoDt(fileAddress);
        }

        /// <summary>
        /// 运算
        /// </summary>
        private void GenerateRecord(DataTable dt)
        {
            //获取表头信息
            _tempdt = generate.Generatetemp(dt);
            //获取表体信息
            _tempdtldt = generate.GeneratetempEmpty(dt,_tempdt);
            //获取结果(若表头与表体都有值的话,就返回true)
            _resultMark = _tempdt.Rows.Count > 0 && _tempdtldt.Rows.Count > 0;
        }

        /// <summary>
        /// 导出
        /// </summary>
        private void ExportDtToExcel(string fileAddress,DataTable tempdt,DataTable tempdtldt)
        {
            _resultMark = export.ExportDtToExcel(fileAddress,tempdt,tempdtldt);
        }
    }
}
