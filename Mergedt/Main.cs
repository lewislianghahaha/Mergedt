using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MergeDt.Logic;

namespace Mergedt
{
    public partial class Main : Form
    {
        TaskLogic task=new TaskLogic();
        Load load = new Load();

        //存放从EXCEL导入的DT
        private DataTable _fromexceldt;

        //存放运算完成的DT(表头)
        private DataTable _exceldt;
        //存放运算完成的DT(表体)
        private DataTable _exceldtdtl;

        public Main()
        {
            InitializeComponent();
            OnRegisterEvents();
        }

        private void OnRegisterEvents()
        {
            btnimport.Click += Btnimport_Click;
            btnGenerate.Click += BtnGenerate_Click;
            btnExport.Click += BtnExport_Click;
            tmclose.Click += Tmclose_Click;
        }

        /// <summary>
        /// 打开EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnimport_Click(object sender, EventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog { Filter = "Xlsx文件|*.xlsx" };
                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileAdd = openFileDialog.FileName;

                //将所需的值赋到Task类内
                task.TaskId = 1;
                task.FileAddress = fileAdd;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();
                _fromexceldt = task.RestulTable;

                if (_fromexceldt.Rows.Count == 0) throw new Exception("不能成功导入EXCEL内容,请检查模板是否正确.");
                else
                {
                    MessageBox.Show($"导入成功,请按'运算'按钮继续", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                if(_fromexceldt.Rows.Count==0)throw new Exception("没有导入EXCEL记录,不能进行运算");
                task.TaskId = 2;
                task.Data = _fromexceldt;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                _exceldt = task.Tempdt;
                _exceldtdtl = task.Tempdtldt;

                if(!task.ResultMark) throw new Exception("运算异常");
                else
                {
                    var clickMessage = $"运算成功,是否进行导出至Excel?";

                    if (MessageBox.Show(clickMessage, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        ExportDttoExcel(_exceldt,_exceldtdtl);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if(_exceldt.Rows.Count==0 || _exceldtdtl.Rows.Count==0) throw new Exception("没有运算成功,不能进行导出");
                ExportDttoExcel(_exceldt,_exceldtdtl);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tmclose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        ///子线程使用(重:用于监视功能调用情况,当完成时进行关闭LoadForm)
        /// </summary>
        private void Start()
        {
            task.StartTask();

            //当完成后将Form2子窗体关闭
            this.Invoke((ThreadStart)(() => {
                load.Close();
            }));
        }

        /// <summary>
        /// 导出DT至EXCEL
        /// </summary>
        /// <param name="tempdt">获取要导出的表头信息 MEASUREMENT_COLOR</param>
        /// <param name="tempdtdtl">获取要导出的表体信息 MEASUREMENT</param>
        void ExportDttoExcel(DataTable tempdt,DataTable tempdtdtl)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog { Filter = "Xlsx文件|*.xlsx" };
                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileAdd = saveFileDialog.FileName;

                //将所需的值赋到Task类内
                task.TaskId = 3;
                task.FileAddress = fileAdd;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                if (!task.ResultMark) throw new Exception("导出异常");
                else
                {
                    MessageBox.Show($"导出成功!可从EXCEL中查阅导出效果", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}

