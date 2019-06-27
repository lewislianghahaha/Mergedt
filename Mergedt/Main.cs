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
        private DataTable _exceldt;

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

                if (task.RestulTable.Rows.Count == 0) throw new Exception("不能成功导入EXCEL内容,请检查模板是否正确.");
                else
                {
                    MessageBox.Show($"导入成功,请按'运算'按钮继续", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

    }
}
