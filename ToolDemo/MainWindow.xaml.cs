using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Excel;
using Action = System.Action;
using DataTable = System.Data.DataTable;
using Window = System.Windows.Window;

namespace ToolDemo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            Init();
        }

        private MSSQLHelper _sqlHelper;
        public MSSQLHelper SqlHelper
        {
            get
            {
                if (_sqlHelper == null)
                {
                    _sqlHelper = new MSSQLHelper();
                }

                return _sqlHelper;
            }
            set { _sqlHelper = value; }
        }

        private List<ResultInfo> _resultInfos;

        public List<ResultInfo> ResultInfos
        {
            get
            {
                if (_resultInfos == null)
                {
                    _resultInfos = new List<ResultInfo>();
                }

                return _resultInfos;
            }
            set
            {
                _resultInfos = value;
            }
        }

        /// <summary>
        /// 初始化
        /// </summary>
        private void Init()
        {
            DatePickerStart.Text = DateTime.Today.ToShortDateString();
            DatePickerEnd.Text = DateTime.Today.ToString("yyyy-MM-dd 23:59:59");

            DataContext = this;
            _sqlHelper = new MSSQLHelper();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBoxIpAddress.Text))
            {
                MessageBox.Show("IP地址不能为空");
                return;
            }

            if (string.IsNullOrEmpty(TextBoxSQLName.Text))
            {
                MessageBox.Show("数据库名不能为空");
                return;
            }

            if (string.IsNullOrEmpty(TextBoxUsr.Text))
            {
                MessageBox.Show("登陆名称不能为空");
                return;
            }

            //连接数据库
            bool bResult = _sqlHelper.Connetction(string.Format("Data Source={0};Initial Catalog={1};User ID={2};Pwd={3}",
                 TextBoxIpAddress.Text, TextBoxSQLName.Text, TextBoxUsr.Text, PasswordBoxPwd.Password));

            //数据库是否连接成功
            if (!bResult)
            {
                MessageBox.Show("连接数据库失败!");
                return;
            }
            MessageBox.Show("数据库连接成功");
        }

        private void ButtonQuery_OnClick(object sender, RoutedEventArgs e)
        {
            if (_sqlHelper == null)
            {
                MessageBox.Show("请先连接数据库！");
                return;
            }

            var ds = _sqlHelper.Query("SELECT " +
                                      "tbl_MemberInfo.UName," +
                                      "tbl_MemberInfo.UTel," +
                                      "tbl_MemberInfo.UGender," +
                                      "tbl_MemberInfo.UNational," +
                                      "tbl_MemberInfo.UCerNo," +
                                      "tbl_MemberRecordInOut.PassTime," +
                                      "tbl_MemberRecordInOut.Result," +
                                      "tbl_MemberRecordInOut.GateID," +
                                      "tb_Lessee.Company " +
                                      "FROM tbl_MemberInfo,tbl_MemberRecordInOut,tb_Lessee " +
                                      "WHERE " +
                                      "tbl_MemberRecordInOut.CardNo = tbl_MemberInfo.CardNo " +
                                      " and tbl_MemberInfo.UCompanyID = tb_Lessee.LesseeID " +
                                      " and tbl_MemberRecordInOut.InOrOut = 0" +
                                      " and tbl_MemberRecordInOut.PassTime between " +
                                      "'" +
                                      DatePickerStart.DisplayDate.ToString("yyyy-MM-dd 00:00:00") +
                                      "' and  '" +
                                      DatePickerEnd.DisplayDate.ToString("yyyy-MM-dd 23:59:59") +
                                      "'" +
                                      " ORDER BY tbl_MemberRecordInOut.PassTime");
            try
            {
                foreach (DataTable dsTable in ds.Tables)
                {
                    _resultInfos = dsTable.AsEnumerable().Select(dataRow => new ResultInfo
                    {
                        Name = dataRow.Field<string>("UName"),
                        Phone = dataRow.Field<string>("UTel"),
                        Sex = dataRow.Field<string>("UGender"),
                        Country = dataRow.Field<string>("UNational"),
                        IDCard = dataRow.Field<string>("UCerNo"),
                        PassTime = dataRow.Field<DateTime>("PassTime").ToString("yyyy-MM-dd hh:mm:ss"),
                        PassState = dataRow.Field<string>("Result"),
                        PassLocation = dataRow.Field<string>("GateID"),
                        Company = dataRow.Field<string>("Company"),
                    }).ToList();
                }

                _resultInfos = _resultInfos.Where((x, i) => _resultInfos.FindIndex(z => z.Name == x.Name) == i).ToList();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                throw;
            }

            DataGrid.ItemsSource = ResultInfos;
        }

        private void ButtonExport_OnClick(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Execel文件|*.xlsx";
            sfd.FileName = (DateTime.Today.ToShortDateString() + "导出结果").Replace('/', '-');

            if (sfd.ShowDialog() == false)
            {
                return;
            }

            Export(DataGrid, sfd.FileName, _resultInfos);
        }

        /// <summary>
        /// 导出实现
        /// </summary>
        /// <param name="dt"></param>
        public static async void Export<T>(DataGrid dt, string saveFilePath, IEnumerable<T> sources)
        {
            //判断本机是否安装了Execel程序
            Type type = Type.GetTypeFromProgID("Excel.Application");
            if (type == null)
            {
                MessageBox.Show("本机未安装Execel程序，无法实现导出以及打印操作！");
            }

            try
            {

                //实例化Execel对象
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                //新建工作簿
                Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Add(true);
                //新建工作表
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workBook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

                ////开启线程运行
                await System.Threading.Tasks.Task.Run(new Action(() =>
                {
                    //获取泛型类型
                    Type t = typeof(T);

                    //跨线程调度控件
                    //首先把列头输出
                    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            worksheet.Cells[1, i + 1] = dt.Columns[i].Header;
                        }
                    }
                    ));
                    //设置表头
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        //设置首列样式
                        //worksheet.Cells[1, i + 1] = dt.Columns[i].Header;
                        Microsoft.Office.Interop.Excel.Range headRange = worksheet.Cells[1, i + 1] as Microsoft.Office.Interop.Excel.Range; //获取表头单元格
                        headRange.Font.Name = "宋体"; //设置字体
                        headRange.Font.Size = 12; //字体大小
                        headRange.Font.Bold = true; //加粗显示
                        headRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; //水平居中
                        headRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter; //垂直居中
                        headRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; //设置边框
                        headRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin; //边框常规粗细
                        headRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; //设置边框

                        System.Windows.Data.Binding binding =
                            dt.Columns[i].ClipboardContentBinding as System.Windows.Data.Binding;
                        if (binding == null)
                        {
                            continue;
                        }

                        string bindingPath = binding.Path.Path;
                        PropertyInfo propertyInfo = t.GetProperty(bindingPath);
                        int row = 2;
                        foreach (var item in sources)
                        {
                            worksheet.Cells[row, i + 1] = "'" + propertyInfo.GetValue(item);

                            headRange = worksheet.Cells[row, i + 1] as Microsoft.Office.Interop.Excel.Range; //获取头单元格
                            headRange.Font.Name = "宋体"; //设置字体
                            headRange.Font.Size = 12; //字体大小
                            headRange.Font.Bold = false; //加粗显示
                            headRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; //水平居中
                            headRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter; //垂直居中
                            headRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; //设置边框
                            headRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin; //边框常规粗细
                            headRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; //设置边框
                            headRange.EntireColumn.AutoFit(); //自动调整列宽度

                            row++;
                        }
                    }
                }));

                //保存Execel
                workBook.SaveAs(saveFilePath);
                excelApp.Quit();
                excelApp = null;
                //获取Excel相关所有后台应用程序
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill(); //没有更好的方法,只有杀掉进程
                }

                GC.Collect();

                if (System.Windows.MessageBox.Show("已成功导出，是否打开文件？", "提示", MessageBoxButton.YesNo) ==
                    System.Windows.MessageBoxResult.No)
                {
                    return;
                }

                System.Diagnostics.Process.Start(saveFilePath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
