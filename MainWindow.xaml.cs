using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System;
using System.Collections.Generic;
using System.IO;
using ADOX;

namespace FindAccess
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private string filePath;
        private string Rowstr;
        private List<string> tableNameList;
        private OleDbConnection Rowtempconn;

        // 选择access数据库
        private void Choose_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "支持的文件格式|*.accdb";
            if (file.ShowDialog() != System.Windows.Forms.DialogResult.OK) {
                return;
            }
            filePath = file.FileName;          // 包括文件的全路径
            string fileExtension = System.IO.Path.GetExtension(filePath).ToLower();           // 获取后缀名
            PROJECT_NAME.Text = filePath;
            Rowstr = @"Provider=Microsoft.Ace.OLEDB.12.0;Jet OLEDB:DataBase Password=;Data Source=" + filePath + ";";
            Rowtempconn = new OleDbConnection(Rowstr);
            Rowtempconn.Open();
            Catalog catalog = new Catalog();
            ADODB.Connection cn = new ADODB.Connection();
            cn.Open("Provider=Microsoft.Ace.OLEDB.12.0;Jet OLEDB:DataBase Password=;Data Source=" + filePath, null, null, -1);
            catalog.ActiveConnection = cn;
            ADOX.Table table = catalog.Tables["dbo_BCF含水层属性表"];
            ADOX.Column col = table.Columns["LYRID"];
            System.Windows.MessageBox.Show(table.Columns[1].Name);
            table.Columns[1].Name = "aaa";
            //col.Properties["Description"].Value = "abc";

            tableNameList = GetTableNameList(SheetNameList, Rowtempconn);
        }


        private void btnCancle_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


        private void btnSure_Click(object sender, RoutedEventArgs e)
        {
            using (SaveFileDialog saveFileDialog1 = new SaveFileDialog()) {
                saveFileDialog1.Title = "另存为";
                saveFileDialog1.FileName = SheetNameList.Text + ".csv"; //设置默认另存为的名字，可选
                saveFileDialog1.Filter = "CSV 文件(*.csv)|";
                saveFileDialog1.AddExtension = true;
                if (saveFileDialog1.ShowDialog().ToString() == "OK") {
                    DataTable table = Rowtempconn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, SheetNameList.Text, null });
                    List<string> name = GetTableFieldNameList(Rowtempconn, SheetNameList.Text);
                    List<string> description = new List<string>();
                    foreach (string s in name) {
                        for (int i = 0; i < table.Rows.Count; i++) {
                            if (s == table.Rows[i]["COLUMN_NAME"].ToString()) {
                                description.Add(table.Rows[i]["DESCRIPTION"].ToString());
                                break;
                            }
                        }
                    }
                    ImportToCSV(name, description, saveFileDialog1.FileName);
                    System.Windows.MessageBox.Show(" [" + SheetNameList.Text + "] 导出成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }


        private void btnallExport_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog().ToString() == "OK") {
                string NPpath = folderBrowserDialog1.SelectedPath;//获取用户选中路径
                foreach (string s in tableNameList) {
                    DataTable table = Rowtempconn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, s, null });
                    List<string> name = GetTableFieldNameList(Rowtempconn, s);
                    List<string> description = new List<string>();
                    foreach (string c in name) {
                        for (int i = 0; i < table.Rows.Count; i++) {
                            if (c == table.Rows[i]["COLUMN_NAME"].ToString()) {
                                description.Add(table.Rows[i]["DESCRIPTION"].ToString());
                                break;
                            }
                        }
                    }
                    ImportToCSV(name, description, NPpath + "\\ " + s + ".csv");
                }
                System.Windows.MessageBox.Show(" 全部导出成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }


        /// <summary>
        /// 取所有表名
        /// </summary>
        /// <returns></returns>
        public List<string> GetTableNameList(System.Windows.Controls.ComboBox combobx, OleDbConnection Rowtempconn)
        {
            List<string> list = new List<string>();
            DataTable dt = Rowtempconn.GetSchema("Tables");
            foreach (DataRow row in dt.Rows) {
                string res = row[2].ToString();
                if (!res.StartsWith("MSys")) {
                    list.Add(res);
                    combobx.Items.Add(res);
                }
            }
            combobx.SelectedIndex = 0;
            return list;
        }


        public static void ImportToCSV(List<string> name, List<string> description, string fileName)
        {
            System.IO.FileStream fs = null;
            StreamWriter sw = null;
            fs = new System.IO.FileStream(fileName, FileMode.Create, FileAccess.Write);
            sw = new StreamWriter(fs, System.Text.Encoding.Default);
            //csv写入数据
            for (int i = 0; i < name.Count; i++) {
                string data2 = name[i] + "," + description[i];
                sw.WriteLine(data2);
            }
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// 取指定表所有字段名称
        /// </summary>
        /// <returns></returns>
        public List<string> GetTableFieldNameList(OleDbConnection Rowtempconn, string TableName)
        {
            List<string> list = new List<string>();
            using (OleDbCommand cmd = new OleDbCommand()) {
                cmd.CommandText = "SELECT TOP 1 * FROM [" + TableName + "]";
                cmd.Connection = Rowtempconn;
                OleDbDataReader dr = cmd.ExecuteReader();
                for (int i = 0; i < dr.FieldCount; i++) {
                    list.Add(dr.GetName(i));
                }
            }
            return list;
        }

        private void inport_Click(object sender, RoutedEventArgs e)
        {
            OleDbCommand oleDbCommand = new OleDbCommand("INSERT INTO dbo_BCF含水层属性表 VALUES(2,1,1,1)", Rowtempconn);
            oleDbCommand.ExecuteNonQuery();
        }
    }
}
