using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Microsoft.DirectX;
using Microsoft.DirectX.DirectSound;
using System.Threading;
using Gma.QrCodeNet.Encoding;
using ZXing.QrCode;
using ZXing;
using ZXing.Common;
using ZXing.Rendering;


namespace EStudio
{
    public partial class FirstDoor : Form
    {
        //获取单词库的路径
        string dir_word = string.Format(System.IO.Directory.GetCurrentDirectory() + "\\WORDlibrary\\");

        //获取录音文件的路径
        string dir_record = string.Format(System.IO.Directory.GetCurrentDirectory() + "\\");

        //定义一套考卷的数目
        public const int question_cout = 20;  
        //定义第一次考试时间
        public const int time_temp = 60;
        //每次闯关的倒计时时间
        int time = time_temp;
        //下一关的关卡时间
        int time_next = time_temp;
        //秒
        string temp_s;
        //分
        string temp_m;
        //时
        string temp_h;
        //及格分数线
        public const int scorePass = 60;
        //计分数
        int score = 0;
        //题的索引，根据index_ti 上一题，下一题
        int index_ti = 1;
        //20道题的答案都存放在answer[][]数组中
        string[][] answer = new string[question_cout][];
        //定义一个临时题目表,0列题目，1列答案
        DataTable data_table_temp = new DataTable();
        //将选中的20道题答案存放到answer_sel
        string[] answer_sel = new string[question_cout];

        // 初始化录音文件
        private SoundRecord recorder = null;

        //定义一个保存导入音乐的名字的数组
        public string[] filename;

        //二维码
        EncodingOptions options = null;
        BarcodeWriter writer = null;

        public FirstDoor()
        {
            InitializeComponent();

            //初始化二维码关联字段
            options = new QrCodeEncodingOptions
            {
                DisableECI = true,
                CharacterSet = "UTF-8",
                Width = 285,
                Height = 207
            };

            writer = new BarcodeWriter();

            writer.Format = BarcodeFormat.QR_CODE;

            writer.Options = options;
        }

        private void FirstDoor_Load(object sender, EventArgs e)
        {
            this.listBox1.Visible = false;
            this.quectionCount.Visible = false;
            this.tbxSentence.Visible = false;
            this.label_Word.Visible = false;

            //获取录音文件名
            this.listBox2.Items.Clear();

            DirectoryInfo mydir = new DirectoryInfo(dir_record);

            FileInfo[] file_name = mydir.GetFiles("*.wav");

            string[] temp_name = new string[file_name.Length];

            for (int i = 0; i < file_name.Length; i++)
            {
                temp_name[i] = file_name[i].Name.Substring(0, file_name[i].Name.Length - 4);

                listBox2.Items.Add(temp_name[i]);
            }

            //隐藏扫除错误功能
            gBxSlipWrong.Visible = false;

            labelRankWord.Visible = false;
            labelRankDate.Visible = false;
            labelRankExplain.Visible = false;
            //加载错误列表
            updateingRanking();


        }
        
        
        /// <summary>
        ///获取表名称
        /// </summary>
        /// <param name="excelFilename">表名</param>
        /// <returns></returns>
        public static DataTable GetExcelTable(string excelFilename)
        {
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;" +
                             "data source=" + excelFilename + ";extended properties='Excel 8.0;HDR=YES;IMEX=1;'";//2007excel

            DataSet ds = new DataSet();

            string tableName;
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                DataTable table = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                tableName = table.Rows[0]["Table_Name"].ToString();

                string strExcel = "select * from " + "[" + tableName + "]";

                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, connectionString);

                adapter.Fill(ds, tableName);

                connection.Close();
            }
            return ds.Tables[tableName];
        }
        //显示当前第几题
        public void index_question()
        {

            quectionCount.Text = "第 " + index_ti.ToString() + " 题 / 共 " + question_cout.ToString() + " 题";
        }
        /// <summary>
        ///随机从词库中取题目
        /// </summary>
        public void get_question()
        {
            //每次取完题后都要index_ti=1；不然下一套考卷的索引 不是从第一题开始
            index_ti = 1;
            //显示当前第几题
            index_question();
            //再次点击词库时要将radioButton清空
            radioBtn1.Checked = false;
            radioBtn2.Checked = false;
            radioBtn3.Checked = false;
            radioBtn4.Checked = false;

            gBxQuection.Visible = true;
            btnPreQuection.Visible = true;
            btnNextQuection.Visible = true;
            btnAnswerTip.Visible = true;
            btnPreQuection.Enabled = true;
            btnNextQuection.Enabled = true;
            btnAnswerTip.Enabled = true;
            messege.Visible = true;
            gBxQuection.Enabled = true;
            gBxErrorQuection.Visible = false;
            label1.Text = null;
            label2.Text = null;
            labeltimer.Text = "时间";

            if (btnAnswerTip.Text == "交卷")
            {
                grade.Visible = true;

                timer1.Interval = 1000;

                timer1.Enabled = true;
                //计时开始
                timer1.Start();
            }
            DataTable data_table = new DataTable();
            //加载题
            data_table = GetExcelTable(dir_word + listBox1.SelectedItem.ToString() + ".xls");

            string timu_ch = data_table.Rows[1].ItemArray[1].ToString();
            //遍历整张表，将缺失空行删掉（尽管已经手动删除还是要注意）
            for (int i = 0; i < data_table.Rows.Count; i++)
            {
                if (null == data_table.Rows[i][1].ToString() || null == data_table.Rows[i][2].ToString())
                {
                    data_table.Rows.RemoveAt(i);
                }
            }
            //判断data_table_temp是否初始化。通过计列数来查看是否有数据
            if (0 == data_table_temp.Columns.Count)
            {
                DataColumn dc = new DataColumn("column1", System.Type.GetType("System.String"));

                //添加一列 此添加行的方法和下面的'再添加2列'是同一个作用
                data_table_temp.Columns.Add(dc);    
 
                dc.AutoIncrement = true;
                //再添加2列
                for (int i = 2; i < 4; i++)
                    data_table_temp.Columns.Add("column" + i.ToString(), System.Type.GetType("System.String"));

                //添加question_cout行
                for (int i = 0; i < question_cout; i++)
                {
                    DataRow dr = data_table_temp.NewRow();

                    dr["Column1"] = "1";

                    data_table_temp.Rows.Add(dr);
                }
            }

            //对表中的行数计数
            int word_cout = data_table.Rows.Count;

            Random ran = new Random();

            int RandKey;
            //初始化数组的数组
            for (int i = 0; i < question_cout; i++)
            {
                answer[i] = new string[4];
            }
            for (int i = 0; i < question_cout; i++)
            {
                RandKey = ran.Next(0, word_cout - 1);
                //0行1列题目
                data_table_temp.Rows[i][1] = data_table.Rows[RandKey].ItemArray[1].ToString().Trim();
                //0行2列答案
                data_table_temp.Rows[i][2] = data_table.Rows[RandKey].ItemArray[2].ToString().Trim();
                //删除一行
                data_table.Rows.RemoveAt(RandKey);  
             
                word_cout--;

                answer[i][0] = data_table_temp.Rows[i].ItemArray[2].ToString();

                //在answer数组中的第1,2,3列存放错误的答案
                for (int j = 1; j < 4; j++)
                {
                    RandKey = ran.Next(0, word_cout - 1);

                    answer[i][j] = data_table.Rows[RandKey][2].ToString().Trim();
                }
                //将answer[][]中每一道题的答案打乱
                RandKey = ran.Next(0, 3);

                string temp_answer;

                temp_answer = answer[i][0];

                answer[i][0] = answer[i][RandKey];

                answer[i][RandKey] = temp_answer;
            }
            labelword.Text = data_table_temp.Rows[0][1].ToString();

            radioBtn1.Text = answer[index_ti - 1][0];
            radioBtn2.Text = answer[index_ti - 1][1];
            radioBtn3.Text = answer[index_ti - 1][2];
            radioBtn4.Text = answer[index_ti - 1][3];


            if (radioBtn1.Checked)
                answer_sel[index_ti - 1] = radioBtn1.Text;

            else if (radioBtn2.Checked)
                answer_sel[index_ti - 1] = radioBtn2.Text;

            else if (radioBtn3.Checked)
                answer_sel[index_ti - 1] = radioBtn3.Text;

            else if (radioBtn4.Checked)
                answer_sel[index_ti - 1] = radioBtn4.Text;

            else
                answer_sel[index_ti - 1] = null;

            radioBtn1.ForeColor = Color.Black;
            radioBtn2.ForeColor = Color.Black;
            radioBtn3.ForeColor = Color.Black;
            radioBtn4.ForeColor = Color.Black;

        }
        /// <summary>
        ///提交考卷，将用户选择的答案和data_table_temp对比
        /// </summary>
        public void Sub_paper()
        {
            btnAnswerTip.Enabled = false;
            btnPreQuection.Enabled = false;
            btnNextQuection.Enabled = false;
            score = 0;
            //统计答错的题数
            int answer_err = question_cout;
            timer1.Stop();
            timer1.Enabled = false;
            for (int i = 0; i < question_cout; i++)
            {
                //将用户选的答案和正确答案比较
                if (answer_sel[i] == data_table_temp.Rows[i][2].ToString())
                {
                    answer_err--;
                    score += 5;
                }

            }
            gBxErrorQuection.Visible = true;
            labelErrorCount.Text = "共答错" + answer_err.ToString() + "题：";
            messege.Text = score.ToString();
            gBxQuection.Enabled = false;
            listBox1.Enabled = false;
            index_ti++;//查看错误答案
            if (score >= scorePass)
            {
                double speed = Math.Round((float)question_cout / (float)(time_next - time), 3);//计算提交试卷时间//保留2位小数
                MessageBox.Show("您的做题速度为:" + speed + "题/秒！\n本关你得分：" + score + "分！\n进入下一关!");
                messege.Text = "分数";
                get_question();
                time = time_next - 5;
                time_next = time_next - 5;
            }
            else
            {
                MessageBox.Show("错得太多啦亲！赶紧检查一下吧！");
                time = time_temp;
                time_next = time_temp;
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            time--;
            //时间以00:00:00样式显示
            if (time / 3600 < 10)
                temp_h = "0" + (time / 3600).ToString();
            else
                temp_h = (time / 3600).ToString();
            if ((time / 60) % 60 < 10)
                temp_m = "0" + ((time / 60) % 60).ToString();
            else
                temp_m = ((time / 60) % 60).ToString();
            if (time % 60 < 10)
                temp_s = "0" + (time % 60).ToString();
            else
                temp_s = (time % 60).ToString();
            labeltimer.Text = temp_h + ":" + temp_m + ":" + temp_s;
            if (time == 0)
            {
                timer1.Stop();
                Sub_paper();
            }
            labelTime.Visible = true;
            labeltimer.Visible = true;   
        }

        //修炼模式
        private void 修炼模式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grade.Visible = false;
            timer1.Stop();
            timer1.Enabled = false;
            btnAnswerTip.Text = "答案提示";
            listBox1.Items.Clear();
            this.listBox1.Visible = true;
            listBox1.Enabled = true;
            btnAnswerTip.Visible = false;
            labelTime.Visible = false;
            labeltimer.Visible = false;
            messege.Visible = false;
            btnPreQuection.Visible = false;
            btnNextQuection.Visible = false;
            gBxErrorQuection.Visible = false;
            gBxQuection.Visible = false;


            DirectoryInfo mydir = new DirectoryInfo(dir_word);

            FileInfo[] file_name = mydir.GetFiles("*.xls");

            string[] temp_name = new string[file_name.Length];

            for (int i = 0; i < file_name.Length; i++)
            {
                temp_name[i] = file_name[i].Name.Substring(0, file_name[i].Name.Length - 4);
                //获取词库中的excel文件名
                listBox1.Items.Add(temp_name[i]);
            }
        }
        //生存模式
        private void 生存模式ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            grade.Visible = false;
            time = time_next;
            timer1.Stop();
            labeltimer.Text = "时间";
            timer1.Enabled = false;
            listBox1.Items.Clear();
            this.listBox1.Visible = true;
            listBox1.Enabled = true;
            btnAnswerTip.Visible = false;
            labelTime.Visible = true;
            labeltimer.Visible = true;
            messege.Visible = false;
            gBxQuection.Visible = false;
            btnPreQuection.Visible = false;
            btnNextQuection.Visible = false;
            gBxErrorQuection.Visible = false;
            btnAnswerTip.Text = "交卷";


            DirectoryInfo mydir = new DirectoryInfo(dir_word);

            FileInfo[] file_name = mydir.GetFiles("*.xls");

            string[] temp_name = new string[file_name.Length];

            for (int i = 0; i < file_name.Length; i++)
            {
                temp_name[i] = file_name[i].Name.Substring(0, file_name[i].Name.Length - 4);

                listBox1.Items.Add(temp_name[i]);
            }
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            //
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            //
        }
        //选择词汇库
        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            this.quectionCount.Visible = true;

            time = time_temp;
            time_next = time_temp;

            for (int i = 0; i < question_cout; i++)
            {
                answer_sel[i] = null;
            }

            if (btnAnswerTip.Text == "答案提示")
            {
                messege.Text = "答案提示信息";
            }
            else
            {
                messege.Text = "分数";
            }
            //获取题目
            get_question();
        }

        private void radioBtn1_CheckedChanged(object sender, EventArgs e)
        {
            answer_sel[index_ti - 1] = radioBtn1.Text;
            if (radioBtn1.Checked == false && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                answer_sel[index_ti - 1] = null;
            }
            if (radioBtn1.Checked == true && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn1.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
        }

        private void radioBtn2_CheckedChanged(object sender, EventArgs e)
        {
            answer_sel[index_ti - 1] = radioBtn2.Text;
            if (radioBtn1.Checked == false && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                answer_sel[index_ti - 1] = null;
            }
            if (radioBtn2.Checked == true && radioBtn1.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn2.ForeColor = Color.Red;
                radioBtn1.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
        }

        private void radioBtn3_CheckedChanged(object sender, EventArgs e)
        {
            answer_sel[index_ti - 1] = radioBtn3.Text;
            if (radioBtn1.Checked == false && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                answer_sel[index_ti - 1] = null;
            }
            if (radioBtn3.Checked == true && radioBtn2.Checked == false && radioBtn1.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn3.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn1.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
        }

        private void radioBtn4_CheckedChanged(object sender, EventArgs e)
        {
            answer_sel[index_ti - 1] = radioBtn4.Text;
            if (radioBtn1.Checked == false && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                answer_sel[index_ti - 1] = null;
            }
            if (radioBtn4.Checked == true && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn1.Checked == false)
            {
                radioBtn4.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn1.ForeColor = Color.Black;
            }
        }
        //上一题
        private void btnPreQuection_Click_1(object sender, EventArgs e)
        {
            --index_ti;

            if (index_ti < 1)
                index_ti = question_cout;
            index_question();
            labelword.Text = data_table_temp.Rows[index_ti - 1][1].ToString();
            radioBtn1.Text = answer[index_ti - 1][0];
            radioBtn2.Text = answer[index_ti - 1][1];
            radioBtn3.Text = answer[index_ti - 1][2];
            radioBtn4.Text = answer[index_ti - 1][3];


            if (radioBtn1.Text == answer_sel[index_ti - 1])
                radioBtn1.Checked = true;
            else if (radioBtn2.Text == answer_sel[index_ti - 1])
                radioBtn2.Checked = true;
            else if (radioBtn3.Text == answer_sel[index_ti - 1])
                radioBtn3.Checked = true;
            else if (radioBtn4.Text == answer_sel[index_ti - 1])
                radioBtn4.Checked = true;
            else
            {
                radioBtn1.Checked = false;
                radioBtn2.Checked = false;
                radioBtn3.Checked = false;
                radioBtn4.Checked = false;
            }
            if (radioBtn1.Checked == true && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn1.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
            else if (radioBtn2.Checked == true && radioBtn1.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn2.ForeColor = Color.Red;
                radioBtn1.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
            else if (radioBtn3.Checked == true && radioBtn1.Checked == false && radioBtn2.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn3.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn1.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
            else if (radioBtn4.Checked == true && radioBtn1.Checked == false && radioBtn3.Checked == false && radioBtn2.Checked == false)
            {
                radioBtn4.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn1.ForeColor = Color.Black;
            }
            else
            {
                radioBtn1.ForeColor = Color.Black;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
        }
        //下一题
        private void btnNextQuection_Click_1(object sender, EventArgs e)
        {
            ++index_ti;
            if (index_ti > question_cout)
                index_ti = 1;
            index_question();
            if (answer_sel[index_ti - 1] == null)
            {
                radioBtn1.Checked = false;
                radioBtn2.Checked = false;
                radioBtn3.Checked = false;
                radioBtn4.Checked = false;
            }
            labelword.Text = data_table_temp.Rows[index_ti - 1][1].ToString();
            radioBtn1.Text = answer[index_ti - 1][0];
            radioBtn2.Text = answer[index_ti - 1][1];
            radioBtn3.Text = answer[index_ti - 1][2];
            radioBtn4.Text = answer[index_ti - 1][3];


            if (radioBtn1.Text == answer_sel[index_ti - 1])
                radioBtn1.Checked = true;
            else if (radioBtn2.Text == answer_sel[index_ti - 1])
                radioBtn2.Checked = true;
            else if (radioBtn3.Text == answer_sel[index_ti - 1])
                radioBtn3.Checked = true;
            else if (radioBtn4.Text == answer_sel[index_ti - 1])
                radioBtn4.Checked = true;
            else
            {
                radioBtn1.Checked = false;
                radioBtn2.Checked = false;
                radioBtn3.Checked = false;
                radioBtn4.Checked = false;
            }
            if (radioBtn1.Checked == true && radioBtn2.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn1.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
            else if (radioBtn2.Checked == true && radioBtn1.Checked == false && radioBtn3.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn2.ForeColor = Color.Red;
                radioBtn1.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
            else if (radioBtn3.Checked == true && radioBtn1.Checked == false && radioBtn2.Checked == false && radioBtn4.Checked == false)
            {
                radioBtn3.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn1.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
            else if (radioBtn4.Checked == true && radioBtn1.Checked == false && radioBtn3.Checked == false && radioBtn2.Checked == false)
            {
                radioBtn4.ForeColor = Color.Red;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn1.ForeColor = Color.Black;
            }
            else
            {
                radioBtn1.ForeColor = Color.Black;
                radioBtn2.ForeColor = Color.Black;
                radioBtn3.ForeColor = Color.Black;
                radioBtn4.ForeColor = Color.Black;
            }
        }
        //答案提示，交卷
        private void btnAnswerTip_Click_1(object sender, EventArgs e)
        {
            if (btnAnswerTip.Text == "答案提示")
            {
                messege.Text = data_table_temp.Rows[index_ti - 1][2].ToString();
                if (data_table_temp.Rows[index_ti - 1][2].ToString() == radioBtn1.Text)
                {
                    radioBtn1.ForeColor = Color.Blue;
                }
                else if (data_table_temp.Rows[index_ti - 1][2].ToString() == radioBtn2.Text)
                {
                    radioBtn2.ForeColor = Color.Blue;
                }
                else if (data_table_temp.Rows[index_ti - 1][2].ToString() == radioBtn3.Text)
                {
                    radioBtn3.ForeColor = Color.Blue;
                }
                else if (data_table_temp.Rows[index_ti - 1][2].ToString() == radioBtn4.Text)
                {
                    radioBtn4.ForeColor = Color.Blue;
                }
            }
            else
            {
                Sub_paper();
            }
        }
        //回顾错误词汇
        private void btnNextTip_Click(object sender, EventArgs e)
        {
            if (index_ti > question_cout)
                index_ti = 1;

            if (answer_sel[index_ti - 1] != data_table_temp.Rows[index_ti - 1][2].ToString())
            {
                label1.Text = data_table_temp.Rows[index_ti - 1][1].ToString();
                label2.Text = data_table_temp.Rows[index_ti - 1][2].ToString();
                string word = data_table_temp.Rows[index_ti - 1][1].ToString();
                string explain = data_table_temp.Rows[index_ti - 1][2].ToString();
                //插入错误单词到错误列表
                DataAccess.InsertData(word, explain, DateTime.Now.ToString());
                //更新错误列表
                updateingRanking();

                index_ti++;

                if (index_ti > question_cout)
                    index_ti = 1;
            }
            while (answer_sel[index_ti - 1] == data_table_temp.Rows[index_ti - 1][2].ToString())
            {
                index_ti++;
                continue;
            }
        }
        //加载错误列表
        private void tabPage2_Click(object sender, EventArgs e)
        {
            updateingRanking();
        }
        //刷新错误列表
        public void updateingRanking(){

            DataTable dt = DataAccess.ReadAllData("select * from WrongList");

            this.dgvRanking.DataSource = dt.DefaultView;
        }   

        private void tabPage7_Click(object sender, EventArgs e)
        {

        }
        
        //搜索词汇库
        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (this.label_Word.Text != null)
            {
                return;
            }
            else
            {
                this.label_Word.Text = "没有找到查询值！";
            }
            this.label_Word.Visible = true;
        }

        private void label_Word_Click(object sender, EventArgs e)
        {
            //
        }
        //即刻搜索
        private void tbxSearch_TextChanged(object sender, EventArgs e)
        {
            Search(this.tbxSearch.Text.Trim());

            this.label_Word.Visible = true;
        }
        /// <summary>
        ///查询access
        /// </summary>
        /// <param name="searchTxt">查询值</param>
        /// <returns></returns>
        public void Search(string searchTxt)
        {
            using (OleDbConnection odcConnection = new OleDbConnection(DataAccess.connectionString))
            {
                //打开连接   
                odcConnection.Open();

                string strSQL = "select * from EnglishLibrary where [Words] = '" + searchTxt + "'";

                using (OleDbCommand cmd = new OleDbCommand(strSQL, odcConnection))
                {
                    try
                    {
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                        DataSet ds = new DataSet();

                        da.Fill(ds);
                        //
                        this.label_Word.Text = ds.Tables[0].Rows[0].ItemArray[3].ToString();

                        this.tbxSentence.Text = ds.Tables[0].Rows[0].ItemArray[4].ToString();

                        this.tbxSentence.Visible = true;
                    }
                    catch
                    {
                        //返回未查询到的结果
                        this.label_Word.Text = "";
                        this.tbxSentence.Visible = false;
                        this.label_Word.Visible = false;
                    }
                    finally 
                    {
                        odcConnection.Close();
                    }
                    
                }
            }
        }

        /// <summary>
        ///显示错误单词详细信息
        ///勾选导出单词表
        /// </summary>
        StringBuilder s;
        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (this.dgvRanking.SelectionMode != DataGridViewSelectionMode.FullColumnSelect)
                {
                    labelRankWord.Visible = true;
                    labelRankDate.Visible = true;
                    labelRankExplain.Visible = true;

                    //获取选中行的行号
                    int index = dgvRanking.SelectedRows[0].Index;

                    this.labelRankWord.Text = dgvRanking.Rows[index].Cells[2].Value.ToString();

                    this.labelRankExplain.Text = dgvRanking.Rows[index].Cells[3].Value.ToString();

                    this.labelRankDate.Text = dgvRanking.Rows[index].Cells[4].Value.ToString();
                }

                s = new StringBuilder();

                int count = Convert.ToInt32(dgvRanking.Rows.Count.ToString());

                //如果DataGridView是可编辑的，将数据提交，否则处于编辑状态的行无法取到
                dgvRanking.EndEdit();

                for (int i = 0; i < count; i++)
                {
                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dgvRanking.Rows[i].Cells["columnAdd"];

                    Boolean flag = Convert.ToBoolean(checkCell.Value);
                    //查找被选择的数据行
                    if (flag == true)     
                    {
                        //从 DATAGRIDVIEW 中获取数据项
                        s.Append(dgvRanking.Rows[i].Cells[2].Value.ToString().Trim());
                        s.Append(":");
                        s.Append(dgvRanking.Rows[i].Cells[3].Value.ToString().Trim());
                        s.Append("//");
                        
                    }
                }
      
            }
            catch {
                labelRankWord.Visible = false;
                labelRankDate.Visible = false;
                labelRankExplain.Visible = false;
            }
        }
        //错误单词下一题
        private void btnNextWrong_Click(object sender, EventArgs e)
        {
            using (OleDbConnection odcConnection = new OleDbConnection(DataAccess.connectionString))
            {
                //打开连接   
                odcConnection.Open();

                string strSQL = "select * from WrongList";

                using (OleDbCommand cmd = new OleDbCommand(strSQL, odcConnection))
                {
                    try
                    {
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                        DataSet ds = new DataSet();

                        da.Fill(ds);

                        this.labelWrongExplain.Text = ds.Tables[0].Rows[0].ItemArray[2].ToString();
                    }
                    catch
                    {
                        this.labelWrongExplain.Text = "不错哦！没有错误的记录";
                    }
                    finally
                    {
                        odcConnection.Close();
                    }

                }
            }
        }
        //扫除错误单词
        private void btnSlipWrong_Click(object sender, EventArgs e)
        {
            using (OleDbConnection odcConnection = new OleDbConnection(DataAccess.connectionString))
            {
                //打开连接   
                odcConnection.Open();

                string strSQL = "select * from WrongList";

                using (OleDbCommand cmd = new OleDbCommand(strSQL, odcConnection))
                {
                    try
                    {
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                        DataSet ds = new DataSet();

                        da.Fill(ds);

                        odcConnection.Close();

                        if (this.tbxSlipWrong.Text.Trim() != "")
                        {
                            if (ds.Tables[0].Rows[0].ItemArray[1].ToString() == this.tbxSlipWrong.Text.Trim()){
                                //删除错误记录
                                DataAccess.DeleteData(ds.Tables[0].Rows[0].ItemArray[0].ToString());

                                this.labelWrongExplain.Text = "成功扫除！";
                            }
                            else {
                                MessageBox.Show("再接再厉！");
                            }
                        }
                        else{
                            MessageBox.Show("单词还未输入！");
                        }
                    }
                    catch
                    {
                        this.labelWrongExplain.Text = "未知异常！";
                    }

                }
            }
        }
        //开始录音
        private void btnRecord_Click(object sender, EventArgs e)
        {
            try
            {
                SaveRecordName srn = new SaveRecordName();

                srn.ShowDialog();
                // 录音设置  
                string wavfile = null;

                wavfile = srn.saveName + ".wav";

                if (wavfile == ".wav")
                {
                    wavfile = null;
                }
                recorder = new SoundRecord();

                recorder.SetFileName(wavfile);

                recorder.RecStart();
            }
            catch
            {
                MessageBox.Show("文件名不能为空！");
                return;
            }
        }
        //结束录音
        private void btnEndRecord_Click(object sender, EventArgs e)
        {
            recorder.RecStop();

            recorder = null;
            //获取录音文件名
            this.listBox2.Items.Clear();

            DirectoryInfo mydir = new DirectoryInfo(dir_record);

            FileInfo[] file_name = mydir.GetFiles("*.wav");

            string[] temp_name = new string[file_name.Length];

            for (int i = 0; i < file_name.Length; i++)
            {
                temp_name[i] = file_name[i].Name.Substring(0, file_name[i].Name.Length - 4);

                listBox2.Items.Add(temp_name[i]);
            }
        }
        //选择播放录音项
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.URL = null;
        }
        //判断音乐文件
        public static bool isMusic(string strIn)  
        {
            return Regex.IsMatch(strIn, @"\.(?i:mp3|wav)$");  
        }   
        //播放录音文件
        private void 播放ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //判断文件是录音文件还是听力材料 
            if (isMusic(listBox2.SelectedItem.ToString()))                //通过匹配
            {
                Regex reg = new Regex(@"[^/\\\\]+$");
                //遍历数组
                foreach (string s in filename)
                {

                    Match result = reg.Match(s);

                    if (result.Success)
                    {
                        //匹配
                        if (result.Value == listBox2.SelectedItem.ToString())
                        {
                            //播放
                            axWindowsMediaPlayer1.URL = s;

                        }

                    }
                    else
                    {
                        MessageBox.Show("文件路径已移动，请重新导入！");
                        axWindowsMediaPlayer1.URL = null;
                    }

                }

            }
            else{
                //是录音文件
                axWindowsMediaPlayer1.URL = dir_record + listBox2.SelectedItem.ToString() + ".wav";
                
                
            }
        }
        //删除录音文件
        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex < 0) return;

            FileInfo file = new FileInfo(dir_record + listBox2.SelectedItem.ToString() + ".wav");

            if (file.Exists)
            {
                //删除源文件
                file.Delete();
            }
            //重新加载
            this.listBox2.Items.Clear();

            DirectoryInfo mydir = new DirectoryInfo(dir_record);

            FileInfo[] file_name = mydir.GetFiles("*.wav");

            string[] temp_name = new string[file_name.Length];

            for (int i = 0; i < file_name.Length; i++)
            {
                temp_name[i] = file_name[i].Name.Substring(0, file_name[i].Name.Length - 4);
                listBox2.Items.Add(temp_name[i]);
            }
        }
        //打开听力材料
        private void btnOpenListenFile_Click(object sender, EventArgs e)
        {

            openFileDialog1.FileName = "";

            openFileDialog1.Filter = "所有文件|*.*|WAV文件|*.wav|MP3文件|*.mp3";

            openFileDialog1.Title = "选择音频文件";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileNames;

                Regex reg = new Regex(@"[^/\\\\]+$");
                
                foreach(string s in filename){
                    //匹配
                    Match result = reg.Match(s);        

                    if (result.Success) {
                        //获取
                        listBox2.Items.Add(result.Value);

                    }

                }

            }
        }
        //进入扫除
        private void btnShowSlip_Click(object sender, EventArgs e)
        {
            gBxSlipWrong.Visible = true;
            btnShowSlip.Visible = false;
            dgvRanking.Visible = false;
            labelRankWord.Visible = false;
            labelRankDate.Visible = false;
            labelRankExplain.Visible = false;
            btnOutExcel.Visible = false;
            btnQr.Visible = false;
        }
        //退出扫除
        private void btnExitSlip_Click(object sender, EventArgs e)
        {
            gBxSlipWrong.Visible = false;
            btnShowSlip.Visible = true;
            dgvRanking.Visible = true;
            labelRankWord.Visible = true;
            labelRankDate.Visible = true;
            labelRankExplain.Visible = true;
            btnOutExcel.Visible = true;
            btnQr.Visible = true;
            updateingRanking();
        }
        /// <summary>
        ///导出错误列表到Excel方法
        /// </summary>
        public static void OutputExcel()
        {
            OleDbConnection con = new OleDbConnection();

            try
            {

                SaveFileDialog saveFile = new SaveFileDialog();
                //指定文件后缀名为Excel 文件。
                saveFile.Filter = ("Excel 文件(*.xls)|*.xls"); 
 
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    string filename = saveFile.FileName;

                    if (System.IO.File.Exists(filename))
                    {
                        //如果文件存在删除文件。
                        System.IO.File.Delete(filename);  
                    }
                    //获取最后一个/的索引
                    int index = filename.LastIndexOf("//");
                    //获取excel名称(新建表的路径相对于SaveFileDialog的路径)
                    filename = filename.Substring(index + 1);  
                    //select * into 建立 新的表。  
                    //[[Excel 8.0;database= excel名].[sheet名] 如果是新建sheet表不能加$,如果向sheet里插入数据要加$.　  
                    //sheet最多存储65535条数据。  
                    string sql = "select top 65535 *  into   [Excel 8.0;database=" + filename + "].[Sheet1] from WrongList";
                    //连接数据库字符串
                    con.ConnectionString = DataAccess.connectionString;  

                    OleDbCommand com = new OleDbCommand(sql, con);

                    con.Open();

                    com.ExecuteNonQuery();

                    MessageBox.Show("导出数据成功", "导出数据", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出Excel失败！" + ex.ToString());
            }
            finally
            {
                con.Close();
            }
        }
        //导出错误列表到Excel
        private void btnOutExcel_Click(object sender, EventArgs e)
        {
            OutputExcel();
        } 
        //生成二维码
        private void btnQr_Click(object sender, EventArgs e)
        {
            int count = Convert.ToInt32(dgvRanking.Rows.Count.ToString());

            int j = 0;//记录未勾选
            int k = 0;//记录勾选数量

            try
            {

                for (int i = 0; i < count; i++)
                {
                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dgvRanking.Rows[i].Cells["columnAdd"];

                    int flag = Convert.ToInt32(checkCell.Value);

                    if (flag == 0){
                        j ++;
                    }
                    if (flag == 1){
                        k++;
                    }
                       
                }
                //未勾选
                if(j == count){
                    MessageBox.Show("还未选择想背诵的单词，无法导出！");
                }
                //勾选值大于预期值
                if (k > 5){
                    MessageBox.Show("背单词奥义：一疗程5个，5分钟一个疗程！妈妈再也不用担心我的学习！");
                }
                else {
                    //获取
                    Bitmap bitmap = writer.Write(s.ToString());
                    //显示
                    FormQr fqr = new FormQr();

                    fqr.sQr = bitmap;

                    fqr.ShowDialog();
                } 
            }
            catch 
            {
                return;
            }

        }
      
    }
}
