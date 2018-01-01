using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

namespace HeiheWaterScenario
{
    public partial class frmWaterScenario : Form
    {
        public frmWaterScenario()
        {
            InitializeComponent();

        }

        #region 全局变量
        GraphPane myPane;
        
        double i = 0;//折线图变量
        int j = 0;//柱状图变量
        int t = 0;//年份变量
        int[] years = new int[]{2006,2007,2008,2009,2010, 2011,2012,2013,2014,2015,2016,2017,2018,2019, 2020, 2021
            ,2022,2023,2024,2025,2026,2027,2028,2029,2030};
        #endregion


        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = false;
            #region 改变checkBox的位置
            this.cbRCP26.Location = new System.Drawing.Point(26, 90);
            this.cbRCP45.Location = new System.Drawing.Point(26, 130);
            this.cbRCP60.Location = new System.Drawing.Point(26, 170);
            this.cbRCP85.Location = new System.Drawing.Point(26, 210);

            this.checkBox12.Location = new System.Drawing.Point(26, 80);
            this.checkBox11.Location = new System.Drawing.Point(26, 110);
            this.checkBox10.Location = new System.Drawing.Point(26, 140);
            this.checkBox9.Location = new System.Drawing.Point(26, 170);
            this.checkBox8.Location = new System.Drawing.Point(26, 200);
            this.checkBox7.Location = new System.Drawing.Point(26, 230);

            this.checkBox18.Location = new System.Drawing.Point(26, 80);
            this.checkBox17.Location = new System.Drawing.Point(26, 110);
            this.checkBox16.Location = new System.Drawing.Point(26, 140);
            this.checkBox15.Location = new System.Drawing.Point(26, 170);
            this.checkBox14.Location = new System.Drawing.Point(26, 200);
            this.checkBox13.Location = new System.Drawing.Point(26, 230);

            this.checkBox24.Location = new System.Drawing.Point(26, 80);
            this.checkBox23.Location = new System.Drawing.Point(26, 110);
            this.checkBox22.Location = new System.Drawing.Point(26, 140);
            this.checkBox21.Location = new System.Drawing.Point(26, 170);
            this.checkBox20.Location = new System.Drawing.Point(26, 200);
            this.checkBox19.Location = new System.Drawing.Point(26, 230);
            #endregion

            #region 改变CheckBox的文本字体
            this.cbRCP26.Font = new System.Drawing.Font("微软雅黑", 11.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbRCP45.Font = new System.Drawing.Font("微软雅黑", 11.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbRCP60.Font = new System.Drawing.Font("微软雅黑", 11.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbRCP85.Font = new System.Drawing.Font("微软雅黑", 11.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
           

            this.checkBox12.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox11.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox10.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox9.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox8.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox7.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            this.checkBox18.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox17.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox16.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox15.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox14.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox13.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            this.checkBox24.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox23.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox22.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox21.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox20.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox19.Font = new System.Drawing.Font("Georgia", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            #endregion

            #region 改变CheckBox的外观
            
            cbRCP85.FlatStyle = FlatStyle.Flat;
            cbRCP60.FlatStyle = FlatStyle.Flat;
            cbRCP45.FlatStyle = FlatStyle.Flat;
            cbRCP26.FlatStyle = FlatStyle.Flat;
            checkBox7.FlatStyle = FlatStyle.Flat;
            checkBox8.FlatStyle = FlatStyle.Flat;
            checkBox9.FlatStyle = FlatStyle.Flat;
            checkBox10.FlatStyle = FlatStyle.Flat;
            checkBox11.FlatStyle = FlatStyle.Flat;
            checkBox12.FlatStyle = FlatStyle.Flat;
            checkBox13.FlatStyle = FlatStyle.Flat;
            checkBox14.FlatStyle = FlatStyle.Flat;
            checkBox15.FlatStyle = FlatStyle.Flat;
            checkBox16.FlatStyle = FlatStyle.Flat;
            checkBox17.FlatStyle = FlatStyle.Flat;
            checkBox18.FlatStyle = FlatStyle.Flat;
            checkBox19.FlatStyle = FlatStyle.Flat;
            checkBox20.FlatStyle = FlatStyle.Flat;
            checkBox21.FlatStyle = FlatStyle.Flat;
            checkBox22.FlatStyle = FlatStyle.Flat;
            checkBox23.FlatStyle = FlatStyle.Flat;
            checkBox24.FlatStyle = FlatStyle.Flat;
            #endregion


            lbTemperature.Text = "0.00" + "℃";
            lbRain.Text = "000.0" + "mm";
            lbUsingWater.Text = "0000.00" + "m³";
            lbGovernment.Text = "00.00" + "亿元";
            lbLabour.Text = "000.00" + "万元";

            button1.Image = new Bitmap(button1.Image, button1.Height - 7, button1.Height - 7);
            button3.Image = new Bitmap(button3.Image, button3.Height - 14, button3.Height - 14);

            TemperatureGraph();
            RainGraph();
            SupplyGraph();
            EnergyGraph();
            AgriWaterGraph();
            UsingWaterGraph();
            JobGragh();
            PopulationGraph();
            CityRate();
            Government();
            tradeGraph();
            LabourGraph();
            
            toolStripStatusLabel1.Text = "简介";//设置状态栏初始状态为第一个模块信息

        }

        #region checkBox勾选同步
        //Green and Pleasant Land
      
        private void checkBox12_Click(object sender, EventArgs e)
        {
            if (checkBox12.Checked)
            {
                checkBox12.Checked = true;
                checkBox18.Checked = true;
                checkBox24.Checked = true;
            }
            else
            {
                checkBox12.Checked = false;
                checkBox18.Checked = false;
                checkBox24.Checked = false;
            }
        }
        private void checkBox18_Click(object sender, EventArgs e)
        {
            if (checkBox18.Checked)
            {
                checkBox12.Checked = true;
                checkBox18.Checked = true;
                checkBox24.Checked = true;
            }
            else
            {
                checkBox12.Checked = false;
                checkBox18.Checked = false;
                checkBox24.Checked = false;
            }
        }
        private void checkBox24_Click(object sender, EventArgs e)
        {
            if (checkBox24.Checked)
            {
                checkBox12.Checked = true;
                checkBox18.Checked = true;
                checkBox24.Checked = true;
            }
            else
            {
                checkBox12.Checked = false;
                checkBox18.Checked = false;
                checkBox24.Checked = false;
            }
        }

        //Nature@Work
        
        private void checkBox11_Click(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {
                checkBox11.Checked = true;
                checkBox17.Checked = true;
                checkBox23.Checked = true;
            }
            else
            {
                checkBox11.Checked = false;
                checkBox17.Checked = false;
                checkBox23.Checked = false;
            }

        }
        private void checkBox17_Click(object sender, EventArgs e)
        {
            if (checkBox17.Checked)
            {
                cbRCP45.Checked = true;
                checkBox11.Checked = true;
                checkBox17.Checked = true;
                checkBox23.Checked = true;
            }
            else
            {
                checkBox11.Checked = false;
                checkBox17.Checked = false;
                checkBox23.Checked = false;
            }
        }
        private void checkBox23_Click(object sender, EventArgs e)
        {
            if (checkBox23.Checked)
            {
                checkBox11.Checked = true;
                checkBox17.Checked = true;
                checkBox23.Checked = true;
            }
            else
            {
                checkBox11.Checked = false;
                checkBox17.Checked = false;
                checkBox23.Checked = false;
            }
        }

        //World Markets
      
        private void checkBox10_Click(object sender, EventArgs e)
        {

            if (checkBox10.Checked)
            {
                checkBox10.Checked = true;
                checkBox16.Checked = true;
                checkBox22.Checked = true;
            }
            else
            {
                checkBox10.Checked = false;
                checkBox16.Checked = false;
                checkBox22.Checked = false;
            }
        }
        private void checkBox16_Click(object sender, EventArgs e)
        {

            if (checkBox16.Checked)
            {
                checkBox10.Checked = true;
                checkBox16.Checked = true;
                checkBox22.Checked = true;
            }
            else
            {
                checkBox10.Checked = false;
                checkBox16.Checked = false;
                checkBox22.Checked = false;
            }
        }
        private void checkBox22_Click(object sender, EventArgs e)
        {

            if (checkBox22.Checked)
            {
                checkBox10.Checked = true;
                checkBox16.Checked = true;
                checkBox22.Checked = true;
            }
            else
            {
                checkBox10.Checked = false;
                checkBox16.Checked = false;
                checkBox22.Checked = false;
            }
        }

        //National Security
        
        private void checkBox9_Click(object sender, EventArgs e)
        {

            if (checkBox9.Checked)
            {
                checkBox9.Checked = true;
                checkBox15.Checked = true;
                checkBox21.Checked = true;
            }
            else
            {
                checkBox9.Checked = false;
                checkBox15.Checked = false;
                checkBox21.Checked = false;
            }
        }
        private void checkBox15_Click(object sender, EventArgs e)
        {
            if (checkBox15.Checked)
            {
                checkBox9.Checked = true;
                checkBox15.Checked = true;
                checkBox21.Checked = true;
            }
            else
            {
                checkBox9.Checked = false;
                checkBox15.Checked = false;
                checkBox21.Checked = false;
            }
        }
        private void checkBox21_Click(object sender, EventArgs e)
        {
            if (checkBox21.Checked)
            {
                checkBox9.Checked = true;
                checkBox15.Checked = true;
                checkBox21.Checked = true;
            }
            else
            {
                checkBox9.Checked = false;
                checkBox15.Checked = false;
                checkBox21.Checked = false;
            }
        }

        //Local Stewardship
        
        private void checkBox8_Click(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                checkBox8.Checked = true;
                checkBox14.Checked = true;
                checkBox20.Checked = true;
            }
            else
            {
                checkBox8.Checked = false;
                checkBox14.Checked = false;
                checkBox20.Checked = false;
            }



        }
        private void checkBox14_Click(object sender, EventArgs e)
        {
            if (checkBox14.Checked)
            {
                checkBox8.Checked = true;
                checkBox14.Checked = true;
                checkBox20.Checked = true;
            }
            else
            {
                checkBox8.Checked = false;
                checkBox14.Checked = false;
                checkBox20.Checked = false;
            }
        }
        private void checkBox20_Click(object sender, EventArgs e)
        {
            if (checkBox20.Checked)
            {
                checkBox8.Checked = true;
                checkBox14.Checked = true;
                checkBox20.Checked = true;
            }
            else
            {
                checkBox8.Checked = false;
                checkBox14.Checked = false;
                checkBox20.Checked = false;
            }
        }

        //Go with the flow
        
        private void checkBox7_Click(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                
                checkBox7.Checked = true;
                checkBox13.Checked = true;
                checkBox19.Checked = true;
            }
            else
            {
                
                checkBox7.Checked = false;
                checkBox13.Checked = false;
                checkBox19.Checked = false;
            }
        }
        private void checkBox13_Click(object sender, EventArgs e)
        {
            if (checkBox13.Checked)
            {
                checkBox7.Checked = true;
                checkBox13.Checked = true;
                checkBox19.Checked = true;
            }
            else
            {
                checkBox7.Checked = false;
                checkBox13.Checked = false;
                checkBox19.Checked = false;
            }






        }
        private void checkBox19_Click(object sender, EventArgs e)
        {
            if (checkBox19.Checked)
            {
                checkBox7.Checked = true;
                checkBox13.Checked = true;
                checkBox19.Checked = true;
            }
            else
            {
                
                checkBox7.Checked = false;
                checkBox13.Checked = false;
                checkBox19.Checked = false;
            }

        }

        #endregion

        #region 进入模块时触发状态栏信息改变
        private void tabPage1_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "简介";
        }

        private void tabPage2_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "情景分析";
        }
        private void tabPage3_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "气候变化模块";
        }
        private void tabPage4_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "生态文明模块";
        }

        private void tabPage5_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "新型城镇化模块";
        }
        private void tabPage6_Enter(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "产业发展模块";
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (cbRCP26.Checked == false && cbRCP45.Checked == false && cbRCP60.Checked == false && cbRCP85.Checked == false && checkBox7.Checked == false && checkBox8.Checked == false && checkBox9.Checked == false && checkBox10.Checked == false
               && checkBox11.Checked == false && checkBox12.Checked == false)
            {
                MessageBox.Show("请先选择一个情景");
            }
            else
            {
                if (cbRCP26.Checked && cbRCP45.Checked || cbRCP26.Checked && cbRCP60.Checked || cbRCP26.Checked && cbRCP85.Checked ||
                    cbRCP45.Checked && cbRCP60.Checked || cbRCP45.Checked && cbRCP85.Checked || cbRCP60.Checked && cbRCP85.Checked ||
                    checkBox7.Checked && checkBox8.Checked || checkBox7.Checked && checkBox9.Checked || checkBox7.Checked && checkBox10.Checked ||
                      checkBox7.Checked && checkBox11.Checked || checkBox7.Checked && checkBox12.Checked || checkBox8.Checked && checkBox9.Checked ||
                      checkBox8.Checked && checkBox10.Checked || checkBox8.Checked && checkBox11.Checked || checkBox8.Checked && checkBox12.Checked ||
                      checkBox9.Checked && checkBox10.Checked || checkBox9.Checked && checkBox11.Checked || checkBox9.Checked && checkBox12.Checked ||
                      checkBox10.Checked && checkBox11.Checked || checkBox10.Checked && checkBox12.Checked || checkBox11.Checked && checkBox12.Checked)
                {
                    MessageBox.Show("暂不支持同时显示多个情景，请重新选择");
                    return;
                    //timer1.Enabled = false;
                    //timer2.Enabled = false;
                }
                else
                {
                    tmYear.Enabled = true;
                    tmYear.Start();
                    t = 0;

                    if (cbRCP26.Checked)
                    {
                        TemperatureChangeRCP26();
                        RainChangeRCP26();

                        CheckBoxUnable();
                    }

                    if (cbRCP45.Checked)
                    {
                        TemperatureChangeRCP45();
                        RainChangeRCP45();

                        CheckBoxUnable();
                    }

                    if (cbRCP60.Checked)
                    {
                        TemperatureChangeRCP60();
                        RainChangeRCP60();

                        CheckBoxUnable();
                    }

                    if (cbRCP85.Checked)
                    {
                        TemperatureChangeRCP85();
                        RainChangeRCP85();

                        CheckBoxUnable();
                    }


                    //Green
                    if (checkBox12.Checked)
                    {
                       
                        SupplyGreenChange();

                        EnergyGreenChange();
                        AgriGreenChange();
                        UsingWaterGreenChange();

                        JobGreenChange();
                        PopulationGreenChange();
                        CityGreenChange();

                        GovernmentGreenChange();
                        TradeGreenChange();
                        LabourGreenChange();

                        CheckBoxUnable();
                    }

                    //NatureWork
                    if (checkBox11.Checked)
                    {
                        SupplyNatureChange();

                        EnergyNatureChange();
                        AgriNatureChange();
                        UsingWaterNatureChange();

                        JobNatureChange();
                        PopulationNatureChange();
                        CityNatureChange();

                        GovernmentNatureChange();
                        TradeNatureChange();
                        LabourNatureChange();

                        CheckBoxUnable();
                    }

                    //Word
                    if (checkBox10.Checked)
                    {
                        SupplyWorldChange();

                        EnergyWorldChange();
                        AgriWorldChange();
                        UsingWaterWorldChange();

                        JobWorldChange();
                        PopulationWorldChange();
                        CityWorldChange();

                        GovernmentWorldChange();
                        TradeWorldChange();
                        LabourWorldChange();
                        CheckBoxUnable();
                    }

                    //National
                    if (checkBox9.Checked)
                    {
                        SupplyNationalChange();

                        EnergyNationalChange();
                        AgriNationalChange();
                        UsingWaterNationalChange();

                        JobNationalChange();
                        PopulationNationalChange();
                        CityNationalChange();

                        GovernmentNationalChange();
                        TradeNationalChange();
                        LabourNationalChange();
                        CheckBoxUnable();
                    }

                    //Local
                    if (checkBox8.Checked)
                    {
                        SupplyLocalChange();

                        EnergyLocalChange();
                        AgriLocalChange();
                        UsingWaterLocalChange();

                        JobLocalChange();
                        PopulationLocalChange();
                        CityLocalChange();

                        GovernmentLocalChange();
                        TradeLocalChange();
                        LabourLocalChange();
                        CheckBoxUnable();
                    }

                    //Go
                    if (checkBox7.Checked)
                    {
                        SupplyGoChange();

                        EnergyGoChange();
                        AgriGoChange();
                        UsingWaterGoChange();

                        JobGoChange();
                        PopulationGoChange();
                        CityGoChange();

                        GovernmentGoChange();
                        TradeGoChange();
                        LabourGoChange();
                        CheckBoxUnable();
                    }
                }



            }
            




        }


        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = false;
            tmYear.Enabled = false;
            allTimer.Enabled = false;
            CheckBoxAble();
        }

        private void allTimer_Tick(object sender, EventArgs e)
        {
            if (cbRCP26.Checked)
            {
                TimerTemperaturRCP26();
                TimerRainRCP26();
            }

            if (cbRCP45.Checked)
            {
                TimerTemperaturRCP45();
                TimerRainRCP45();
            }
            if (cbRCP60.Checked)
            {
                TimerTemperaturRCP60();
                TimerRainRCP60();
            }
            if (cbRCP85.Checked)
            {
                TimerTemperaturRCP85();
                TimerRainRCP85();
            }


        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            //temperature--Green
            if (checkBox12.Checked)
            {
               

                EnergyTimerGreen();
                UsingWaterTimerGreen();

                JobTimerGreen();

                GovernmentTimerGreen();
                LabourTimerGreen();
            }


            //temperature--Nature
            if (checkBox11.Checked)
            {
            

                EnergyTimerNature();
                UsingWaterTimerNature();

                JobTimerNature();

                GovernmentTimerNature();
                LabourTimerNature();
            }

            //temperature--Word
            if (checkBox10.Checked)
            {
                EnergyTimerWorld();
                UsingWaterTimerWorld();

                JobTimerWorld();

                GovernmentTimerWorld();
                LabourTimerWorld();
            }

            //temperature--National
            if (checkBox9.Checked)
            {
                EnergyTimerNational();
                UsingWaterTimerNational();

                JobTimerNational();

                GovernmentTimerNational();
                LabourTimerNational();
            }

            //temperature--Local
            if (checkBox8.Checked)
            {
              

                EnergyTimerLocal();
                UsingWaterTimerLocal();

                JobTimerLocal();

                GovernmentTimerLocal();
                LabourTimerLocal();
            }

            //temperature--Go
            if (checkBox7.Checked)
            {
                

                EnergyTimerGo();
                UsingWaterTimerGo();

                JobTimerGo();

                GovernmentTimerGo();
                LabourTimerGo();
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (checkBox12.Checked)
            {
                SupplyTimerGreen();
                AgriTimerGreen();
                TradeTimerGreen();
            }

            if (checkBox11.Checked)
            {
                SupplyTimerNature();
                AgriTimerNature();
                TradeTimerNature();
            }
            if (checkBox10.Checked)
            {
                SupplyTimerWorld();
                AgriTimerWorld();
                TradeTimerWorld();
            }
            if (checkBox9.Checked)
            {
                SupplyTimerNational();
                AgriTimerNational();
                TradeTimerNational();
            }
            if (checkBox8.Checked)
            {
                SupplyTimerLocal();
                AgriTimerLocal();
                TradeTimerLocal();
            }
            if (checkBox7.Checked)
            {
                SupplyTimerGo();
                AgriTimerGo();
                TradeTimerGo();
            }
        }

        private void tmYear_Tick(object sender, EventArgs e)
        {
            if (t <= 24)
            {
                int year = years[t];
                lbYearClimate.Text = year.ToString();
                lbYearEcology.Text = year.ToString();
                lbYearCity.Text = year.ToString();
                lbYearRoad.Text = year.ToString();
                t++;
            }
            else
            {
                tmYear.Enabled = false;
            }
        }


        private void CheckBoxUnable()
        {
            cbRCP26.Enabled = false;
            cbRCP45.Enabled = false;
            cbRCP60.Enabled = false;
            cbRCP85.Enabled = false;
            checkBox7.Enabled = false;
            checkBox8.Enabled = false;
            checkBox9.Enabled = false;
            checkBox10.Enabled = false;
            checkBox11.Enabled = false;
            checkBox12.Enabled = false;
            checkBox13.Enabled = false;
            checkBox14.Enabled = false;
            checkBox15.Enabled = false;
            checkBox16.Enabled = false;
            checkBox17.Enabled = false;
            checkBox18.Enabled = false;
            checkBox19.Enabled = false;
            checkBox20.Enabled = false;
            checkBox21.Enabled = false;
            checkBox22.Enabled = false;
            checkBox23.Enabled = false;
            checkBox24.Enabled = false;

        }

        private void CheckBoxAble()
        {
            cbRCP26.Enabled = true;
            cbRCP45.Enabled = true;
            cbRCP60.Enabled = true;
            cbRCP85.Enabled = true;
            checkBox7.Enabled = true;
            checkBox8.Enabled = true;
            checkBox9.Enabled = true;
            checkBox10.Enabled = true;
            checkBox11.Enabled = true;
            checkBox12.Enabled = true;
            checkBox13.Enabled = true;
            checkBox14.Enabled = true;
            checkBox15.Enabled = true;
            checkBox16.Enabled = true;
            checkBox17.Enabled = true;
            checkBox18.Enabled = true;
            checkBox19.Enabled = true;
            checkBox20.Enabled = true;
            checkBox21.Enabled = true;
            checkBox22.Enabled = true;
            checkBox23.Enabled = true;
            checkBox24.Enabled = true;
        }

        private void CannotChooseTogether()
        {
            if (checkBox7.Checked && checkBox8.Checked || checkBox7.Checked && checkBox9.Checked || checkBox7.Checked && checkBox10.Checked ||
                checkBox7.Checked && checkBox11.Checked || checkBox7.Checked && checkBox12.Checked || checkBox8.Checked && checkBox9.Checked ||
                checkBox8.Checked && checkBox10.Checked || checkBox8.Checked && checkBox11.Checked || checkBox8.Checked && checkBox12.Checked ||
                checkBox9.Checked && checkBox10.Checked || checkBox9.Checked && checkBox11.Checked || checkBox9.Checked && checkBox12.Checked ||
                checkBox10.Checked && checkBox11.Checked || checkBox10.Checked && checkBox12.Checked || checkBox11.Checked && checkBox12.Checked)
            {
                MessageBox.Show("暂不支持同时显示两个情景...尚在开发中");
                return;
                //timer1.Enabled = false; 
                //timer2.Enabled = false;

            }

        }



       
        
        #region 气候变化模块

        #region 气候数据图表
        PointPairList TLa = new PointPairList();
        PointPairList TLb = new PointPairList();
        PointPairList TLc = new PointPairList();
        PointPairList TLd = new PointPairList();

        int iTemRCP = 0;//rcp-temperature

        private void TemperatureGraph()
        {
            myPane = zedGraphControl1.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "气温年变化";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "温度(℃)";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 2005;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 2030;    //X轴最大100   

            myPane.XAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 4.5;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 7;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 0.1;//Y轴小步长0.1,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 0.5;//Y轴大步长为0.3，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl1.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl1.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl1.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl1.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.AliceBlue);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl1.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem ta = myPane.AddCurve("", TLa, Color.Green, SymbolType.Circle);
            LineItem tb = myPane.AddCurve("", TLb, Color.Yellow, SymbolType.Circle);
            LineItem tc = myPane.AddCurve("", TLc, Color.SkyBlue, SymbolType.Circle);
            LineItem td = myPane.AddCurve("", TLd, Color.Orange, SymbolType.Circle);
            
            //显示网格线
            zedGraphControl1.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl1.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();

            timer1.Enabled = false;





        }    //气温图表

        #region RCP2.6
        List<double> listTemperatureRCP26 = new List<double>();
        private void TemperatureChangeRCP26()
        {
            TLa.Clear();
            TLb.Clear();
            TLc.Clear();
            TLd.Clear();
            
            listTemperatureRCP26.Clear();
           
            
            lbTemperature.Text = "00.00" + "℃";
            GetTemperatureRCP26();
            allTimer.Enabled = true;
            allTimer.Start();
            iTemRCP = 0;

        }

        private void GetTemperatureRCP26()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                 + "\\Data\\TempRCP_regionamean.xls", 1);
            for (int i = 6; i <= 30; i++)
            {
                listTemperatureRCP26.Add(Convert.ToDouble(dt.Rows[i][1].ToString()));
            }
        }

        private void TimerTemperaturRCP26()
        {
            if (iTemRCP < 25)
            {

                TLa.Add(iTemRCP + 2006, listTemperatureRCP26[iTemRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();//立即显示
            if (iTemRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbTemperature.Text = Convert.ToDouble(listTemperatureRCP26[iTemRCP].ToString("0.00 ")) + "℃";//保留两位小数
            }

            iTemRCP = iTemRCP + 1;
        }
        #endregion

        #region RCP4.5
        List<double> listTemperatureRCP45 = new List<double>();
        private void TemperatureChangeRCP45()
        {
            TLa.Clear();
            TLb.Clear();
            TLc.Clear();
            TLd.Clear();
           
            listTemperatureRCP45.Clear();

            lbTemperature.Text = "00.00" + "℃";
            GetTemperatureRCP45();
            allTimer.Enabled = true;
            allTimer.Start();
            iTemRCP = 0;

        }

        private void GetTemperatureRCP45()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                 + "\\Data\\TempRCP_regionamean.xls", 1);
            for (int i = 6; i <= 30; i++)
            {
                listTemperatureRCP45.Add(Convert.ToDouble(dt.Rows[i][2].ToString()));
            }
        }

        private void TimerTemperaturRCP45()
        {
            if (iTemRCP < 25)
            {

                TLb.Add(iTemRCP + 2006, listTemperatureRCP45[iTemRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();//立即显示
            if (iTemRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbTemperature.Text = Convert.ToDouble(listTemperatureRCP45[iTemRCP].ToString("0.00 ")) + "℃";//保留两位小数
            }
            iTemRCP = iTemRCP + 1;
        }
        #endregion

        #region RCP6.0
        List<double> listTemperatureRCP60 = new List<double>();
        private void TemperatureChangeRCP60()
        {
            TLa.Clear();
            TLb.Clear();
            TLc.Clear();
            TLd.Clear();
            
            listTemperatureRCP60.Clear();

            lbTemperature.Text = "00.00" + "℃";
            GetTemperatureRCP60();
            allTimer.Enabled = true;
            allTimer.Start();
            iTemRCP = 0;

        }

        private void GetTemperatureRCP60()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                 + "\\Data\\TempRCP_regionamean.xls", 1);
            for (int i = 6; i <= 30; i++)
            {
                listTemperatureRCP60.Add(Convert.ToDouble(dt.Rows[i][3].ToString()));
            }
        }

        private void TimerTemperaturRCP60()
        {
            if (iTemRCP < 25)
            {

                TLc.Add(iTemRCP + 2006, listTemperatureRCP60[iTemRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();//立即显示
            if (iTemRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbTemperature.Text = Convert.ToDouble(listTemperatureRCP60[iTemRCP].ToString("0.00 ")) + "℃";//保留两位小数
            }
            iTemRCP = iTemRCP + 1;
        }
        #endregion

        #region RCP8.5
        List<double> listTemperatureRCP85 = new List<double>();
        private void TemperatureChangeRCP85()
        {
            TLa.Clear();
            TLb.Clear();
            TLc.Clear();
            TLd.Clear();
           
            listTemperatureRCP85.Clear();

            lbTemperature.Text = "00.00" + "℃";
            GetTemperatureRCP85();
            allTimer.Enabled = true;
            allTimer.Start();
            iTemRCP = 0;

        }

        private void GetTemperatureRCP85()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                 + "\\Data\\TempRCP_regionamean.xls", 1);
            for (int i = 6; i <= 30; i++)
            {
                listTemperatureRCP85.Add(Convert.ToDouble(dt.Rows[i][4].ToString()));
            }
        }

        private void TimerTemperaturRCP85()
        {
            if (iTemRCP < 25)
            {

                TLd.Add(iTemRCP + 2006, listTemperatureRCP85[iTemRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl1.AxisChange();
            zedGraphControl1.Refresh();//立即显示
            if (iTemRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbTemperature.Text = Convert.ToDouble(listTemperatureRCP85[iTemRCP].ToString("0.00 ")) + "℃";//保留两位小数
            }
            iTemRCP = iTemRCP + 1;
        }
        #endregion
        
        #endregion

        #region 降水数据图表
        PointPairList RLa = new PointPairList();
        PointPairList RLb = new PointPairList();
        PointPairList RLc = new PointPairList();
        PointPairList RLd = new PointPairList();

        int iRainRCP = 0;//rcp-temperature
        
        private void RainGraph()
        {
            myPane = zedGraphControl2.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "降水数据";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "降水量(mm)";//纵坐标


            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl2.GraphPane.XAxis.Scale.FontSpec.Size = 17f;
            zedGraphControl2.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl2.GraphPane.YAxis.Scale.FontSpec.Size = 17f;
            zedGraphControl2.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.AliceBlue);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl2.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem ra = myPane.AddCurve("", RLa, Color.Brown, SymbolType.Circle);
            LineItem rb = myPane.AddCurve("", RLb, Color.DarkOrchid, SymbolType.Circle);
            LineItem rc = myPane.AddCurve("", RLc, Color.YellowGreen, SymbolType.Circle);
            LineItem rd = myPane.AddCurve("", RLd, Color.Crimson, SymbolType.Circle);

            //显示网格线
            zedGraphControl2.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl2.GraphPane.XAxis.MajorGrid.IsVisible = true;


            //X轴
            myPane.XAxis.Scale.Min = 2005;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 2030;    //X轴最大100   

            myPane.XAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   



            //Y轴
            myPane.YAxis.Scale.Min = 260;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 360;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 2;//Y轴小步长5,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 10;//Y轴大步长为5，也就是显示文字的大间隔   

            zedGraphControl2.AxisChange();
            zedGraphControl2.Refresh();




        }

        #region RCP26
        List<double> listRainRCP26 = new List<double>();

        private void RainChangeRCP26()
        {
            RLa.Clear();
            RLb.Clear();
            RLc.Clear();
            RLd.Clear();

            listRainRCP26.Clear();

            lbRain.Text = "000.00" + "mm";
            GetRainRCP26();
            allTimer.Enabled = true;
            allTimer.Start();
            iRainRCP = 0;
        }  //降水变化---Green，Click事件触发

        private void GetRainRCP26()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                   + "\\Data\\TempRCP_regionamean.xls", 0);
            for (int i = 0; i <= 24; i++)
            {
                listRainRCP26.Add(Convert.ToDouble(dt.Rows[i][1].ToString()));
            }
        }

        private void TimerRainRCP26()
        {
            if (iRainRCP < 25)
            {
                RLa.Add(iRainRCP + 2006, listRainRCP26[iRainRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl2.AxisChange();
            zedGraphControl2.Refresh();//立即显示
            if (iRainRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbRain.Text = Convert.ToDouble(listRainRCP26[iRainRCP].ToString("0.00 ")) + "mm";//保留两位小数
            }
            iRainRCP = iRainRCP + 1;

        }

        #endregion

        #region RCP45
        List<double> listRainRCP45 = new List<double>();

        private void RainChangeRCP45()
        {
            RLa.Clear();
            RLb.Clear();
            RLc.Clear();
            RLd.Clear();

            listRainRCP45.Clear();

            lbRain.Text = "000.00" + "mm";
            GetRainRCP45();
            allTimer.Enabled = true;
            allTimer.Start();
            iRainRCP = 0;
        }  //降水变化---Green，Click事件触发

        private void GetRainRCP45()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                   + "\\Data\\TempRCP_regionamean.xls", 0);
            for (int i = 0; i < 25; i++)
            {
                listRainRCP45.Add(Convert.ToDouble(dt.Rows[i][2].ToString()));
            }
        }

        private void TimerRainRCP45()
        {
            if (iRainRCP < 25)
            {
                RLb.Add(iRainRCP + 2006, listRainRCP45[iRainRCP]);
            }
            zedGraphControl2.AxisChange();
            zedGraphControl2.Refresh();//立即显示
            if (iRainRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbRain.Text = Convert.ToDouble(listRainRCP45[iRainRCP].ToString("0.00 ")) + "mm";//保留两位小数
            }
            iRainRCP = iRainRCP + 1;

        }

        #endregion

        #region RCP60
        List<double> listRainRCP60 = new List<double>();

        private void RainChangeRCP60()
        {
            RLa.Clear();
            RLb.Clear();
            RLc.Clear();
            RLd.Clear();

            listRainRCP60.Clear();

            lbRain.Text = "000.00" + "mm";
            GetRainRCP60();
            allTimer.Enabled = true;
            allTimer.Start();
            iRainRCP = 0;
        }  //降水变化---Green，Click事件触发

        private void GetRainRCP60()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                   + "\\Data\\TempRCP_regionamean.xls", 0);
            for (int i = 0; i < 25; i++)
            {
                listRainRCP60.Add(Convert.ToDouble(dt.Rows[i][3].ToString()));
            }
        }

        private void TimerRainRCP60()
        {
            if (iRainRCP < 25)
            {
                RLc.Add(iRainRCP + 2006, listRainRCP60[iRainRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl2.AxisChange();
            zedGraphControl2.Refresh();//立即显示
            if (iRainRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbRain.Text = Convert.ToDouble(listRainRCP60[iRainRCP].ToString("0.00 ")) + "mm";//保留两位小数
            }
            iRainRCP = iRainRCP + 1;

        }

        #endregion

        #region RCP85
        List<double> listRainRCP85 = new List<double>();

        private void RainChangeRCP85()
        {
            RLa.Clear();
            RLb.Clear();
            RLc.Clear();
            RLd.Clear();

            listRainRCP85.Clear();

            lbRain.Text = "000.00" + "mm";
            GetRainRCP85();
            allTimer.Enabled = true;
            allTimer.Start();
            iRainRCP = 0;
        }  //降水变化---Green，Click事件触发

        private void GetRainRCP85()
        {
            DataTable dt = ReadExcel.ExcelToDataTable(System.IO.Directory.GetCurrentDirectory().ToString()
                   + "\\Data\\TempRCP_regionamean.xls", 0);
            for (int i = 0; i < 25; i++)
            {
                listRainRCP85.Add(Convert.ToDouble(dt.Rows[i][4].ToString()));
            }
        }

        private void TimerRainRCP85()
        {
            if (iRainRCP < 25)
            {
                RLd.Add(iRainRCP + 2006, listRainRCP85[iRainRCP]);
            }
            else
            {
                CheckBoxAble();
            }
            zedGraphControl2.AxisChange();
            zedGraphControl2.Refresh();//立即显示
            if (iRainRCP < 25)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbRain.Text = Convert.ToDouble(listRainRCP85[iRainRCP].ToString("0.00 ")) + "mm";//保留两位小数
            }
            iRainRCP = iRainRCP + 1;

        }

        #endregion

        #endregion

        #region zedgraph3
        PointPairList PPLa = new PointPairList();
        PointPairList PPLb = new PointPairList();
        PointPairList PPLc = new PointPairList();




        List<double> list1 = new List<double> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };
        List<double> list2 = new List<double> { 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20 };
        List<double> list3 = new List<double> { 0, 3, 5, 6, 8, 8, 9, 10, 12 };
        private void SupplyGraph()
        {
            GraphPane myPane = zedGraphControl3.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框

            zedGraphControl3.GraphPane.YAxis.MajorGrid.IsVisible = true;//显示网格线

            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.AliceBlue);

            zedGraphControl3.GraphPane.Chart.Fill = new Fill(Color.AliceBlue);  //设置图表内部颜色

            myPane.Title.Text = "标题"; //设计图表的标题    
            myPane.XAxis.Title.Text = "";//X轴标题    
            myPane.YAxis.Title.Text = " ";//Y轴标题   
            myPane.Title.IsVisible = false;//标题隐藏

            myPane.BarSettings.ClusterScaleWidth = 2;

            BarItem a = myPane.AddBar("A", PPLa, Color.Green);
            BarItem b = myPane.AddBar("B", PPLb, Color.Gray);
            BarItem c = myPane.AddBar("C", PPLc, Color.Blue);
           
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 4;    //X轴最大30   
            myPane.XAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 1;//X轴大步长为5，也就是显示文字的大间隔   
            myPane.YAxis.Scale.Min = 0;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 20;   //Y轴最大值3
            myPane.YAxis.Scale.MinorStep = 5;//Y轴小步长3
            myPane.YAxis.Scale.MajorStep = 5;//Y轴大步长为3

            zedGraphControl3.GraphPane.XAxis.Scale.FontSpec.Size = 20f;
            zedGraphControl3.GraphPane.YAxis.Scale.FontSpec.Size = 20f;


            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//这句话非常重要，否则不会立即显示
            timer2.Enabled = false;



        }

        private void SupplyGreenChange()
        {
            PPLa.Clear();
            PPLb.Clear();
            PPLc.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void SupplyTimerGreen()
        {
            if (j <= 8)
            {
                PPLa.Add(1.5, list1[j]);
            }

            if (j <= 8)
            {
                PPLb.Add(2, list2[j]);

            }

            if (j <= 7)
            {
                PPLc.Add(2.5, list3[j]);
            }
            j++;
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//立即显示  
        }

        private void SupplyNatureChange()
        {
            PPLa.Clear();
            PPLb.Clear();
            PPLc.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void SupplyTimerNature()
        {
            if (j <= 3)
            {
                PPLa.Add(1.5, list1[j]);
            }

            if (j <= 8)
            {
                PPLb.Add(2, list2[j]);

            }

            if (j <= 7)
            {
                PPLc.Add(2.5, list3[j]);
            }
            j++;
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//立即显示  
        }

        private void SupplyWorldChange()
        {
            PPLa.Clear();
            PPLb.Clear();
            PPLc.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void SupplyTimerWorld()
        {
            if (j <= 7)
            {
                PPLa.Add(1.5, list1[j]);
            }

            if (j <= 9)
            {
                PPLb.Add(2, list2[j]);

            }

            if (j <= 6)
            {
                PPLc.Add(2.5, list3[j]);
            }
            j++;
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//立即显示  
        }

        private void SupplyNationalChange()
        {
            PPLa.Clear();
            PPLb.Clear();
            PPLc.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void SupplyTimerNational()
        {
            if (j <= 7)
            {
                PPLa.Add(1.5, list1[j]);
            }

            if (j <= 3)
            {
                PPLb.Add(2, list2[j]);

            }

            if (j <= 8)
            {
                PPLc.Add(2.5, list3[j]);
            }
            j++;
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//立即显示  
        }

        private void SupplyLocalChange()
        {
            PPLa.Clear();
            PPLb.Clear();
            PPLc.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void SupplyTimerLocal()
        {
            if (j <= 3)
            {
                PPLa.Add(1.5, list1[j]);
            }

            if (j <= 7)
            {
                PPLb.Add(2, list2[j]);

            }

            if (j <= 7)
            {
                PPLc.Add(2.5, list3[j]);
            }
            j++;
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//立即显示  
        }

        private void SupplyGoChange()
        {
            PPLa.Clear();
            PPLb.Clear();
            PPLc.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void SupplyTimerGo()
        {
            if (j <= 9)
            {
                PPLa.Add(1.5, list1[j]);
            }

            if (j <= 6)
            {
                PPLb.Add(2, list2[j]);

            }

            if (j <= 7)
            {
                PPLc.Add(2.5, list3[j]);
            }
            j++;
            zedGraphControl3.AxisChange();
            zedGraphControl3.Refresh();//立即显示  
        }

        #endregion

        #endregion

        #region 生态文明模块

        #region 能源图表
        PointPairList ELa = new PointPairList();
        PointPairList ELb = new PointPairList();
        PointPairList ELc = new PointPairList();

        private void EnergyGraph()
        {
            myPane = zedGraphControl4.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "中国能源供需";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "Mtoe";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 2000;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 4500;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 500;//Y轴小步长，也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 500;//Y轴大步长，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl4.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl4.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl4.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl4.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.Honeydew);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl4.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem ea = myPane.AddCurve("能源供应", ELa, Color.Green);
            LineItem eb = myPane.AddCurve("能源需求", ELb, Color.Yellow);
            LineItem ec = myPane.AddCurve("IEa需求", ELc, Color.SkyBlue);

            //显示网格线
            zedGraphControl4.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl4.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();

            timer1.Enabled = false;





        }    //能源图表
        private void EnergyGreenChange()
        {
            ELa.Clear();
            ELb.Clear();
            ELc.Clear();

            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void EnergyTimerGreen()
        {
            if (i <= 30)
            {
                ELa.Add(i, 200 * Math.Sqrt(i) + 1900);
                ELb.Add(i, 220 * Math.Sqrt(i) + 2100);
                ELc.Add(i, 220 * Math.Sqrt(i) + 2200);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 1;
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();//立即显示 

        }

        private void EnergyNatureChange()
        {
            ELa.Clear();
            ELb.Clear();
            ELc.Clear();

            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void EnergyTimerNature()
        {
            if (i <= 30)
            {
                ELa.Add(i, 260 * Math.Sqrt(i) + 2000);
                ELb.Add(i, 220 * Math.Sqrt(i) + 2200);
                ELc.Add(i, 220 * Math.Sqrt(i) + 2500);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 1;
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();//立即显示 

        }

        private void EnergyWorldChange()
        {
            ELa.Clear();
            ELb.Clear();
            ELc.Clear();

            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void EnergyTimerWorld()
        {
            if (i <= 30)
            {
                ELa.Add(i, 200 * Math.Sqrt(i) + 2000);
                ELb.Add(i, 220 * Math.Sqrt(i) + 2200);
                ELc.Add(i, 280 * Math.Sqrt(i) + 2500);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 1;
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();//立即显示 

        }

        private void EnergyNationalChange()
        {
            ELa.Clear();
            ELb.Clear();
            ELc.Clear();

            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void EnergyTimerNational()
        {
            if (i <= 30)
            {
                ELa.Add(i, 200 * Math.Sqrt(i) + 2500);
                ELb.Add(i, 220 * Math.Sqrt(i) + 2200);
                ELc.Add(i, 220 * Math.Sqrt(i) + 2500);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 1;
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();//立即显示 

        }

        private void EnergyLocalChange()
        {
            ELa.Clear();
            ELb.Clear();
            ELc.Clear();

            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void EnergyTimerLocal()
        {
            if (i <= 30)
            {
                ELa.Add(i, 200 * Math.Sqrt(i) + 2000);
                ELb.Add(i, 300 * Math.Sqrt(i) + 2200);
                ELc.Add(i, 220 * Math.Sqrt(i) + 2500);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 1;
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();//立即显示 

        }

        private void EnergyGoChange()
        {
            ELa.Clear();
            ELb.Clear();
            ELc.Clear();

            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void EnergyTimerGo()
        {
            if (i <= 30)
            {
                ELa.Add(i, 200 * Math.Sqrt(i) + 2000);
                ELb.Add(i, 220 * Math.Sqrt(i) + 2200);
                ELc.Add(i, 250 * Math.Sqrt(i) + 3000);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 1;
            zedGraphControl4.AxisChange();
            zedGraphControl4.Refresh();//立即显示 

        }

        #endregion

        #region 农业灌溉用水效率
        List<double> elist = new List<double> { 0, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100 };

        PointPairList APLa = new PointPairList();
        PointPairList APLb = new PointPairList();
        PointPairList APLc = new PointPairList();
        PointPairList APLd = new PointPairList();
        PointPairList APLe = new PointPairList();
        PointPairList APLf = new PointPairList();

        private void AgriWaterGraph()
        {
            GraphPane myPane = zedGraphControl5.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框

            zedGraphControl5.GraphPane.YAxis.MajorGrid.IsVisible = true;//显示网格线

            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.Honeydew);

            zedGraphControl5.GraphPane.Chart.Fill = new Fill(Color.Honeydew);  //设置图表内部颜色

            myPane.Title.Text = "标题"; //设计图表的标题    
            myPane.XAxis.Title.Text = "";//X轴标题    
            myPane.YAxis.Title.Text = " ";//Y轴标题   
            myPane.Title.IsVisible = false;//标题隐藏

            myPane.BarSettings.ClusterScaleWidth = 5;

            //BarItem a = myPane.AddBar("", APLa, Color.Green);
            //BarItem b = myPane.AddBar("", APLb, Color.Green);
            //BarItem c= myPane.AddBar("", APLc, Color.Green);
            //BarItem d = myPane.AddBar("", APLd, Color.Green);
            //BarItem e = myPane.AddBar("", APLe, Color.Green);
            //BarItem f = myPane.AddBar("", APLf, Color.Green);
            myPane.AddBar("", APLa, Color.Blue).Bar.Fill = new Fill(Color.Red, Color.White, Color.Red, 0f);
            myPane.AddBar("", APLb, Color.Blue).Bar.Fill = new Fill(Color.Cyan, Color.White, Color.Cyan, 0f);
            myPane.AddBar("", APLc, Color.Blue).Bar.Fill = new Fill(Color.PeachPuff, Color.White, Color.PeachPuff, 0f);
            myPane.AddBar("", APLd, Color.Blue).Bar.Fill = new Fill(Color.Gold, Color.White, Color.Gold, 0f);
            myPane.AddBar("", APLe, Color.Blue).Bar.Fill = new Fill(Color.DarkOrange, Color.White, Color.DarkOrange, 0f);
            myPane.AddBar("", APLf, Color.Blue).Bar.Fill = new Fill(Color.ForestGreen, Color.White, Color.ForestGreen, 0f);




            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 2;    //X轴最大30   
            myPane.XAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 1;//X轴大步长为5，也就是显示文字的大间隔   
            myPane.YAxis.Scale.Min = 0;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 100;   //Y轴最大值3
            myPane.YAxis.Scale.MinorStep = 5;//Y轴小步长3
            myPane.YAxis.Scale.MajorStep = 20;//Y轴大步长为3

            zedGraphControl5.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl5.GraphPane.YAxis.Scale.FontSpec.Size = 15f;


            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//这句话非常重要，否则不会立即显示
            timer2.Enabled = false;



        }
        private void AgriGreenChange()
        {
            APLa.Clear();
            APLb.Clear();
            APLc.Clear();
            APLd.Clear();
            APLe.Clear();
            APLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void AgriTimerGreen()
        {
            if (j <= 16)
            {
                APLa.Add(2.85, elist[j]);
            }
            j++;
            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//立即显示  
        }

        private void AgriNatureChange()
        {
            APLa.Clear();
            APLb.Clear();
            APLc.Clear();
            APLd.Clear();
            APLe.Clear();
            APLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void AgriTimerNature()
        {
            if (j <= 15)
            {
                APLb.Add(2.1, elist[j]);
            }
            j++;
            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//立即显示  
        }

        private void AgriWorldChange()
        {
            APLa.Clear();
            APLb.Clear();
            APLc.Clear();
            APLd.Clear();
            APLe.Clear();
            APLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void AgriTimerWorld()
        {
            if (j <= 16)
            {
                APLc.Add(1.4, elist[j]);
            }
            j++;
            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//立即显示  
        }

        private void AgriNationalChange()
        {
            APLa.Clear();
            APLb.Clear();
            APLc.Clear();
            APLd.Clear();
            APLe.Clear();
            APLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void AgriTimerNational()
        {
            if (j <= 14)
            {
                APLd.Add(0.59, elist[j]);
            }
            j++;
            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//立即显示  
        }

        private void AgriLocalChange()
        {
            APLa.Clear();
            APLb.Clear();
            APLc.Clear();
            APLd.Clear();
            APLe.Clear();
            APLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void AgriTimerLocal()
        {
            if (j <= 17)
            {
                APLe.Add(-0.15, elist[j]);
            }
            j++;
            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//立即显示  
        }

        private void AgriGoChange()
        {
            APLa.Clear();
            APLb.Clear();
            APLc.Clear();
            APLd.Clear();
            APLe.Clear();
            APLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void AgriTimerGo()
        {
            if (j <= 17)
            {
                APLe.Add(-0.2, elist[j]);
            }
            j++;
            zedGraphControl5.AxisChange();
            zedGraphControl5.Refresh();//立即显示
        }
        #endregion

        #region 全国用水总量
        PointPairList ULa = new PointPairList();
        PointPairList ULb = new PointPairList();
        PointPairList ULc = new PointPairList();
        PointPairList ULd = new PointPairList();
        PointPairList ULe = new PointPairList();
        PointPairList ULf = new PointPairList();

        private void UsingWaterGraph()
        {
            myPane = zedGraphControl6.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "全国用水总量";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "总量（m³）";//纵坐标


            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl6.GraphPane.XAxis.Scale.FontSpec.Size = 19f;
            zedGraphControl6.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl6.GraphPane.YAxis.Scale.FontSpec.Size = 19f;
            zedGraphControl6.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.Honeydew);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl6.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem ua = myPane.AddCurve("", ULa, Color.Brown);
            LineItem ub = myPane.AddCurve("", ULb, Color.DarkSalmon);
            LineItem uc = myPane.AddCurve("", ULc, Color.YellowGreen);
            LineItem ud = myPane.AddCurve("", ULd, Color.Crimson);
            LineItem ue = myPane.AddCurve("", ULe, Color.Plum);
            LineItem uf = myPane.AddCurve("", ULf, Color.Magenta);

            //显示网格线
            zedGraphControl6.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl6.GraphPane.XAxis.MajorGrid.IsVisible = true;


            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 5300;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 6300;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 100;//Y轴小步长5,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 100;//Y轴大步长为5，也就是显示文字的大间隔   

            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();




        }
        private void UsingWaterGreenChange()
        {
            ULa.Clear();
            ULb.Clear();
            ULc.Clear();
            ULd.Clear();
            ULe.Clear();
            ULf.Clear();
            lbUsingWater.Text = "" + "m³";
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void UsingWaterTimerGreen()
        {
            if (i <= 30)
            {
                ULa.Add(i, 20 * i + 5600);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.01;
            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();//立即显示 
            if (i <= 30)
            {
                lbUsingWater.Text = (20 * i + 5600).ToString("00.00") + "m³";//保留两位小数
            }

        }

        private void UsingWaterNatureChange()
        {
            ULa.Clear();
            ULb.Clear();
            ULc.Clear();
            ULd.Clear();
            ULe.Clear();
            ULf.Clear();
            lbUsingWater.Text = "" + "m³";
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void UsingWaterTimerNature()
        {
            if (i <= 30)
            {
                ULb.Add(i, 15 * i + 5700);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.01;
            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();//立即显示 
            if (i <= 30)
            {
                lbUsingWater.Text = (15 * i + 5700).ToString("00.00") + "m³";//保留两位小数
            }

        }

        private void UsingWaterWorldChange()
        {
            ULa.Clear();
            ULb.Clear();
            ULc.Clear();
            ULd.Clear();
            ULe.Clear();
            ULf.Clear();
            lbUsingWater.Text = "" + "m³";
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void UsingWaterTimerWorld()
        {
            if (i <= 30)
            {
                ULc.Add(i, 30 * i + 5300);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.01;
            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();//立即显示 
            if (i <= 30)
            {
                lbUsingWater.Text = (30 * i + 5300).ToString("00.00") + "m³";//保留两位小数
            }

        }


        private void UsingWaterNationalChange()
        {
            ULa.Clear();
            ULb.Clear();
            ULc.Clear();
            ULd.Clear();
            ULe.Clear();
            ULf.Clear();
            lbUsingWater.Text = "" + "m³";
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void UsingWaterTimerNational()
        {
            if (i <= 30)
            {
                ULd.Add(i, 28 * i + 5400);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.01;
            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();//立即显示 
            if (i <= 30)
            {
                lbUsingWater.Text = (28 * i + 5400).ToString("00.00") + "m³";//保留两位小数
            }

        }


        private void UsingWaterLocalChange()
        {
            ULa.Clear();
            ULb.Clear();
            ULc.Clear();
            ULd.Clear();
            ULe.Clear();
            ULf.Clear();
            lbUsingWater.Text = "" + "m³";
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void UsingWaterTimerLocal()
        {
            if (i <= 30)
            {
                ULe.Add(i, 0.8 * i * i + 5600);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.01;
            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();//立即显示 
            if (i <= 30)
            {
                lbUsingWater.Text = (0.8 * i * i + 5600).ToString("00.00") + "m³";//保留两位小数
            }

        }


        private void UsingWaterGoChange()
        {
            ULa.Clear();
            ULb.Clear();
            ULc.Clear();
            ULd.Clear();
            ULe.Clear();
            ULf.Clear();
            lbUsingWater.Text = "" + "m³";
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void UsingWaterTimerGo()
        {
            if (i <= 30)
            {
                ULf.Add(i, i * i + 5500);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.01;
            zedGraphControl6.AxisChange();
            zedGraphControl6.Refresh();//立即显示 
            if (i <= 30)
            {
                lbUsingWater.Text = (i * i + 5500).ToString("00.00") + "m³";//保留两位小数
            }

        }

        #endregion

        #endregion

        #region 新型城镇化模块

        #region 就业率
        PointPairList JLa = new PointPairList();
        PointPairList JLb = new PointPairList();
        PointPairList JLc = new PointPairList();
        PointPairList JLd = new PointPairList();
        PointPairList JLe = new PointPairList();
        PointPairList JLf = new PointPairList();
        private void JobGragh()
        {
            myPane = zedGraphControl10.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "就业率（%）";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 0;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 100;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 10;//Y轴小步长0.1,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 10;//Y轴大步长为0.3，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl10.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl10.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl10.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl10.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.LemonChiffon);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl1.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem ja = myPane.AddCurve("", JLa, Color.Brown);
            LineItem jb = myPane.AddCurve("", JLb, Color.Yellow);
            LineItem jc = myPane.AddCurve("", JLc, Color.SkyBlue);
            LineItem jd = myPane.AddCurve("", JLd, Color.Blue);
            LineItem je = myPane.AddCurve("", JLe, Color.Orange);
            LineItem jf = myPane.AddCurve("", JLf, Color.Purple);
            //显示网格线
            zedGraphControl10.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl10.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();

            timer1.Enabled = false;

        }
        private void JobGreenChange()
        {
            JLa.Clear();
            JLb.Clear();
            JLc.Clear();
            JLd.Clear();
            JLe.Clear();
            JLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void JobTimerGreen()
        {
            if (i <= 30)
            {
                JLa.Add(i, -0.3 * i + 70);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();//立即显示 

        }

        private void JobNatureChange()
        {
            JLa.Clear();
            JLb.Clear();
            JLc.Clear();
            JLd.Clear();
            JLe.Clear();
            JLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void JobTimerNature()
        {
            if (i <= 30)
            {
                JLb.Add(i, -0.4 * i + 60);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();//立即显示 

        }

        private void JobWorldChange()
        {
            JLa.Clear();
            JLb.Clear();
            JLc.Clear();
            JLd.Clear();
            JLe.Clear();
            JLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void JobTimerWorld()
        {
            if (i <= 30)
            {
                JLc.Add(i, -0.5 * i + 70);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();//立即显示 

        }

        private void JobNationalChange()
        {
            JLa.Clear();
            JLb.Clear();
            JLc.Clear();
            JLd.Clear();
            JLe.Clear();
            JLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void JobTimerNational()
        {
            if (i <= 30)
            {
                JLd.Add(i, -0.3 * i + 65);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();//立即显示 

        }

        private void JobLocalChange()
        {
            JLa.Clear();
            JLb.Clear();
            JLc.Clear();
            JLd.Clear();
            JLe.Clear();
            JLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void JobTimerLocal()
        {
            if (i <= 30)
            {
                JLe.Add(i, -0.3 * i + 60);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();//立即显示 

        }

        private void JobGoChange()
        {
            JLa.Clear();
            JLb.Clear();
            JLc.Clear();
            JLd.Clear();
            JLe.Clear();
            JLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void JobTimerGo()
        {
            if (i <= 30)
            {
                JLf.Add(i, -0.4 * i + 65);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl10.AxisChange();
            zedGraphControl10.Refresh();//立即显示 

        }
        #endregion

        #region 农民工总量
        PointPairList PLa = new PointPairList();
        PointPairList PLb = new PointPairList();
        PointPairList PLc = new PointPairList();
        PointPairList PLd = new PointPairList();
        PointPairList PLe = new PointPairList();
        PointPairList PLf = new PointPairList();
        private void PopulationGraph()
        {
            myPane = zedGraphControl9.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "人数（10²）";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 200;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 300;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 4;//Y轴小步长0.1,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 20;//Y轴大步长为0.3，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl9.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl9.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl9.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl9.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.LemonChiffon);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl9.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            BarItem pa = myPane.AddBar("", PLa, Color.Brown);
            BarItem pb = myPane.AddBar("", PLb, Color.Green);
            BarItem pc = myPane.AddBar("", PLc, Color.Blue);
            BarItem pd = myPane.AddBar("", PLd, Color.Red);
            BarItem pe = myPane.AddBar("", PLe, Color.Orange);
            BarItem pf = myPane.AddBar("", PLf, Color.Purple);
            myPane.BarSettings.ClusterScaleWidth = 10;
            //显示网格线

            zedGraphControl9.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl9.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();

        }
        private void PopulationGreenChange()
        {
            PLa.Clear();
            PLb.Clear();
            PLc.Clear();
            PLd.Clear();
            PLe.Clear();
            PLf.Clear();
            PLa.Add(8.8, 260);
            PLb.Add(12.3, 280);
            PLc.Add(15.7, 278);
            PLd.Add(19.2, 279);
            PLe.Add(22.8, 289);
            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();
        }
        private void PopulationNatureChange()
        {
            PLa.Clear();
            PLb.Clear();
            PLc.Clear();
            PLd.Clear();
            PLe.Clear();
            PLf.Clear();
            PLa.Add(8.8, 260);
            PLb.Add(12.3, 280);
            PLc.Add(15.7, 278);
            PLd.Add(19.2, 256);
            PLe.Add(22.8, 299);
            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();
        }

        private void PopulationWorldChange()
        {
            PLa.Clear();
            PLb.Clear();
            PLc.Clear();
            PLd.Clear();
            PLe.Clear();
            PLf.Clear();
            PLa.Add(8.8, 260);
            PLb.Add(12.3, 280);
            PLc.Add(15.7, 278);
            PLd.Add(19.2, 269);
            PLe.Add(22.8, 299);
            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();
        }
        private void PopulationNationalChange()
        {
            PLa.Clear();
            PLb.Clear();
            PLc.Clear();
            PLd.Clear();
            PLe.Clear();
            PLf.Clear();
            PLa.Add(8.8, 260);
            PLb.Add(12.3, 280);
            PLc.Add(15.7, 278);
            PLd.Add(19.2, 260);
            PLe.Add(22.8, 290);
            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();
        }
        private void PopulationLocalChange()
        {
            PLa.Clear();
            PLb.Clear();
            PLc.Clear();
            PLd.Clear();
            PLe.Clear();
            PLf.Clear();
            PLa.Add(8.8, 260);
            PLb.Add(12.3, 280);
            PLc.Add(15.7, 278);
            PLd.Add(19.2, 277);
            PLe.Add(22.8, 289);
            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();
        }
        private void PopulationGoChange()
        {
            PLa.Clear();
            PLb.Clear();
            PLc.Clear();
            PLd.Clear();
            PLe.Clear();
            PLf.Clear();
            PLa.Add(8.8, 260);
            PLb.Add(12.3, 280);
            PLc.Add(15.7, 278);
            PLd.Add(19.2, 267);
            PLe.Add(22.8, 281);
            zedGraphControl9.AxisChange();
            zedGraphControl9.Refresh();
        }
        #endregion 

        #region 城镇化率
        PointPairList CLa = new PointPairList();
        PointPairList CLb = new PointPairList();
        PointPairList CLc = new PointPairList();
        PointPairList CLd = new PointPairList();
        PointPairList CLe = new PointPairList();
        PointPairList CLf = new PointPairList();
        private void CityRate()
        {
            myPane = zedGraphControl7.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "城镇化率（%）";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 0;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 100;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 4;//Y轴小步长0.1,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 20;//Y轴大步长为0.3，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl7.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl7.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl7.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl7.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.LemonChiffon);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl7.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            myPane.AddBar("", CLa, Color.Blue).Bar.Fill = new Fill(Color.Red, Color.White, Color.Red, 0f);
            myPane.AddBar("", CLb, Color.Blue).Bar.Fill = new Fill(Color.Cyan, Color.White, Color.Cyan, 0f);
            myPane.AddBar("", CLc, Color.Blue).Bar.Fill = new Fill(Color.PeachPuff, Color.White, Color.PeachPuff, 0f);
            myPane.AddBar("", CLd, Color.Blue).Bar.Fill = new Fill(Color.Gold, Color.White, Color.Gold, 0f);
            myPane.AddBar("", CLe, Color.Blue).Bar.Fill = new Fill(Color.DarkOrange, Color.White, Color.DarkOrange, 0f);
            myPane.AddBar("", CLf, Color.Blue).Bar.Fill = new Fill(Color.ForestGreen, Color.White, Color.ForestGreen, 0f);
            myPane.BarSettings.ClusterScaleWidth = 10;
            //显示网格线

            zedGraphControl7.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl7.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();

        }

        private void CityGreenChange()
        {
            CLa.Clear();
            CLb.Clear();
            CLc.Clear();
            CLd.Clear();
            CLe.Clear();
            CLf.Clear();
            CLa.Add(8.8, 40);
            CLb.Add(12.3, 45);
            CLc.Add(15.7, 50);
            CLd.Add(19.2, 55);
            CLe.Add(22.8, 59);
            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();
        }
        private void CityNatureChange()
        {
            CLa.Clear();
            CLb.Clear();
            CLc.Clear();
            CLd.Clear();
            CLe.Clear();
            CLf.Clear();
            CLa.Add(8.8, 40);
            CLb.Add(12.3, 45);
            CLc.Add(15.7, 50);
            CLd.Add(19.2, 51);
            CLe.Add(22.8, 65);
            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();
        }
        private void CityWorldChange()
        {
            CLa.Clear();
            CLb.Clear();
            CLc.Clear();
            CLd.Clear();
            CLe.Clear();
            CLf.Clear();
            CLa.Add(8.8, 40);
            CLb.Add(12.3, 45);
            CLc.Add(15.7, 50);
            CLd.Add(19.2, 52);
            CLe.Add(22.8, 59);
            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();
        }
        private void CityNationalChange()
        {
            CLa.Clear();
            CLb.Clear();
            CLc.Clear();
            CLd.Clear();
            CLe.Clear();
            CLf.Clear();
            CLa.Add(8.8, 40);
            CLb.Add(12.3, 45);
            CLc.Add(15.7, 50);
            CLd.Add(19.2, 65);
            CLe.Add(22.8, 77);
            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();
        }
        private void CityLocalChange()
        {
            CLa.Clear();
            CLb.Clear();
            CLc.Clear();
            CLd.Clear();
            CLe.Clear();
            CLf.Clear();
            CLa.Add(8.8, 40);
            CLb.Add(12.3, 45);
            CLc.Add(15.7, 50);
            CLd.Add(19.2, 47);
            CLe.Add(22.8, 65);
            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();
        }
        private void CityGoChange()
        {
            CLa.Clear();
            CLb.Clear();
            CLc.Clear();
            CLd.Clear();
            CLe.Clear();
            CLf.Clear();
            CLa.Add(8.8, 40);
            CLb.Add(12.3, 45);
            CLc.Add(15.7, 50);
            CLd.Add(19.2, 49);
            CLe.Add(22.8, 54);
            zedGraphControl7.AxisChange();
            zedGraphControl7.Refresh();
        }
        #endregion

        #endregion

        #region 产业发展模块

        #region 政府投资
        PointPairList GLa = new PointPairList();
        PointPairList GLb = new PointPairList();
        PointPairList GLc = new PointPairList();
        PointPairList GLd = new PointPairList();
        PointPairList GLe = new PointPairList();
        PointPairList GLf = new PointPairList();
        private void Government()
        {
            myPane = zedGraphControl8.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "亿元";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 50;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 100;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 5;//Y轴小步长0.1,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 10;//Y轴大步长为0.3，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl8.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl8.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl8.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl8.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.LavenderBlush);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl8.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem ga = myPane.AddCurve("", GLa, Color.Aquamarine);
            LineItem gb = myPane.AddCurve("", GLb, Color.LimeGreen);
            LineItem gc = myPane.AddCurve("", GLc, Color.DeepSkyBlue);
            LineItem gd = myPane.AddCurve("", GLd, Color.GreenYellow);
            LineItem ge = myPane.AddCurve("", GLe, Color.LightCoral);
            LineItem gf = myPane.AddCurve("", GLf, Color.LightGray);
            //显示网格线
            zedGraphControl8.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl8.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();

            timer1.Enabled = false;
        }

        private void GovernmentGreenChange()
        {
            GLa.Clear();
            GLb.Clear();
            GLc.Clear();
            GLd.Clear();
            GLe.Clear();
            GLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void GovernmentTimerGreen()
        {
            if (i <= 30)
            {
                GLa.Add(i, 0.03*i*i + 70);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbGovernment.Text = (0.03 * i * i + 70).ToString(("0.00 ")) + "亿元";//保留两位小数
            }

        }

        private void GovernmentNatureChange()
        {
            GLa.Clear();
            GLb.Clear();
            GLc.Clear();
            GLd.Clear();
            GLe.Clear();
            GLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void GovernmentTimerNature()
        {
            if (i <= 30)
            {
                GLb.Add(i, 0.7 * i + 70);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbGovernment.Text = (0.7 * i + 70).ToString(("0.00 ")) + "亿元";//保留两位小数
            }


        }


        private void GovernmentWorldChange()
        {
            GLa.Clear();
            GLb.Clear();
            GLc.Clear();
            GLd.Clear();
            GLe.Clear();
            GLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void GovernmentTimerWorld()
        {
            if (i <= 30)
            {
                GLc.Add(i, 0.8 * i + 60);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbGovernment.Text = (0.8 * i + 60).ToString(("0.00 ")) + "亿元";//保留两位小数
            }

        }


        private void GovernmentNationalChange()
        {
            GLa.Clear();
            GLb.Clear();
            GLc.Clear();
            GLd.Clear();
            GLe.Clear();
            GLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void GovernmentTimerNational()
        {
            if (i <= 30)
            {
                GLd.Add(i, 0.02 * i * i + 70);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbGovernment.Text = (0.02 * i * i + 70).ToString(("0.00 ")) + "亿元";//保留两位小数
            }

        }


        private void GovernmentLocalChange()
        {
            GLa.Clear();
            GLb.Clear();
            GLc.Clear();
            GLd.Clear();
            GLe.Clear();
            GLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void GovernmentTimerLocal()
        {
            if (i <= 30)
            {
                GLe.Add(i, 0.5 * i + 80);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbGovernment.Text = (0.5 * i + 80).ToString(("0.00 ")) + "亿元";//保留两位小数
            }

        }


        private void GovernmentGoChange()
        {
            GLa.Clear();
            GLb.Clear();
            GLc.Clear();
            GLd.Clear();
            GLe.Clear();
            GLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void GovernmentTimerGo()
        {
            if (i <= 30)
            {
                GLf.Add(i, 1.3*Math.Sqrt(i)+70);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl8.AxisChange();
            zedGraphControl8.Refresh();//立即显示
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbGovernment.Text = (1.3 * Math.Sqrt(i) + 70).ToString(("0.00 ")) + "亿元";//保留两位小数
            }

        }
        #endregion

        #region 贸易
        PointPairList TPLa = new PointPairList();
        PointPairList TPLb = new PointPairList();
        PointPairList TPLc = new PointPairList();
        PointPairList TPLd = new PointPairList();
        PointPairList TPLe = new PointPairList();
        PointPairList TPLf = new PointPairList();

        List<double> tradelist1 = new List<double> { 0, 1, 20, 25, 30,35, 40, 45, 50, 55, 60, 65, 70, 75, 80};
        List<double> tradelist2 = new List<double> { 0, 1, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20 };
        List<double> tradelist3 = new List<double> { 0, 1, 3, 6, 9, 12, 15, 18, 21, 24,27, 30};
        List<double> tradelist4 = new List<double> { 0, 1, 2, 4, 6, 8, 10, 11, 12, 14, 16, 18, 20} ;
        List<double> tradelist5 = new List<double> { 0, 1, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20 };
        List<double> tradelist6 = new List<double> { 0, 1, 3, 5, 6, 8, 8, 9, 10, 12 };

       
        private void tradeGraph()
        {
            GraphPane myPane = zedGraphControl11.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框

            zedGraphControl11.GraphPane.YAxis.MajorGrid.IsVisible = true;//显示网格线

            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.LavenderBlush);

            zedGraphControl11.GraphPane.Chart.Fill = new Fill(Color.LavenderBlush);  //设置图表内部颜色

            myPane.Title.Text = "标题"; //设计图表的标题    
            myPane.XAxis.Title.Text = "";//X轴标题    
            myPane.YAxis.Title.Text = "百分比（%）";//Y轴标题   
            myPane.Title.IsVisible = false;//标题隐藏

            myPane.BarSettings.ClusterScaleWidth = 2;

            BarItem a = myPane.AddBar("Asia", TPLa, Color.Green);
            BarItem b = myPane.AddBar("Africa", TPLb, Color.Gray);
            BarItem c = myPane.AddBar("American", TPLc, Color.Blue);
            BarItem d = myPane.AddBar("Europe", TPLd, Color.Orange);
            BarItem e = myPane.AddBar("Antarctic", TPLe, Color.Red);
            BarItem f = myPane.AddBar("Oceania", TPLf, Color.Pink);

            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 7;    //X轴最大30   
            myPane.XAxis.Scale.MinorStep = 1;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 1;//X轴大步长为5，也就是显示文字的大间隔   
            myPane.YAxis.Scale.Min = 0;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 80;   //Y轴最大值3
            myPane.YAxis.Scale.MinorStep = 10;//Y轴小步长3
            myPane.YAxis.Scale.MajorStep = 10;//Y轴大步长为3

            zedGraphControl11.GraphPane.XAxis.Scale.FontSpec.Size = 20f;
            zedGraphControl11.GraphPane.YAxis.Scale.FontSpec.Size = 20f;


            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//这句话非常重要，否则不会立即显示
            timer2.Enabled = false;

        }
        
        private void TradeGreenChange()
        {
            TPLa.Clear();
            TPLb.Clear();
            TPLc.Clear();
            TPLd.Clear();
            TPLe.Clear();
            TPLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void TradeTimerGreen()
        {
            if (j <= 14)
            {
                TPLa.Add(1.75, tradelist1[j]);
            }

            if (j <= 10)
            {
                TPLb.Add(2.45, tradelist2[j]);

            }

            if (j <= 10)
            {
                TPLc.Add(3.15, tradelist3[j]);
            }
            if (j <= 6)
            {
                TPLd.Add(3.85, tradelist4[j]);
            }

            if (j <= 10)
            {
                TPLe.Add(4.55, tradelist5[j]);

            }

            if (j <= 6)
            {
                TPLf.Add(5.25, tradelist6[j]);
            }
            j++;
            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//立即显示  
        }

        private void TradeNatureChange()
        {
            TPLa.Clear();
            TPLb.Clear();
            TPLc.Clear();
            TPLd.Clear();
            TPLe.Clear();
            TPLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void TradeTimerNature()
        {
            if (j <= 13)
            {
                TPLa.Add(1.75, tradelist1[j]);
            }

            if (j <= 10)
            {
                TPLb.Add(2.45, tradelist2[j]);

            }

            if (j <= 7)
            {
                TPLc.Add(3.15, tradelist3[j]);
            }
            if (j <= 8)
            {
                TPLd.Add(3.85, tradelist4[j]);
            }

            if (j <= 10)
            {
                TPLe.Add(4.55, tradelist5[j]);

            }

            if (j <= 8)
            {
                TPLf.Add(5.25, tradelist6[j]);
            }
            j++;
            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//立即显示  
        }

        private void TradeWorldChange()
        {
            TPLa.Clear();
            TPLb.Clear();
            TPLc.Clear();
            TPLd.Clear();
            TPLe.Clear();
            TPLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void TradeTimerWorld()
        {
            if (j <= 13)
            {
                TPLa.Add(1.75, tradelist1[j]);
            }

            if (j <= 11)
            {
                TPLb.Add(2.45, tradelist2[j]);

            }

            if (j <= 10)
            {
                TPLc.Add(3.15, tradelist3[j]);
            }
            if (j <= 7)
            {
                TPLd.Add(3.85, tradelist4[j]);
            }

            if (j <= 8)
            {
                TPLe.Add(4.55, tradelist5[j]);

            }

            if (j <= 9)
            {
                TPLf.Add(5.25, tradelist6[j]);
            }
            j++;
            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//立即显示  
        }

        private void TradeNationalChange()
        {
            TPLa.Clear();
            TPLb.Clear();
            TPLc.Clear();
            TPLd.Clear();
            TPLe.Clear();
            TPLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void TradeTimerNational()
        {
            if (j <= 13)
            {
                TPLa.Add(1.75, tradelist1[j]);
            }

            if (j <= 7)
            {
                TPLb.Add(2.45, tradelist2[j]);

            }

            if (j <= 10)
            {
                TPLc.Add(3.15, tradelist3[j]);
            }
            if (j <= 10)
            {
                TPLd.Add(3.85, tradelist4[j]);
            }

            if (j <= 7)
            {
                TPLe.Add(4.55, tradelist5[j]);

            }

            if (j <= 9)
            {
                TPLf.Add(5.25, tradelist6[j]);
            }
            j++;
            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//立即显示  
        }

        private void TradeLocalChange()
        {
            TPLa.Clear();
            TPLb.Clear();
            TPLc.Clear();
            TPLd.Clear();
            TPLe.Clear();
            TPLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void TradeTimerLocal()
        {
            if (j <= 11)
            {
                TPLa.Add(1.75, tradelist1[j]);
            }

            if (j <= 10)
            {
                TPLb.Add(2.45, tradelist2[j]);

            }

            if (j <= 9)
            {
                TPLc.Add(3.15, tradelist3[j]);
            }
            if (j <= 7)
            {
                TPLd.Add(3.85, tradelist4[j]);
            }

            if (j <= 8)
            {
                TPLe.Add(4.55, tradelist5[j]);

            }

            if (j <= 9)
            {
                TPLf.Add(5.25, tradelist6[j]);
            }
            j++;
            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//立即显示  
        }

        private void TradeGoChange()
        {
            TPLa.Clear();
            TPLb.Clear();
            TPLc.Clear();
            TPLd.Clear();
            TPLe.Clear();
            TPLf.Clear();
            timer2.Enabled = true;
            timer2.Start();
            j = 0;
        }
        private void TradeTimerGo()
        {
            if (j <= 12)
            {
                TPLa.Add(1.75, tradelist1[j]);
            }

            if (j <= 10)
            {
                TPLb.Add(2.45, tradelist2[j]);

            }

            if (j <= 10)
            {
                TPLc.Add(3.15, tradelist3[j]);
            }
            if (j <= 10)
            {
                TPLd.Add(3.85, tradelist4[j]);
            }

            if (j <= 7)
            {
                TPLe.Add(4.55, tradelist5[j]);

            }

            if (j <= 9)
            {
                TPLf.Add(5.25, tradelist6[j]);
            }
            j++;
            zedGraphControl11.AxisChange();
            zedGraphControl11.Refresh();//立即显示  
        }
        #endregion

        #region 劳动力数量
        PointPairList LLa = new PointPairList();
        PointPairList LLb = new PointPairList();
        PointPairList LLc = new PointPairList();
        PointPairList LLd = new PointPairList();
        PointPairList LLe = new PointPairList();
        PointPairList LLf = new PointPairList();
        private void LabourGraph()
        {
            myPane = zedGraphControl12.GraphPane;
            myPane.Border.IsVisible = false;//去掉边框
            myPane.Title.Text = "";//标题
            myPane.XAxis.Title.Text = "年份";//横坐标
            myPane.YAxis.Title.Text = "万人";//纵坐标

            //X轴
            myPane.XAxis.Scale.Min = 0;        //X轴最小值0   
            myPane.XAxis.Scale.Max = 30;    //X轴最大30   

            myPane.XAxis.Scale.MinorStep = 5;//X轴小步长1,也就是小间隔   
            myPane.XAxis.Scale.MajorStep = 5;//X轴大步长为5，也就是显示文字的大间隔   


            //Y轴
            myPane.YAxis.Scale.Min = 100;     //Y轴最小值0
            myPane.YAxis.Scale.Max = 150;   //Y轴最大值3

            myPane.YAxis.Scale.MinorStep = 5;//Y轴小步长0.1,也就是小间隔   
            myPane.YAxis.Scale.MajorStep = 5;//Y轴大步长为0.3，也就是显示文字的大间隔   

            //改变标题字体两种方法
            //myPane.Title.FontSpec = new FontSpec("Arial", 20, Color.Red,false,false,false);//第一种
            //zedGraphControl1.GraphPane.Title.FontSpec.FontColor = Color.Green;//第二种
            //zedGraphControl1.GraphPane.Title.FontSpec.Size = 23f;


            //去掉标题边框
            //myPane.Title.FontSpec.Border.IsVisible=false;  //去掉标题边框
            myPane.Title.IsVisible = false;//标题隐藏

            //设置轴，Y轴标题大小和颜色
            //myPane.XAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);
            //myPane.YAxis.Title.FontSpec = new FontSpec("楷体", 20, Color.Black, false, false, false);

            //myPane.XAxis.Title.FontSpec.Border.IsVisible = false;  //去掉X轴标题边框
            //myPane.YAxis.Title.FontSpec.Border.IsVisible = false;  //去掉Y轴标题边框


            //设置X轴，Y轴坐标大小和颜色
            zedGraphControl12.GraphPane.XAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl12.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            zedGraphControl12.GraphPane.YAxis.Scale.FontSpec.Size = 15f;
            zedGraphControl12.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Black;


            // 设置图表的背景颜色
            myPane.Fill = new Fill(Color.LavenderBlush);
            //myPane.Fill = new Fill(Color.White, Color.PaleTurquoise);//渐变色
            //myPane.Chart.Fill.Type = FillType.None;  //如果设置为FillType.None,则图表内部颜色和背景色相同
            zedGraphControl12.GraphPane.Chart.Fill = new Fill(Color.White);  //设置图表内部颜色

            LineItem la = myPane.AddCurve("", LLa, Color.LawnGreen);
            LineItem lb = myPane.AddCurve("", LLb, Color.LightBlue);
            LineItem lc = myPane.AddCurve("", LLc, Color.Indigo);
            LineItem ld = myPane.AddCurve("", LLd, Color.DarkSalmon);
            LineItem le = myPane.AddCurve("", LLe, Color.DeepPink);
            LineItem lf = myPane.AddCurve("", LLf, Color.DarkOrange);
            //显示网格线
            zedGraphControl12.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl12.GraphPane.XAxis.MajorGrid.IsVisible = true;

            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();

            timer1.Enabled = false;
        }

        private void LabourGreenChange()
        {
            LLa.Clear();
            LLb.Clear();
            LLc.Clear();
            LLd.Clear();
            LLe.Clear();
            LLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void LabourTimerGreen()
        {
            if (i <= 30)
            {
                LLa.Add(i, -0.04*(i-23)*(i-23)+135);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbLabour.Text = (-0.04 * (i - 23) * (i - 23) + 135).ToString(("0.00 ")) + "万人";//保留两位小数
            }

        }

        private void LabourNatureChange()
        {
            LLa.Clear();
            LLb.Clear();
            LLc.Clear();
            LLd.Clear();
            LLe.Clear();
            LLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void LabourTimerNature()
        {
            if (i <= 30)
            {
                LLb.Add(i, -0.03 * (i - 25) * (i - 25) + 135);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbLabour.Text = (-0.03 * (i - 25) * (i - 25) + 135).ToString("0.00 ") + "万人";//保留两位小数
            }

        }

        private void LabourWorldChange()
        {
            LLa.Clear();
            LLb.Clear();
            LLc.Clear();
            LLd.Clear();
            LLe.Clear();
            LLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void LabourTimerWorld()
        {
            if (i <= 30)
            {
                LLc.Add(i, -0.09 * (i - 18) * (i - 18) + 140);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbLabour.Text = (0.3 * (i - 10) * (i - 10) + 120).ToString(("0.00 ")) + "万人";//保留两位小数
            }

        }

        private void LabourNationalChange()
        {
            LLa.Clear();
            LLb.Clear();
            LLc.Clear();
            LLd.Clear();
            LLe.Clear();
            LLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void LabourTimerNational()
        {
            if (i <= 30)
            {
                LLd.Add(i, -0.075 * (i - 25) * (i - 23) + 145);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbLabour.Text = (-0.075 * (i - 25) * (i - 23) + 145).ToString(("0.00 ")) + "万人";//保留两位小数
            }

        }

        private void LabourLocalChange()
        {
            LLa.Clear();
            LLb.Clear();
            LLc.Clear();
            LLd.Clear();
            LLe.Clear();
            LLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void LabourTimerLocal()
        {
            if (i <= 30)
            {
                LLe.Add(i, -0.02 * (i - 25) * (i - 25) + 125);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbLabour.Text = (-0.02 * (i - 25) * (i - 25) + 125).ToString(("0.00 ")) + "万人";//保留两位小数
            }

        }

        private void LabourGoChange()
        {
            LLa.Clear();
            LLb.Clear();
            LLc.Clear();
            LLd.Clear();
            LLe.Clear();
            LLf.Clear();
            timer1.Enabled = true;
            timer1.Start();
            i = 0;
        }
        private void LabourTimerGo()
        {
            if (i <= 30)
            {
                LLf.Add(i, -0.078 * (i - 20) * (i - 20) + 140);
            }
            else
            {
                CheckBoxAble();
            }

            i = i + 0.001;
            zedGraphControl12.AxisChange();
            zedGraphControl12.Refresh();//立即显示 
            if (i <= 30)
            {
                //lbYear.Text = "20"+((int)i).ToString();
                lbLabour.Text = (-0.078 * (i - 20) * (i - 20) + 140).ToString("0.00 ") + "万人";//保留两位小数
            }

        }
        #endregion









        #endregion












    }
}


