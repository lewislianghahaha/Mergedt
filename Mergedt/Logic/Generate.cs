using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MergeDt.DB;

namespace Mergedt.Logic
{
    public class Generate
    {
        DtList dtList=new DtList();

        /// <summary>
        /// 运算-获取要生成的表头信息(MEASUREMENT_COLOR)
        /// </summary>
        /// <param name="dt">从模板EXCEL处获取的DT</param>
        /// <returns></returns>
        public DataTable Generatetemp(DataTable dt)
        {
            var resultdt = new DataTable();

            try
            {
                //获取对应临时表(表头)
                resultdt = dtList.Get_MeasureMentColordt();
                //循环从模板EXCEL获取的DT
                foreach (DataRow row in dt.Rows)
                {
                    var newrow = resultdt.NewRow();
                    newrow[0] = row[0];             //BMMEASUREMENTID(主键)
                    newrow[1] = row[1];             //COLORCODE(内部色号)
                    newrow[2] = row[112];           //FORMULAVERSIONDATE(版本日期)
                    newrow[3] = row[2];             //DIFFUSECOARSENESS(颗粒度)
                    newrow[4] = DateTime.Now.Date;  //CREATEDDATE(创建日期)
                    newrow[5] = row[111];           //MEASUREMENTTIME(测量时间)
                    resultdt.Rows.Add(newrow);
                }
            }
            catch (Exception)
            {
                resultdt.Rows.Clear();
                resultdt.Columns.Clear();
            }
            return resultdt;
        }

        /// <summary>
        /// 运算-获取要生成的表体信息(MEASUREMENT)
        /// </summary>
        /// <param name="dt">从模板EXCEL处获取的DT</param>
        /// <param name="tempdt">获取已运算成功的表头信息</param>
        /// <returns></returns>
        public DataTable GeneratetempEmpty(DataTable dt,DataTable tempdt)
        {
            var resultdt = new DataTable();
            var bmaid = 1; //记录表体的BMAID主键信息

            try
            {
                //获取对应临时表(表体)
                resultdt = dtList.Get_MeasureMentEmptydt();
                //循环获取已运算成功的表头信息
                foreach (DataRow row in tempdt.Rows)
                {
                    //根据表头的ID信息查询从EXCEL模板得出的DT内的相关内容
                    var rows = dt.Select("ID='" + Convert.ToInt32(row[0]) + "'");
                    //当查询的结果有值的话,就按相关角度(15 45 110)进行对临时表赋值
                    if (rows.Length > 0)
                    {
                        //循环将记录插入至resultdt临时表内(注:每循环一次角度 主键 及相关值都要转换)
                        for (var i = 0; i < 3; i++)
                        {
                            var newrow = resultdt.NewRow();
                            switch (i)
                            {
                                //角度:15
                                case 0:
                                    newrow = GetRecordToRows(newrow, bmaid,15,
                                                            Convert.ToInt32(row[0]),
                                                            Convert.ToDecimal(rows[0][3]),   //L
                                                            Convert.ToDecimal(rows[0][4]),   //A
                                                            Convert.ToDecimal(rows[0][5]),   //B
                                                            Convert.ToDecimal(rows[0][6]),   //C
                                                            Convert.ToDecimal(rows[0][7]),   //H

                                                            Convert.ToDecimal(rows[0][8]),Convert.ToDecimal(rows[0][9]),Convert.ToDecimal(rows[0][10]),Convert.ToDecimal(rows[0][11]),
                                                            Convert.ToDecimal(rows[0][12]),Convert.ToDecimal(rows[0][13]),Convert.ToDecimal(rows[0][14]),Convert.ToDecimal(rows[0][15]),
                                                            Convert.ToDecimal(rows[0][16]),Convert.ToDecimal(rows[0][17]),Convert.ToDecimal(rows[0][18]),Convert.ToDecimal(rows[0][19]),
                                                            Convert.ToDecimal(rows[0][20]),Convert.ToDecimal(rows[0][21]),Convert.ToDecimal(rows[0][22]),Convert.ToDecimal(rows[0][23]),
                                                            Convert.ToDecimal(rows[0][24]),Convert.ToDecimal(rows[0][25]),Convert.ToDecimal(rows[0][26]),Convert.ToDecimal(rows[0][27]),
                                                            Convert.ToDecimal(rows[0][28]),Convert.ToDecimal(rows[0][29]),Convert.ToDecimal(rows[0][30]),Convert.ToDecimal(rows[0][31]),
                                                            Convert.ToDecimal(rows[0][32]),Convert.ToDecimal(rows[0][33]),Convert.ToDecimal(rows[0][34]),Convert.ToDecimal(rows[0][35]),
                                                            Convert.ToDecimal(rows[0][36]),Convert.ToDecimal(rows[0][37]),Convert.ToDecimal(rows[0][38])
                                                            );
                                    break;
                                //角度:45
                                case 1:
                                    newrow = GetRecordToRows(newrow, bmaid, 45, 
                                                            Convert.ToInt32(row[0]),
                                                            Convert.ToDecimal(rows[0][39]),   //L
                                                            Convert.ToDecimal(rows[0][40]),   //A
                                                            Convert.ToDecimal(rows[0][41]),   //B
                                                            Convert.ToDecimal(rows[0][42]),   //C
                                                            Convert.ToDecimal(rows[0][43]),   //H

                                                            Convert.ToDecimal(rows[0][44]), Convert.ToDecimal(rows[0][45]), Convert.ToDecimal(rows[0][46]), Convert.ToDecimal(rows[0][47]),
                                                            Convert.ToDecimal(rows[0][48]), Convert.ToDecimal(rows[0][49]), Convert.ToDecimal(rows[0][50]), Convert.ToDecimal(rows[0][51]),
                                                            Convert.ToDecimal(rows[0][52]), Convert.ToDecimal(rows[0][53]), Convert.ToDecimal(rows[0][54]), Convert.ToDecimal(rows[0][55]),
                                                            Convert.ToDecimal(rows[0][56]), Convert.ToDecimal(rows[0][57]), Convert.ToDecimal(rows[0][58]), Convert.ToDecimal(rows[0][59]),
                                                            Convert.ToDecimal(rows[0][60]), Convert.ToDecimal(rows[0][61]), Convert.ToDecimal(rows[0][62]), Convert.ToDecimal(rows[0][63]),
                                                            Convert.ToDecimal(rows[0][64]), Convert.ToDecimal(rows[0][65]), Convert.ToDecimal(rows[0][66]), Convert.ToDecimal(rows[0][67]),
                                                            Convert.ToDecimal(rows[0][68]), Convert.ToDecimal(rows[0][69]), Convert.ToDecimal(rows[0][70]), Convert.ToDecimal(rows[0][71]),
                                                            Convert.ToDecimal(rows[0][72]), Convert.ToDecimal(rows[0][73]), Convert.ToDecimal(rows[0][74])
                                                              );
                                    break;
                                //角度:110
                                default:
                                    newrow = GetRecordToRows(newrow, bmaid, 110,
                                                            Convert.ToInt32(row[0]),
                                                            Convert.ToDecimal(rows[0][75]),   //L
                                                            Convert.ToDecimal(rows[0][76]),   //A
                                                            Convert.ToDecimal(rows[0][77]),   //B
                                                            Convert.ToDecimal(rows[0][78]),   //C
                                                            Convert.ToDecimal(rows[0][79]),   //H

                                                            Convert.ToDecimal(rows[0][80]), Convert.ToDecimal(rows[0][81]), Convert.ToDecimal(rows[0][82]), Convert.ToDecimal(rows[0][83]),
                                                            Convert.ToDecimal(rows[0][84]), Convert.ToDecimal(rows[0][85]), Convert.ToDecimal(rows[0][86]), Convert.ToDecimal(rows[0][87]),
                                                            Convert.ToDecimal(rows[0][88]), Convert.ToDecimal(rows[0][89]), Convert.ToDecimal(rows[0][90]), Convert.ToDecimal(rows[0][91]),
                                                            Convert.ToDecimal(rows[0][92]), Convert.ToDecimal(rows[0][93]), Convert.ToDecimal(rows[0][94]), Convert.ToDecimal(rows[0][95]),
                                                            Convert.ToDecimal(rows[0][96]), Convert.ToDecimal(rows[0][97]), Convert.ToDecimal(rows[0][98]), Convert.ToDecimal(rows[0][99]),
                                                            Convert.ToDecimal(rows[0][100]), Convert.ToDecimal(rows[0][101]), Convert.ToDecimal(rows[0][102]), Convert.ToDecimal(rows[0][103]),
                                                            Convert.ToDecimal(rows[0][104]), Convert.ToDecimal(rows[0][105]), Convert.ToDecimal(rows[0][106]), Convert.ToDecimal(rows[0][107]),
                                                            Convert.ToDecimal(rows[0][108]), Convert.ToDecimal(rows[0][109]), Convert.ToDecimal(rows[0][110])
                                                              );
                                    break;
                            }
                            bmaid++;
                            resultdt.Rows.Add(newrow);
                        }
                        
                    }
                }
            }
            catch (Exception)
            {
                resultdt.Rows.Clear();
                resultdt.Columns.Clear();
            }
            return resultdt;
        }

        /// <summary>
        /// 将相关内容插入至DataRow内
        /// </summary>
        /// <returns></returns>
        private DataRow GetRecordToRows(DataRow newrow,int bmaid,int angle,
                                        int bmmeasurementid,decimal L,decimal A,decimal B,decimal C,decimal H,
                                        decimal BMA_R400,decimal BMA_R410,decimal BMA_R420,decimal BMA_R430,decimal BMA_R440,
                                        decimal BMA_R450,decimal BMA_R460,decimal BMA_R470,decimal BMA_R480,decimal BMA_R490,
                                        decimal BMA_R500,decimal BMA_R510,decimal BMA_R520,decimal BMA_R530,decimal BMA_R540,
                                        decimal BMA_R550,decimal BMA_R560,decimal BMA_R570,decimal BMA_R580,decimal BMA_R590,
                                        decimal BMA_R600,decimal BMA_R610,decimal BMA_R620,decimal BMA_R630,decimal BMA_R640,
                                        decimal BMA_R650,decimal BMA_R660,decimal BMA_R670,decimal BMA_R680,decimal BMA_R690,
                                        decimal BMA_R700)
        {
            newrow[0] = bmaid;                          //BMAID
            newrow[1] = angle;                          //ANGLE
            newrow[2] = bmmeasurementid;                //BMMEASUREMENTID
            newrow[3] = L;                              //L
            newrow[4] = A;                              //A
            newrow[5] = B;                              //B
            newrow[6] = C;                              //C
            newrow[7] = H;                              //H
            newrow[8] = BMA_R400;                       //BMA_R400
            newrow[9] = BMA_R410;                       //BMA_R410
            newrow[10] = BMA_R420;                      //BMA_R420
            newrow[11] = BMA_R430;                      //BMA_R430
            newrow[12] = BMA_R440;                      //BMA_R440
            newrow[13] = BMA_R450;                      //BMA_R450
            newrow[14] = BMA_R460;                      //BMA_R460
            newrow[15] = BMA_R470;                      //BMA_R470
            newrow[16] = BMA_R480;                      //BMA_R480
            newrow[17] = BMA_R490;                      //BMA_R490
            newrow[18] = BMA_R500;                      //BMA_R500
            newrow[19] = BMA_R510;                      //BMA_R510
            newrow[20] = BMA_R520;                      //BMA_R520
            newrow[21] = BMA_R530;                      //BMA_R530
            newrow[22] = BMA_R540;                      //BMA_R540
            newrow[23] = BMA_R550;                      //BMA_R550
            newrow[24] = BMA_R560;                      //BMA_R560
            newrow[25] = BMA_R570;                      //BMA_R570
            newrow[26] = BMA_R580;                      //BMA_R580
            newrow[27] = BMA_R590;                      //BMA_R590
            newrow[28] = BMA_R600;                      //BMA_R600
            newrow[29] = BMA_R610;                      //BMA_R610
            newrow[30] = BMA_R620;                      //BMA_R620
            newrow[31] = BMA_R630;                      //BMA_R630
            newrow[32] = BMA_R640;                      //BMA_R640
            newrow[33] = BMA_R650;                      //BMA_R650
            newrow[34] = BMA_R660;                      //BMA_R660
            newrow[35] = BMA_R670;                      //BMA_R670
            newrow[36] = BMA_R680;                      //BMA_R680
            newrow[37] = BMA_R690;                      //BMA_R690
            newrow[38] = BMA_R700;                      //BMA_R700
            return newrow;
        }
    }
}
