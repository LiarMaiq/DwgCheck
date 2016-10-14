using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace ClassLibCAD
{
    public class Class1
    {
        //判断一个点和一条多段线位置关系
        static int IsIntersect(Point3d p, Polyline pl)
        {
            int count = 0;
            for (int i = 0; i < pl.NumberOfVertices; i++)
            {
                Point2d s = pl.GetPoint2dAt(i);
                Point2d e = pl.GetPoint2dAt((i + 1) % pl.NumberOfVertices);
                if (s.Y == e.Y)//直线平行于X轴
                {
                    continue;
                }
                if (p.Y < (s.Y < e.Y ? s.Y : e.Y))//p点Y坐标小于直线最低点Y坐标
                {
                    continue;
                }
                if (p.Y > (s.Y > e.Y ? s.Y : e.Y))//p点Y坐标大于直线最高点Y坐标
                {
                    continue;
                }
                //点在线上，包括点在直线的端点
                //判断double类型值相等应该考虑容差的问题，直接判定绝对相等可能会出现误判
                //X1000表示判定到小数点后三位
                if ((long)((e.Y - p.Y) * (e.X - s.X) * 1000) == (long)((e.Y - s.Y) * (e.X - p.X) * 1000))
                {
                    return 0;
                }
                double X = e.X - ((e.X - s.X) * (e.Y - p.Y)) / (e.Y - s.Y);
                if (X > p.X)
                {
                    count++;
                    //当多边形的顶点位于直线Y=p.X上时，判定为1个交点
                    if (p.Y == e.Y)
                    {
                        count--;
                    }
                    //排除仅有一个顶点位于直线Y=p.X上的多边形
                    if (p.Y == e.Y && (s.Y - p.Y) * (pl.GetPoint2dAt((i + 2) % pl.NumberOfVertices).Y - p.Y) > 0)
                    {
                        count--;
                    }
                }
            }
            if (count % 2 == 1)
            {
                return 1;
            }
            return 0;
        }

        // 两点距离
        static double pointDistance(Point2d p1, Point2d p2)
        {
            double distance = 0;
            distance = Math.Sqrt((p1.Y - p2.Y) * (p1.Y - p2.Y) + (p1.X - p2.X) * (p1.X - p2.X));
            return distance;
        }

        // 三点计算面积
        static double area(Point2d p1, Point2d p2, Point2d p3)
        {
            double area = 0;
            double a = 0, b = 0, c = 0, s = 0;
            a = pointDistance(p1, p2);
            b = pointDistance(p2, p3);
            c = pointDistance(p1, p3);
            s = 0.5 * (a + b + c);
            area = Math.Sqrt(s * (s - a) * (s - b) * (s - c));
            return area;
        }

        //
        // 点到线段最短距离
        static double pointToLine(Point2d p1, Point2d p2, Point2d p)
        {
            double ans = 0;
            double a, b, c;
            a = pointDistance(p1, p2);
            b = pointDistance(p1, p);
            c = pointDistance(p2, p);
            if (c + b == a)
            {//点在线段上
                ans = 0;
                return ans;
            }
            if (a <= 0.00001)
            {//不是线段，是一个点
                ans = b;
                return ans;
            }
            if (c * c >= a * a + b * b)
            { //组成直角三角形或钝角三角形，p1为直角或钝角
                ans = b;
                return ans;
            }
            if (b * b >= a * a + c * c)
            {// 组成直角三角形或钝角三角形，p2为直角或钝角
                ans = c;
                return ans;
            }
            // 组成锐角三角形，则求三角形的高
            double p0 = (a + b + c) / 2;// 半周长
            double s = Math.Sqrt(p0 * (p0 - a) * (p0 - b) * (p0 - c));// 海伦公式求面积
            ans = 2 * s / a;// 返回点到线的距离（利用三角形面积公式求高）
            return ans;
        }



        [CommandMethod("check")]
        public static void Chack()
        {
            Form1 f = new Form1();
            f.openFileDialog1.ShowDialog();
            foreach (string item in f.openFileDialog1.FileNames)
            {
                string fileName = item;
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                //新建一个数据库对象以读取Dwg文件
                Database db = new Database(false, true);

                //如果指定文件名的文件存在
                if (System.IO.File.Exists(fileName))
                {
                    //把文件读入到数据库中
                    db.ReadDwgFile(fileName, System.IO.FileShare.ReadWrite, true, null);
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        //获取数据库的图层表对象
                        LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                        //获取数据库的块对象
                        BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        //循环遍历每个图层
                        bool existJZD = false;
                        bool existBH = false;

                        //
                        // 复合命名规则图层
                        List<string> validLayers = new List<string>();


                        foreach (ObjectId layerId in lt)
                        {
                            LayerTableRecord ltr = (LayerTableRecord)trans.GetObject(layerId, OpenMode.ForWrite);
                            ed.WriteMessage("\n检查图层:" + ltr.Name);
                            if (ltr != null)
                            {
                                if (ltr.Name == "JZD")
                                {
                                    existJZD = true;
                                }
                                if (ltr.Name == "编号")
                                {
                                    existBH = true;
                                }

                                //ed.WriteMessage(ltr.Name.PadLeft(5));
                                if (ltr.Name.Length >= 5 && ltr.Name.Substring(0, 5) == "CBDK_")
                                {
                                    validLayers.Add(ltr.Name);
                                }
                            }
                            else
                            {
                                //using (System.IO.StreamWriter file = new System.IO.StreamWriter(item + ".txt", true))
                                //{
                                //    file.WriteLine("此文件为空");
                                //}
                                //return;
                            }
                        }



                        ////判断是否存在必要图层
                        //if (existJZD == false)
                        //{
                        //    using (System.IO.StreamWriter file = new System.IO.StreamWriter(item + ".txt", true))
                        //    {
                        //        file.WriteLine("此文件不存在[JZD]图层");
                        //        return;//JZD图层不存在，函数直接退出
                        //    }
                        //}
                        //if (existBH == false)
                        //{
                        //    using (System.IO.StreamWriter file = new System.IO.StreamWriter(item + ".txt", true))
                        //    {
                        //        file.WriteLine("此文件不存在[编号]图层");
                        //        return;//编号图层不存在，函数直接退出
                        //    }
                        //}



                        //循环遍历所有块记录
                        List<int> DKBM = new List<int>();//用于保存地块编码p
                        //List<Point3d> DKBM_Position = new List<Point3d>();//用于保存DKBM的位置
                        List<Polyline> DK = new List<Polyline>();
                        List<DBText> BM = new List<DBText>();
                        ed.WriteMessage("\n检查文件块表...");
                        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        //Entity ent = (Entity)trans.GetObject(objId, OpenMode.ForRead);
                        //              using (System.IO.StreamWriter file = new System.IO.StreamWriter(item + ".txt", true))
                        {
                            //file.WriteLine(btr.Name);
                            foreach (ObjectId obj in btr)
                            {
                                Entity ent = trans.GetObject(obj, OpenMode.ForWrite) as Entity;
                                //file.WriteLine(ent);//测试用
                                
                                if (ent == null)
                                {
                                    continue;
                                }
                                
                                //else if (ent.Layer == "DK")//检查JZD图层相关信息
                                else if (validLayers.Contains(ent.Layer))
                                {
                                    DBText dbt = ent as DBText;
                                    DBPoint dbp = ent as DBPoint;
                                    Line l = ent as Line;
                                    Polyline pl = ent as Polyline;
                                    Circle c = ent as Circle;

                                    if (dbt != null)
                                    {
                                        dbt.ColorIndex = 50;
                                        dbt.LineWeight = LineWeight.LineWeight070;
                                        //                                      file.WriteLine("[JZD] 图层中存在内容为 [" + dbt.TextString + "] 的文本实体【黄色】");
                                    }
                                    if (dbp != null)
                                    {
                                        dbp.ColorIndex = 50;
                                        //                                       file.WriteLine("[JZD] 图层中存在位置为 [" + dbp.Position.X + "," + dbp.Position.Y + "] 的点实体【黄色】");
                                    }
                                    if (l != null)
                                    {
                                        l.ColorIndex = 50;
                                        l.LineWeight = LineWeight.LineWeight070;
                                        //                                      file.WriteLine("[JZD] 图层中存在起点为 [" + l.StartPoint.X + "," + l.StartPoint.Y + "] 的直线实体【黄色】");
                                    }
                                    if (pl != null)
                                    {
                                        if (!pl.Closed)
                                        {
                                            pl.ColorIndex = 5;//未闭合多段线用蓝色标记
                                            pl.LineWeight = LineWeight.LineWeight070;
                                            //                                           file.WriteLine("[JZD] 图层中存在非闭合多段线实体【蓝色】");
                                        }
                                        else
                                        {
                                            DK.Add(pl);//将闭合的多段线添加进入DK链表
                                        }
                                    }
                                    if (c != null)
                                    {
                                        c.ColorIndex = 50;
                                        c.LineWeight = LineWeight.LineWeight070;
                                        //                                       file.WriteLine("[JZD] 图层中存在圆心为 [" + c.Center.X + "," + c.Center.Y + "] 的圆实体【黄色】");
                                    }
                                    if (dbt == null && dbp == null && l == null && pl == null && c == null)
                                    {
                                        ent.ColorIndex = 50;
                                        ent.LineWeight = LineWeight.LineWeight070;
                                        //                                       file.WriteLine("[JZD] 图层中存在非多段线实体【黄色】");
                                    }
                                }



                                //                               else if (ent.Layer == "编号")
                                //                               {
                                //                                   DBText dbt = ent as DBText;
                                //                                   DBPoint dbp = ent as DBPoint;
                                //                                   Line l = ent as Line;
                                //                                   Polyline pl = ent as Polyline;
                                //                                   Circle c = ent as Circle;

                                //                                   if (dbt != null)
                                //                                   {
                                //                                       int dkbm = 0;
                                //                                       //首先判断长度是否为5
                                //                                       if (dbt.TextString.Length != 5)
                                //                                       {
                                ////                                           file.WriteLine("[编号] 图层中的编号 [" + dbt.TextString + "] 不符合规范");
                                //                                           dbt.LineWeight = LineWeight.LineWeight070;
                                //                                           dbt.ColorIndex = 130;
                                //                                       }
                                //                                       //如果长度为5，判断是否能够转换为int
                                //                                       else if (!(int.TryParse(dbt.TextString, out dkbm)))
                                //                                       {
                                ////                                           file.WriteLine("[编号] 图层中的编号 [" + dbt.TextString + "] 不符合规范");
                                //                                           dbt.LineWeight = LineWeight.LineWeight070;
                                //                                           dbt.ColorIndex = 130;
                                //                                       }
                                //                                       //长度为5，并且可以转换为int，则执行下面语句
                                //                                       else
                                //                                       {
                                //                                           DKBM.Add(dkbm);
                                //                                           BM.Add(dbt);
                                //                                           //将符合条件的地块编码添加进链表
                                //                                       }
                                //                                   }

                                //                                   if (dbp != null)
                                //                                   {
                                //                                       dbp.ColorIndex = 80;
                                ////                                       file.WriteLine("[编号] 图层中存在位置为 [" + dbp.Position.X + "," + dbp.Position.Y + "] 的点实体【绿色】");
                                //                                   }
                                //                                   if (l != null)
                                //                                   {
                                //                                       l.ColorIndex = 80;
                                //                                       l.LineWeight = LineWeight.LineWeight070;
                                // //                                      file.WriteLine("[编号] 图层中存在起点为 [" + l.StartPoint.X + "," + l.StartPoint.Y + "] 的直线实体【绿色】");
                                //                                   }
                                //                                   if (pl != null)
                                //                                   {
                                //                                       pl.ColorIndex = 80;
                                //                                       pl.LineWeight = LineWeight.LineWeight070;
                                ////                                       file.WriteLine("[编号] 图层中存在起点为 [" + pl.StartPoint.X + "," + pl.StartPoint.Y + "] 的多段线实体【绿色】");
                                //                                   }
                                //                                   if (c != null)
                                //                                   {
                                //                                       c.ColorIndex = 80;
                                //                                       c.LineWeight = LineWeight.LineWeight070;
                                ////                                       file.WriteLine("[编号] 图层中存在圆心为 [" + c.Center.X + "," + c.Center.Y + "] 的圆实体【绿色】");
                                //                                   }
                                //                                   if (dbt == null && dbp == null && l == null && pl == null && c == null)
                                //                                   {
                                //                                       ent.ColorIndex = 80;
                                //                                       ent.LineWeight = LineWeight.LineWeight070;
                                ////                                       file.WriteLine("[编号] 图层中存在非文字实体【绿色】");
                                //                                   }
                                //                               }



                            }
                        }
                        DKBM.Sort();//排序地块编码
                        //using (System.IO.StreamWriter file = new System.IO.StreamWriter(item + ".txt", true))
                        {
                            //                            ed.WriteMessage("\n检查地块编码...");
                            //                            //检查地块编码
                            //                            for (int i = 0; i < DKBM.Count - 1; i++)
                            //                            {
                            //                                if (i == 0)
                            //                                {
                            //                                    if (DKBM[i] - 1 == 1)
                            //                                    {
                            // //                                       file.WriteLine("地块编码 [00001] 缺漏。");
                            //                                    }
                            //                                    else if (DKBM[i] - 1 > 1)
                            //                                    {
                            // //                                       file.WriteLine("地块编码 [00001] 至 [" + (DKBM[i] - 1).ToString().PadLeft(5, '0') + "] 缺漏。");
                            //                                    }
                            //                                }

                            //                                int temp = DKBM[i + 1] - DKBM[i];
                            //                                if (temp > 2)
                            //                                {
                            // //                                   file.WriteLine("地块编码 [" + (DKBM[i] + 1).ToString().PadLeft(5, '0') + "] 至 [" +
                            // //                                       (DKBM[i + 1] - 1).ToString().PadLeft(5, '0') + "] 缺漏。");
                            //                                }
                            //                                else if (temp == 0)
                            //                                {
                            // //                                   file.WriteLine("地块编码 [" + DKBM[i].ToString().PadLeft(5, '0') + "] 存在重号。");
                            //                                }
                            //                                else if (temp == 2)
                            //                                {
                            // //                                   file.WriteLine("地块编码 [" + (DKBM[i] + 1).ToString().PadLeft(5, '0') + "] 缺漏。");
                            //                                }
                            //                            }
                            //                            ed.WriteMessage("\n检查多段线与地块编码对应关系...");
                            //                            //拓扑检查-是否全部实体均存在地块编码与之对应
                            //                            foreach (Polyline pl in DK)
                            //                            {
                            //                                //ed.WriteMessage("\n"+pl.ObjectId);
                            //                                int BMcount = 0;
                            //                                foreach (DBText p in BM)
                            //                                {
                            //                                    if (IsIntersect(p.Position, pl)==1)
                            //                                    {
                            //                                        BMcount++;
                            //                                    }
                            //                                }
                            //                                if (BMcount == 0)//多段线没有编号
                            //                                {
                            //                                    pl.ColorIndex = 210;
                            //                                    pl.LineWeight = LineWeight.LineWeight070;
                            ////                                    file.WriteLine("[JZD] 图层中存在未编号多段线实体【洋红】");
                            //                                }
                            //                                else if (BMcount >= 2)//多段线存在多个编号
                            //                                {
                            //                                    pl.ColorIndex = 210;
                            //                                    pl.LineWeight = LineWeight.LineWeight070;
                            // //                                   file.WriteLine("[JZD] 图层中存在多个编号的多段线实体【洋红】");
                            //                                }
                            //                                BMcount = 0;//重现初始化为0
                            //                            }
                            //                            ed.WriteMessage("\n检查地块编码与多段线对应关系...");
                            //                            //拓扑检查-是否全部地块编码均存在多段线与之对应
                            //                            foreach (DBText p in BM)
                            //                            {
                            //                                //ed.WriteMessage("\n" + p.ObjectId);
                            //                                int BMcount = 0;
                            //                                foreach (Polyline pl in DK)
                            //                                {
                            //                                    if (IsIntersect(p.Position,pl)==1)
                            //                                    {
                            //                                        BMcount++;
                            //                                    }
                            //                                }
                            //                                if (BMcount == 0)//编号没有多段线
                            //                                {
                            //                                    p.ColorIndex = 7;
                            //                                    p.LineWeight = LineWeight.LineWeight070;
                            ////                                    file.WriteLine("[编号] 图层中存在地块编码没有多段线与之对应【白色】");
                            //                                }
                            //                                else if (BMcount >= 2)//编号存在多个多段线与之对应
                            //                                {
                            //                                    p.ColorIndex = 7;
                            //                                    p.LineWeight = LineWeight.LineWeight070;
                            ////                                    file.WriteLine("[编号] 图层中存在多个多段线实体对应的地块编码【白色】");
                            //                                }
                            //                                BMcount = 0;//重现初始化为0
                            //                            }
                            ed.WriteMessage("\n拓扑检查...");
                            //拓扑检查
                            //通过比较一个多段线的每个顶点是否均能在另外一个多段线的顶点集合中找到，如果均能找到则说明另个多段线
                            //重复，采用此种方式的原因在于，CAD中的多段线是存在起点终点和方向的。
                            foreach (Polyline pl in DK)
                            {
                                //重复节点检查
                                for (int i = 0; i < pl.NumberOfVertices; i++)
                                {
                                    int j = i + 1;
                                    for (; j < pl.NumberOfVertices; j++)
                                    {
                                        if (pl.GetPoint2dAt(i).X == pl.GetPoint2dAt(j).X &&
                                            pl.GetPoint2dAt(i).Y == pl.GetPoint2dAt(j).Y)
                                        {
                                            pl.ColorIndex = 190;
                                            pl.LineWeight = LineWeight.LineWeight070;
                                            //                                            file.WriteLine("[JZD] 图层中多段线存在重复节点【紫色】");
                                            break;
                                        }
                                    }
                                    if (j != pl.NumberOfVertices)
                                    {
                                        break;
                                    }
                                }
                                //检查重复多段线
                                foreach (Polyline subpl in DK)
                                {
                                    //
                                    // 检查一条多段线的节点是否是另一条多段线节点的子集
                                    // 也可认为此种情况为压盖
                                    if (pl != subpl && pl.NumberOfVertices > subpl.NumberOfVertices)
                                    {
                                        int i = 0;
                                        for (; i < subpl.NumberOfVertices; i++)
                                        {
                                            int j = 0;
                                            for (; j < pl.NumberOfVertices; j++)
                                            {
                                                if (pl.GetPoint2dAt(j).X == subpl.GetPoint2dAt(i).X && pl.GetPoint2dAt(j).Y == subpl.GetPoint2dAt(i).Y)
                                                {
                                                    break;
                                                }
                                            }
                                            if (j == pl.NumberOfVertices)
                                            {
                                                break;
                                            }
                                        }
                                        if (i == subpl.NumberOfVertices)
                                        {
                                            subpl.ColorIndex = 10;
                                            subpl.LineWeight = LineWeight.LineWeight070;
                                            //                                            file.WriteLine("[JZD] 图层中存在重复多段线实体【红色】");
                                        }
                                    }


                                    //
                                    // 两条多段线完全重复的情况
                                    if (pl == subpl || pl.NumberOfVertices != subpl.NumberOfVertices)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        int i = 0;
                                        for (; i < pl.NumberOfVertices; i++)
                                        {
                                            int j = 0;
                                            for (; j < subpl.NumberOfVertices; j++)
                                            {
                                                if (pl.GetPoint2dAt(i).X == subpl.GetPoint2dAt(j).X && pl.GetPoint2dAt(i).Y == subpl.GetPoint2dAt(j).Y)
                                                {
                                                    break;
                                                }
                                            }
                                            if (j == subpl.NumberOfVertices)
                                            {
                                                break;
                                            }
                                        }
                                        if (i == pl.NumberOfVertices)
                                        {
                                            pl.ColorIndex = 10;
                                            pl.LineWeight = LineWeight.LineWeight070;
                                            //                                            file.WriteLine("[JZD] 图层中存在重复多段线实体【红色】");
                                        }
                                    }
                                }
                            }
                            //检查多段线压盖
                            foreach (Polyline pl in DK)
                            {
                                ed.WriteMessage("\n" + pl.ObjectId);
                                foreach (Polyline subpl in DK)
                                {
                                    if (pl != subpl)
                                    {
                                        for (int i = 0; i < pl.NumberOfVertices; i++)
                                        {
                                            if (IsIntersect(pl.GetPoint3dAt(i), subpl) == 1)
                                            {
                                                pl.ColorIndex = 30;
                                                subpl.ColorIndex = 30;
                                                pl.LineWeight = LineWeight.LineWeight070;
                                                subpl.LineWeight = LineWeight.LineWeight070;
                                                //                                                file.WriteLine("[JZD] 图层中存在相互压盖的多段线实体【橙色】");
                                                break;//只要发现一处存在压盖就可以停止循环
                                            }
                                        }
                                    }
                                }
                            }
                            //检查界址点5cm容差
                            List<Point2d> JZD = new List<Point2d>();//建立JZD链表用于存储界址点坐标
                            //遍历DK链表将所有界址点添加到JZD链表当中
                            foreach (Polyline dk in DK)
                            {
                                for (int i = 0; i < dk.NumberOfVertices; i++)
                                {
                                    //CAD中获取坐标时，设置容差为1MM，避免double类型数值精度过高而导致的程序误判
                                    //此处不能使用int类型，坐标值X*1000后可能超出了int类型所表示的范围
                                    Point2d p = new Point2d(((long)(dk.GetPoint2dAt(i).X * 10000)) / 10000.0,
                                        ((long)(dk.GetPoint2dAt(i).Y * 10000)) / 10000.0);
                                    //如果JZD中已经包含坐标相同的点，则不再添加进链表
                                    if (!JZD.Contains(p))
                                    {
                                        JZD.Add(p);
                                    }
                                }
                            }
                            ////遍历JZD链表，寻找距离小于5CM的界址点并做出标记
                            //foreach (Point2d jzd in JZD)
                            //{
                            //    foreach (Point2d subjzd in JZD)
                            //    {
                            //        //确保每个点不与自己本身做比较
                            //        if (jzd.X == subjzd.X && jzd.Y == subjzd.Y)
                            //        {
                            //            continue;
                            //        }
                            //        else
                            //        {
                            //            if (((jzd.X - subjzd.X) * (jzd.X - subjzd.X) + (jzd.Y - subjzd.Y) * (jzd.Y - subjzd.Y)) < 0.025)
                            //            {
                            //                Circle c = new Circle();
                            //                c.Center = new Point3d(jzd.X, jzd.Y, 0);
                            //                c.Radius = 0.8;
                            //                c.LineWeight = LineWeight.LineWeight070;
                            //                c.ColorIndex = 50;
                            //                btr.AppendEntity(c);
                            //                //这是一句很诡异的代码，弄不明白其作用是什么，但如果没有这一句，就会在保存和另存时候出现错误
                            //                trans.AddNewlyCreatedDBObject(c, true);
                            //                //将错误信息记录到文本当中
                            //                //                                            file.WriteLine("地块界址点间距离小于5CM【黄色圆圈】");
                            //                //只要出现画出一次圆，就可以结束本次循环
                            //                break;
                            //            }
                            //        }
                            //    }
                            //}

                            //
                            // 点到线段距离小于15cm检查
                            foreach (Point2d jzd in JZD)
                            {
                                foreach (Polyline dk in DK)
                                {
                                    for (int i = 0; i < dk.NumberOfVertices; i++)
                                    {
                                        Point2d s = new Point2d(((long)(dk.GetPoint2dAt(i).X * 10000)) / 10000.0,
                                        ((long)(dk.GetPoint2dAt(i).Y * 10000)) / 10000.0);
                                        Point2d e = new Point2d(((long)(dk.GetPoint2dAt((i + 1) % dk.NumberOfVertices).X * 10000)) / 10000.0,
                                        ((long)(dk.GetPoint2dAt((i + 1) % dk.NumberOfVertices).Y * 10000)) / 10000.0);

                                        //double len = PointToSegDist(jzd.X, jzd.Y, s.X, s.Y, e.X, e.Y);
                                        
                                        if (s != jzd && e != jzd)
                                        {
                                            double len = pointToLine(s, e, jzd);
                                            if (len < 0.15)
                                            {
                                                Line line = new Line();
                                                line.StartPoint = new Point3d(dk.GetPoint2dAt(i).X, dk.GetPoint2dAt(i).Y,0.0);
                                                line.EndPoint= new Point3d(dk.GetPoint2dAt((i + 1) % dk.NumberOfVertices).X, 
                                                    dk.GetPoint2dAt((i + 1) % dk.NumberOfVertices).Y, 0.0);

                                                line.LineWeight = LineWeight.LineWeight070;
                                                line.ColorIndex = 50;
                                                btr.AppendEntity(line);
                                                trans.AddNewlyCreatedDBObject(line, true);

                                                break;
                                            }
                                        }

                                    }
                                    
                                }
                            }


                        }
                        trans.Commit();
                    }
                    db.SaveAs(fileName.Insert(fileName.Length - 4, "_检查"), db.OriginalFileVersion);
                }
                ed.WriteMessage("\n" + item + " 完成");
            }
        }


        [CommandMethod("merge")]
        public static void Merge()
        {
            Form1 f = new Form1();
            f.openFileDialog1.ShowDialog();
        }
    }
}
