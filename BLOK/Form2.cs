using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace BLOK
{



#if AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
    using DbS = Autodesk.AutoCAD.DatabaseServices;
    using EdI = Autodesk.AutoCAD.EditorInput;
#elif NANOCAD
    using Teigha.DatabaseServices;
    using Teigha.Geometry;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
    using DbS = Teigha.DatabaseServices;
    using EdI = HostMgd.EditorInput;
#endif

    public partial class Form2 : Form
    {

        public struct PLOSpr
        {
            public string OSI;
            public double Xmin;
            public double Ymin;
            public double Xmax;
            public double Ymax;
            public Point3d PSK;
            public Point3d MSK;
            public void NomOSI(string i) { OSI = i; }
            public void NomXmin(double i) { Xmin = i; }
            public void NomYmin(double i) { Ymin = i; }
            public void NomXmax(double i) { Xmax = i; }
            public void NomYmax(double i) { Ymax = i; }
            public void NomPSK(Point3d i) { PSK = i; }
            public void NomMSK(Point3d i) { MSK = i; }
        };
        public string Sozd_Izm;
        //
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string HANDL = "";
            CreateLayer("Выноски");
            CreateLayer("ТРАССЫ");
            CreateLayer("ТРАССЫскрытые");
            CreateLayer("Плоскости");
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            Point3d PSKtl = new Point3d(0, 0, 0);
            Point3d MSKtl = new Point3d(0, 0, 0);
            string OSItl = "";
            if (Sozd_Izm == "Izm")
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed =
                        Application.DocumentManager.MdiActiveDocument.Editor;
                    try
                    {
                        // Просим пользователя выбрать примитив
                        PromptEntityResult ers = ed.GetEntity("Укажите примитив ");
                        // Открываем выбранный примитив
                        BlockReference bref = tr.GetObject(ers.ObjectId, OpenMode.ForRead) as BlockReference;
                        BP = bref.Position;
                        if (bref.IsDynamicBlock)
                        {
                            DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                                if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                                if (prop.PropertyName == "Положение4 X") { x2 = Convert.ToDouble(prop.Value.ToString()); }
                                if (prop.PropertyName == "Положение4 Y") { y2 = Convert.ToDouble(prop.Value.ToString()); }
                                if (prop.PropertyName == "Видимость1") { OSItl = prop.Value.ToString(); }
                            }
                        }
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "X") { xMir = Convert.ToDouble(atrRef.TextString); }
                                    if (atrRef.Tag == "Y") { yMir = Convert.ToDouble(atrRef.TextString); }
                                    if (atrRef.Tag == "Z") { zMir = Convert.ToDouble(atrRef.TextString); }
                                }
                            }
                        }
                        PSKtl = new Point3d(BP.X + x1, BP.Y + y1, 0);
                        MSKtl = new Point3d(xMir, yMir, zMir);
                        tr.Commit();
                    }
                    catch
                    {
                        tr.Abort();
                    }
                    this.textBox1.Text = Convert.ToInt64(MSKtl.X).ToString();
                    this.textBox2.Text = Convert.ToInt64(MSKtl.Y).ToString();
                    this.textBox3.Text = Convert.ToInt64(MSKtl.Z).ToString();
                    this.label6.Text = PSKtl.ToString();
                    this.label7.Text = HANDL;
                    this.button1.Visible = false;
                }
            }

        }//загрузка формы
        private void button4_Click(object sender, EventArgs e)
        {
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            Point3d PSKtl = new Point3d(0, 0, 0);
            Point3d MSKtl = new Point3d(0, 0, 0);
            string OSItl = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            this.Hide();
            using (DocumentLock docLock = doc.LockDocument())
            {
                Editor ed = doc.Editor;
                PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
                PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
                Transaction tr = db.TransactionManager.StartTransaction();
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                using (tr)
                {
                    //Entity bref = tr.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                    BlockReference bref = tr.GetObject(prEntResult.ObjectId, OpenMode.ForRead) as BlockReference;
                    BP = bref.Position;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение4 X") { x2 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение4 Y") { y2 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Видимость1") { OSItl = prop.Value.ToString(); }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "X") { xMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Y") { yMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Z") { zMir = Convert.ToDouble(atrRef.TextString); }
                            }
                        }
                    }
                    PSKtl = new Point3d(BP.X + x1, BP.Y + y1, 0);
                    MSKtl = new Point3d(xMir, yMir, zMir);
                    string OSiIsx = this.label7.Text;
                    double delX = Convert.ToDouble(this.textBox1.Text)- MSKtl.X;
                    double delY = Convert.ToDouble(this.textBox2.Text)- MSKtl.Y;
                    double delZ = Convert.ToDouble(this.textBox3.Text) - MSKtl.Z;
                    double Xnow = PSKtl.X;
                    double Ynow = PSKtl.Y;
                    double Znow = PSKtl.Z;


                    if (OSItl == "ZY" ) { Xnow = Xnow - delY; Ynow = Ynow + delZ; }
                    if (OSItl == "YZ" ) { Xnow = Xnow + delY; Ynow = Ynow + delZ; }
                    
                    if (OSItl == "XZ" ) { Xnow = Xnow + delX; Ynow = Ynow + delZ; }
                    if (OSItl == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; }
                   
                    if (OSItl == "XY" ) { Xnow = Xnow + delX; Ynow = Ynow + delY; }

                    Point3d IskTP = new Point3d(Xnow, Ynow, 0);
                    Point3d IskTMir = new Point3d(Convert.ToDouble(this.textBox1.Text), Convert.ToDouble(this.textBox2.Text), Convert.ToDouble(this.textBox3.Text));
                    BlockTable acBlkTbl;
                    BlockTableRecord acBlkTblRec;
                    acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = IskTP;
                    Poz.Height = 50;
                    Poz.ColorIndex = 3;
                    Poz.Layer = "Плоскости";
                    Poz.TextString = IskTMir.ToString();
                    acBlkTblRec.AppendEntity(Poz);
                    tr.AddNewlyCreatedDBObject(Poz, true);
                    tr.Commit();
                }
            }
            this.Show();
        }//найти точку на другом виде
        private void button5_Click(object sender, EventArgs e)
        {
#region переменные и фильтр
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            List<PLOSpr> SPPal = new List<PLOSpr>();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "POLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Плоскости"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет плоскостей...");
                this.Show();
                return;
            }
            SelectionSet acSSet = selRes.Value;
            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку для извлечения координат");
            this.Hide();
            using (DocumentLock docLock = doc.LockDocument())
            {
#endregion
#region список палуб
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    pPtRes = doc.Editor.GetPoint(pPtOpts);
                    Point3d Toch = pPtRes.Value;
                    foreach (SelectedObject sobj in acSSet)
                    {
                        double XminTL = 999999999999999999;
                        double YminTL = 999999999999999999;
                        double XmaxTL = -999999999999999999;
                        double YmaxTL = -999999999999999999;
                        string OSItl = "";
                        Point3d PSKtl = new Point3d(0, 0, 0);
                        Point3d MSKtl = new Point3d(0, 0, 0);
                        Polyline3d ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline3d;
                        if (ln != null)
                        {
                            Point3dCollection acPts3d = new Point3dCollection();
                            foreach (ObjectId acObjIdVert in ln)
                            {
                                PolylineVertex3d acPolVer3d;
                                acPolVer3d = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                                //acPts3d.Add(acPolVer3d.Position);
                                if (acPolVer3d.Position.X < XminTL) { XminTL = acPolVer3d.Position.X; }
                                if (acPolVer3d.Position.Y < YminTL) { YminTL = acPolVer3d.Position.Y; }
                                if (acPolVer3d.Position.X > XmaxTL) { XmaxTL = acPolVer3d.Position.X; }
                                if (acPolVer3d.Position.Y > YmaxTL) { YmaxTL = acPolVer3d.Position.Y; }
                            }
                            //Application.ShowAlertDialog("1");
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            int Schet;
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { OSItl = value.Value.ToString(); }
                                    if (Schet == 2) { MSKtl = (Point3d)value.Value; }
                                    if (Schet == 3) { PSKtl = (Point3d)value.Value; }
                                    Schet = Schet + 1;
                                }
                            }
                            PLOSpr Tpal = new PLOSpr();
                            Tpal.NomOSI(OSItl);
                            Tpal.NomXmax(XmaxTL);
                            Tpal.NomXmin(XminTL);
                            Tpal.NomYmax(YmaxTL);
                            Tpal.NomYmin(YminTL);
                            Tpal.NomMSK(MSKtl);
                            Tpal.NomPSK(PSKtl);
                            SPPal.Add(Tpal);
                        }// if (ln != null)
                    }
                    this.Show();
                    PLOSpr TPlos = new PLOSpr();
                    TPlos.NomOSI("мимо");
                    foreach (PLOSpr value in SPPal)
                    {
                        if (Toch.X > value.Xmin && Toch.X < value.Xmax && Toch.Y > value.Ymin && Toch.Y < value.Ymax) { TPlos = value; }
                        //Application.ShowAlertDialog("Плоскость-" + value.Xmin.ToString() + " " + value.Xmax.ToString() + " " + value.Ymin.ToString() + " " + value.Ymax.ToString() + " точка-" + Toch.ToString());
                    }
                    //Application.ShowAlertDialog(TPlos.OSI);
                    if (TPlos.OSI != "мимо")
                    {
                        double delX = Toch.X - TPlos.PSK.X;
                        double delY = Toch.Y - TPlos.PSK.Y;
                        double delZ = Toch.Z - TPlos.PSK.Z;
                        double Xnow = TPlos.MSK.X;
                        double Ynow = TPlos.MSK.Y;
                        double Znow = TPlos.MSK.Z;
                        if (TPlos.OSI == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ; }
                        if (TPlos.OSI == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                        if (TPlos.OSI == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                        if (TPlos.OSI == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                        if (TPlos.OSI == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                        Point3d T1 = new Point3d(Convert.ToDouble(this.textBox1.Text), Convert.ToDouble(this.textBox2.Text), Convert.ToDouble(this.textBox3.Text));
                        Point3d T2 = new Point3d(Xnow, Ynow, Znow);
                        this.label6.Text = T1.DistanceTo(T2).ToString();
                    }
                    else
                    { Application.ShowAlertDialog("мимо"); }
                    tr.Commit();
                }
            }
#endregion
        }//расстояние от точки до точки
        private void button2_Click(object sender, EventArgs e)
        {
#region переменные и фильтр
            this.Hide();
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            //Point3d PSKtl = new Point3d(0, 0, 0);
            //Point3d MSKtl = new Point3d(0, 0, 0);
            //string OSItl = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            List<PLOSpr> SPPal = new List<PLOSpr>();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Плоскости"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
                if (selRes.Status != PromptStatus.OK)
                {

                    ed.WriteMessage("Нет плоскостей...");
                    this.Show();
                    return;
                }
            SelectionSet acSSet = selRes.Value;
            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку для извлечения координат");
            //this.Hide();
            using (DocumentLock docLock = doc.LockDocument())
            {
#endregion
#region список палуб
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    pPtRes = doc.Editor.GetPoint(pPtOpts);
                    Point3d Toch = pPtRes.Value;
                    foreach (SelectedObject sobj in acSSet)
                    {
                        string OSItl = "";
                        Point3d PSKtl = new Point3d(0, 0, 0);
                        Point3d MSKtl = new Point3d(0, 0, 0);
                        //Polyline3d ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline3d;
                        BlockReference ln = tr.GetObject(sobj.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (ln != null)
                        {
                            BP = ln.Position;
                            if (ln.IsDynamicBlock)
                            {
                                DynamicBlockReferencePropertyCollection props = ln.DynamicBlockReferencePropertyCollection;
                                foreach (DynamicBlockReferenceProperty prop in props)
                                {
                                    object[] values = prop.GetAllowedValues();
                                    if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                                    if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                                    if (prop.PropertyName == "Положение4 X") { x2 = Convert.ToDouble(prop.Value.ToString()); }
                                    if (prop.PropertyName == "Положение4 Y") { y2 = Convert.ToDouble(prop.Value.ToString()); }
                                    if (prop.PropertyName == "Видимость1") { OSItl = prop.Value.ToString(); }
                                }
                            }
                            foreach (ObjectId idAtrRef in ln.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "X") { xMir = Convert.ToDouble(atrRef.TextString); }
                                        if (atrRef.Tag == "Y") { yMir = Convert.ToDouble(atrRef.TextString); }
                                        if (atrRef.Tag == "Z") { zMir = Convert.ToDouble(atrRef.TextString); }
                                    }
                                }
                            }
                            PSKtl = new Point3d(BP.X + x1, BP.Y + y1, 0);
                            MSKtl = new Point3d(xMir, yMir, zMir);
                            if (OSItl != "")
                            {
                                PLOSpr Tpal = new PLOSpr();
                                Tpal.NomOSI(OSItl);
                                Tpal.NomXmax(ln.GeometricExtents.MaxPoint.X);
                                Tpal.NomXmin(ln.GeometricExtents.MinPoint.X);
                                Tpal.NomYmax(ln.GeometricExtents.MaxPoint.Y);
                                Tpal.NomYmin(ln.GeometricExtents.MinPoint.Y);
                                Tpal.NomMSK(MSKtl);
                                Tpal.NomPSK(PSKtl);
                                SPPal.Add(Tpal);
                            }
                        }// if (ln != null)
                    }
                    this.Show();
                    PLOSpr TPlos = new PLOSpr();
                    TPlos.NomOSI("мимо");
                        foreach (PLOSpr value in SPPal)
                        {
                            if (Toch.X > value.Xmin && Toch.X < value.Xmax && Toch.Y > value.Ymin && Toch.Y < value.Ymax) { TPlos = value; }
                            //Application.ShowAlertDialog("Плоскость-" + value.Xmin.ToString() + " " + value.Xmax.ToString() + " " + value.Ymin.ToString() + " " + value.Ymax.ToString() + " точка-" + Toch.ToString());
                        }
                    //Application.ShowAlertDialog(TPlos.OSI);
                        if (TPlos.OSI != "мимо")
                        {
                            double delX = Toch.X - TPlos.PSK.X;
                            double delY = Toch.Y - TPlos.PSK.Y;
                            double delZ = Toch.Z - TPlos.PSK.Z;
                            double Xnow = TPlos.MSK.X;
                            double Ynow = TPlos.MSK.Y;
                            double Znow = TPlos.MSK.Z;
                            if (TPlos.OSI == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ; }
                            if (TPlos.OSI == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                            if (TPlos.OSI == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                            if (TPlos.OSI == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                            if (TPlos.OSI == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                            this.textBox1.Text = Convert.ToInt64(Xnow).ToString();
                            this.textBox2.Text = Convert.ToInt64(Ynow).ToString();
                            this.textBox3.Text = Convert.ToInt64(Znow).ToString();
                            //this.label8.Text = TPlos.OSI;
                        }
                        else
                        { Application.ShowAlertDialog("мимо"); }
                    tr.Commit();
                }
            }
#endregion
        }//точка из другой плоскости
        private void button1_Click(object sender, EventArgs e)
        {
            string Adr = this.textBox4.Text;
            string Name = Adr.Split('\\').Last().Split('.')[0];
            this.Hide();
            //Application.ShowAlertDialog(@Adr + " " + Name);
            if (System.IO.File.Exists(@Adr) == true)
                InsBlockRef(@Adr, Name, "", "Плоскость", "", "");
            else
                Application.ShowAlertDialog(Adr + "-  не найден");
            //InsBlockRef(@"D:\C#\TRASER_NC_AC19\Блоки\Плоскость_пр.dwg", "Плоскость_пр", "");
            //Vinoski.Clear();
            //PEREX.Clear();
            //SpPLOS.Clear();
            //using (DocumentLock docLock = doc.LockDocument()) { SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS); }
            //this.dataGridView5.Rows.Clear();
            //foreach (PLOS TV in SpPLOS) { this.dataGridView5.Rows.Add(TV.Vid, TV.List, TV.Osi, TV.Msk.ToString()); }
            this.Show();
        }//выставить плоскость
        public void FPoly(Point3dCollection TkoorPL)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            using (tr1)
            {
                //Polyline2d poly = new Polyline2d();
                Polyline3d poly = new Polyline3d();
                //Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = 5;
                poly.Layer = "Плоскости";
                poly.Closed = true;
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);
                    foreach (Point3d pt in TkoorPL) 
                    {
                        PolylineVertex3d vex3d = new PolylineVertex3d(pt);
                        //Vertex2d vex2d = new Vertex2d(pt,0,0,0,0);
                        //poly.AppendVertex(vex2d);
                        poly.AppendVertex(vex3d);
                        tr1.AddNewlyCreatedDBObject(vex3d, true);
                    }
                tr1.Commit();
            }
        }//создать полилинию
        public void CreateLayer(String layerName)
        {
            ObjectId layerId;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                // Проверяем нет ли еще слоя с таким именем в чертеже
                if (lt.Has(layerName))
                {
                    layerId = lt[layerName];
                }
                else
                {
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName; // Задаем имя слоя
                    layerId = lt.Add(ltr);
                    trans.AddNewlyCreatedDBObject(ltr, true);
                }
                trans.Commit();
            }
        }//создать слой
        public void NSozd_Izm(string i) { Sozd_Izm = i; }

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    //PromptPointResult pPtRes;
        //    Point3d TochMir = new Point3d(Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text), Convert.ToDouble(textBox3.Text));
        //    this.Hide();
        //    using (DocumentLock docLock = doc.LockDocument())
        //    {
        //        using (Transaction Tx = db.TransactionManager.StartTransaction())
        //        {

        //                // Открываем примитив
        //                Entity ent = Tx.GetObject(ID, OpenMode.ForWrite) as Entity;
        //                // Получаем таблицу зарегистрированных приложений
        //                RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForRead);
        //                if (!regTable.Has("LAUNCH01"))
        //                {
        //                    regTable.UpgradeOpen();
        //                    // Добавляем имя приложения, которое мы будем
        //                    // использовать в расширенных данных
        //                    RegAppTableRecord app =
        //                            new RegAppTableRecord();
        //                    app.Name = "LAUNCH01";
        //                    regTable.Add(app);
        //                    Tx.AddNewlyCreatedDBObject(app, true);
        //                }
        //                // Добавляем расширенные данные к примитиву
        //                ent.XData = new ResultBuffer(new TypedValue(1001, "LAUNCH01"), new TypedValue(1000, OSItl), new TypedValue(1011, TochMir), new TypedValue(1011, PSKtl), new TypedValue(1000, textBox4.Text), new TypedValue(1000, textBox5.Text));
        //            Tx.Commit();
        //        }
        //        this.Show();
        //    }
        //}//кнопка изменить плоскость

        public void InsBlockRef(string BlockPath, string NAME, string Vin, string TipDBL, string SUF, string ID)
        {
            // Активный документ в редакторе AutoCAD
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // База данных чертежа (в данном случае - активного документа)
            Database db = doc.Database;
            // Редактор базы данных чертежа
            // Запускаем транзакцию
            using (DocumentLock docLock = doc.LockDocument())
            {

                using (DbS.Transaction tr = db.TransactionManager.StartTransaction())
                {
                    EdI.Editor ed = doc.Editor;
                    EdI.PromptPointOptions pPtOpts;
                    pPtOpts = new EdI.PromptPointOptions("\nУкажите точку вставки блока: ");
                    // Выбор точки пользователем
                    var pPtRes = doc.Editor.GetPoint(pPtOpts);
                    if (pPtRes.Status != EdI.PromptStatus.OK)
                        return;
                    var ptStart = pPtRes.Value;

                    DbS.BlockTable bt = tr.GetObject(db.BlockTableId, DbS.OpenMode.ForRead) as DbS.BlockTable;
                    DbS.BlockTableRecord model = tr.GetObject(bt[DbS.BlockTableRecord.ModelSpace], DbS.OpenMode.ForWrite) as DbS.BlockTableRecord;
                    // Создаем новую базу
                    using (DbS.Database db1 = new DbS.Database(false, false))
                    {
                        // Получаем базу чертежа-донора
                        db1.ReadDwgFile(BlockPath, System.IO.FileShare.Read, true, null);
                        // Получаем ID нового блока
                        DbS.ObjectId BlkId = db.Insert(NAME, db1, false);
                        DbS.BlockReference bref = new DbS.BlockReference(ptStart, BlkId);
                        // Дефолтные свойства блока (слой, цвет и пр.)
                        //bref.SetDatabaseDefaults();
                        //bref.Layer = "Плоскости";
                        // Добавляем блок в модель
                        model.AppendEntity(bref);
                        // Добавляем блок в транзакцию
                        tr.AddNewlyCreatedDBObject(bref, true);
                        // Расчленяем блок     
                        bref.ExplodeToOwnerSpace();
                        bref.Erase();
                        //bref.Layer = "Плоскости";
                        // Закрываем транзакцию
                        tr.Commit();
                    }
                }
                SetDynamicBlkProperty();
            }
        }//Выставмить дин блок
        static public void SetDynamicBlkProperty( )
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForWrite);
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    bref.Layer = "Плоскости";
                    //bref.Rotation = Ugol;
                    //foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    //{
                    //    using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                    //    {
                    //        if (atrRef != null)
                    //        {
                    //            if (atrRef.Tag == "Помещение") { atrRef.TextString = Pom; }
                    //            if (atrRef.Tag == "Высота_установки") { atrRef.TextString = Visota; }
                    //            if (atrRef.Tag == "Раздел_спецификации") { atrRef.TextString = RazdelSp; }
                    //            if (atrRef.Tag == "Исполнение") { atrRef.TextString = Compon; }
                    //        }
                    //    }
                    //}
                }
                Tx.Commit();
            }
        }//расширеные данные
    }
    }

