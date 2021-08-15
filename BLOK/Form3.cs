
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;

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

    public partial class Form3 : Form
    {
        enum Napravl {forvard,bak};
        ObjectId ID;
        string HANDL = "";
        string stKomp = "";
        string stPom = "";
        string stRazd = "";
        string stKEI = "";
        string stHtoEto = "";
        string stHOZ = "";
        string stRab = "";
        string sLINK = "";
        string dDlin = "";
        string dVisot = "";

        double IDles = 0;

        public List<CWAY> SpCWAY = new List<CWAY>();
        List<Point3d> SpT = new List<Point3d>();
        List<TPosrL> SpPoint_poln = new List<TPosrL>();
        List<TPosrL> SpPoint1 = new List<TPosrL>();
        List<TPosrL> SpPoint1L2 = new List<TPosrL>();
        List<TPosrL> SpPointVN = new List<TPosrL>();
        List<TPosrL> SpPoint11 = new List<TPosrL>();
        List<TPosrL> SpPoint12 = new List<TPosrL>();
        List<TPosrL> SpPoint11_12 = new List<TPosrL>();
        List<Lestn> SpLESTN = new List<Lestn>();
        List<Prepad> SpPREP = new List<Prepad>();

        Point3d Tn = new Point3d();
        Point3d Tk = new Point3d();
        public struct Noda
        {
            public string Nom, Otkuda, Hoz, Vin, KoordMOD, Vid;
            public string SpSmNod;
            public double Ves;
            public double DlinP;
            public double VisT;
            public double Param;
            public Point3d Koord;
            public Point3d Koord3D;
            public void NVid(string i) { Vid = i; }
            public void NomVin(string i) { Vin = i; }
            public void NomNod(string i) { Nom = i; }
            public void NomNOtk(string i) { Otkuda = i; }
            public void NomHoz(string i) { Hoz = i; }
            public void NSpSmNod(string i) { SpSmNod = i; }
            public void NKoordMOD(string i) { KoordMOD = i; }
            public void NVes(double i) { Ves = i; }
            public void NDlinP(double i) { DlinP = i; }
            public void NVisT(double i) { VisT = i; }
            public void NParam(double i) { Param = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoor3D(Point3d i) { Koord3D = i; }
        }
        public struct Kriv
        {
            public string Name;
            public List<Noda> SpNod;
            public void NomName(string i) { Name = i; }
            public void NSpNod(List<Noda> i) { SpNod = i; }
        }
        public struct Prepad
        {
            public string Tip;
            public Point3d T1,T2,Tp;
            public double Rad,Param,Visot, Pram, Gib, PramGib, KVisot, ZazorVir, Zazor;
            public void NTip(string i) { Tip = i; }
            public void NT1(Point3d i) { T1 = i; }
            public void NT2(Point3d i) { T2 = i; }
            public void NTp(Point3d i) { Tp = i; }
            public void NRad(double i) { Rad = i; }
            public void NParam(double i) { Param = i; }
            public void NVisot(double i) { Visot = i; }
            public void GlDat(double pram, double gib, double pramGib, double kVisot, double zazorVir, double zazor) 
            {
                Pram = pram;
                Gib = gib;
                PramGib = pramGib;
                KVisot = kVisot;
                ZazorVir = zazorVir;
                Zazor = zazor;
            }

        }
        public struct Part
        {
            public string CWName, CompName, StartPoint, EndPoint, Coordinates, ModuleName;
            public void NCWName(string i) { CWName = i; }
            public void NCompName(string i) { CompName = i; }
            public void NStartPoint(string i) { StartPoint = i; }
            public void NEndPoint(string i) { EndPoint = i; }
            public void NCoordinates(string i) { Coordinates = i; }
            public void NModuleName(string i) { ModuleName = i; }
        }
        public struct CWAY
        {
            public string CWName, ModuleName, CompName, length;
            public List<Part> SpPart;
            public void NCWName(string i) { CWName = i; }
            public void NSpPart(List<Part> i) { SpPart = i; }
            public void NCompName(string i) { CompName = i; }
            public void NModuleName(string i) { ModuleName = i; }
            public void Nlength(string i) { length = i; }
        }
        public struct TPosrL
        {
            public ObjectId OId;
            public Point3d TPoint;
            public string Visot, Komp, Tip, DopDet, LINK;
            public double NomT;
            public void NVisot(string i) { Visot = i; }
            public void NKomp(string i) { Komp = i; }
            public void NTip(string i) { Tip = i; }
            public void NDopDet(string i) { DopDet = i; }
            public void NLINK(string i) { LINK = i; }
            public void NNomT(double i) { NomT = i; }
            public void NTPoint(Point3d i) { TPoint = i; }
            public void NOId(ObjectId i) { OId = i; }

        }
        public struct Lestn
        {
            public string Name, Povorot, RazdelSP, Soed, Hvost, DopSoed, DopHvost, Component, LINK,Room;
            public double Hsir, HsagXv, HsagSoed,IDCWey;
            public List<TPosrL> SpPoint1 , SpPoint_poln;
           
            public void NName(string i) { Name = i; }
            public void NPovorot(string i) { Povorot = i; }
            public void NRazdelSP(string i) { RazdelSP = i; }
            public void NSoed(string i) { Soed = i; }
            public void NHvost(string i) { Hvost = i; }
            public void NDopSoed(string i) { DopSoed = i; }
            public void NDopHvost(string i) { DopHvost = i; }
            public void NHsir(double i) { Hsir = i; }
            public void NHsagXv(double i) { HsagXv = i; }
            public void NHsagSoed(double i) { HsagSoed = i; }
            public void ADDAtr(string component, string chapter_Sp, string Find, string lINC, string connection, string shank, string room, double ID_CWEY, double Steep_Shank) 
            {
                Component = component;
                RazdelSP = chapter_Sp;
                Name = Find;
                LINK = lINC;
                Soed = connection;
                Hvost = shank;
                Room = room;
                IDCWey = ID_CWEY;
                HsagXv = Steep_Shank;
            }
            public void ADDSp(List<TPosrL> spPoint1, List<TPosrL> spPoint_poln) 
            {
                SpPoint1 = spPoint1;
                SpPoint_poln = spPoint_poln;
            }

        }
        public struct KONTUR
        {
            public string Name;
            public ObjectId ID;
            public List<Point3d> SpKoor;
            public List<Point3d> SpTper;
            public void NomName(string i) { Name = i; }
            public void NSpKoor(List<Point3d> i) { SpKoor = i; }
            public void NSpTper(List<Point3d> i) { SpTper = i; }
            public void NID(ObjectId i) { ID = i; }
        }
        public Form3()
        {
            InitializeComponent();
        }

        public void Form3_Load(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            int Schet = 0;

            Point3d Tser = new Point3d();
            Point3d TserPl1 = new Point3d();
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
                    Entity ent = (Entity)tr.GetObject(ers.ObjectId, OpenMode.ForWrite);
                    ID = ent.ObjectId;
                    HANDL = ent.Handle.ToString();
                    ResultBuffer buffer = ent.GetXDataForApplication("LAUNCH01");
                    // Если есть расширенные данные – удалим их.
                    // Для этого в качестве расширенных данных
                    // передаём только имя приложения.
                    // Только связанные с ним данные будут удалены.
                    //string strDan="";
                    if (buffer != null)
                    {
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 2) { dVisot = value.Value.ToString(); }
                            if (Schet == 4) { stPom = value.Value.ToString(); }
                            if (Schet == 5) { stRazd = value.Value.ToString(); }
                            if (Schet == 6) { stKEI = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 8) { stHOZ = value.Value.ToString(); }
                            if (Schet == 9) { dDlin = value.Value.ToString(); }
                            if (Schet == 10) { stRab = value.Value.ToString(); }
                            if (Schet == 11) { sLINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                        ////Application.ShowAlertDialog(strDan);
                    }
                    if (ent.GetType() == typeof(Polyline))
                    {
                        Polyline ln = tr.GetObject(ers.ObjectId, OpenMode.ForWrite) as Polyline;
                        double KolV = ln.EndParam;
                        for (int i = 0; i <= KolV; i++) SpT.Add(ln.GetPointAtParameter(i));
                        //T0 = SpT[0];
                        double Shir = SpT.Last().DistanceTo(SpT[0]);
                        Tser = SpT[(SpT.Count / 2) - 1];
                        TserPl1 = SpT[SpT.Count / 2];
                        Tn = SpT[SpT.Count - 2] + SpT[SpT.Count - 2].GetVectorTo(SpT[SpT.Count - 1]) / 2;
                        Tk = Tser + Tser.GetVectorTo(TserPl1) / 2;
                        //this.label27.Text = Tn.ToString() + Tk.ToString();
                    }
                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
                //Form4 form1 = new Form4();
                //form1.Show();
                this.label3.Text = HANDL;
                this.label7.Text = dDlin.ToString();
                this.textBox3.Text = stRab;
                this.label14.Text = stKomp;
                this.textBox2.Text = dVisot.ToString();
                this.textBox4.Text = stPom;
                this.textBox5.Text = stRazd;
                this.textBox6.Text = stKEI;
                this.textBox7.Text = stHtoEto;
                this.textBox8.Text = stHOZ;
                this.textBox9.Text = sLINK;
            }
            ADDLes("#ЛестУсил200", 208, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("#ЛестУсил300", 308, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("#ЛестУсил400", 408, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("#ЛестУсил500", 508, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("#ЛестУсил600", 608, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("#ЛестУсил700", 708, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("#ЛестУсил800", 808, "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 1200, 0, ref SpLESTN);
            ADDLes("ТрубаЦ25х3,2#01420986011", 34, "Тр", "", "42*60#СоедТрМуфта25#01001902005", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-06*1*796#Тр*КЛИГ.754141.007-16*1*796", "Тр*01002702005*1*796", 1500, 3000, ref SpLESTN);
            ADDLes("Труба32х4ГОСТ8734-75#01362760346", 32, "Тр", "", "45*38#СоедТрШтуцер#ИТШЛ.302615.086-05", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-06*1*796#Тр*КЛИГ.754141.007-16*1*796", "Тр*ИТШЛ.753081.012-06*1*796", 1500, 3000, ref SpLESTN);
            ADDLes("Труба34х4ГОСТ8734-75#01270807175", 34, "Тр", "", "42*60#СоедТрМуфта25#01001902005", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-06*1*796#Тр*КЛИГ.754141.007-16*1*796", "Тр*01002702005*1*796", 1500, 3000, ref SpLESTN);
            foreach (Lestn Tles in SpLESTN) this.dataGridView1.Rows.Add(Tles.Name, Tles.Hsir, Tles.RazdelSP, Tles.Povorot, Tles.Soed, Tles.Hvost, Tles.DopHvost, Tles.DopSoed, Tles.HsagXv, Tles.HsagSoed);
            List<TPosrL> SpOBJID_Poln = new List<TPosrL>();
            SBORpl(ref SpCWAY, ref SpOBJID_Poln, stRab);
            double IDt = 0;
            foreach (TPosrL TT in SpOBJID_Poln)
            {
                if (TT.Tip == "хв" & TT.DopDet != "") { this.textBox11.Text = TT.DopDet; }
                if (TT.Tip == "соед" & TT.DopDet != "") { this.textBox12.Text = TT.DopDet; }
                if (TT.Tip == "хв") { if (SpLESTN.Exists(x => x.Hvost.Contains(TT.Komp))) this.textBox15.Text = SpLESTN.Find(x => x.Hvost.Contains(TT.Komp)).Hvost; }
                if (TT.Tip == "соед") { if (SpLESTN.Exists(x => x.Soed.Contains(TT.Komp))) this.textBox16.Text = SpLESTN.Find(x => x.Soed.Contains(TT.Komp)).Soed; }
            }
            if (stRazd == "Лес") { this.textBox13.Text = "1200"; }
            if (stRazd == "Тр") { this.textBox13.Text = "1500"; this.textBox14.Text = "3000"; }
            foreach (CWAY tCWAY in SpCWAY)
            {
                if (Double.TryParse(tCWAY.CWName, out IDt)) IDt = Convert.ToDouble(tCWAY.CWName);
                if (IDt > IDles) IDles = IDt;
            }
            this.label17.Text = IDles.ToString() + "(" + SpCWAY.Count.ToString() + ")";
        }
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (this.dataGridView1.CurrentRow.Cells[0].Value == null) return;
            {
                this.label14.Text = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
                this.textBox11.Text = this.dataGridView1.CurrentRow.Cells[6].Value.ToString();
                this.textBox12.Text = this.dataGridView1.CurrentRow.Cells[7].Value.ToString();
                this.textBox13.Text = this.dataGridView1.CurrentRow.Cells[8].Value.ToString();
                this.textBox14.Text = this.dataGridView1.CurrentRow.Cells[9].Value.ToString();
                this.textBox15.Text = this.dataGridView1.CurrentRow.Cells[5].Value.ToString();
            }
        }
        public void button1_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
            SpPoint1.Clear();
            SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln,ref SpPREP);
            Perestr(SpPoint1, SpOBJID, SpPoint_poln);
            }
            this.Close();
        }//перестроить
        private void button2_Click(object sender, EventArgs e)
        {
            double tID = 0;
            double ttID = 0;
            double Hsirin = 108;
            double HSag_Hvost = 1000;
            double Zazor = Convert.ToDouble(this.textBox1.Text);
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            double Gib = Convert.ToDouble(this.textBox18.Text);
            double Pram = Convert.ToDouble(this.textBox10.Text);
            string Adr = "";
            string Xvost = "";
            double ParRazr = 0;
            double j = 0;
            bool Perv = true;
            bool RekXvost = this.checkBox1.Checked;
            Lestn TLest = new Lestn();
            if (SpLESTN.Exists(x => x.Name == this.label14.Text))
            {
                TLest = SpLESTN.Find(x => x.Name == this.label14.Text);
                Hsirin = TLest.Hsir;
            }
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                EdI.Editor ed = doc.Editor;
                EdI.PromptPointOptions pPtOpts;
                pPtOpts = new EdI.PromptPointOptions("\nУкажите первую точку ");
                var pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status == EdI.PromptStatus.OK) T1 = pPtRes.Value; else return;
                pPtOpts = new EdI.PromptPointOptions("\nУкажите вторую точку ");
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = T1;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status == EdI.PromptStatus.OK) T2 = pPtRes.Value; else return;
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                Find_KONTUR(ref SpPoint1, ref SpPointVN, Tn, Tk, SpT, T1);
                foreach (TPosrL ID in SpPoint_poln)
                {
                    if (ID.Tip == "Лес" | ID.Tip == "часть")
                    { ttID = Convert.ToDouble(ID.LINK.Split('-').Last());
                        if (ttID > tID) tID = ttID;
                    }
                }
                tID = tID + 1;
                SpPoint11.Add(SpPointVN[0]);
                for (int i = 0; i <= SpPointVN.Count - 2; i++)
                {
                    Point3d Tt1 = SpPointVN[i].TPoint;
                    Point3d Tt2 = SpPointVN[i + 1].TPoint;
                    Point3d TTper = TPer1(Tt1, Tt2, T1, T2);
                    Vector3d Vek1 = Tt1.GetVectorTo(Tt2);
                    Vector3d Vek2 = Tt2.GetVectorTo(Tt1);
                    double Ugol1 = Vek1.GetAngleTo(new Vector3d(1, 0, 0));
                    double Ugol2 = Ugol1 + Math.PI;
                    if (TTper != new Point3d())
                    {
                        if (i != 0) SpPoint11.Add(SpPointVN[i]);
                        j = j + 1;
                        Tper = TTper;
                        Point3d Tk = polar1(Tper, Vek2, Zazor / 2);
                        Point3d Tn = polar1(Tper, Vek1, Zazor / 2);
                        TPosrL TkL = new TPosrL();
                        if (Perepad > 0 & Gib > 0 & Pram > 0)
                        {
                            TPosrL TkLPr = new TPosrL();
                            TPosrL TkLGib = new TPosrL();
                            Point3d TkPr = polar1(Tper, Vek2, (Zazor / 2) + Pram);
                            Point3d TkGib = polar1(Tper, Vek2, (Zazor / 2) + Pram + Gib);
                            TkLGib.NNomT(SpPointVN[i].NomT + j / 10);
                            TkLGib.NTPoint(TkGib);
                            TkLGib.NVisot(SpPointVN[i].Visot);
                            SpPoint11.Add(TkLGib);
                            SpPoint1.Add(TkLGib);
                            j = j + 1;
                            TkLPr.NNomT(SpPointVN[i].NomT + j / 10);
                            TkLPr.NTPoint(TkPr);
                            TkLPr.NVisot((Convert.ToDouble(SpPointVN[i].Visot) - Perepad).ToString());
                            SpPoint11.Add(TkLPr);
                            SpPoint1.Add(TkLPr);
                            j = j + 1;
                        }
                        TkL.NNomT(i + j / 10);
                        TkL.NTPoint(Tk);
                        TkL.NVisot((Convert.ToDouble(SpPointVN[i].Visot) - Perepad).ToString());
                        TPosrL TnL = new TPosrL();
                        TnL.NNomT(0);
                        TnL.NTPoint(Tn);
                        TnL.NVisot((Convert.ToDouble(SpPointVN[i + 1].Visot) - Perepad).ToString());
                        SpPoint11.Add(TkL);
                        SpPoint12.Add(TnL);
                        if (Perepad > 0 & Gib > 0 & Pram > 0)
                        {
                            j = j + 1;
                            TPosrL TnLPr = new TPosrL();
                            TPosrL TnLGib = new TPosrL();
                            Point3d TnPr = polar1(Tper, Vek1, (Zazor / 2) + Pram);
                            Point3d TnGib = polar1(Tper, Vek1, (Zazor / 2) + Pram + Gib);
                            TnLPr.NNomT(j);
                            TnLPr.NTPoint(TnPr);
                            TnLPr.NVisot((Convert.ToDouble(SpPointVN[i].Visot) - Perepad).ToString());
                            SpPoint12.Add(TnLPr);
                            TnLPr.NNomT(SpPointVN[i].NomT + j / 10);
                            SpPoint1.Add(TnLPr);
                            j = j + 1;
                            TnLGib.NNomT(j);
                            TnLGib.NTPoint(TnGib);
                            TnLGib.NVisot(SpPointVN[i].Visot);
                            SpPoint12.Add(TnLGib);
                            TnLGib.NNomT(SpPointVN[i].NomT + j / 10);
                            SpPoint1.Add(TnLGib);
                        }
                        Perv = false;
                        ParRazr = i;
                    }
                    if (Perv)
                    {
                        if (i != 0) SpPoint11.Add(SpPointVN[i]);
                    }
                    else
                    {
                        if (i != 0 & i != ParRazr) SpPoint12.Add(SpPointVN[i]);
                    }
                    j = j + 1;
                }
                if (Perv) SpPoint11.Add(SpPointVN.Last()); else SpPoint12.Add(SpPointVN.Last());
                if (this.textBox5.Text == "Лес")
                {
                    if (SpPoint11.Count > 1) Lest_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, this.textBox4.Text, RekXvost);
                    tID = tID + 1;
                    if (SpPoint12.Count > 1) Lest_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, this.textBox4.Text, RekXvost);
                }
                if (this.textBox5.Text == "Тр")
                {
                    if (SpPoint11.Count > 1) TRub_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, RekXvost);
                    tID = tID + 1;
                    if (SpPoint12.Count > 1) TRub_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, RekXvost);
                }
                Transaction tr1 = db.TransactionManager.StartTransaction();
                using (tr1)
                {
                    foreach (TPosrL ID in SpPoint_poln)
                    {
                        if (ID.LINK == this.textBox9.Text)
                        {
                            Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                            Obj.Erase();
                        }
                    }
                    if (Perepad > 0 | Gib > 0 | Pram > 0)
                    {
                        double n = 0;
                        foreach (TPosrL ID in SpPoint_poln)
                        {
                            if (ID.Tip == "Точка построения")
                            {
                                Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                                Obj.Erase();
                            }
                        }
                        SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        foreach (TPosrL TT in SpPoint1)
                        {
                            if (TT.Visot != "")
                            {
                                KrugIText(TT, SpPoint1, n.ToString(), this.textBox3.Text);
                                n = n + 1;
                            }
                        }
                    }
                    Entity Obj1 = tr1.GetObject(ID, OpenMode.ForWrite) as Entity;
                    Obj1.Erase();
                    tr1.Commit();
                }
            }
            this.Close();
        }//создать разрыв
        private void button5_Click(object sender, EventArgs e)
        {
            double tID = 0;
            double ttID = 0;
            double Hsirin = 108;
            double HSag_Hvost = 1000;
            double Zazor = Convert.ToDouble(this.textBox1.Text);
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            double Gib = Convert.ToDouble(this.textBox18.Text);
            double Pram = Convert.ToDouble(this.textBox19.Text);
            string Adr = "";
            string Xvost = "";
            double ParRazr = 0;
            double j = 0;
            bool Perv = true;
            bool RekXvost = this.checkBox1.Checked;
            Lestn TLest = new Lestn();
            if (SpLESTN.Exists(x => x.Name == this.label14.Text))
            {
                TLest = SpLESTN.Find(x => x.Name == this.label14.Text);
                Hsirin = TLest.Hsir;
            }
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                EdI.Editor ed = doc.Editor;
                EdI.PromptPointOptions pPtOpts;
                pPtOpts = new EdI.PromptPointOptions("\nУкажите первую точку ");
                var pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status == EdI.PromptStatus.OK) T1 = pPtRes.Value; else return;
                pPtOpts = new EdI.PromptPointOptions("\nУкажите вторую точку ");
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = T1;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status == EdI.PromptStatus.OK) T2 = pPtRes.Value; else return;
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                Find_KONTUR(ref SpPoint1, ref SpPointVN, Tn, Tk, SpT, T1);
                foreach (TPosrL ID in SpPoint_poln)
                {
                    if (ID.Tip == "Лес" | ID.Tip == "часть")
                    {
                        ttID = Convert.ToDouble(ID.LINK.Split('-').Last());
                        if (ttID > tID) tID = ttID;
                    }
                }
                tID = tID + 1;
                SpPoint11.Add(SpPointVN[0]);
                for (int i = 0; i <= SpPointVN.Count - 2; i++)
                {
                    Point3d Tt1 = SpPointVN[i].TPoint;
                    Point3d Tt2 = SpPointVN[i + 1].TPoint;
                    Point3d TTper = TPer1(Tt1, Tt2, T1, T2);
                    Vector3d Vek1 = Tt1.GetVectorTo(Tt2);
                    Vector3d Vek2 = Tt2.GetVectorTo(Tt1);
                    double Ugol1 = Vek1.GetAngleTo(new Vector3d(1, 0, 0));
                    double Ugol2 = Ugol1 + Math.PI;
                    if (TTper != new Point3d())
                    {
                        if (i != 0) SpPoint11.Add(SpPointVN[i]);
                        j = j + 1;
                        Tper = TTper;
                        TPosrL TkLGib = new TPosrL();
                        Point3d TkGib = polar1(Tper, Vek2, (Pram / 2) + Gib);
                        TkLGib.NNomT(SpPointVN[i].NomT + j / 10);
                        TkLGib.NTPoint(TkGib);
                        TkLGib.NVisot(SpPointVN[i].Visot);
                        SpPoint11.Add(TkLGib);
                        SpPoint1.Add(TkLGib);
                        j = j + 1;
                        TPosrL TkL = new TPosrL();
                        Point3d Tk = polar1(Tper, Vek2, Pram / 2);
                        TkL.NNomT(SpPointVN[i].NomT + j / 10);
                        TkL.NTPoint(Tk);
                        TkL.NVisot((Convert.ToDouble(SpPointVN[i].Visot) - Perepad).ToString());
                        SpPoint11.Add(TkL);
                        SpPoint1.Add(TkL);
                        j = j + 1;
                        TPosrL TnL = new TPosrL();
                        Point3d Tn = polar1(Tper, Vek1, Pram / 2);
                        TnL.NNomT(SpPointVN[i].NomT + j / 10);
                        TnL.NTPoint(Tn);
                        TnL.NVisot((Convert.ToDouble(SpPointVN[i + 1].Visot) - Perepad).ToString());
                        SpPoint11.Add(TnL);
                        SpPoint1.Add(TnL);
                        j = j + 1;
                        TPosrL TnLGib = new TPosrL();
                        Point3d TnGib = polar1(Tper, Vek1, (Pram / 2) + Gib);
                        TnLGib.NNomT(i + j / 10);
                        TnLGib.NTPoint(TnGib);
                        TnLGib.NVisot(SpPointVN[i].Visot);
                        SpPoint12.Add(TnLGib);
                        TnLGib.NNomT(SpPointVN[i].NomT + j / 10);
                        SpPoint11.Add(TnLGib);
                        SpPoint1.Add(TnLGib);
                        Perv = false;
                        ParRazr = i;
                    }
                    if (i != 0) SpPoint11.Add(SpPointVN[i]);
                }
                SpPoint11.Add(SpPointVN.Last());
                if (this.textBox5.Text == "Лес")
                {
                    if (SpPoint11.Count > 1) Lest_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, this.textBox4.Text, RekXvost);
                }
                if (this.textBox5.Text == "Тр")
                {
                    if (SpPoint11.Count > 1) TRub_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, RekXvost);
                }
                Transaction tr1 = db.TransactionManager.StartTransaction();
                using (tr1)
                {
                    foreach (TPosrL ID in SpPoint_poln)
                    {
                        if (ID.LINK == this.textBox9.Text)
                        {
                            Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                            Obj.Erase();
                        }
                    }
                    if (Perepad > 0 | Gib > 0 | Pram > 0)
                    {
                        double n = 0;
                        foreach (TPosrL ID in SpPoint_poln)
                        {
                            if (ID.Tip == "Точка построения")
                            {
                                Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                                Obj.Erase();
                            }
                        }
                        SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        foreach (TPosrL TT in SpPoint1)
                        {
                            if (TT.Visot != "")
                            {
                                KrugIText(TT, SpPoint1, n.ToString(), this.textBox3.Text);
                                n = n + 1;
                            }
                        }
                    }
                    Entity Obj1 = tr1.GetObject(ID, OpenMode.ForWrite) as Entity;
                    Obj1.Erase();
                    tr1.Commit();
                }
            }
            this.Close();
        }//Создать месный гиб
        private void button4_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                EdI.Editor ed = doc.Editor;
                EdI.PromptPointOptions pPtOpts;
                pPtOpts = new EdI.PromptPointOptions("\nУкажите первую точку ");
                var pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status == EdI.PromptStatus.OK) T1 = pPtRes.Value; else return;
                pPtRes = ed.GetCorner("\nУкажите вторую точку: ", T1);
                if (pPtRes.Status == EdI.PromptStatus.OK) T2 = pPtRes.Value; else return;
                Point3dCollection Kontur = CoorKontur(T1, T2);
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                RazrSekR(Kontur, SpPoint1, SpOBJID, SpPoint_poln, Perepad, SpPoint1L2);
            }
            this.Close();
        }//разрыв секущей рамкой заданой прямоугольником
        private void button9_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3dCollection Kontur = new Point3dCollection();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
            PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            using (DocumentLock docLock = doc.LockDocument())
            {
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    Entity bref1 = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                    string[] TYPEm = bref1.GetType().ToString().Split('.');
                    string TYPE = TYPEm.Last();
                        if (TYPE == "Polyline")
                        {
                        Polyline ln = Tx.GetObject(bref1.ObjectId, OpenMode.ForWrite) as Polyline;
                        double KolV = ln.EndParam;
                        for (int i = 0; i <= KolV; i++) Kontur.Add(ln.GetPointAtParameter(i));
                        }
                    }
                    SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                RazrSekR(Kontur, SpPoint1, SpOBJID, SpPoint_poln, Perepad, SpPoint1L2);
            }
            this.Close();
        }//Создать разрыв указав контур
        private void button11_Click(object sender, EventArgs e)
        {
            int Schet = 0;
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3dCollection Kontur = new Point3dCollection();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
            PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            using (DocumentLock docLock = doc.LockDocument())
            {
            SpPoint1.Clear();
            SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    Entity bref1 = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                    string[] TYPEm = bref1.GetType().ToString().Split('.');
                    string TYPE = TYPEm.Last();
                    if (TYPE == "Polyline")
                    {
                        Polyline ln = Tx.GetObject(bref1.ObjectId, OpenMode.ForWrite) as Polyline;
                        double KolV = ln.EndParam;
                        for (int i = 0; i <= KolV; i++) Kontur.Add(ln.GetPointAtParameter(i));
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        // Если есть расширенные данные – удалим их.
                        // Для этого в качестве расширенных данных
                        // передаём только имя приложения.
                        // Только связанные с ним данные будут удалены.
                        //string strDan="";
                        dVisot = "0";
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = value.Value.ToString(); }
                                if (Schet == 10) { stRab = value.Value.ToString(); }
                                if (Schet == 11) { sLINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                        List<ObjectId> SpOBJIDL2 = new List<ObjectId>();
                        List<TPosrL> SpPoint_polnL2 = new List<TPosrL>();
                        SBORBl_FIND(ref SpPoint1L2, stRab, ref SpOBJIDL2, ref SpPoint_polnL2, ref SpPREP);
                        //Application.ShowAlertDialog(SpPoint1L2.Count.ToString());
                    }
                }
                RazrSekR(Kontur, SpPoint1, SpOBJID, SpPoint_poln, Perepad, SpPoint1L2);
            }
            this.Close();
        }//подрезать по другой лестнице
        private void button6_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                EdI.Editor ed = doc.Editor;
                EdI.PromptPointOptions pPtOpts;
                pPtOpts = new EdI.PromptPointOptions("\nУкажите первую точку ");
                var pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status == EdI.PromptStatus.OK) T1 = pPtRes.Value; else return;
                pPtRes = ed.GetCorner("\nУкажите вторую точку: ", T1);
                if (pPtRes.Status == EdI.PromptStatus.OK) T2 = pPtRes.Value; else return;
                Point3dCollection Kontur = CoorKontur(T1, T2);
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                GibSekR(Kontur, SpPoint1, SpOBJID, SpPoint_poln, SpPoint1L2);
            }
            this.Close();
        }//Создать гиб секущей рамкой
        private void button10_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3dCollection Kontur = new Point3dCollection();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
            PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
            using (DocumentLock docLock = doc.LockDocument())
            {
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    Entity bref1 = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                    string[] TYPEm = bref1.GetType().ToString().Split('.');
                    string TYPE = TYPEm.Last();
                    if (TYPE == "Polyline")
                    {
                        Polyline ln = Tx.GetObject(bref1.ObjectId, OpenMode.ForWrite) as Polyline;
                        double KolV = ln.EndParam;
                        for (int i = 0; i <= KolV; i++) Kontur.Add(ln.GetPointAtParameter(i));
                    }
                }
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                GibSekR(Kontur, SpPoint1, SpOBJID, SpPoint_poln, SpPoint1L2);
            }
            this.Close();
        }//Создать гиб указав контур
        private void button12_Click(object sender, EventArgs e)
        {
            int Schet = 0;
            List<ObjectId> SpOBJID = new List<ObjectId>();
            this.Hide();
            Point3dCollection Kontur = new Point3dCollection();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d Tper = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
            PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            using (DocumentLock docLock = doc.LockDocument())
            {
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    Entity bref1 = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                    string[] TYPEm = bref1.GetType().ToString().Split('.');
                    string TYPE = TYPEm.Last();
                    if (TYPE == "Polyline")
                    {
                        Polyline ln = Tx.GetObject(bref1.ObjectId, OpenMode.ForWrite) as Polyline;
                        double KolV = ln.EndParam;
                        for (int i = 0; i <= KolV; i++) Kontur.Add(ln.GetPointAtParameter(i));
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        // Если есть расширенные данные – удалим их.
                        // Для этого в качестве расширенных данных
                        // передаём только имя приложения.
                        // Только связанные с ним данные будут удалены.
                        //string strDan="";
                        dVisot = "0";
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = value.Value.ToString(); }
                                if (Schet == 10) { stRab = value.Value.ToString(); }
                                if (Schet == 11) { sLINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                        List<ObjectId> SpOBJIDL2 = new List<ObjectId>();
                        List<TPosrL> SpPoint_polnL2 = new List<TPosrL>();
                        SBORBl_FIND(ref SpPoint1L2, stRab, ref SpOBJIDL2, ref SpPoint_polnL2, ref SpPREP);
                        //Application.ShowAlertDialog(SpPoint1L2.Count.ToString());
                    }
                }
                GibSekR(Kontur, SpPoint1, SpOBJID, SpPoint_poln, SpPoint1L2);
            }
            this.Close();
        }//гиб по другой лестнице
        private void button3_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                this.dataGridView2.Rows.Clear();
                foreach (TPosrL Tles in SpPoint1)
                {
                    if (Tles.Tip == "Точка построения")
                        this.dataGridView2.Rows.Add(Tles.NomT, Tles.Visot);
                }
            }
        }//Список связаных обьектов
        private void button7_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                List<TPosrL> spTpostr = SpPoint_poln.FindAll(x => x.Tip == "Точка построения");
                UdalToch(ref spTpostr);
                Perestr(spTpostr, SpOBJID, SpPoint_poln);
            }
        this.Close();
        }//Удалить лишние точки
        private void button8_Click(object sender, EventArgs e)
        {
            if (this.dataGridView2.CurrentRow.Cells[0].Value != null) 
            {
                this.Hide();
                List<ObjectId> SpOBJID = new List<ObjectId>();
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    SpPoint1.Clear();
                    SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                    List<TPosrL> SpPoint = new List<TPosrL>();
                    TPosrL T0 = SpPoint1.Find(x => x.NomT == Convert.ToDouble(this.dataGridView2.CurrentRow.Cells[0].Value));
                    int nom = SpPoint1.FindIndex(x => x.NomT == Convert.ToDouble(this.dataGridView2.CurrentRow.Cells[0].Value));
                    double Rast = 1;
                    if (T0.NomT != SpPoint1.Last().NomT) Rast = SpPoint1[nom + 1].NomT - SpPoint1[nom].NomT;
                    SpPoint.Add(T0);
                    string TVis = T0.Visot;
                    while (TVis != "")
                    {
                        TPosrL TPoint = FTpostr(TVis, SpPoint, ID.ToString());
                        TVis = TPoint.Visot;
                        if (TPoint.Visot != "") { SpPoint.Add(TPoint); }
                    }
                    SpPoint.Remove(T0);
                    double nadb = 0;
                    double Plus = Rast / SpPoint.Count();
                    foreach (TPosrL TT in SpPoint)
                    {
                        nadb = nadb + Plus;
                        TT.NNomT(T0.NomT + nadb);
                        SpPoint1.Add(TT);
                    }
                    SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    Perestr(SpPoint1, SpOBJID, SpPoint_poln);
                }
                this.Close();
            }
        }//Добавить точки после выбраной
        private void button13_Click(object sender, EventArgs e)
        {
            double tID = 0;
            double ttID = 0;
            double Zazor = Convert.ToDouble(this.textBox1.Text)/2;
            double Gib = Convert.ToDouble(this.textBox18.Text);
            double Pram = Convert.ToDouble(this.textBox10.Text);
            double PramGib = Convert.ToDouble(this.textBox19.Text)/2;
            double KVisot = Convert.ToDouble(this.textBox21.Text);
            double ZazorVir = Convert.ToDouble(this.textBox20.Text);
            Prepad GlObcl = new Prepad();
            GlObcl.GlDat(Pram, Gib, PramGib, KVisot, ZazorVir, Zazor);
            double Steep_Shank = Convert.ToDouble(this.textBox13.Text);
            double ID_CWEY = Convert.ToDouble(this.textBox3.Text);
            string Find = this.textBox3.Text;
            string chapter_Sp = this.textBox5.Text;
            string Component = this.label14.Text;
            string shank = this.textBox15.Text;
            string connection = this.textBox16.Text;
            string LINK = this.textBox9.Text;
            string Room = this.textBox4.Text;
            List<ObjectId> SpOBJID = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SpPoint1.Clear();
                SBORBl_FIND(ref SpPoint1, Find, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                Lestn CWEYGlob = new Lestn();
                CWEYGlob.ADDAtr(Component, chapter_Sp, Find, LINK, connection, shank, Room, ID_CWEY, Steep_Shank);
                CWEYGlob.ADDSp(SpPoint1, SpPoint_poln);
                apply_obstacles(SpPREP, SpLESTN, CWEYGlob, GlObcl,checkBox1.Checked,ID);
            }
        }//Применить препядствия
        private void button14_Click(object sender, EventArgs e)
        {
            List<ObjectId> SpOBJID = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SBORBl_FIND(ref SpPoint1, this.textBox3.Text, ref SpOBJID, ref SpPoint_poln, ref SpPREP);
                this.dataGridView3.Rows.Clear();
                foreach (Prepad TP in SpPREP) this.dataGridView3.Rows.Add(TP.Tip, TP.Rad);
            }
        }//Показать список препядствий
        public TPosrL FTpostr(string TVis, List<TPosrL> SpPoint, string ID)
        {
            string Rez = "";
            TPosrL Tpoint = new TPosrL();
            Tpoint.NVisot("");
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (DocumentLock docLock = doc.LockDocument())
            {
                EdI.Editor ed = doc.Editor;
                EdI.PromptPointOptions pPtOpts;

                string Nom = "0";
                if (SpPoint.Count > 0) Nom = SpPoint.Count.ToString();
                pPtOpts = new EdI.PromptPointOptions("\nУкажите точку ");

                if (SpPoint.Count > 0)
                {
                    pPtOpts.UseBasePoint = true;
                    pPtOpts.BasePoint = SpPoint.Last().TPoint;
                }
                string strOk = "Введи высоту точки <" + TVis + ">";
                var pPtRes = doc.Editor.GetPoint(pPtOpts);
                if (pPtRes.Status != EdI.PromptStatus.OK) return Tpoint;
                var Visot = doc.Editor.GetString(strOk);
                Rez = Visot.StringResult;
                if (Rez != "") TVis = Visot.StringResult;
                Point3d ptStart = pPtRes.Value;
                Tpoint.NTPoint(ptStart);
                Tpoint.NVisot(TVis);
            }
            return Tpoint;
        }//добовление точки в список 
        public static void Lest_N_INtr_SP(List<TPosrL> SpPointVH, string Tipor, double Hsirin, string Zagal, string Povorot, string Xvost, string Soed, string Adr, double Otstup, double HSag_Hvost, double ID_Les, ref double ID, bool StrTosh,string Room,bool RekXv)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForWrite);
                    if (!regTable.Has("LAUNCH01"))
                    {
                        regTable.UpgradeOpen();
                        // Добавляем имя приложения, которое мы будем
                        // использовать в расширенных данных
                        RegAppTableRecord app =
                                new RegAppTableRecord();
                        app.Name = "LAUNCH01";
                        regTable.Add(app);
                        Tx.AddNewlyCreatedDBObject(app, true);
                    }
                    Tx.Commit();
                }
                 List<string> SpID = new List<string>();
                List<TPosrL> SpPoint = new List<TPosrL>();
                //this.Hide();
                string TVis = "0";
                foreach (TPosrL TPoint in SpPointVH)
                {
                    TVis = TPoint.Visot;
                    if (TPoint.Visot != "")
                    {
                        if (SpPoint.Count() > 1)
                        {
                            TPosrL PostT = SpPoint.Last();
                            Vector3d Vek = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
                            double tan1 = (TPoint.TPoint.Y - SpPoint.Last().TPoint.Y) / (TPoint.TPoint.X - SpPoint.Last().TPoint.X);
                            double tan = (SpPoint.Last().TPoint.Y - SpPoint[SpPoint.Count - 2].TPoint.Y) / (SpPoint.Last().TPoint.X - SpPoint[SpPoint.Count - 2].TPoint.X);
                            double alf = Math.Abs(Math.Atan(tan1) - Math.Atan(tan));
                            double Ugol = 0;
                            if (alf == Math.PI / 2 | alf == 3 * Math.PI / 2)
                            {
                                Vector3d Vek1 = SpPoint.Last().TPoint.GetVectorTo(SpPoint[SpPoint.Count - 2].TPoint);
                                Vector3d Vek2 = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
                                //Application.ShowAlertDialog(Vek1.ToString() + " " + Vek2.ToString());
                                if ((Vek1.X < 0 & Vek2.Y < 0) | (Vek1.Y < 0 & Vek2.X < 0)) Ugol = 3 * (Math.PI) / 2;
                                if (Vek1.X > 0 & Vek2.Y > 0 | (Vek1.Y > 0 & Vek2.X > 0)) Ugol = (Math.PI) / 2;
                                if (Vek1.X < 0 & Vek2.Y > 0 | (Vek1.Y > 0 & Vek2.X < 0)) Ugol = Math.PI;
                                FPovorot(SpPoint.Last().TPoint, "", "#ПоворотГориз90ЛестУсил" + (Hsirin - 8).ToString(), Ugol, "0", Hsirin, Room, ID_Les.ToString());
                                double alf1 = Math.Atan(tan);
                                double alf2 = Math.Atan(tan1);
                                if (alf1 == 0) alf1 = Math.PI;
                                if (alf1 == Math.PI) alf1 = 0;
                                if (alf1 == Math.PI / 2) alf1 = -1 * Math.PI / 2;
                                if (alf1 == 3 * Math.PI / 2) alf1 = Math.PI / 2;
                                //Point3d PoslT = polar(SpPoint.Last().TPoint, alf1, alf1, Otstup);
                                Point3d PoslT = polar1(SpPoint.Last().TPoint, Vek1, Otstup);
                                TPosrL TPoslT = new TPosrL();
                                TPoslT.NTPoint(PoslT);
                                TPoslT.NVisot(SpPoint.Last().Visot);
                                SpPoint[SpPoint.Count - 1] = TPoslT;
                                Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, Room, ID_Les.ToString(), Zagal, RekXv);
                                //ID_Les = ID_Les + 1;
                                ID = ID + 1;
                                //Point3d PervT = polar(PostT.TPoint, alf2, alf2, Otstup);
                                Point3d PervT = polar1(PostT.TPoint, Vek2, Otstup);
                                TPosrL TPervT = new TPosrL();
                                TPervT.NTPoint(PervT);
                                TPervT.NVisot(PostT.Visot);
                                SpPoint.Clear();
                                SpPoint.Add(TPervT);
                            }
                        }
                        SpPoint.Add(TPoint);
                    }
                }
                if (SpPoint.Count > 1) Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, Room, ID_Les.ToString(), Zagal, RekXv);
                double n = 0;
                if (StrTosh)
                    foreach (TPosrL TT in SpPointVH)
                    {
                        if (TT.Visot != "")
                        {
                            //KrugIText(TT, SpPoint, TT.NomT.ToString(), ID_Les.ToString());
                            KrugIText(TT, SpPoint, n.ToString(), ID_Les.ToString());
                            n = n + 1;
                        }
                    }
            }
            //this.Show();
        }//не интерактивный способ построения лестниц
        public void TRub_N_INtr_SP(List<TPosrL> SpPointVH, string Tipor, double Hsirin, string Zagal, string Povorot, string Xvost, string Soed, string Adr, double Otstup, double HSag_Hvost, string DobavSoed, string DobavKr, double HSag_Kr, double ID_Les, ref double ID, bool StrTosh, bool RekXv)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForWrite);
                    if (!regTable.Has("LAUNCH01"))
                    {
                        regTable.UpgradeOpen();
                        // Добавляем имя приложения, которое мы будем
                        // использовать в расширенных данных
                        RegAppTableRecord app =
                                new RegAppTableRecord();
                        app.Name = "LAUNCH01";
                        regTable.Add(app);
                        Tx.AddNewlyCreatedDBObject(app, true);
                    }
                    Tx.Commit();
                }
                List<string> SpID = new List<string>();
                List<TPosrL> SpPoint = new List<TPosrL>();
                this.Hide();
                //double ID = 0;
                string TVis = "0";
                foreach (TPosrL TPoint in SpPointVH)
                {
                    //TPosrL TPoint = FTpostr(TVis, SpPoint, ID.ToString());
                    TVis = TPoint.Visot;
                    if (TPoint.Visot != "")
                    {
                        if (SpPoint.Count() > 1)
                        {
                            TPosrL PostT = SpPoint.Last();
                            Vector3d Vek = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
                            double tan1 = (TPoint.TPoint.Y - SpPoint.Last().TPoint.Y) / (TPoint.TPoint.X - SpPoint.Last().TPoint.X);
                            double tan = (SpPoint.Last().TPoint.Y - SpPoint[SpPoint.Count - 2].TPoint.Y) / (SpPoint.Last().TPoint.X - SpPoint[SpPoint.Count - 2].TPoint.X);
                            double Delt = 0;
                            Vector3d Vek1 = SpPoint[SpPoint.Count - 2].TPoint.GetVectorTo(SpPoint[SpPoint.Count - 1].TPoint);
                            Vector3d Vek2 = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
                            Vector3d Vek3 = SpPoint.Last().TPoint.GetVectorTo(SpPoint[SpPoint.Count - 2].TPoint);
                            //double alf = Vek3.GetAngleTo(Vek);
                            double alf11 = Vek1.GetAngleTo(new Vector3d(1, 0, 0));
                            double alf21 = Vek2.GetAngleTo(new Vector3d(1, 0, 0));
                            if (Math.Abs(alf11 - alf21) > 0.05)
                            {
                                FPovorot_TR(SpPoint[SpPoint.Count - 2].TPoint, SpPoint.Last().TPoint, TPoint.TPoint, ID_Les.ToString(), "ПоворотТр" + Vek3.GetAngleTo(Vek2).ToString("0.###") + "-" + Tipor, "0", Hsirin, this.textBox1.Text, 6, 1, ref Delt, ID_Les.ToString(), Tipor);
                                double alf1 = Math.Atan(tan);
                                double alf2 = Math.Atan(tan1);
                                if (alf1 == 0) alf1 = Math.PI;
                                if (alf1 == Math.PI) alf1 = 0;
                                if (alf1 == Math.PI / 2) alf1 = -1 * Math.PI / 2;
                                if (alf1 == 3 * Math.PI / 2) alf1 = Math.PI / 2;
                                Point3d PoslT = polar1(SpPoint.Last().TPoint, Vek3, Delt);
                                TPosrL TPoslT = new TPosrL();
                                TPoslT.NTPoint(PoslT);
                                TPoslT.NVisot(SpPoint.Last().Visot);
                                SpPoint[SpPoint.Count - 1] = TPoslT;
                                Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, this.textBox1.Text, ID_Les.ToString(), Zagal, RekXv);
                                ID = ID + 1;
                                Point3d PervT = polar1(PostT.TPoint, Vek2, Delt);
                                TPosrL TPervT = new TPosrL();
                                TPervT.NTPoint(PervT);
                                TPervT.NVisot(PostT.Visot);
                                SpPoint.Clear();
                                SpPoint.Add(TPervT);
                            }
                        }
                        SpPoint.Add(TPoint);
                    }
                }
                if (SpPoint.Count > 1) Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, this.textBox4.Text, ID_Les.ToString(), Zagal, RekXv);
                double Dlin = 0;
                double n = 0;
                double Ang = 0;
                double Ang1 = 0;
                double DlinStor1 = 0;
                string Visot = "";
                Point3dCollection TkoorPL = new Point3dCollection();
                Point3d Point1 = SpPointVH[0].TPoint;
                if (StrTosh) foreach (TPosrL TT in SpPointVH)
                    {
                        if (TT.Visot != "")
                        {
                            KrugIText(TT, SpPoint, n.ToString(), ID_Les.ToString());
                            TkoorPL.Add(TT.TPoint);
                            Dlin = Dlin + Point1.DistanceTo(TT.TPoint);
                            Point1 = TT.TPoint;
                            n = n + 1;
                        }
                    }
                Point1 = new Point3d();
                if (this.checkBox2.Checked == true)
                {
                    for (double RastTT = 40; RastTT < Dlin; RastTT = RastTT + HSag_Hvost)
                    {
                        Point_Ang(ref Point1, ref Ang, TkoorPL, RastTT, SpPoint, ref Visot, ref Ang1);
                        FXvost(Point1, "", ID_Les.ToString(), Ang, Visot, this.textBox1.Text, Xvost, DobavKr);
                    }
                    Point_Ang(ref Point1, ref Ang, TkoorPL, Dlin - 40, SpPoint, ref Visot, ref Ang1);
                    FXvost(Point1, "", ID_Les.ToString(), Ang, Visot, this.textBox1.Text, Xvost, DobavKr);
                }
                if (this.checkBox3.Checked == true)
                    for (double RastTT = HSag_Kr; RastTT < Dlin; RastTT = RastTT + HSag_Kr)
                    {
                        Point_Ang(ref Point1, ref Ang, TkoorPL, RastTT, SpPoint, ref Visot, ref Ang1);
                        FSoedin(Point1, "", Soed, Ang, Visot, 0, this.textBox1.Text, ID_Les.ToString(), Zagal, DobavSoed);
                    }
            }
            this.Show();
        }//не интерактивный способ построения лестниц
        public static void Postr(List<TPosrL> SpPoint, double Dles, string ID, double HSag_Hvost, string Adr, string Xvost, string COMPON, string POM, string InD, string Razdel,bool RekreXvos)
        {
            double IDd = 0;
            double Dlin = 0;
            double DeltX = 0;
            double DeltY = 0;
            double DeltX1 = 0;
            double DeltX2 = 0;
            double tan = 0;
            double alf = 0;
            double alf1 = 0;
            double alf2 = 0;
            Point3dCollection TkoorPL1 = new Point3dCollection();
            Point3dCollection TkoorPL2 = new Point3dCollection();
            Point3dCollection TkoorPL11 = new Point3dCollection();
            Point3dCollection TkoorPL22 = new Point3dCollection();
            Point3d T1k = new Point3d();
            Point3d T2k = new Point3d();
            Point3d T1 = new Point3d();
            Point3d T2 = new Point3d();
            Point3d T3 = new Point3d();
            Point3d T4 = new Point3d();
            Point3d Tipl1 = new Point3d();
            Point3d Ti = new Point3d();
            Point3d T0 = SpPoint[0].TPoint;
            for (int i = 0; i < SpPoint.Count - 1; i++)
            {
                Dlin = Dlin + SpPoint[i].TPoint.DistanceTo(SpPoint[i + 1].TPoint);
                //KrugIText(SpPoint[i], SpPoint, IDd.ToString(), InD);
                IDd = IDd + 1;
                //Vector3d Vek = SpPoint[i].TPoint.GetVectorTo(SpPoint[i + 1].TPoint);
                //Vector3d Vek1 = Vek + new Vector3d(1,0,0);
                //Vector3d Vek2 = Vek + new Vector3d(-1, 0, 0);
                tan = (SpPoint[i + 1].TPoint.Y - SpPoint[i].TPoint.Y) / (SpPoint[i + 1].TPoint.X - SpPoint[i].TPoint.X);
                alf = Math.Atan(tan);
                alf1 = Math.Atan(tan) + Math.PI / 2;
                alf2 = Math.Atan(tan) - Math.PI / 2;
                //Application.ShowAlertDialog("alf=" + alf + " alf1=" +alf1 + " alf2=" + alf2);
                T1 = polar(SpPoint[i].TPoint, alf1, alf1, Dles);
                T2 = polar(SpPoint[i].TPoint, alf2, alf2, Dles);
                T3 = polar(SpPoint[i + 1].TPoint, alf1, alf1, Dles);
                T4 = polar(SpPoint[i + 1].TPoint, alf2, alf2, Dles);
                //T1 = polar1(SpPoint[i].TPoint, Vek1, Dles);
                //T2 = polar1(SpPoint[i].TPoint, Vek2, Dles);
                //T3 = polar1(SpPoint[i + 1].TPoint, Vek1, Dles);
                //T4 = polar1(SpPoint[i + 1].TPoint, Vek2, Dles);
                if (TkoorPL1.Count == 0)
                {
                    TkoorPL1.Add(T1);
                    TkoorPL2.Add(T2);
                    TkoorPL1.Add(T3);
                    TkoorPL2.Add(T4);
                }
                else
                {
                    double dist1 = T1.DistanceTo(TkoorPL1[TkoorPL1.Count - 1]);
                    double dist2 = T1.DistanceTo(TkoorPL2[TkoorPL2.Count - 1]);
                    if (dist1 < dist2)
                    {
                        TkoorPL1.Add(T1);
                        TkoorPL2.Add(T2);
                        TkoorPL1.Add(T3);
                        TkoorPL2.Add(T4);
                    }
                    else
                    {
                        TkoorPL1.Add(T2);
                        TkoorPL2.Add(T1);
                        TkoorPL1.Add(T4);
                        TkoorPL2.Add(T3);
                    }
                }
            }
            double RastT = 40;
            double Ang = 0;
            double Ang1 = 0;
            double DlinStor1 = 0;
            string Visot = "";
            Point3d Point1 = new Point3d();
            Point3d Point2 = new Point3d();
            if (SpPoint.Count > 2)
            {
                TkoorPL11 = FTkoorPL(TkoorPL1);
                TkoorPL22 = FTkoorPL(TkoorPL2);
            }
            else
            {
                TkoorPL11 = TkoorPL1;
                TkoorPL22 = TkoorPL2;
            }
            for (int i = 0; i < TkoorPL11.Count - 1; i++) DlinStor1 = DlinStor1 + TkoorPL11[i].DistanceTo(TkoorPL11[i + 1]);
            if (Razdel == "Лес" & !RekreXvos)
            {
                for (double RastTT = 40; RastTT < Dlin; RastTT = RastTT + HSag_Hvost)
                {
                    Point_Ang(ref Point1, ref Ang, TkoorPL11, RastTT, SpPoint, ref Visot, ref Ang1);
                    Point_Ang(ref Point2, ref Ang, TkoorPL22, RastTT, SpPoint, ref Visot, ref Ang1);
                    FXvost(Point1, "", InD + "-" + ID, Ang, Visot, POM, Xvost, "");
                    FXvost(Point2, "", InD + "-" + ID, Ang1, Visot, POM, Xvost, "");
                }
                Point_Ang(ref Point1, ref Ang, TkoorPL11, DlinStor1 - 40, SpPoint, ref Visot, ref Ang1);
                Point_Ang(ref Point2, ref Ang, TkoorPL22, DlinStor1 - 40, SpPoint, ref Visot, ref Ang1);
                FXvost(Point1, "", InD + "-" + ID, Ang, Visot, POM, Xvost, "");
                FXvost(Point2, "", InD + "-" + ID, Ang1, Visot, POM, Xvost, "");
            }
            Zagib(TkoorPL11, TkoorPL22, SpPoint, InD + "-" + ID);
            for (int i = TkoorPL22.Count - 1; i >= 0; i--) TkoorPL11.Add(TkoorPL22[i]);
            //FPoly(TkoorPL11, SpPoint, COMPON, POM, InD, Razdel, InD + "-" + ID);
            if (Razdel != "Тр")
                FPoly(TkoorPL11, SpPoint, COMPON, POM, InD, Razdel, InD + "-" + ID);
            else
                FPoly(TkoorPL11, SpPoint, COMPON, POM, InD, Razdel, InD);
        }///определение координат для построения лесницы полилинией
        public static Point3d polar(Point3d XYZ1, double ugolX, double ugolY, double Rast)
        {
            double X = Rast * Math.Cos(ugolX);
            double Y = Rast * Math.Sin(ugolX);
            Point3d XYZ2 = new Point3d(XYZ1.X + X, XYZ1.Y + Y, XYZ1.Z);
            return XYZ2;
        }//поиск точки отстаящей от заданной на растояние и угол заданный 
        public static Point3d polar1(Point3d XYZ1, Vector3d Vek, double Rast)
        {
            double X = 0;
            double Y = 0;
            double alfX = Vek.GetAngleTo(new Vector3d(1, 0, 0));
            double alfY = Vek.GetAngleTo(new Vector3d(0, 1, 0));
            //Application.ShowAlertDialog(Vek.ToString() + " alfX=" + alfX.ToString() + " alfY=" + alfY.ToString());
            X = Rast * Math.Cos(alfX);
            Y = Rast * Math.Sin(alfX);
            Point3d XYZ2 = new Point3d(XYZ1.X + X, XYZ1.Y + Y, XYZ1.Z);
            if (alfY > Math.PI / 2) XYZ2 = new Point3d(XYZ1.X + X, XYZ1.Y - Y, XYZ1.Z);
            return XYZ2;
        }//поиск точки отстаящей от заданной на растояние и угол заданный 
        public static void KrugIText(TPosrL TPostr, List<TPosrL> SpPoint, string ID, string Name)
        {
            string Nom = "0";
            string blkName = "CWPoint";
            if (SpPoint.Count > 0) Nom = SpPoint.Count.ToString();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            EdI.Editor ed = doc.Editor;
            EdI.PromptPointOptions pPtOpts;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                if (!acBlkTbl.Has(blkName))
                {
                    Circle KrugPOZ = new Circle();
                    KrugPOZ.SetDatabaseDefaults();
                    KrugPOZ.Center = TPostr.TPoint;
                    KrugPOZ.Radius = 10;
                    KrugPOZ.Layer = "0";
                    acBlkTblRec.AppendEntity(KrugPOZ);
                    Tx.AddNewlyCreatedDBObject(KrugPOZ, true);
                }
                Tx.Commit();
            }
            Creat_BL(TPostr.TPoint, ID, TPostr.Visot, Name, "CWPoint", "CWAYPoint", 0, "", "");
        }
        public static void Creat_BL(Point3d BazeP, string ID, string Visot, string Name, string blkName, string Sloi, double Ugol, string Pom, string strDobav)
        {
            //this.Hide();
            //string blkName = "CWPoint";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //Point3d BazeP = new Point3d();
            using (DocumentLock docLock = doc.LockDocument())
            {
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(blkName))
                    {
                        //ObjectId BlkId = db.Insert(blkName, db, false);
                        InsertBlock(BazeP, blkName);
                        DEFAtr_Cr_Bl(blkName);
                        if (blkName.Contains("Уголок") == true)
                            SetDynamicBlkProperty(blkName, Visot, strDobav, Pom, "Доизол.д", "006", "хв", "", 0, "", Name, Ugol, Sloi);
                        else
                            SetDynamicBlkProperty(ID, Visot, "", "", "", "", "", "", 0, "", Name, Ugol, Sloi);
                        tr.Commit();
                        return;
                    }
                    // Create our new block table record...
                    BlockTableRecord btr = new BlockTableRecord();
                    // ... and set its properties
                    btr.Name = blkName;
                    // Add the new block to the block table
                    bt.UpgradeOpen();
                    ObjectId btrId = bt.Add(btr);
                    tr.AddNewlyCreatedDBObject(btr, true);
                    PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
                    if (acSSPrompt.Status != PromptStatus.OK) return;
                    SelectionSet acSSet = acSSPrompt.Value;
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        // Check to make sure a valid SelectedObject object was returned
                        if (acSSObj != null)
                        {
                            // Open the selected object for write
                            Entity acEnt = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                            //BazeP = acEnt.GeometricExtents.MaxPoint;
                            Entity acEnt_Cl = acEnt.Clone() as Entity;
                            btr.AppendEntity(acEnt_Cl);
                            tr.AddNewlyCreatedDBObject(acEnt_Cl, true);
                        }
                    }
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        // Check to make sure a valid SelectedObject object was returned
                        if (acSSObj != null)
                        {
                            // Open the selected object for write
                            Entity acEnt = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                            acEnt.Erase();
                        }
                    }
                    btr.Origin = BazeP;
                    bt.UpgradeOpen();
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    BlockReference br = new BlockReference(BazeP, btrId);
                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    ADDAtr_Cr_Bl(blkName);
                    if (blkName.Contains("Уголок") == true)
                        SetDynamicBlkProperty(blkName, Visot, strDobav, Pom, "Доизол.д", "006", "хв", "", 0, "", Name, Ugol, Sloi);
                    else
                        SetDynamicBlkProperty(ID, Visot, "", "", "", "", "", "", 0, "", Name, Ugol, Sloi);
                    tr.Commit();
                }
            }
            //this.Show();
        }//создание блока из последнено созданного примитива
        public static void FPoly(Point3dCollection TkoorPL, List<TPosrL> SpPoint, string COMPON, string POM, string ID, string Razdel, string LINK)
        {
            double DlinL = 0;
            Point3d Tt = new Point3d(SpPoint[0].TPoint.X, SpPoint[0].TPoint.Y, Convert.ToDouble(SpPoint[0].Visot));
            foreach (TPosrL tPostr in SpPoint)
            {
                DlinL = DlinL + Tt.DistanceTo(new Point3d(tPostr.TPoint.X, tPostr.TPoint.Y, Convert.ToDouble(tPostr.Visot)));
                Tt = new Point3d(tPostr.TPoint.X, tPostr.TPoint.Y, Convert.ToDouble(tPostr.Visot));
            }
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            int CollorInd = 0;
            // Append the point to the database
            using (tr1)
            {
                Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = CollorInd;
                poly.Closed = true;
                poly.Layer = "Насыщение";
                //poly.Close();
                int i = 0;
                foreach (Point3d pt in TkoorPL)
                {
                    poly.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    i = i + 1;
                }
                poly.XData = new ResultBuffer
                    (
                    new TypedValue(1001, "LAUNCH01"),
                    new TypedValue(1000, COMPON),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, POM),
                    new TypedValue(1000, Razdel),
                    new TypedValue(1000, "006"),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, DlinL.ToString("##.")),
                    new TypedValue(1000, ID),
                    new TypedValue(1000, LINK)
                    );
                //Application.ShowAlertDialog(DlinL.ToString());
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);
                btr.Dispose();
                tr1.Commit();
            }
        }//построение полилинии по списку координат
        public static Point3dCollection FTkoorPL(Point3dCollection TkoorTL)
        {
            Point3dCollection NTkoorPL = new Point3dCollection();
            NTkoorPL.Add(TkoorTL[0]);
            for (int i = 0; i <= TkoorTL.Count - 4; i = i + 2)
            {
                Point3d T1 = TkoorTL[i];
                Point3d T2 = TkoorTL[i + 1];
                Point3d T3 = TkoorTL[i + 2];
                Point3d T4 = TkoorTL[i + 3];
                Point3d Tper = TPer(T1, T2, T3, T4);
                if (T2 == T3) Tper = T2;
                NTkoorPL.Add(Tper);
            }
            NTkoorPL.Add(TkoorTL[TkoorTL.Count - 1]);
            return NTkoorPL;
        }//пересчет координат
        public Point3dCollection CoorKontur(Point3d T1, Point3d T2)
        {
            Point3dCollection kontur = new Point3dCollection();
            kontur.Add(T1);
            kontur.Add(new Point3d(T1.X, T2.Y, 0));
            kontur.Add(T2);
            kontur.Add(new Point3d(T2.X, T1.Y, 0));
            kontur.Add(T1);
            return kontur; }//контур по двум точкам
        public List<TPosrL> DopT(Point3dCollection CoorKontur, List<TPosrL> SpPoint,double Perepad, List<TPosrL> SpPoint_Poln, List<TPosrL> SpPoint1L2)
        {
            List<TPosrL> SpDopT = new List<TPosrL>();
            for (int i = 0; i <= SpPoint_Poln.Count - 2; i++)
            {
                Point3d Tt1 = SpPoint_Poln[i].TPoint;
                Point3d Tt2 = SpPoint_Poln[i + 1].TPoint;
                for (int j = 0; j <= CoorKontur.Count - 2; j++)
                {
                    Point3d Tk1 = CoorKontur[j];
                    Point3d Tk2 = CoorKontur[j + 1];
                    Point3d Tper = TPer1(Tt1, Tt2, Tk1, Tk2);
                    if (Tper != new Point3d())
                    {
                        TPosrL nT = new TPosrL();
                        nT.NNomT(SpPoint_Poln[i].NomT + Tt1.DistanceTo(Tper) / Tt1.DistanceTo(Tt2));
                        nT.NTPoint(Tper);
                        nT.NVisot((Convert.ToDouble(SpPoint_Poln[i].Visot) - Perepad).ToString());
                            if (SpPoint1L2.Count > 0) 
                            {
                            double Rast = 999999999999;
                            string Visot="";
                            foreach (TPosrL TT in SpPoint1L2) if (TT.TPoint.DistanceTo(Tper) < Rast) { Rast = TT.TPoint.DistanceTo(Tper);Visot = TT.Visot;}
                            nT.NVisot(Visot);
                            }
                        if(nT.NomT> SpPoint[0].NomT & nT.NomT < SpPoint.Last().NomT) SpDopT.Add(nT);
                    }
                }
            }
            return SpDopT;
        }//Точки перересечения секущего контура с линией трассы
        public static Point3d TPer(Point3d T11, Point3d T12, Point3d T21, Point3d T22)
        {
            Point3d Tperes = new Point3d();
            double k1 = (T12.Y - T11.Y) / (T12.X - T11.X);
            double k2 = (T22.Y - T21.Y) / (T22.X - T21.X);
            double b1 = T12.Y - k1 * T12.X;
            double b2 = T22.Y - k2 * T22.X;
            double X = 0;
            double Y = 0;
            //Application.ShowAlertDialog("K1=" + k1.ToString() + " K2=" + k2.ToString());
            if ((k1 - k2) != 0)
            {
                X = (b2 - b1) / (k1 - k2);
                Y = k1 * X + b1;
                if (Math.Abs(T12.X - T11.X) < 0.5) { X = T12.X; Y = k2 * X + b2; }
                if (Math.Abs(T22.X - T21.X) < 0.5) { X = T22.X; Y = k1 * X + b1; }
                Tperes = new Point3d(X, Y, 0);
            }
            return Tperes;
        }//Определение точки пересечения отрезки продливаются
        public static Point3d TPer1(Point3d T11, Point3d T12, Point3d T21, Point3d T22)
        {
            double DistA_B = 0;
            double DistA_C = 0;
            double DistC_B = 0;
            double DistA1_B1 = 0;
            double DistA1_C1 = 0;
            double DistC1_B1 = 0;
            double DELT = 0;
            double DELT1 = 0;
            Point3d Tperes = new Point3d();
            double k1 = (T12.Y - T11.Y) / (T12.X - T11.X);
            double k2 = (T22.Y - T21.Y) / (T22.X - T21.X);
            double b1 = T12.Y - k1 * T12.X;
            double b2 = T22.Y - k2 * T22.X;
            double X = 0;
            double Y = 0;
            if ((k1 - k2) != 0)
            {
                X = (b2 - b1) / (k1 - k2);
                Y = k1 * X + b1;
                if (Math.Abs(T12.X - T11.X) < 0.5) { X = T12.X; Y = k2 * X + b2; }
                if (Math.Abs(T22.X - T21.X) < 0.5) { X = T22.X; Y = k1 * X + b1; }
                //if (Math.Abs(T22.X - T21.X) < 0.5) { X = T12.X; Y = k2 * X + b2; }
                Point3d VTperes = new Point3d(X, Y, 0);
                DistA_B = T11.DistanceTo(T12);
                DistA_C = T11.DistanceTo(VTperes);
                DistC_B = VTperes.DistanceTo(T12);
                DistA1_B1 = T21.DistanceTo(T22);
                DistA1_C1 = T21.DistanceTo(VTperes);
                DistC1_B1 = VTperes.DistanceTo(T22);
                DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                DELT1 = Math.Abs((DistA1_C1 + DistC1_B1) - DistA1_B1);
                //Application.ShowAlertDialog(VTperes.ToString() + " T12=" + T11.ToString() + " T22=" + T12.ToString() + " T21=" + T21.ToString() + " T22=" + T22.ToString());
                if (DELT < 0.5 & DELT1 < 0.5)
                { Tperes = new Point3d(X, Y, 0); }
            }
            return Tperes;
        }//Определение точки пересечения отрезки не продливаются 
        public Point3d TPerOtr(Point3d T11, Point3d T12, Point3d T21, Point3d T22)
        {
            Point3d Tperes1 = new Point3d();
            Point3d Tperes = new Point3d();
            double k1 = (T12.Y - T11.Y) / (T12.X - T11.X);
            double k2 = (T22.Y - T21.Y) / (T22.X - T21.X);
            double b1 = T12.Y - k1 * T12.X;
            double b2 = T22.Y - k2 * T22.X;
            double X = 0;
            double Y = 0;
            //Application.ShowAlertDialog("K1=" + k1.ToString() + " K2=" + k2.ToString());
            if ((k1 - k2) != 0)
            {
                X = (b2 - b1) / (k1 - k2);
                Y = k1 * X + b1;
                if (Math.Abs(T12.X - T11.X) < 0.5) { X = T12.X; Y = k2 * X + b2; }
                if (Math.Abs(T22.X - T21.X) < 0.5) { X = T22.X; Y = k1 * X + b1; }
                Tperes = new Point3d(X, Y, 0);
                double AB = T11.DistanceTo(T12);
                double AC = T11.DistanceTo(Tperes);
                double CB = T12.DistanceTo(Tperes);
                if (Math.Abs(AB - (AC + CB))<0.5) Tperes1 = Tperes;
            }
            //Application.ShowAlertDialog(Tperes1.ToString() + " " + Tperes.ToString());
            return Tperes1;
        }//Определение точки пересечения
        static public void SBORBl_FIND(ref List<TPosrL> SpOBJID, string FIND_N,ref List<ObjectId> SpID, ref List<TPosrL> SpOBJID_Poln, ref List<Prepad> SpPREP)
        {
            int Schet = 0;
            string stKomp = "";
            string stRez = "";
            string stRad = "";
            string stHtoEto = "";
            string LINK = "";
            string stTip = "";
            string strTipSvOb = "";
            string stPOLnNaim = "";
            string stDobav = "";
            string dVisot = "";
            Point3d PointPostr = new Point3d();
            double NomT = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[11];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 1);
            acTypValAr.SetValue(new TypedValue(0, "Circle"), 2);
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 3);
            acTypValAr.SetValue(new TypedValue(0, "LINE"), 4);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 6);
            acTypValAr.SetValue(new TypedValue(8, "CWAYPoint"), 7);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 8);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 9);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 10);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            //PromptSelectionOptions pso = new PromptSelectionOptions();
            //pso.
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stRez = "";
                    stRad = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stDobav = "";
                    stTip = "";
                    LINK = "";
                    dVisot = "0";
                    NomT = 0;
                    Schet = 0;
                    Entity Obj = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    //Application.ShowAlertDialog(Obj.GetType().ToString());
                    if (Obj.GetType()==typeof(BlockReference))
                    {
                        BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        stKomp = bref.Name;
                        PointPostr = bref.Position;
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "Исполнение") { if (Double.TryParse(atrRef.TextString, out NomT)) NomT = Convert.ToDouble(atrRef.TextString); }
                                    if (atrRef.Tag == "Ссылка") { stHtoEto = atrRef.TextString; }
                                    if (atrRef.Tag == "Высота_установки") { dVisot = atrRef.TextString; }
                                    if (atrRef.Tag == "ДопДетали") { stDobav = atrRef.TextString; }
                                    if (atrRef.Tag == "Что_это") { stTip = atrRef.TextString; }
                                }
                            }
                        }
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = value.Value.ToString(); }
                                if (Schet == 8) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 10) { LINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                            if (stHtoEto == FIND_N & stKomp == "CWPoint")
                            {
                                TPosrL TPostr = new TPosrL();
                                TPostr.NKomp(stKomp);
                                TPostr.NNomT(NomT);
                                TPostr.NVisot(dVisot);
                                TPostr.NTPoint(PointPostr);
                                TPostr.NTip("Точка построения");
                                SpOBJID.Add(TPostr);
                                SpID.Add(acSSObj.ObjectId);
                            }
                            if ((stHtoEto == FIND_N | stHtoEto.Contains(FIND_N + "-")) & stKomp != "CWPoint") { SpID.Add(acSSObj.ObjectId); }
                            if (stHtoEto == FIND_N & stKomp != "")
                            {
                            TPosrL TPostr1 = new TPosrL();
                            TPostr1.NKomp(stKomp);
                            TPostr1.NNomT(NomT);
                            TPostr1.NVisot(dVisot);
                            TPostr1.NDopDet(stDobav);
                            TPostr1.NTPoint(PointPostr);
                            TPostr1.NTip(stTip);
                            TPostr1.NOId(acSSObj.ObjectId);
                            if (stKomp == "CWPoint") TPostr1.NTip("Точка построения");
                            SpOBJID_Poln.Add(TPostr1);
                            }
                            if (stHtoEto.Contains(FIND_N + "-") & stKomp != "")
                            {
                            TPosrL TPostr1 = new TPosrL();
                            TPostr1.NKomp(stKomp);
                            TPostr1.NNomT(NomT);
                            TPostr1.NVisot(dVisot);
                            TPostr1.NTPoint(PointPostr);
                            TPostr1.NTip(stTip);
                            TPostr1.NLINK(stHtoEto);
                            TPostr1.NOId(acSSObj.ObjectId);
                            if (stKomp == "CWPoint") TPostr1.NTip("Точка построения");
                            SpOBJID_Poln.Add(TPostr1);
                            }
                    }
                    if (Obj.GetType() == typeof(Circle))
                    {
                        Circle bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Circle;
                        PointPostr = bref.Center;
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = value.Value.ToString(); }
                                if (Schet == 8) { stHtoEto = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                        if (stHtoEto == FIND_N & stKomp != "")
                        {
                            TPosrL TPostr = new TPosrL();
                            string[] stKompM = stKomp.Split('-');
                            TPostr.NKomp(stKompM.Last());
                            TPostr.NNomT(NomT);
                            TPostr.NVisot(dVisot);
                            TPostr.NTPoint(PointPostr);
                            TPostr.NOId(acSSObj.ObjectId);
                            SpOBJID.Add(TPostr);
                            SpID.Add(acSSObj.ObjectId);
                            SpOBJID_Poln.Add(TPostr);
                        }
                    }
                    if (Obj.GetType() == typeof(Polyline))
                    {
                        Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 10) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 11) { LINK = value.Value.ToString(); }
                                if (Schet == 9) { dVisot = value.Value.ToString(); }
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                        if (stHtoEto == FIND_N & stKomp != "")
                        {
                            TPosrL TPostr = new TPosrL();
                            string[] stKompM = stKomp.Split('-');
                            TPostr.NKomp(stKompM.Last());
                            TPostr.NNomT(NomT);
                            TPostr.NVisot(dVisot);
                            TPostr.NTPoint(PointPostr);
                            TPostr.NLINK(LINK);
                            TPostr.NTip("часть");
                            TPostr.NOId(acSSObj.ObjectId);
                            //SpOBJID.Add(TPostr);
                            SpID.Add(acSSObj.ObjectId);
                            SpOBJID_Poln.Add(TPostr);
                        }
                    }
                    if (Obj.GetType() == typeof(Line))
                    {
                        Line bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Line;
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = value.Value.ToString(); }
                                if (Schet == 4) { stRez = value.Value.ToString(); }
                                if (Schet == 5) { stRad = value.Value.ToString(); }
                                if (Schet == 11) { LINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                        if (LINK.Contains(FIND_N + "-"))
                        {
                            TPosrL TPostr = new TPosrL();
                            TPostr.NKomp("Линия гиба");
                            TPostr.NNomT(0);
                            TPostr.NVisot(dVisot);
                            //TPostr.NTPoint(PointPostr);
                            TPostr.NLINK(LINK);
                            TPostr.NTip("Линия гиба");
                            TPostr.NOId(acSSObj.ObjectId);
                            //SpOBJID.Add(TPostr);
                            SpID.Add(acSSObj.ObjectId);
                            SpOBJID_Poln.Add(TPostr);
                        }
                        if (stKomp == "Препядствие") 
                        {
                            Prepad NPerep = new Prepad();
                            NPerep.NTip(stRez);
                            NPerep.NRad(Convert.ToDouble(stRad));
                            NPerep.NVisot(Convert.ToDouble(dVisot));
                            NPerep.NT1(bref.StartPoint);
                            NPerep.NT2(bref.EndPoint);
                            SpPREP.Add(NPerep);
                        }
                    }
                }
                Tx.Commit();
            }
            SpOBJID.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
        }//поиск блоков по компоненту
        static public void SBORnBl_FIND(ref List<TPosrL> SpOBJID, string FIND_N, ref List<ObjectId> SpID)
        {
            string stKomp = "";
            string stHtoEto = "";
            string stPOLnNaim = "";
            string dVisot = "";
            Point3d PointPostr = new Point3d();
            int Schet;
            double NomT = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            //PromptSelectionOptions pso = new PromptSelectionOptions();
            //pso.

            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    dVisot = "0";
                    NomT = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    PointPostr = bref.Position;
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "Исполнение") { if (Double.TryParse(atrRef.TextString, out NomT)) NomT = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Ссылка") { stHtoEto = atrRef.TextString; }
                                if (atrRef.Tag == "Высота_установки") { dVisot = atrRef.TextString; }
                            }
                        }
                    }
                    if (stHtoEto == FIND_N & stKomp == "CWPoint")
                    {
                        TPosrL TPostr = new TPosrL();
                        TPostr.NNomT(NomT);
                        TPostr.NVisot(dVisot);
                        TPostr.NTPoint(PointPostr);
                        SpOBJID.Add(TPostr);
                        SpID.Add(acSSObj.ObjectId);
                    }
                    if (stHtoEto == FIND_N & stKomp != "CWPoint")
                        SpID.Add(acSSObj.ObjectId);
                }
                Tx.Commit();
            }
            SpOBJID.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
        }//поиск блоков по компоненту
        public static void Point_Ang(ref Point3d Poin, ref double Ang, Point3dCollection TkoorPL1, double Rast, List<TPosrL> SpPoint, ref string Visot, ref double Ang2)
        {
            double TDistLest = 0;
            double TDist = 0;
            double TDistOtr = 0;
            Vector3d VekT = new Vector3d();
            double TAngX = 0;
            double TAngY = 0;
            for (int i = 0; i < TkoorPL1.Count - 1; i++)
            {
                TDistLest = 0;
                TDistOtr = TkoorPL1[i].DistanceTo(TkoorPL1[i + 1]);
                VekT = TkoorPL1[i].GetVectorTo(TkoorPL1[i + 1]);
                TAngX = VekT.GetAngleTo(new Vector3d(1, 0, 0));
                TAngY = VekT.GetAngleTo(new Vector3d(0, 1, 0));
                if (Rast > TDist & Rast < TDist + TDistOtr)
                {
                    for (int i1 = 0; i1 < SpPoint.Count - 1; i1++)
                    {
                        if (Rast > TDistLest & Rast < TDistLest + SpPoint[i1].TPoint.DistanceTo(SpPoint[i1 + 1].TPoint))
                        //{ Visot = SpPoint[i1].Visot; Application.ShowAlertDialog("Visot-" + Visot + "Rast-" + Rast.ToString()  + "TDistLest -" + TDistLest.ToString() + "TDistLest+1" + (TDistLest + SpPoint[i1].TPoint.DistanceTo(SpPoint[i1 + 1].TPoint)).ToString()); }
                        { Visot = SpPoint[i1].Visot; }
                        TDistLest = TDistLest + SpPoint[i1].TPoint.DistanceTo(SpPoint[i1 + 1].TPoint);
                    }
                    //Visot = SpPoint[i].Visot;
                    Poin = polar1(TkoorPL1[i], VekT, Rast - TDist);
                    Ang = TAngX;
                    if (VekT.X >= 0 & VekT.Y >= 0) { Ang = TAngX; Ang2 = Ang + Math.PI / 2; };
                    if (VekT.X >= 0 & VekT.Y < 0) { Ang = VekT.GetAngleTo(new Vector3d(0, -1, 0)) + Math.PI; Ang2 = Ang - Math.PI / 2; };
                    if (VekT.X < 0 & VekT.Y <= 0) { Ang = VekT.GetAngleTo(new Vector3d(-1, 0, 0)); Ang2 = Ang + Math.PI / 2; };
                    if (VekT.X < 0 & VekT.Y > 0) { Ang = TAngX + Math.PI / 2; Ang2 = Ang - Math.PI / 2; }
                    return;
                }
                TDist = TDist + TDistOtr;
            }
        }//угол в точке кривой 
        public static void FXvost(Point3d Nkoor, string ID, string Name, double Ugol, string Visot, string Pom, string Xvost, string DopDet)
        {
            string Nom = "0";
            string[] XvostM = Xvost.Split('#');
            string[] ABC = XvostM[0].Split('*');
            string blkName = XvostM[1] + "#" + XvostM[2];
            //if (SpPoint.Count > 0) Nom = SpPoint.Count.ToString();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            EdI.Editor ed = doc.Editor;
            EdI.PromptPointOptions pPtOpts;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                //Point3d Nkoor = TPostr.TPoint;
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                if (!acBlkTbl.Has(blkName))
                {
                    Polyline poly = new Polyline();
                    poly.SetDatabaseDefaults();
                    poly.Closed = true;
                    poly.Layer = "0";
                    poly.AddVertexAt(0, new Point2d(Nkoor.X, Nkoor.Y), 0, 0, 0);
                    poly.AddVertexAt(0, new Point2d(Nkoor.X, Nkoor.Y + Convert.ToDouble(ABC[0])), 0, 0, 0);
                    poly.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[2]), Nkoor.Y + Convert.ToDouble(ABC[0])), 0, 0, 0);
                    poly.AddVertexAt(0, new Point2d(Nkoor.X - 4, Nkoor.Y + Convert.ToDouble(ABC[2])), 0, 0, 0);
                    poly.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[1]), Nkoor.Y + Convert.ToDouble(ABC[2])), 0, 0, 0);
                    poly.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[1]), Nkoor.Y), 0, 0, 0);
                    acBlkTblRec.AppendEntity(poly);
                    Tx.AddNewlyCreatedDBObject(poly, true);
                    acBlkTblRec.Dispose();
                }
                Tx.Commit();
            }
            Creat_BL(Nkoor, ID, Visot, Name, blkName, "Насыщение", Ugol, Pom, DopDet);
        }//отрисовка хвостовиков
        public void FSoedin(Point3d Nkoor, string ID, string Name, double Ugol, string Visot, double Hscir, string Pom, string LINK, string Razdel, string Dobav)
        {
            //if (SpPLOS.Count == 0) return;
            string Nom = "0";
            string[] XvostM = Name.Split('#');
            string[] ABC = XvostM[0].Split('*');
            string blkName = XvostM[1] + "#" + XvostM[2];
            //if (SpPoint.Count > 0) Nom = SpPoint.Count.ToString();
            List<ObjectId> SpPrim = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            EdI.Editor ed = doc.Editor;
            EdI.PromptPointOptions pPtOpts;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                //Point3d Nkoor = TPostr.TPoint;
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                if (!acBlkTbl.Has(blkName))
                {
                    if (blkName.Contains("Муфта"))
                    {
                        Polyline poly1 = new Polyline();
                        poly1.SetDatabaseDefaults();
                        poly1.Layer = "0";
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[0]) / 2, Nkoor.Y + Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[0]) / 2, Nkoor.Y - Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        acBlkTblRec.AppendEntity(poly1);
                        Tx.AddNewlyCreatedDBObject(poly1, true);

                        Polyline poly2 = new Polyline();
                        poly2.SetDatabaseDefaults();
                        poly2.Layer = "0";
                        poly2.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[0]) / 2, Nkoor.Y + Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        poly2.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[0]) / 2, Nkoor.Y - Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        acBlkTblRec.AppendEntity(poly2);
                        Tx.AddNewlyCreatedDBObject(poly2, true);
                        acBlkTblRec.Dispose();

                        SpPrim.Add(poly1.ObjectId);
                        SpPrim.Add(poly2.ObjectId);
                    }
                    if (blkName.Contains("Штуцер"))
                    {
                        Polyline poly1 = new Polyline();
                        poly1.SetDatabaseDefaults();
                        poly1.Layer = "0";
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[0]) / 2, Nkoor.Y - Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[0]) / 2, Nkoor.Y), 0, 0, 0);
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[0]) / 2, Nkoor.Y), 0, 0, 0);
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[0]) / 2, Nkoor.Y - Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        acBlkTblRec.AppendEntity(poly1);
                        Tx.AddNewlyCreatedDBObject(poly1, true);

                        Polyline poly2 = new Polyline();
                        poly2.SetDatabaseDefaults();
                        poly2.Layer = "0";
                        poly2.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[1]) / 2, Nkoor.Y - Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        poly2.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[1]) / 2, Nkoor.Y - Convert.ToDouble(ABC[1]) / 2), 0, 0, 0);
                        acBlkTblRec.AppendEntity(poly2);
                        Tx.AddNewlyCreatedDBObject(poly2, true);
                        acBlkTblRec.Dispose();

                        SpPrim.Add(poly1.ObjectId);
                        SpPrim.Add(poly2.ObjectId);
                    }
                    if (blkName.Contains("Фланец"))
                    {
                        Polyline poly1 = new Polyline();
                        poly1.SetDatabaseDefaults();
                        poly1.Layer = "0";
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[1]) / 2, Nkoor.Y - Convert.ToDouble(ABC[0]) / 2), 0, 0, 0);
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[1]) / 2, Nkoor.Y - Convert.ToDouble(ABC[0]) / 2), 0, 0, 0);
                        acBlkTblRec.AppendEntity(poly1);
                        Tx.AddNewlyCreatedDBObject(poly1, true);

                        Polyline poly2 = new Polyline();
                        poly2.SetDatabaseDefaults();
                        poly2.Layer = "0";
                        poly2.AddVertexAt(0, new Point2d(Nkoor.X + Convert.ToDouble(ABC[1]) / 2, Nkoor.Y + Convert.ToDouble(ABC[0]) / 2), 0, 0, 0);
                        poly2.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[1]) / 2, Nkoor.Y + Convert.ToDouble(ABC[0]) / 2), 0, 0, 0);
                        acBlkTblRec.AppendEntity(poly2);
                        Tx.AddNewlyCreatedDBObject(poly2, true);
                        acBlkTblRec.Dispose();

                        SpPrim.Add(poly1.ObjectId);
                        SpPrim.Add(poly2.ObjectId);
                    }

                }
                Tx.Commit();

            }
            Creat_BL_SpOID(Nkoor, ID, Visot, LINK, blkName, "Насыщение", Ugol, SpPrim, Pom, Razdel, blkName, Dobav);
        }//отрисовка поворотов 90 граусов у лестниц
        public static void FPovorot(Point3d Nkoor, string ID, string Name, double Ugol, string Visot, double Hscir, string Pom,string LINK)
        {
            string Nom = "0";
            string blkName = Name;
            //if (SpPoint.Count > 0) Nom = SpPoint.Count.ToString();
            List<ObjectId> SpPrim = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            EdI.Editor ed = doc.Editor;
            EdI.PromptPointOptions pPtOpts;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                //Point3d Nkoor = TPostr.TPoint;
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                if (!acBlkTbl.Has(blkName))
                {
                    Arc Ark1 = new Arc(new Point3d(Nkoor.X + 250 + Hscir / 2, Nkoor.Y - 250 - Hscir / 2, 0), 250, Math.PI / 2, Math.PI);
                    acBlkTblRec.AppendEntity(Ark1);
                    Tx.AddNewlyCreatedDBObject(Ark1, true);

                    Arc Ark2 = new Arc(new Point3d(Nkoor.X + 250 + Hscir / 2, Nkoor.Y - 250 - Hscir / 2, 0), 250 + Hscir, Math.PI / 2, Math.PI);
                    acBlkTblRec.AppendEntity(Ark2);
                    Tx.AddNewlyCreatedDBObject(Ark2, true);

                    Polyline poly1 = new Polyline();
                    poly1.SetDatabaseDefaults();
                    poly1.Layer = "0";
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X + Hscir / 2, Nkoor.Y - 250 - Hscir / 2), 0, 0, 0);
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X + Hscir / 2, Nkoor.Y - 400 - Hscir / 2), 0, 0, 0);
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X - Hscir / 2, Nkoor.Y - 400 - Hscir / 2), 0, 0, 0);
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X - Hscir / 2, Nkoor.Y - 250 - Hscir / 2), 0, 0, 0);
                    acBlkTblRec.AppendEntity(poly1);
                    Tx.AddNewlyCreatedDBObject(poly1, true);

                    Polyline poly2 = new Polyline();
                    poly2.SetDatabaseDefaults();
                    poly2.Layer = "0";
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 250 + Hscir / 2, Nkoor.Y + Hscir / 2), 0, 0, 0);
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 400 + Hscir / 2, Nkoor.Y + Hscir / 2), 0, 0, 0);
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 400 + Hscir / 2, Nkoor.Y - Hscir / 2), 0, 0, 0);
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 250 + Hscir / 2, Nkoor.Y - Hscir / 2), 0, 0, 0);
                    acBlkTblRec.AppendEntity(poly2);
                    Tx.AddNewlyCreatedDBObject(poly2, true);
                    acBlkTblRec.Dispose();

                    SpPrim.Add(Ark1.ObjectId);
                    SpPrim.Add(Ark2.ObjectId);
                    SpPrim.Add(poly1.ObjectId);
                    SpPrim.Add(poly2.ObjectId);
                }
                Tx.Commit();

            }
            Creat_BL_SpOID(Nkoor, ID, Visot, LINK, blkName, "Насыщение", Ugol, SpPrim, Pom,"Лес", blkName,"");
        }//отрисовка хвостовиков
        public void FPovorot_TR(Point3d T1, Point3d T2, Point3d T3, string ID, string Name, string Visot, double Hscir, string Pom, double NdiamIZgib, double NdiamPrUHC, ref double Delt, string LINK, string Compon)
        {
            Vector3d Vek1 = T1.GetVectorTo(T2);
            Vector3d Vek2 = T2.GetVectorTo(T3);
            double Povorot = 0;
            double alfMV = Math.PI / 2 - Vek1.GetAngleTo(Vek2);
            double tan1 = 0;
            double tan2 = 0;
            double alf = 0;
            double alf1 = 0;
            double alf2 = 0;
            string blkName = Name;
            double DeltRad = Hscir * NdiamIZgib;
            double DeltPrU = Hscir * NdiamPrUHC;
            Point3d Tper1 = new Point3d();
            Point3d Tper2 = new Point3d();
            tan1 = (T2.Y - T1.Y) / (T2.X - T1.X);
            alf1 = Math.Atan(tan1) + Math.PI / 2;
            alf2 = Math.Atan(tan1) - Math.PI / 2;
            Point3d T11 = polar(T1, alf1, DeltRad);
            Point3d T21 = polar(T1, alf2, DeltRad);
            Point3d T31 = polar(T2, alf1, DeltRad);
            Point3d T41 = polar(T2, alf2, DeltRad);
            tan2 = (T3.Y - T2.Y) / (T3.X - T2.X);
            alf1 = Math.Atan(tan2) + Math.PI / 2;
            alf2 = Math.Atan(tan2) - Math.PI / 2;
            Point3d T12 = polar(T2, alf1, DeltRad);
            Point3d T22 = polar(T2, alf2, DeltRad);
            Point3d T32 = polar(T3, alf1, DeltRad);
            Point3d T42 = polar(T3, alf2, DeltRad);
            double dist1 = T31.DistanceTo(T12);
            double dist2 = T31.DistanceTo(T22);
            //Application.ShowAlertDialog("Vek1=" + Vek1.GetAngleTo(Vek2));
            if (dist1 < dist2)
            {
                Tper1 = TPer(T11, T31, T12, T32);
                Tper2 = TPer(T21, T41, T22, T42);
            }
            else
            {
                Tper1 = TPer(T11, T31, T22, T42);
                Tper2 = TPer(T21, T41, T12, T32);
            }
            Point3d Tper = new Point3d();
            if (Math.Abs(T11.DistanceTo(T31) - (T11.DistanceTo(Tper1) + Tper1.DistanceTo(T31))) < 0.1) { Tper = Tper1; }
            if (Math.Abs(T21.DistanceTo(T41) - (T21.DistanceTo(Tper2) + Tper2.DistanceTo(T41))) < 0.1) { Tper = Tper2; }
            if (Vek1.GetAngleTo(Vek2) == Math.PI / 2) Tper = FTper90(T1, T2, T3, DeltRad);
            List<ObjectId> SpPrim = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            EdI.Editor ed = doc.Editor;
            EdI.PromptPointOptions pPtOpts;
            double cosALF = DeltRad / T2.DistanceTo(Tper);
            Delt = T2.DistanceTo(Tper) * Math.Sqrt(1 - Math.Pow(cosALF, 2));
            Povorot = OpredUgl(Vek1, Vek2);
            alfMV = 2 * Math.Acos(cosALF);
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                if (!acBlkTbl.Has(blkName))
                {
                    Arc Ark1 = new Arc(Tper, DeltRad - Hscir / 2, 0, alfMV);
                    acBlkTblRec.AppendEntity(Ark1);
                    Tx.AddNewlyCreatedDBObject(Ark1, true);

                    Arc Ark2 = new Arc(Tper, DeltRad + Hscir / 2, 0, alfMV);
                    acBlkTblRec.AppendEntity(Ark2);
                    Tx.AddNewlyCreatedDBObject(Ark2, true);

                    SpPrim.Add(Ark1.ObjectId);
                    SpPrim.Add(Ark2.ObjectId);
                }
                Tx.Commit();
            }
            Creat_BL_SpOID(Tper, ID, (Delt * 2).ToString("0.##"), LINK, blkName, "Насыщение", Povorot, SpPrim, Pom, "Тр", Compon, "");
        }//отрисовка поворотов у труб
        public Point3d FTper90(Point3d T1, Point3d T2, Point3d T3, double DeltRad)
        {
            Point3d Tper90 = T2;
            Vector3d Vek1 = T1.GetVectorTo(T2);
            Vector3d Vek2 = T2.GetVectorTo(T3);
            if (Vek1.X > 0 & Vek1.Y == 0 & Vek2.X == 0 & Vek2.Y > 0) Tper90 = new Point3d(T2.X - DeltRad, T2.Y + DeltRad, T2.Z);
            if (Vek1.X > 0 & Vek1.Y == 0 & Vek2.X == 0 & Vek2.Y < 0) Tper90 = new Point3d(T2.X - DeltRad, T2.Y - DeltRad, T2.Z);
            if (Vek1.X < 0 & Vek1.Y == 0 & Vek2.X == 0 & Vek2.Y < 0) Tper90 = new Point3d(T2.X + DeltRad, T2.Y - DeltRad, T2.Z);
            if (Vek1.X < 0 & Vek1.Y == 0 & Vek2.X == 0 & Vek2.Y > 0) Tper90 = new Point3d(T2.X + DeltRad, T2.Y + DeltRad, T2.Z);
            if (Vek1.X == 0 & Vek1.Y > 0 & Vek2.X > 0 & Vek2.Y == 0) Tper90 = new Point3d(T2.X + DeltRad, T2.Y - DeltRad, T2.Z);
            if (Vek1.X == 0 & Vek1.Y < 0 & Vek2.X > 0 & Vek2.Y == 0) Tper90 = new Point3d(T2.X + DeltRad, T2.Y + DeltRad, T2.Z);
            if (Vek1.X == 0 & Vek1.Y > 0 & Vek2.X < 0 & Vek2.Y == 0) Tper90 = new Point3d(T2.X - DeltRad, T2.Y - DeltRad, T2.Z);
            if (Vek1.X == 0 & Vek1.Y > 0 & Vek2.X < 0 & Vek2.Y == 0) Tper90 = new Point3d(T2.X - DeltRad, T2.Y - DeltRad, T2.Z);
            if (Vek1.X == 0 & Vek1.Y < 0 & Vek2.X < 0 & Vek2.Y == 0) Tper90 = new Point3d(T2.X - DeltRad, T2.Y + DeltRad, T2.Z);
            return Tper90;
        }
        public double OpredUgl(Vector3d Vek1, Vector3d Vek2)
        {
            double Ugol = 0;
            double V1X = Vek1.GetAngleTo(new Vector3d(1, 0, 0));
            double V1Y = Vek1.GetAngleTo(new Vector3d(0, 1, 0));
            double V2X = Vek2.GetAngleTo(new Vector3d(1, 0, 0));
            double V2Y = Vek2.GetAngleTo(new Vector3d(0, 1, 0));
            if (Vek1.X >= 0 & Vek1.Y >= 0 & Vek2.X >= 0 & Vek2.Y >= 0) { if (V2X >= Math.PI / 4) Ugol = 0 - V1Y; else Ugol = Math.PI / 2 + V2X; }//1++++
            if (Vek1.X >= 0 & Vek1.Y <= 0 & Vek2.X >= 0 & Vek2.Y >= 0) Ugol = 2 * Math.PI - V1Y;//2+-++
            if (Vek1.X >= 0 & Vek1.Y >= 0 & Vek2.X >= 0 & Vek2.Y <= 0) Ugol = Math.PI / 2 - V2X;//3+++-
            if (Vek1.X >= 0 & Vek1.Y <= 0 & Vek2.X >= 0 & Vek2.Y <= 0) { if (V1Y < V2Y) Ugol = Math.PI / 2 - V2X; else Ugol = 0 - V1Y; }//4+-+-
            if (Vek1.X <= 0 & Vek1.Y >= 0 & Vek2.X <= 0 & Vek2.Y >= 0) { if (V1Y < V2Y) Ugol = V1Y; else Ugol = Math.PI + V2Y; }//5-+-+
            if (Vek1.X <= 0 & Vek1.Y <= 0 & Vek2.X <= 0 & Vek2.Y >= 0) Ugol = Math.PI + V2Y;//6---+
            if (Vek1.X <= 0 & Vek1.Y >= 0 & Vek2.X <= 0 & Vek2.Y <= 0) Ugol = V1Y;//7-+--
            if (Vek1.X <= 0 & Vek1.Y <= 0 & Vek2.X <= 0 & Vek2.Y <= 0) { if (V1Y > V2Y) Ugol = Math.PI + V2Y; else Ugol = V1Y; }//8----
            //9--++
            //10++--
            if (Vek1.X <= 0 & Vek1.Y >= 0 & Vek2.X >= 0 & Vek2.Y >= 0) Ugol = Math.PI / 2 + V2X;//11-+++
            if (Vek1.X >= 0 & Vek1.Y >= 0 & Vek2.X <= 0 & Vek2.Y >= 0) Ugol = 0 + V1Y;//12++-+
            if (Vek1.X >= 0 & Vek1.Y <= 0 & Vek2.X <= 0 & Vek2.Y <= 0) Ugol = 0 - (Math.PI - V2Y);//13+---
            //14-++-
            //15+--+
            //16+-++
            //17+--+
            if (Vek1.X <= 0 & Vek1.Y <= 0 & Vek2.X >= 0 & Vek2.Y <= 0) Ugol = V1Y;//18--+-
            return Ugol;
        }
        public Point3d polar(Point3d XYZ1, double ugolX, double Rast)
        {
            double X = Rast * Math.Cos(ugolX);
            double Y = Rast * Math.Sin(ugolX);
            Point3d XYZ2 = new Point3d(XYZ1.X + X, XYZ1.Y + Y, XYZ1.Z);
            return XYZ2;
        }//поиск точки отстаящей от заданной на растояние и угол заданный 
        public static void Creat_BL_SpOID(Point3d BazeP, string ID, string Visot, string Name, string blkName, string Sloi, double Ugol, List<ObjectId> SpOID, string Pom, string Razd, string Compon, string Dobav)
        {
            //this.Hide();
            //string blkName = "CWPoint";
            blkName = blkName.Replace(',', '_');
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //Point3d BazeP = new Point3d();
            using (DocumentLock docLock = doc.LockDocument())
            {
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(blkName))
                    {
                        //ObjectId BlkId = db.Insert(blkName, db, false);
                        InsertBlock(BazeP, blkName);
                        DEFAtr_Cr_Bl(blkName);
                        if (blkName.Contains("Уголок"))
                            SetDynamicBlkProperty(Compon, Visot, Dobav, Pom, Razd, "006", "хв", "", 0, "", Name, Ugol, Sloi);
                        else if (blkName.Contains("ПоворотГор"))
                            SetDynamicBlkProperty(Compon, Visot, Dobav, Pom, Razd, "796", "", "", 0, "", Name, Ugol, Sloi);
                        else if (blkName.Contains("СоедТр"))
                            SetDynamicBlkProperty(Compon, Visot, Dobav, Pom, Razd, "796", "соед", "", 0, "", Name, Ugol, Sloi);
                        else if (blkName.Contains("ПоворотТр"))
                            SetDynamicBlkProperty(Compon, Visot, Dobav, Pom, Razd, "006", "Тр", "", 0, "", Name, Ugol, Sloi);
                        else
                            SetDynamicBlkProperty(ID, Visot, "", "", "", "", "", "", 0, "", Name, Ugol, Sloi);
                        tr.Commit();
                        return;
                    }
                    // Create our new block table record...
                    BlockTableRecord btr = new BlockTableRecord();
                    // ... and set its properties
                    btr.Name = blkName;
                    // Add the new block to the block table
                    bt.UpgradeOpen();
                    ObjectId btrId = bt.Add(btr);
                    tr.AddNewlyCreatedDBObject(btr, true);
                    //PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
                    //if (acSSPrompt.Status != PromptStatus.OK) return;
                    //SelectionSet acSSet = acSSPrompt.Value;
                    foreach (ObjectId acSSObj in SpOID)
                    {
                        // Check to make sure a valid SelectedObject object was returned
                        if (acSSObj != null)
                        {
                            // Open the selected object for write
                            Entity acEnt = tr.GetObject(acSSObj, OpenMode.ForWrite) as Entity;
                            //BazeP = acEnt.GeometricExtents.MaxPoint;
                            Entity acEnt_Cl = acEnt.Clone() as Entity;
                            btr.AppendEntity(acEnt_Cl);
                            tr.AddNewlyCreatedDBObject(acEnt_Cl, true);
                        }
                    }
                    foreach (ObjectId acSSObj in SpOID)
                    {
                        // Check to make sure a valid SelectedObject object was returned
                        if (acSSObj != null)
                        {
                            // Open the selected object for write
                            Entity acEnt = tr.GetObject(acSSObj, OpenMode.ForWrite) as Entity;
                            acEnt.Erase();
                        }
                    }
                    btr.Origin = BazeP;
                    bt.UpgradeOpen();
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    BlockReference br = new BlockReference(BazeP, btrId);
                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    ADDAtr_Cr_Bl(blkName);
                    if (blkName.Contains("Уголок"))
                        SetDynamicBlkProperty(Compon, Visot, "", Pom, Razd, "006", "хв", "", 0, "", Name, Ugol, Sloi);
                    else if (blkName.Contains("ПоворотГор"))
                        SetDynamicBlkProperty(Compon, Visot, "", Pom, Razd, "796", "", "", 0, "", Name, Ugol, Sloi);
                    else if (blkName.Contains("ПоворотТр"))
                        SetDynamicBlkProperty(Compon, Visot, "", Pom, Razd, "006", "Тр", "", 0, "", Name, Ugol, Sloi);
                    else if (blkName.Contains("СоедТр"))
                        SetDynamicBlkProperty(Compon, Visot, Dobav, Pom, Razd, "796", "соед", "", 0, "", Name, Ugol, Sloi);
                    else
                        SetDynamicBlkProperty(ID, Visot, "", "", "", "", "", "", 0, "", Name, Ugol, Sloi);
                    tr.Commit();
                }
            }
            //this.Show();
        }//создание блока из списка примитивов
        public void InsBlockRef_NI(string BlockPath, string Block, Point3d Tvstav, double Rot)
        {
            // Активный документ в редакторе AutoCAD
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // База данных чертежа (в данном случае - активного документа)
            Database db = doc.Database;
            // Редактор базы данных чертежа
            // Запускаем транзакцию
            //using (DocumentLock docLock = doc.LockDocument())
            //{
            using (DbS.Transaction tr = db.TransactionManager.StartTransaction())
            {
                //EdI.Editor ed = doc.Editor;
                //EdI.PromptPointOptions pPtOpts;
                //pPtOpts = new EdI.PromptPointOptions("\nУкажите точку вставки блока: ");
                //// Выбор точки пользователем

                //var pPtRes = doc.Editor.GetPoint(pPtOpts);
                //if (pPtRes.Status != EdI.PromptStatus.OK)
                //    return;
                //var ptStart = pPtRes.Value;

                DbS.BlockTable bt = tr.GetObject(db.BlockTableId, DbS.OpenMode.ForRead) as DbS.BlockTable;
                DbS.BlockTableRecord model = tr.GetObject(bt[DbS.BlockTableRecord.ModelSpace], DbS.OpenMode.ForWrite) as DbS.BlockTableRecord;
                // Создаем новую базу
                using (DbS.Database db1 = new DbS.Database(false, false))
                {
                    // Получаем базу чертежа-донора
                    db1.ReadDwgFile(BlockPath, System.IO.FileShare.Read, true, null);
                    // Получаем ID нового блока
                    DbS.ObjectId BlkId = db.Insert(Block, db1, false);
                    DbS.BlockReference bref = new DbS.BlockReference(Tvstav, BlkId);
                    bref.Rotation = Rot;
                    // Дефолтные свойства блока (слой, цвет и пр.)
                    bref.SetDatabaseDefaults();
                    // Добавляем блок в модель
                    model.AppendEntity(bref);
                    // Добавляем блок в транзакцию
                    tr.AddNewlyCreatedDBObject(bref, true);
                    // Расчленяем блок
                    bref.ExplodeToOwnerSpace();
                    bref.Erase();
                    // Закрываем транзакцию
                    tr.Commit();
                }
            }
            //}
        }//Блок из файла
        public static void Zagib(Point3dCollection TkoorTL1, Point3dCollection TkoorTL2, List<TPosrL> SpPoint, string LINK)
        {
            if (SpPoint.Count < 3) return;
            for (int i = 1; i <= SpPoint.Count - 2; i++)
            {
                if (SpPoint[i - 1].Visot != SpPoint[i + 1].Visot)
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    Database db = doc.Database;
                    using (Transaction Tx = db.TransactionManager.StartTransaction())
                    {
                        BlockTable acBlkTbl;
                        BlockTableRecord acBlkTblRec;
                        // Open Model space for write
                        acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        Line acLine = new Line(TkoorTL1[i], TkoorTL2[i]);
                        acLine.SetDatabaseDefaults();
                        acLine.ColorIndex = 1;
                        acLine.Layer = "Насыщение";
                        acLine.XData = new ResultBuffer
                        (
                        new TypedValue(1001, "LAUNCH01"),
                        new TypedValue(1000, LINK + "-" + i.ToString()),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, SpPoint[i].Visot),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, LINK)
                        );
                        acBlkTblRec.AppendEntity(acLine);
                        Tx.AddNewlyCreatedDBObject(acLine, true);
                        Tx.Commit();
                    }
                }
            }
        }//Построение линии гиба 
        public static void MestnGib(ref List<TPosrL> SpPointVH,double PramU,double Gib,double Param, double Visot, bool napr,ref List<TPosrL> SpPointPoln, double Razr) 
        {
            if (Razr > 0)
            {
                TPosrL TRazr = Tisk(napr, Param, Razr, SpPointVH, Visot, SpPointPoln);         
                if (napr)
                { SpPointVH.RemoveAt(0); 
                  SpPointVH.Insert(0, TRazr);
                  Param = TRazr.NomT;
                }
                else
                { SpPointVH.Remove(SpPointVH.Last());
                  SpPointVH.Add(TRazr);
                  Param = TRazr.NomT;
                }
                SpPointVH.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                SpPointPoln.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
            }
            if (PramU > 0)
            {
                if(Razr==0) Visot = Convert.ToDouble(SpPointVH.Find(x => x.NomT == Param).Visot);
                TPosrL Tpr = Tisk(napr, Param,  PramU, SpPointVH, Visot, SpPointPoln);
                SpPointVH.Add(Tpr);
                SpPointPoln.Add(Tpr);
                SpPointVH.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                SpPointPoln.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
            }
            if (Gib > 0)
            {
                double ParZ = Math.Truncate(Param);
                Visot = Convert.ToDouble(SpPointPoln.Find(x => x.NomT == ParZ).Visot);
                if (napr) Visot = Convert.ToDouble(SpPointPoln.Find(x => x.NomT == ParZ + 1).Visot);
                TPosrL Tgib = Tisk(napr, Param,  PramU + Gib, SpPointVH, Visot, SpPointPoln);
                SpPointVH.Add(Tgib);
                SpPointPoln.Add(Tgib);
                SpPointVH.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                SpPointPoln.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
            }
        }
        public static TPosrL Tisk(bool napr,double Param, double Delt, List<TPosrL> SpPointVH, double Visot, List<TPosrL> SpPoint_Poln) 
        {
            double ParZ = Math.Truncate(Param);
            double Dist = 0;
            Vector3d Vek2 = new Vector3d();
            Point3d Tper = SpPointVH.Find(x => x.NomT == Param).TPoint;
            int nom = SpPointVH.FindIndex(x => x.NomT == Param);
            int nom1 = SpPoint_Poln.FindIndex(x => x.NomT == ParZ);
            Dist = SpPoint_Poln.Find(x=>x.NomT== ParZ).TPoint.DistanceTo(SpPoint_Poln.Find(x => x.NomT == ParZ+1).TPoint);
            if (napr)
            {
                Vek2 = Tper.GetVectorTo(SpPointVH[nom + 1].TPoint);
            }
            else
            {
                Vek2 = Tper.GetVectorTo(SpPointVH[nom - 1].TPoint);
            }
         double ParPr = Param;
         TPosrL TkLPr = new TPosrL();
         Point3d TkPr = polar1(Tper, Vek2, Delt);
         double DeltParPr = Tper.DistanceTo(TkPr) / Dist;
                    if (napr)
                    ParPr = Param + DeltParPr;
                    else
                    ParPr = Param - DeltParPr;
         TkLPr.NNomT(ParPr);
         TkLPr.NTPoint(TkPr);
         TkLPr.NVisot(Visot.ToString());
         return TkLPr;
        }//поиск отчки на трассе
        public void Izm() 
        {
            //Application.ShowAlertDialog(ID.ToString());
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                try
                {
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                    Entity ent = (Entity)tr.GetObject(ID, OpenMode.ForWrite);
                    if (!regTable.Has("LAUNCH01"))
                    {
                        regTable.UpgradeOpen();
                        // Добавляем имя приложения, которое мы будем
                        // использовать в расширенных данных
                        RegAppTableRecord app =
                                new RegAppTableRecord();
                        app.Name = "LAUNCH01";
                        regTable.Add(app);
                        tr.AddNewlyCreatedDBObject(app, true);
                    }
                    //Добавляем расширенные данные к примитиву
                    ent.XData = new ResultBuffer(
                    new TypedValue(1001, "LAUNCH01"),
                    new TypedValue(1000, this.label14.Text),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, this.textBox4.Text),
                    new TypedValue(1000, this.textBox5.Text),
                    new TypedValue(1000, "006"),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, this.textBox3.Text),
                    new TypedValue(1000, this.label7.Text)
                    );
                    //Application.ShowAlertDialog("Должно поменятся");
                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }

        }//изменить расширеные данные
        public static void InsertBlock(Point3d insPt, string blockName)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;


            using (var tr = db.TransactionManager.StartTransaction())
            {
                // check if the block table already has the 'blockName'" block
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                // create a new block reference
                using (var br = new BlockReference(insPt, bt[blockName]))
                {
                    var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);
                    //if (acBlkTblRec.HasAttributeDefinitions)
                    //{
                    //    // Add attributes from the block table record
                    //    foreach (ObjectId objID in acBlkTblRec)
                    //    {
                    //        DBObject dbObj = Tx.GetObject(objID, OpenMode.ForRead) as DBObject;
                    //        if (dbObj is AttributeDefinition)
                    //        {
                    //            AttributeDefinition acAtt = dbObj as AttributeDefinition;
                    //            if (!acAtt.Constant)
                    //            {
                    //                using (AttributeReference acAttRef = new AttributeReference())
                    //                {
                    //                    acAttRef.SetAttributeFromBlock(acAtt, acBlkRef.BlockTransform);
                    //                    acAttRef.Position = acAtt.Position.TransformBy(acBlkRef.BlockTransform);
                    //                    acAttRef.TextString = acAtt.TextString;
                    //                    acBlkRef.AttributeCollection.AppendAttribute(acAttRef);
                    //                    Tx.AddNewlyCreatedDBObject(acAttRef, true);
                    //                }
                    //            }
                    //        }
                    //    }
                    //}

                }
                //ADDAtr_Cr_Bl(blockName);
                tr.Commit();
            }
        }//вставка блока если он есть в базе данных чертежа
        public static void DEFAtr_Cr_Bl(string Blname)
        {
            Point3d Pos = new Point3d();
            Point3d PosBL = new Point3d();
            ObjectId IDBl = ObjectId.Null;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockReference bref1;
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
                    SelectionSet acSSet = acSSPrompt.Value;
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        //BlockReference bref1 = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        bref1 = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        Pos = bref1.Position;
                        IDBl = bref1.ObjectId;
                        //bref1.Erase();
                    }
                    BlockReference bref = Tx.GetObject(IDBl, OpenMode.ForRead) as BlockReference;
                    Blname = bref.Name;
                    BlockTable acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    ObjectId blkRecId = ObjectId.Null;
                    using (BlockTableRecord acBlkTblRec = Tx.GetObject(acBlkTbl[Blname], OpenMode.ForWrite) as BlockTableRecord)
                    {
                        blkRecId = acBlkTblRec.Id;
                        PosBL = Pos;
                    }
                    //Application.ShowAlertDialog("1");
                    if (blkRecId != ObjectId.Null)
                    {
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = Tx.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;
                        // Create and insert the new block reference
                        using (BlockReference acBlkRef = new BlockReference(PosBL, blkRecId))
                        {
                            BlockTableRecord acCurSpaceBlkTblRec;
                            acCurSpaceBlkTblRec = Tx.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                            Tx.AddNewlyCreatedDBObject(acBlkRef, true);
                            // Verify block table record has attribute definitions associated with it
                            if (acBlkTblRec.HasAttributeDefinitions)
                            {
                                // Add attributes from the block table record
                                foreach (ObjectId objID in acBlkTblRec)
                                {
                                    DBObject dbObj = Tx.GetObject(objID, OpenMode.ForRead) as DBObject;
                                    if (dbObj is AttributeDefinition)
                                    {
                                        AttributeDefinition acAtt = dbObj as AttributeDefinition;
                                        if (!acAtt.Constant)
                                        {
                                            using (AttributeReference acAttRef = new AttributeReference())
                                            {
                                                acAttRef.SetAttributeFromBlock(acAtt, acBlkRef.BlockTransform);
                                                acAttRef.Position = acAtt.Position.TransformBy(acBlkRef.BlockTransform);
                                                acAttRef.TextString = acAtt.TextString;
                                                acBlkRef.AttributeCollection.AppendAttribute(acAttRef);
                                                Tx.AddNewlyCreatedDBObject(acAttRef, true);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        foreach (SelectedObject acSSObj in acSSet)
                        {
                            BlockReference bref2 = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                            bref2.Erase();
                        }
                    }
                    Tx.Commit();
                }
            }
        }//определение атрибутов у блока
        static public void SetDynamicBlkProperty(string Compon, string Visota, string StrDobav, string Pom, string RazdelSp, string KEI, string HCTOEto, string Hozain, double Dlinna, string Rab, string LINK, double Ugol, string Layer)
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
                    bref.Layer = Layer;
                    bref.Rotation = Ugol;
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "Помещение") { atrRef.TextString = Pom; }
                                if (atrRef.Tag == "Высота_установки") { atrRef.TextString = Visota; }
                                if (atrRef.Tag == "Раздел_спецификации") { atrRef.TextString = RazdelSp; }
                                if (atrRef.Tag == "Исполнение") { atrRef.TextString = Compon; }
                                if (atrRef.Tag == "КЕИ") { atrRef.TextString = KEI; }
                                if (atrRef.Tag == "Что_это") { atrRef.TextString = HCTOEto; }
                                if (atrRef.Tag == "Ссылка") { atrRef.TextString = LINK; }
                            }
                        }
                    }
                }
                Tx.Commit();
            }
        }//расширеныеасширеные данные
        public static void ADDAtr_Cr_Bl(string Blname)
        {
            Point3d Pos = new Point3d();
            Point3d PosBL = new Point3d();
            ObjectId IDBl = ObjectId.Null;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockReference bref1;
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
                    SelectionSet acSSet = acSSPrompt.Value;
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        //BlockReference bref1 = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        bref1 = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        Pos = bref1.Position;
                        IDBl = bref1.ObjectId;
                        //bref1.Erase();
                    }
                    BlockReference bref = Tx.GetObject(IDBl, OpenMode.ForRead) as BlockReference;
                    Blname = bref.Name;
                    BlockTable acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    ObjectId blkRecId = ObjectId.Null;
                    using (BlockTableRecord acBlkTblRec = Tx.GetObject(acBlkTbl[Blname], OpenMode.ForWrite) as BlockTableRecord)
                    {
#region Атрибуты
                        PosBL = Pos;
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Исполнение #: ";
                            acAttDef.Tag = "Исполнение";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Что_это #: ";
                            acAttDef.Tag = "Что_это";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Высота_установки #: ";
                            acAttDef.Tag = "Высота_установки";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            //acAttDef.Position = new Point3d(0, 39, 0);
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "КЕИ #: ";
                            acAttDef.Tag = "КЕИ";
                            acAttDef.TextString = "796";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Раздел_спецификации #: ";
                            acAttDef.Tag = "Раздел_спецификации";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Помещение #: ";
                            acAttDef.Tag = "Помещение";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Примечание #: ";
                            acAttDef.Tag = "Примечание";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "Ссылка #: ";
                            acAttDef.Tag = "Ссылка";
                            acAttDef.TextString = "";
                            acAttDef.Height = 2.5;
                            //acAttDef.Justify = AttachmentPoint.MiddleCenter;
                            acBlkTblRec.AppendEntity(acAttDef);
                            acBlkTbl.UpgradeOpen();
                        }
#endregion
                        blkRecId = acBlkTblRec.Id;
                    }
                    //Application.ShowAlertDialog("1");
                    if (blkRecId != ObjectId.Null)
                    {
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = Tx.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;
                        // Create and insert the new block reference
                        using (BlockReference acBlkRef = new BlockReference(PosBL, blkRecId))
                        {
                            BlockTableRecord acCurSpaceBlkTblRec;
                            acCurSpaceBlkTblRec = Tx.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                            Tx.AddNewlyCreatedDBObject(acBlkRef, true);
                            // Verify block table record has attribute definitions associated with it
                            if (acBlkTblRec.HasAttributeDefinitions)
                            {
                                // Add attributes from the block table record
                                foreach (ObjectId objID in acBlkTblRec)
                                {
                                    DBObject dbObj = Tx.GetObject(objID, OpenMode.ForRead) as DBObject;
                                    if (dbObj is AttributeDefinition)
                                    {
                                        AttributeDefinition acAtt = dbObj as AttributeDefinition;
                                        if (!acAtt.Constant)
                                        {
                                            using (AttributeReference acAttRef = new AttributeReference())
                                            {
                                                acAttRef.SetAttributeFromBlock(acAtt, acBlkRef.BlockTransform);
                                                acAttRef.Position = acAtt.Position.TransformBy(acBlkRef.BlockTransform);
                                                acAttRef.TextString = acAtt.TextString;
                                                acBlkRef.AttributeCollection.AppendAttribute(acAttRef);
                                                Tx.AddNewlyCreatedDBObject(acAttRef, true);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        foreach (SelectedObject acSSObj in acSSet)
                        {
                            BlockReference bref2 = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                            bref2.Erase();
                        }
                    }
                    Tx.Commit();
                }
            }
        }//добавление атрибутов к блоку
        public void ADDLes(string Comp,double SCHir, string RazdelSP, string Povorot, string soed, string Hvost, string DopSoed, string DopHvost,double HsagXv, double HsagSoed, ref List<Lestn> SpLESTN) 
        {
            Lestn NLestn = new Lestn();
            NLestn.NName(Comp);
            NLestn.NHsir(SCHir);
            NLestn.NRazdelSP(RazdelSP);
            NLestn.NPovorot(Povorot);
            NLestn.NSoed(soed);
            NLestn.NHvost(Hvost);
            NLestn.NDopSoed(DopSoed);
            NLestn.NDopHvost(DopHvost);
            NLestn.NHsagXv(HsagXv);
            NLestn.NHsagSoed(HsagSoed);
            SpLESTN.Add(NLestn);
        }//добавление лестницы в спецификацию
        static public void SBORpl(ref List<CWAY> SpCWAYcshert, ref List<TPosrL> SpOBJID_Poln,string FIND_N)
        {
            string stKomp = "";
            string stHtoEto = "";
            string stTip = "";
            string strTipSvOb = "";
            string stPOLnNaim = "";
            string stDobav = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[10];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 1);
            acTypValAr.SetValue(new TypedValue(0, "Circle"), 2);
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 5);
            acTypValAr.SetValue(new TypedValue(8, "CWAYPoint"), 6);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 7);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 8);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 9);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей изображенных полилинией...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stDobav = "";
                    stKomp = "";
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stHOZ = "";
                    dVisot = 0;
                    dDlin = 0;
                    Entity Obj = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    if (Obj.GetType() == typeof(BlockReference))
                    {
                        BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        stKomp = bref.Name;
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                    if (atrRef.Tag == "Ссылка") { stHtoEto = atrRef.TextString; }
                                    if (atrRef.Tag == "ДопДетали") { stDobav = atrRef.TextString; }
                                    if (atrRef.Tag == "Что_это") { stTip = atrRef.TextString; }
                                }
                            }
                        }
                        if ((stHtoEto.Contains(FIND_N + "-") | stHtoEto == FIND_N) & stKomp != "")
                        {
                            TPosrL TPostr1 = new TPosrL();
                            TPostr1.NKomp(stKomp);
                            TPostr1.NDopDet(stDobav);
                            TPostr1.NTip(stTip);
                            SpOBJID_Poln.Add(TPostr1);
                        }
                    }
                    ResultBuffer buffer = Obj.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                            if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out dVisot)) dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            if (Schet == 4) { stPom = value.Value.ToString(); }
                            if (Schet == 5) { stRazd = value.Value.ToString(); }
                            if (Schet == 6) { stKEI = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 10) { stHOZ = value.Value.ToString(); }
                            if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Math.Ceiling(Convert.ToDouble(value.Value.ToString()) / 100) / 10; }
                            Schet = Schet + 1;
                        }
                    }
                    CWAY nCWAY = new CWAY();
                    nCWAY.NCompName(stKomp);
                    nCWAY.NCWName(stHOZ);
                    nCWAY.Nlength(dDlin.ToString());
                    nCWAY.NModuleName(stPom);
                    if ((stKomp == "") == false){ SpCWAYcshert.Add(nCWAY);}
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных полилиниями
        static public void SBORpl_ID(ref List<string> SpCWAY_ID)
        {
            string stHOZ = "";
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[4];
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 3);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей изображенных полилинией...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stHOZ = "";
                    Entity bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 10) { stHOZ = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    SpCWAY_ID.Add(stHOZ);
                }
                Tx.Commit();
            }
        }//Создание списков ID деталей изображенных полилиниями
        public void SozdSpKONTUR(ref List<KONTUR> spKONTIR,string Find)
        {
            string Name = "", Tip = "", Pl1 = "", Pl2 = "", Dlin = "", Svazi = "";
            int Schet = 0;
            string smNOD1 = "";
            string smNOD2 = "";
            string spSmNod = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr;
            acTypValAr = new TypedValue[3];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Насыщение"), 1);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "НасыщениеСкрытое"), 2);
            List<Kriv> SpSkKriv = new List<Kriv>();
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    foreach (SelectedObject sobj in acSSet)
                    {
                        Name = "";
                        Pl1 = "";
                        Polyline ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline;
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 3) { Name = value.Value.ToString(); }
                                Schet = Schet + 1;
                                if (Schet > 3) { break; }
                            }
                        }
                        KONTUR NKontur = new KONTUR();
                        NKontur.NomName(Name);
                        NKontur.NID(sobj.ObjectId);
                        var KolV = ln.EndParam;
                        List<Point3d> SpT = new List<Point3d>();
                        for (int i = 0; i <= KolV; i++) SpT.Add(ln.GetPointAtParameter(i));
                        NKontur.NSpKoor(SpT);
                        spKONTIR.Add(NKontur);
                    }
                    tr.Commit();
                }
            }
        }//создания списка контуров помещений
        public void Find_KONTUR(ref List<TPosrL> SpPoint1,ref List<TPosrL> SpPoint_vn, Point3d Tn, Point3d Tk, List<Point3d> SpT,Point3d Tnar)
        {
            double DistA_B;
            double DistA_C;
            double DistC_B;
            double DELT;
            TPosrL TNPl = new TPosrL();
            TPosrL TKPl = new TPosrL();
            List<TPosrL> SpPoint_vn1 = new List<TPosrL>();
            //Application.ShowAlertDialog("Сейчас можно запивывать");
            for (int j = 0; j < SpPoint1.Count - 1; j++)
             {
                List<Point3d> SpToch = new List<Point3d>();
                Point3d TTr1 = SpPoint1[j].TPoint;
                Point3d TTr2 = SpPoint1[j + 1].TPoint;
                DistA_B = TTr1.DistanceTo(TTr2);
                DistA_C = TTr1.DistanceTo(Tn);
                DistC_B = Tn.DistanceTo(TTr2);
                DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                    if (DELT < 0.5) 
                    {
                    TNPl.NKomp("CWPoint");
                    TNPl.NNomT(SpPoint1[j].NomT + DistA_C/ DistA_B);
                    TNPl.NVisot(SpPoint1[j].Visot);
                    TNPl.NTPoint(Tn);
                    TNPl.NTip("Точка построения");
                    }
                DistA_B = TTr1.DistanceTo(TTr2);
                DistA_C = TTr1.DistanceTo(Tk);
                DistC_B = Tk.DistanceTo(TTr2);
                DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                    if (DELT < 0.5)
                    {
                    TKPl.NKomp("CWPoint");
                    TKPl.NNomT(SpPoint1[j].NomT + DistA_C / DistA_B);
                    TKPl.NVisot(SpPoint1[j].Visot);
                    TKPl.NTPoint(Tk);
                    TKPl.NTip("Точка построения");
                }
                if (TVnKont(SpT, SpPoint1[j].TPoint, Tnar)) { SpPoint_vn1.Add(SpPoint1[j]); }
            }
            SpPoint_vn.Add(TNPl);
            foreach (TPosrL TT in SpPoint_vn1) { SpPoint_vn.Add(TT);  }
            SpPoint_vn.Add(TKPl);
        }//список точек внутри контура детали изображеиейнной полилин
        public bool TVnKont(List<Point3d> SpT ,Point3d Tnar, Point3d Vn) 
        {
            List<Point3d> SpToch = new List<Point3d>();
            bool Y_N=false;
            for (int i = 0; i < SpT.Count - 1; i++)
            {
                Point3d TOtr1 = SpT[i];
                Point3d TOtr2 = SpT[i + 1];
                Point3d TperK = TPer1(TOtr1, TOtr2, Vn, Tnar);
                if (TperK != new Point3d() & SpToch.Exists(X=>X== TperK)==false) {SpToch.Add(TperK); }
            }
            if (SpToch.Count > 0 & (SpToch.Count % 2 != 0)) Y_N = true;
        return Y_N;     
        }//точки внутри контура
        public void Perestr(List<TPosrL> SpPoint1, List<ObjectId> SpOBJID, List<TPosrL> SpPoint_poln) 
        {
            double tID = 0;
            double Hsirin = 108;
            double HSag_Hvost = 1000;
            string Adr = "";
            string Xvost = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Lestn TLest = new Lestn();
            if (SpLESTN.Exists(x => x.Name == this.label14.Text))
            {
                TLest = SpLESTN.Find(x => x.Name == this.label14.Text);
                Hsirin = TLest.Hsir;
            }
            if (SpPoint1.Count > 1)
                {
                    if (this.textBox5.Text == "Тр") TRub_N_INtr_SP(SpPoint1, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, true,  checkBox1.Checked);
                    if (this.textBox5.Text == "Лес") Lest_N_INtr_SP(SpPoint1, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, true, this.textBox4.Text, checkBox1.Checked);
                }
                Database db = doc.Database;
                Transaction tr1 = db.TransactionManager.StartTransaction();
                SpOBJID.Add(ID);
                using (tr1)
                {
                    if (this.checkBox1.Checked) SpPoint_poln.RemoveAll(x => x.Tip == "хв");
                    foreach (TPosrL ID in SpPoint_poln)
                    { Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity; Obj.Erase(); }
                    tr1.Commit();
                }
         } //перестроить трассу или трубу
        public void UdalToch(ref List<TPosrL> spTpostr) 
        {
            List<TPosrL> spTUdal = new List<TPosrL>();
            for (int i = 0; i < spTpostr.Count() - 2; i++)
            {
                TPosrL TT0 = spTpostr[i];
                TPosrL TT1 = spTpostr[i + 1];
                TPosrL TT2 = spTpostr[i + 2];
                Vector3d vek1 = TT0.TPoint.GetVectorTo(TT1.TPoint);
                Vector3d vek2 = TT1.TPoint.GetVectorTo(TT2.TPoint);
                if (vek1.GetAngleTo(new Vector3d(1, 0, 0)) == vek2.GetAngleTo(new Vector3d(1, 0, 0)) & TT0.Visot == TT2.Visot & TT0.Visot == TT1.Visot) spTUdal.Add(TT1);
            }
            foreach (TPosrL TT in spTUdal) spTpostr.Remove(TT);
        }//удаление лишних точек в траектории постраения
        public void RazrSekR(Point3dCollection Kontur, List<TPosrL> SpPoint1, List<ObjectId> SpOBJID, List<TPosrL> SpPoint_poln, double Perepad, List<TPosrL> SpPoint1L2)
        {
            double tID = 0;
            double ttID = 0;
            double Hsirin = 108;
            double HSag_Hvost = 1000;
            double Zazor = Convert.ToDouble(this.textBox1.Text);
            double Gib = Convert.ToDouble(this.textBox18.Text);
            double Pram = Convert.ToDouble(this.textBox10.Text);
            string Adr = "";
            string Xvost = "";
            double ParRazr = 0;
            double j = 0;
            bool Perv = true;
            bool RekXvost = this.checkBox1.Checked;
            Lestn TLest = new Lestn();
            if (SpLESTN.Exists(x => x.Name == this.label14.Text))
            {
                TLest = SpLESTN.Find(x => x.Name == this.label14.Text);
                Hsirin = TLest.Hsir;
            }
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Find_KONTUR(ref SpPoint1, ref SpPointVN, Tn, Tk, SpT, Kontur[0]);
                List<TPosrL> SpDopT = DopT(Kontur, SpPointVN, Perepad, SpPoint1, SpPoint1L2);
                if (SpDopT.Count == 0) { this.Show(); return; }
                foreach (TPosrL ID in SpPoint_poln)
                {
                    if (ID.Tip == "Лес" | ID.Tip == "часть")
                    {
                        ttID = Convert.ToDouble(ID.LINK.Split('-').Last());
                        if (ttID > tID) tID = ttID;
                    }
                }
                tID = tID + 1;
                if (SpDopT.Count > 1)
                {
                    SpDopT.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    SpPoint11 = SpPointVN.FindAll(x => x.NomT < SpDopT[0].NomT);
                    SpPoint11.Add(SpDopT[0]);
                    if (Pram > 0 & Gib > 0) 
                    {
                        MestnGib(ref SpPoint11, Pram, Gib, SpPoint11.Last().NomT, Perepad, false, ref SpPoint1,0);
                        SpPoint11.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        if (SpPoint1L2.Count>0) PodgonVisot(ref SpPoint1,ref SpPoint11, SpPoint1L2,false);
                    }
                    //SpPoint11.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    SpPoint12 = SpPointVN.FindAll(x => x.NomT > SpDopT[1].NomT);
                    SpPoint12.Insert(0, SpDopT[1]);
                    if (Pram > 0 & Gib > 0) 
                    {
                        MestnGib(ref SpPoint12, Pram, Gib, SpPoint12[0].NomT, Perepad, true, ref SpPoint1,0);
                        SpPoint12.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        if (SpPoint1L2.Count > 0) PodgonVisot(ref SpPoint1, ref SpPoint12, SpPoint1L2,false);
                    }
                    //SpPoint12.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                }
                else
                {
                    SpPoint11 = SpPointVN.FindAll(x => x.NomT < SpDopT[0].NomT);
                    SpPoint12 = SpPointVN.FindAll(x => x.NomT > SpDopT[0].NomT);
                    List<Point3d> KonturT = new List<Point3d>();
                    foreach (Point3d tt in Kontur) KonturT.Add(tt);
                    Point3d Tnar = polar1(SpPoint12[0].TPoint,new Vector3d(1,0,0),9999999999);
                    if (TVnKont(KonturT, Tnar, SpPoint12[0].TPoint))
                    {
                        SpPoint11.Add(SpDopT[0]);
                        if (Pram > 0 & Gib > 0) MestnGib(ref SpPoint11, Pram, Gib, SpPoint11.Last().NomT, Perepad, false, ref SpPoint1,0);
                        SpPoint11.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        if (SpPoint1L2.Count > 0) PodgonVisot(ref SpPoint1, ref SpPoint11, SpPoint1L2,true);
                    }
                    else
                    {
                        SpPoint12.Insert(0, SpDopT[0]);
                        if (Pram > 0 & Gib > 0) MestnGib(ref SpPoint12, Pram, Gib, SpPoint12[0].NomT, Perepad, true, ref SpPoint1,0);
                        SpPoint12.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        if (SpPoint1L2.Count > 0) PodgonVisot(ref SpPoint1, ref SpPoint12, SpPoint1L2,true);
                    }
                }
                if (this.textBox5.Text == "Лес")
                {
                    if (SpPoint11.Count > 1) Lest_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, this.textBox4.Text, RekXvost);
                    tID = tID + 1;
                    if (SpPoint12.Count > 1) Lest_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, this.textBox4.Text, RekXvost);
                    //if (SpPoint11.Count > 1) Lest_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, true);
                    //tID = tID + 1;
                    //if (SpPoint12.Count > 1) Lest_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, true);
                }
                if (this.textBox5.Text == "Тр")
                {
                    if (SpPoint11.Count > 1) TRub_N_INtr_SP(SpPoint11, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, RekXvost);
                    tID = tID + 1;
                    if (SpPoint12.Count > 1) TRub_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, RekXvost);
                }
                Transaction tr1 = db.TransactionManager.StartTransaction();
                using (tr1)
                {
                    foreach (TPosrL ID in SpPoint_poln)
                    {
                        if (ID.LINK == this.textBox9.Text)
                        {
                            Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                            Obj.Erase();
                        }
                    }
                    if (Perepad > 0 | Gib > 0 | Pram > 0)
                    {
                        double n = 0;
                        foreach (TPosrL ID in SpPoint_poln)
                        {
                            if (ID.Tip == "Точка построения")
                            {
                                Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                                Obj.Erase();
                            }
                        }
                        SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        foreach (TPosrL TT in SpPoint1)
                        {
                            if (TT.Visot != "")
                            {
                                KrugIText(TT, SpPoint1, n.ToString(), this.textBox3.Text);
                                //KrugIText(TT, SpPoint1, TT.NomT.ToString(), this.textBox3.Text);
                                n = n + 1;
                            }
                        }
                    }
                    Entity Obj1 = tr1.GetObject(ID, OpenMode.ForWrite) as Entity;
                    Obj1.Erase();
                    tr1.Commit();
                }
            }
        }//создать разрыв секущей рамкой
        public void GibSekR(Point3dCollection Kontur, List<TPosrL> SpPoint1, List<ObjectId> SpOBJID, List<TPosrL> SpPoint_poln, List<TPosrL> SpPoint1L2) 
        {
            double tID = 0;
            double ttID = 0;
            double Hsirin = 108;
            double HSag_Hvost = 1000;
            double Zazor = Convert.ToDouble(this.textBox1.Text);
            double Perepad = Convert.ToDouble(this.textBox17.Text);
            double Gib = Convert.ToDouble(this.textBox18.Text);
            double Pram = Convert.ToDouble(this.textBox10.Text);
            string Adr = "";
            string Xvost = "";
            double ParRazr = 0;
            double j = 0;
            bool Perv = true;
            Lestn TLest = new Lestn();
            if (SpLESTN.Exists(x => x.Name == this.label14.Text))
            {
                TLest = SpLESTN.Find(x => x.Name == this.label14.Text);
                Hsirin = TLest.Hsir;
            }
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Find_KONTUR(ref SpPoint1, ref SpPointVN, Tn, Tk, SpT, Kontur[0]);
                List<TPosrL> SpDopT = DopT(Kontur, SpPointVN, Perepad, SpPoint1, SpPoint1L2);
                if (SpDopT.Count == 0) { this.Show(); return; }
                foreach (TPosrL ID in SpPoint_poln)
                {
                    if (ID.Tip == "Лес" | ID.Tip == "часть")
                    {
                        ttID = Convert.ToDouble(ID.LINK.Split('-').Last());
                        if (ttID > tID) tID = ttID;
                    }
                }
                tID = tID + 1;
                if (SpDopT.Count > 1)
                {
                    SpDopT.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    SpPointVN.Add(SpDopT[0]);
                    SpPointVN.Add(SpDopT[1]);
                    SpPoint1.Add(SpDopT[0]);
                    SpPoint1.Add(SpDopT[1]);
                    SpPoint11_12 = SpPointVN.FindAll(x => x.NomT > SpDopT[0].NomT & x.NomT < SpDopT[1].NomT);
                        if (SpPoint1L2.Count() > 0) 
                        {
                        double Rast = 99999999999;
                        foreach (TPosrL TT in SpPoint1L2) 
                        {
                            if (SpDopT[0].TPoint.DistanceTo(TT.TPoint) < Rast) 
                            {
                                Rast = SpDopT[0].TPoint.DistanceTo(TT.TPoint);
                                Perepad = Convert.ToDouble(TT.Visot);
                            }
                        }
                        IzmVisot(SpPoint11_12, ref SpPoint1, ref SpPointVN, Perepad,false);
                        }
                        else
                        IzmVisot(SpPoint11_12,ref SpPoint1,ref SpPointVN, Perepad,true);
                    SpPointVN.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    if (Gib > 0) MestnGib(ref SpPointVN, 0, Gib, SpDopT[0].NomT, 0, false, ref SpPoint1,0);
                    if (Gib > 0) MestnGib(ref SpPointVN, 0, Gib, SpDopT[1].NomT, 0, true, ref SpPoint1,0);
                    SpPointVN.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                }
                else
                {
                    SpPointVN.Add(SpDopT[0]);
                    SpPoint1.Add(SpDopT[0]);
                    SpPointVN.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    SpPoint11 = SpPointVN.FindAll(x => x.NomT < SpDopT[0].NomT);
                    SpPoint12 = SpPointVN.FindAll(x => x.NomT > SpDopT[0].NomT);
                    List<Point3d> KonturT = new List<Point3d>();
                    foreach (Point3d tt in Kontur) KonturT.Add(tt);
                    Point3d Tnar = polar1(SpPoint12[0].TPoint, new Vector3d(1, 0, 0), 9999999999);
                    if (TVnKont(KonturT, Tnar, SpPoint12[0].TPoint))
                    {
                        if (Gib > 0) MestnGib(ref SpPointVN, 0, Gib, SpDopT[0].NomT, Perepad, false, ref SpPoint1,0);
                        IzmVisot(SpPoint12, ref SpPoint1, ref SpPointVN, Perepad,true);
                        SpPointVN.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    }
                    else
                    {
                        if (Gib > 0) MestnGib(ref SpPointVN, 0, Gib, SpDopT[0].NomT, Perepad, true, ref SpPoint1,0);
                        IzmVisot(SpPoint11, ref SpPoint1, ref SpPointVN, Perepad,true);
                        SpPointVN.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    }
                }
                if (this.textBox5.Text == "Лес")
                {
                    if (SpPointVN.Count > 1) Lest_N_INtr_SP(SpPointVN, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, this.textBox4.Text, checkBox1.Checked);
                    //tID = tID + 1;
                    //if (SpPoint12.Count > 1) Lest_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false);
                }
                if (this.textBox5.Text == "Тр")
                {
                    if (SpPointVN.Count > 1) TRub_N_INtr_SP(SpPointVN, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false, checkBox1.Checked);
                    //tID = tID + 1;
                    //if (SpPoint12.Count > 1) TRub_N_INtr_SP(SpPoint12, this.label14.Text, Hsirin, this.textBox5.Text, "", this.textBox15.Text, this.textBox16.Text, "", 400 + Hsirin / 2, Convert.ToDouble(this.textBox13.Text), this.textBox12.Text, this.textBox11.Text, Convert.ToDouble(this.textBox14.Text), Convert.ToDouble(this.textBox3.Text), ref tID, false);
                }
                Transaction tr1 = db.TransactionManager.StartTransaction();
                using (tr1)
                {
                    foreach (TPosrL ID in SpPoint_poln)
                    {
                        if (ID.LINK == this.textBox9.Text)
                        {
                            Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                            Obj.Erase();
                        }
                    }
                    if (Perepad > 0 & Gib > 0)
                    {
                        double n = 0;
                        foreach (TPosrL ID in SpPoint_poln)
                        {
                            if (ID.Tip == "Точка построения")
                            {
                                Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                                Obj.Erase();
                            }
                        }
                        SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                        foreach (TPosrL TT in SpPoint1)
                        {
                            if (TT.Visot != "")
                            {
                                KrugIText(TT, SpPoint1, n.ToString(), this.textBox3.Text);
                                //KrugIText(TT, SpPoint1, TT.NomT.ToString(), this.textBox3.Text);
                                n = n + 1;
                            }
                        }
                    }
                    Entity Obj1 = tr1.GetObject(ID, OpenMode.ForWrite) as Entity;
                    Obj1.Erase();
                    tr1.Commit();
                }
            }
            this.Close();
        }//гиб секущей рамкой
        public void IzmVisot(List<TPosrL> SpPoint11_12, ref List<TPosrL> SpPoint1,ref List<TPosrL> SpPointVN, double Perepad,bool Umen) 
        {
            foreach (TPosrL TT in SpPoint11_12)
            {
                SpPointVN.RemoveAll(x => x.NomT == TT.NomT);
                SpPoint1.RemoveAll(x => x.NomT == TT.NomT);
                TPosrL NT = new TPosrL();
                NT.NNomT(TT.NomT);
                NT.NTPoint(TT.TPoint);
                    if(Umen)
                    NT.NVisot((Convert.ToDouble(TT.Visot) - Perepad).ToString());
                    else
                    NT.NVisot(Perepad.ToString());
                SpPointVN.Add(NT);
                SpPoint1.Add(NT);
            }
        }//уменьшение высоты
        public void PodgonVisot(ref List<TPosrL> SpPoint1, ref List<TPosrL> SpPointVN , List<TPosrL> SpPointNAR,bool last) 
        {
            bool TnTk = true;
            TPosrL Tn = SpPointVN[0];
            TPosrL Tk = SpPointVN.Last();
            double Dist = 999999999999;
            double Visot = 0;
            int Kiol = SpPointVN.Count - 1;
            int Kiol1 = SpPoint1.Count - 1;
            foreach (TPosrL TT in SpPointNAR) 
            {
                if (Tn.TPoint.DistanceTo(TT.TPoint) < Dist) { Dist = Tn.TPoint.DistanceTo(TT.TPoint); Visot = Convert.ToDouble(TT.Visot); TnTk = true; }
                if (Tk.TPoint.DistanceTo(TT.TPoint) < Dist) { Dist = Tk.TPoint.DistanceTo(TT.TPoint); Visot = Convert.ToDouble(TT.Visot); TnTk = false; }
            }
            if (TnTk) 
            {
                TPosrL T00 = SpPointVN[0];
                TPosrL T01 = SpPointVN[1];
                TPosrL T00n = new TPosrL();
                T00n.NNomT(T00.NomT);
                T00n.NTPoint(T00.TPoint);
                T00n.NVisot(Visot.ToString());
                TPosrL T01n = new TPosrL();
                T01n.NNomT(T01.NomT);
                T01n.NTPoint(T01.TPoint);
                T01n.NVisot(Visot.ToString());
                SpPointVN.Remove(T00);
                SpPointVN.Remove(T01);
                SpPointVN.Insert(0,T00n);
                SpPointVN.Insert(1,T01n);
                if (last)
                {
                    TPosrL T0 = SpPoint1[0];
                    TPosrL T0n = new TPosrL();
                    T0n.NNomT(T0.NomT);
                    T0n.NTPoint(T0.TPoint);
                    T0n.NVisot(Visot.ToString());
                    SpPoint1.Remove(T0);
                    SpPoint1.Insert(0, T0n);
                }
                int Nom = SpPoint1.FindIndex(x => x.NomT == T01.NomT);
                TPosrL T1 = SpPoint1[Nom];
                TPosrL T1n = new TPosrL();
                T1n.NNomT(T1.NomT);
                T1n.NTPoint(T1.TPoint);
                T1n.NVisot(Visot.ToString());
                SpPoint1.Remove(T1);
                SpPoint1.Insert(Nom, T1n);
            }
            else
            {
                TPosrL T00 = SpPointVN[Kiol];
                TPosrL T01 = SpPointVN[Kiol - 1];
                TPosrL T00n = new TPosrL();
                T00n.NNomT(T00.NomT);
                T00n.NTPoint(T00.TPoint);
                T00n.NVisot(Visot.ToString());
                TPosrL T01n = new TPosrL();
                T01n.NNomT(T01.NomT);
                T01n.NTPoint(T01.TPoint);
                T01n.NVisot(Visot.ToString());
                SpPointVN.Remove(T00);
                SpPointVN.Remove(T01);
                SpPointVN.Add(T01n);
                SpPointVN.Add(T00n);
                //Application.ShowAlertDialog("Последняя точка00-" + T00n.NomT.ToString() + "Последняя точка01-" + T01n.NomT.ToString());
                if (SpPoint1.Exists(x => x.NomT == T01.NomT))
                {
                    int Nom = SpPoint1.FindIndex(x => x.NomT == T01.NomT);
                    TPosrL T1 = SpPoint1.Find(x => x.NomT == T01.NomT);
                    TPosrL T1n = new TPosrL();
                    T1n.NNomT(T1.NomT);
                    T1n.NTPoint(T1.TPoint);
                    T1n.NVisot(Visot.ToString());
                    SpPoint1.Remove(T1);
                    SpPoint1.Insert(Nom,T1n);
                }
                if (last)
                {
                    TPosrL T0 = SpPoint1[Kiol1];
                    TPosrL T0n = new TPosrL();
                    T0n.NNomT(T0.NomT);
                    T0n.NTPoint(T0.TPoint);
                    T0n.NVisot(Visot.ToString());
                    SpPoint1.Remove(T0);
                    SpPoint1.Add(T0n);
                }//
            }
        }//изменение высоыточек
        public static void AvtoGib( List<Prepad> SpPREP_VZ, ref List<TPosrL> Tsp,double Pram,double Gib,double Visot, ref List<TPosrL> SpPoint1, double Zazor) 
        {
            //double Zazor = Convert.ToDouble(this.textBox22.Text);
            double par1 = Tsp[0].NomT;
            double par2 = Tsp.Last().NomT;
            List<Prepad> SpPREP_Gib = SpPREP_VZ.FindAll(x => x.Tip == "N" & par1 < x.Param & par2 > x.Param);
            //this.textBox23.Text = this.textBox23.Text + " par1= " + par1.ToString() + " par2=" + par2.ToString() + " SpPREP_Gib.count=" + SpPREP_Gib.Count +  '\r' + '\n';
            if (SpPREP_Gib.Count > 0)
            { 
                foreach (Prepad Tpr in SpPREP_Gib)
                {              
                    TPosrL nT = new TPosrL();
                    nT.NNomT(Tpr.Param);
                    nT.NTPoint(Tpr.Tp);
                    nT.NVisot((Convert.ToDouble(Tpr.Visot) + Zazor + 40).ToString());
                    SpPoint1.Add(nT);
                    Tsp.Add(nT);
                    SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    Visot = Convert.ToDouble(nT.Visot) + Zazor + 40 ;
                    Tsp.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    MestnGib(ref Tsp, Pram, Gib, Tpr.Param, Visot, false, ref SpPoint1, 0);
                    MestnGib(ref Tsp, Pram, Gib, Tpr.Param, Visot, true, ref SpPoint1, 0);
                    Tsp.Remove(nT);
                    SpPoint1.Remove(nT);
                }
            }
        }
        public static void apply_obstacles( List<Prepad> SpPREP,  List<Lestn> SpLESTN, Lestn GlCWEY, Prepad GlObc,bool RekXv,ObjectId IDthis)
        {
            double Zazor = GlObc.Zazor;
            double Pram = GlObc.Pram;
            double Gib = GlObc.Gib;
            double PramGib = GlObc.PramGib;
            double KVisot = GlObc.KVisot;
            double ZazorVir = GlObc.ZazorVir;
            double tID = 0;
            double ttID = 0;
            double Visot = 0;
            List<Prepad> SpPREP_VZ = new List<Prepad>();
            List<Prepad> SpPREP_Razr = new List<Prepad>();
            List<Prepad> SpPREP_Gib = new List<Prepad>();
            List<List<TPosrL>> SpSpPoint = new List<List<TPosrL>>();
            List<ObjectId> SpOBJID = new List<ObjectId>();
            double Steep_Shank = GlCWEY.HsagXv;
            double ID_CWEY = GlCWEY.IDCWey;
            string Find = GlCWEY.Name;
            string chapter_Sp = GlCWEY.RazdelSP;
            string Component = GlCWEY.Component;
            string shank = GlCWEY.Hvost;
            string connection = GlCWEY.Soed;
            string LINK = GlCWEY.LINK;
            string Room = GlCWEY.Room;
            List<TPosrL> SpPoint_poln = GlCWEY.SpPoint_poln;
            List<TPosrL> SpPoint1 = GlCWEY.SpPoint1;
            foreach (TPosrL ID in SpPoint_poln)
            {
                if (ID.Tip == "Лес" | ID.Tip == "часть")
                {
                    ttID = Convert.ToDouble(ID.LINK.Split('-').Last());
                    if (ttID > tID) tID = ttID;
                }
            }
            tID = tID + 1;
            tID = tID + 1;
            for (int i = 0; i <= SpPoint1.Count - 2; i++)
            {
                Point3d Tt1 = SpPoint1[i].TPoint;
                Point3d Tt2 = SpPoint1[i + 1].TPoint;
                foreach (Prepad TP in SpPREP)
                {
                    Point3d TTper = TPer1(Tt1, Tt2, TP.T1, TP.T2);
                    if (TTper != new Point3d())
                    {
                        Prepad TPre = new Prepad();
                        TPre.NTip(TP.Tip);
                        TPre.NParam(i + Tt1.DistanceTo(TTper) / Tt1.DistanceTo(Tt2));
                        TPre.NTp(TTper);
                        TPre.NVisot(TP.Visot);
                        TPre.NRad(TP.Rad);
                        SpPREP_VZ.Add(TPre);
                    }
                }
            }
            double par1 = 0;
            double par2 = 0;
            SpPREP_Razr = SpPREP_VZ.FindAll(x => x.Tip == "Y");
            if (SpPREP_Razr.Count > 0) SpPREP_Razr.Sort(delegate (Prepad x, Prepad y) { return x.Param.CompareTo(y.Param); });
            par1 = 0;
            par2 = 0;
            foreach (Prepad Tpr in SpPREP_Razr)
            {
                par2 = Tpr.Param;
                List<TPosrL> TSpT = SpPoint1.FindAll(x => x.NomT >= par1 & x.NomT <= par2);
                if (SpSpPoint.Count > 0)
                {
                    TPosrL T1 = new TPosrL();
                    T1.NNomT(SpSpPoint.Last().Last().NomT);
                    T1.NVisot(SpSpPoint.Last().Last().Visot);
                    T1.NTPoint(SpSpPoint.Last().Last().TPoint);
                    TSpT.Insert(0, T1);
                }
                TPosrL T2 = new TPosrL();
                T2.NNomT(Tpr.Param);
                T2.NVisot(Tpr.Visot.ToString());
                T2.NTPoint(Tpr.Tp);
                TSpT.Add(T2);
                if (KVisot != 0)
                {
                    Visot = Math.Round(Tpr.Visot - (Tpr.Visot * KVisot) - ZazorVir);
                    if (SpSpPoint.Count > 0) { MestnGib(ref TSpT, Pram, Gib, TSpT[0].NomT, Visot, true, ref SpPoint1, Zazor * 2); }
                    MestnGib(ref TSpT, Pram, Gib, TSpT.Last().NomT, Visot, false, ref SpPoint1, Zazor);
                }
                else
                    MestnGib(ref TSpT, 0, 0, TSpT.Last().NomT, Visot, false, ref SpPoint1, Zazor);
                SpSpPoint.Add(TSpT);
                par1 = par2;
            }
            if (SpPREP_Razr.Count > 0)
            {
                par1 = SpPREP_Razr.Last().Param;
                List<TPosrL> TSpT = SpPoint1.FindAll(x => x.NomT >= par1);
                TPosrL T2 = new TPosrL();
                T2.NNomT(SpPREP_Razr.Last().Param);
                T2.NVisot(SpPREP_Razr.Last().Visot.ToString());
                T2.NTPoint(SpPREP_Razr.Last().Tp);
                TSpT.Insert(0, T2);
                if (KVisot != 0)
                { MestnGib(ref TSpT, Pram, Gib, TSpT[0].NomT, Visot, true, ref SpPoint1, Zazor); }
                else
                    MestnGib(ref TSpT, 0, 0, TSpT[0].NomT, Visot, true, ref SpPoint1, Zazor);
                SpSpPoint.Add(TSpT);
            }
            else
            {
                SpSpPoint.Add(SpPoint1);
            }
            double Hsirin = 108;
            double HSag_Hvost = 1000;
            string Adr = "";
            string Xvost = "";
            Lestn TLest = new Lestn();
            if (SpLESTN.Exists(x => x.Name == Component))
            {
                TLest = SpLESTN.Find(x => x.Name == Component);
                Hsirin = TLest.Hsir;
            }
            foreach (List<TPosrL> Tsp in SpSpPoint)
            {
                List<TPosrL> TSpT = Tsp.ToList();
                AvtoGib(SpPREP_VZ, ref TSpT, PramGib, Gib, Visot, ref SpPoint1, Zazor);
                if (chapter_Sp == "Лес") Lest_N_INtr_SP(TSpT, Component, Hsirin, chapter_Sp, "", shank, connection, "", 400 + Hsirin / 2, Steep_Shank, ID_CWEY, ref tID, false, Room, RekXv);
                tID = tID + 1;
            }
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Transaction tr1 = doc.TransactionManager.StartTransaction();
            using (tr1)
            {
                foreach (TPosrL ID in SpPoint_poln)
                {
                    if (ID.LINK == LINK)
                    {
                        Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                        Obj.Erase();
                    }
                }
                //if (Perepad > 0 | Gib > 0 | Pram > 0)
                {
                    double n = 0;
                    foreach (TPosrL ID in SpPoint_poln)
                    {
                        if (ID.Tip == "Точка построения")
                        {
                            Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity;
                            Obj.Erase();
                        }
                    }
                    SpPoint1.Sort(delegate (TPosrL x, TPosrL y) { return x.NomT.CompareTo(y.NomT); });
                    foreach (TPosrL TT in SpPoint1)
                    {
                        if (TT.Visot != "")
                        {
                            KrugIText(TT, SpPoint1, n.ToString(), Find);
                            //KrugIText(TT, SpPoint1, TT.NomT.ToString(), this.textBox3.Text);
                            n = n + 1;
                        }
                    }
                }
                Entity Obj1 = tr1.GetObject(IDthis, OpenMode.ForWrite) as Entity;
                Obj1.Erase();
                tr1.Commit();
            }

        }
    }
}
