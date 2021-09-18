using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.CSharp.RuntimeBinder;

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


    public partial class Form1 : Form
    {
        public enum Side { LeftUp, LeftBottom, RigthUp, RigthBottom };
        public enum orientation { gor, vert, notDef };
        public static string DBPath;
        public static SQLiteConnection connection;
        public static SQLiteCommand command;
        public DataSet DS1 = new DataSet();
        public System.Data.DataTable DT1 = new System.Data.DataTable();

        public struct TPodk
        {
            public string IND, Naim, Pom, Bort, Pal, Hpan, D3D2, Shem;
            public double VisT;
            public Point3d Koord, Koord3D;
            public ObjectId ID;
            public string KoordMod;
            public int Vstr;
            public string Sist;
            public string BlNOD;
            public void ND3D2(string i) { D3D2 = i; }
            public void NIND(string i) { IND = i; }
            public void NSist(string i) { Sist = i; }
            public void NBlNOD(string i) { BlNOD = i; }
            public void NVisT(double i) { VisT = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoor3D(Point3d i) { Koord3D = i; }
            public void NKoorMod(string i) { KoordMod = i; }
            public void NVstr(int i) { Vstr = i; }
            public void NNaim(string i) { Naim = i; }
            public void NShem(string i) { Shem = i; }
            public void NID(ObjectId i) { ID = i; }
            public Point3dCollection ListPointCweyDim(ObjectId Cwey)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Point3dCollection ListPoint = new Point3dCollection();
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    Polyline bref = Tx.GetObject(Cwey, OpenMode.ForWrite) as Polyline;
                    for (int i = 0; i <= bref.EndParam; i++) { ListPoint.Add(bref.GetPointAtParameter(i)); }
                }
                return ListPoint;
            }
        }
        public struct PLOS
        {
            //public  Side PlaneSide;
            public string Vid, Mas, Spravka, List, Osi, Nomer, Krad;
            public Point3d Psk, Msk, max, min, max_n, min_v;
            public void NOsi(string i) { Osi = i; }
            public void NVid(string i) { Vid = i; }
            public void NMas(string i) { Mas = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NList(string i) { List = i; }
            public void NNomer(string i) { Nomer = i; }
            public void NKrad(string i) { Krad = i; }
            public void NPsk(Point3d i) { Psk = i; }
            public void NMsk(Point3d i) { Msk = i; }
            public void Nmax(Point3d i) { max = i; }
            public void Nmin(Point3d i) { min = i; }
            public void Nmin_max() { max_n = new Point3d(max.X, min.Y, 0); min_v = new Point3d(min.X, max.Y, 0); }
            public Point3d FFindPoint(Point3dCollection ListPoint, out Side PlaneSide)
            {
                double MinDist = 99999999999999999;
                Point3d FindPoint = new Point3d();
                PlaneSide = Side.LeftBottom;
                foreach (Point3d TP in ListPoint) { if (TP.DistanceTo(max) < MinDist) { MinDist = TP.DistanceTo(max); FindPoint = TP; PlaneSide = Side.RigthUp; } }
                foreach (Point3d TP in ListPoint) { if (TP.DistanceTo(max_n) < MinDist) { MinDist = TP.DistanceTo(max_n); FindPoint = TP; PlaneSide = Side.RigthBottom; } }
                foreach (Point3d TP in ListPoint) { if (TP.DistanceTo(min) < MinDist) { MinDist = TP.DistanceTo(min); FindPoint = TP; PlaneSide = Side.LeftBottom; } }
                foreach (Point3d TP in ListPoint) { if (TP.DistanceTo(min_v) < MinDist) { MinDist = TP.DistanceTo(min_v); FindPoint = TP; PlaneSide = Side.LeftUp; } }
                return FindPoint;
            }
        };
        public struct POZIZIA
        {
            public string Compon, Dobav, Pom, RazdelSp, KEI, HtoEto, NOMpoz, Ind, shabl, rNAS, TiPObor, Hozain, RAB, LINK, polnNAIM;
            public double Dlin, Visot, Kol;
            public double KolF;
            public void NNOMpoz(string i) { NOMpoz = i; }
            public void NCompon(string i) { Compon = i; }
            public void NDobav(string i) { Dobav = i; }
            public void NPom(string i) { Pom = i; }
            public void NRazdelSp(string i) { RazdelSp = i; }
            public void NKEI(string i) { KEI = i; }
            public void NHtoEto(string i) { HtoEto = i; }
            public void NDlin(double i) { Dlin = i; }
            public void NVisot(double i) { Visot = i; }
            public void NKolF(double i) { KolF = i; }
            public void NKol(double i) { Kol = i; }
            public void NInd(string i) { Ind = i; }
            public void Nshabl(string i) { shabl = i; }
            public void NrNAS(string i) { rNAS = i; }
            public void NTiPObor(string i) { TiPObor = i; }
            public void NHozain(string i) { Hozain = i; }
            public void NRAB(string i) { RAB = i; }
            public void NLINK(string i) { LINK = i; }
            public void NpolnNAIM(string i) { polnNAIM = i; }
        };
        public struct Zag
        {
            public string Zagal;
            public int Graf;
            public void NomZagal(string i) { Zagal = i; }
            public void NomGraf(int i) { Graf = i; }
        }
        public struct TPosrL
        {
            public ObjectId OId;
            public Point3d TPoint, TPoint1, TPointZ, MaxPoint ,MinPoint;
            public Point3dCollection SpKoor;
            public string Visot, Komp, Tip, DopDet, LINK, NameCWAY;
            public double NomT, Delt, AngelV, AngelG;
            public void NDelt(double i) { Delt = i; }
            public void NVisot(string i) { Visot = i; }
            public void NKomp(string i) { Komp = i; }
            public void NTip(string i) { Tip = i; }
            public void NDopDet(string i) { DopDet = i; }
            public void NLINK(string i) { LINK = i; }
            public void NNomT(double i) { NomT = i; }
            public void NTPoint(Point3d i) { TPoint = i; }
            public void NTPoint1(Point3d i) { TPoint1 = i; }
            public void Dimensions(Point3d max, Point3d min) { MaxPoint = max; MinPoint = min; }
            public void NTPointZ()
            {
                double Dist = TPoint.DistanceTo(TPoint1);
                Vector3d Vekt1 = TPoint.GetVectorTo(TPoint1) / Dist;
                TPointZ = TPoint + Vekt1 * (Dist / 2);
            }
            public void getAngel()
            {
                Vector3d Vekt1 = TPoint.GetVectorTo(TPoint1);
                AngelV = Vekt1.GetAngleTo(new Vector3d(0, 1, 0));
                AngelG = Vekt1.GetAngleTo(new Vector3d(1, 0, 0));
            }
            public void NOId(ObjectId i) { OId = i; }
            public void NSpKoor(Point3dCollection i) { SpKoor = i; }
            public void NNameCWAY(string nameCWAY) { NameCWAY = nameCWAY; }
        }
        public struct Lestn
        {
            public string Name, Povorot, RazdelSP, Soed, Hvost,AddidDetHvpst, AddidDetSoed,StepHvost,StepSoed;
            public double Hsir;
            public void NName(string i) { Name = i; }
            public void NPovorot(string i) { Povorot = i; }
            public void NRazdelSP(string i) { RazdelSP = i; }
            public void NSoed(string i) { Soed = i; }
            public void NHvost(string i) { Hvost = i; }
            public void NHsir(double i) { Hsir = i; }
            public void Creat(string name,double hir,string razdelSP, string povorot,string soed,string hvost,string addidDetHvpst, string addidDetSoed, string stepHvost, string stepSoed) 
            {
                Name = name;
                Hsir = hir;
                RazdelSP = razdelSP;
                Povorot = povorot;
                Soed = soed;
                Hvost = hvost;
                AddidDetHvpst = addidDetHvpst;
                AddidDetSoed = addidDetSoed;
                StepHvost = stepHvost;
                StepSoed = stepSoed;
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
            public void NCompName(string i)
            {
                CompName = i;
            }
            public void NModuleName(string i) { ModuleName = i; }
            public void Nlength(string i) { length = i; }
        }
        public struct Lest
        {
            public string Nazv, Tip_Pom, Baza, poz, dobav, xvost, perem, pom, Comp;
            public List<TPodk> SpT;
            public List<TPosrL> SpTCW;
            public List<string> SpXvost;
            public void NNazv(string i) { Nazv = i; }
            public void NTip_Pom(string i) { Tip_Pom = i; }
            public void NBaza(string i) { Baza = i; }
            public void Npoz(string i) { poz = i; }
            public void Nxvost(string i) { xvost = i; }
            public void Nperem(string i) { perem = i; }
            public void Ndobav(string i) { dobav = i; }
            public void Npom(string i) { pom = i; }
            public void NComp(string i) { Comp = i; }
            public void NSpT(List<TPodk> i) { SpT = i; }
            public void NSpT(List<TPosrL> i) { SpTCW = i; }
            public void NSpXvost(List<string> i) { SpXvost = i; }

        }

        public List<PLOS> SpPLOS = new List<PLOS>();
        public List<Part> SpPart = new List<Part>();
        public List<CWAY> SpCWAY = new List<CWAY>();
        public List<Lestn> SpLest = new List<Lestn>();
        public List<ObjectId> SpOBJID = new List<ObjectId>();
        public List<POZIZIA> SpPOZ = new List<POZIZIA>();
        public List<POZIZIA> SpPOZsPOZ = new List<POZIZIA>();
        public List<POZIZIA> SpNamenkl = new List<POZIZIA>();
        public List<CWAY> SpCWAYcshert = new List<CWAY>();
        public List<Lestn> ListLader = new List<Lestn>();
        public string Blok = "";
        public string RazdSP = "";
        public string TipBLOK = "";
        public string Mas;
        public string Rasp;
        public string RaspIZ;
        BindingSource BS = new BindingSource();

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            CreateLayer("Насыщение");
            CreateLayer("НасыщениеСкрытое");
            CreateLayer("CWAYPoint");
            CreateLayer("Плоскости");
            CreateLayer("АвтоРазмеры");
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            this.textBox9.Text = HCtenSlov("DERIZ", @"\\tbserv.vympel.local\home\\41 отдел\Ерофеев А.А\\Программы\ПодборУзлов\Изображения\");
            RaspIZ = this.textBox9.Text;
            this.textBox7.Text = HCtenSlov("DERBL", @"\\tbserv.vympel.local\home\41 отдел\Ерошенко\BAZA\");
            Rasp = this.textBox7.Text;
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
            if (Directory.Exists(@Rasp + @"796DB\") == true)
            {
                string[] files = Directory.GetFiles(@Rasp + @"796DB\", "*.dwg");
                for (int i = 0; i < files.Length; i++)
                {
                    string[] blokM = files[i].Split('\\');
                    Blok = blokM.Last().TrimEnd('.', 'd', 'w', 'g');
                    this.listBox1.Items.Add(Blok);
                }
            }
            SOZDlov("MAS#");
            Mas = HCtenSlov("MAS#", "1");
            //Mas = "1";
            if (Mas == "0.4") { this.radioButton1.Checked = true; }
            if (Mas == "1") { this.radioButton2.Checked = true; }
            if (Mas == "2") { this.radioButton3.Checked = true; }
            this.textBox1.Text = HCtenSlov("POM", "");
            this.textBox4.Text = HCtenSlov("DOC", "");
            this.label9.Text = HCtenSlov("IstokNm", "");
            this.label8.Text = HCtenSlov("sp", "");
            this.comboBox1.Items.Add("Детали поштучно (КЕИ=796)");
            this.comboBox1.Items.Add("Детали метражем (КЕИ=006)");
            this.comboBox1.Items.Add("Детали поштучно (КЕИ=796 статические блоки)");
            if (System.IO.File.Exists(@"C:\МАРШРУТ\SpIzNANOsPoz.txt") == true) HCteniPOZTXT(ref SpPOZsPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
            ZapolnTabl_Nam();
            ZapolnTabl_Sp_Les();
            SBORBl(ref SpPOZ, SpNamenkl, ref SpPLOS);
            SBORpl(ref SpPOZ, SpNamenkl, ref SpCWAYcshert);
            ZapolnTabl(SpPOZ, SpNamenkl);
        }//загрузка формы заполнение списков
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RaspIZ = this.textBox9.Text;
            Blok = this.listBox1.SelectedItem.ToString();
            string[] stKompM = Blok.Split('#');
            string Kort = stKompM.Last().Replace('@', '.');
            string strAdrRis = RaspIZ + @"Слесарка\" + Kort + ".jpg";
            string strAdrErr = RaspIZ + "Desert.jpg";
            //this.pictureBox1.Image = System.Drawing.Image.FromFile(strAdrRis);
            if (System.IO.File.Exists(strAdrRis))
            {
                this.pictureBox1.Load(strAdrRis);
            }
            else
            {
                this.pictureBox1.Image = null;
            }
            this.label4.Text = strAdrRis;
        }//выбор из какой папки будут братся блоки
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Rasp = this.textBox7.Text;
            this.listBox1.Items.Clear();
            TipBLOK = this.comboBox1.SelectedItem.ToString();
            if (TipBLOK == "Детали поштучно (КЕИ=796)")
            {
                if (Directory.Exists(@Rasp + @"796DB\") == true)
                {
                    string[] files = Directory.GetFiles(@Rasp + @"796DB\", "*.dwg");
                    for (int i = 0; i < files.Length; i++)
                    {
                        string[] blokM = files[i].Split('\\');
                        Blok = blokM.Last().TrimEnd('.', 'd', 'w', 'g');
                        this.listBox1.Items.Add(Blok); }
                }
            }
            if (TipBLOK == "Детали метражем (КЕИ=006)")
            {
                if (Directory.Exists(@Rasp + @"006DB\") == true)
                {
                    string[] files = Directory.GetFiles(@Rasp + @"006DB\", "*.dwg");
                    for (int i = 0; i < files.Length; i++)
                    {
                        string[] blokM = files[i].Split('\\');
                        Blok = blokM.Last().TrimEnd('.', 'd', 'w', 'g');
                        this.listBox1.Items.Add(Blok);
                    }
                }
            }
            if (TipBLOK == "Детали поштучно (КЕИ=796 статические блоки)")
            {
                if (Directory.Exists(@Rasp + @"796\") == true)
                {
                    string[] files = Directory.GetFiles(@Rasp + @"796\", "*.dwg");
                    for (int i = 0; i < files.Length; i++)
                    {
                        string[] blokM = files[i].Split('\\');
                        Blok = blokM.Last().TrimEnd('.', 'd', 'w', 'g');
                        this.listBox1.Items.Add(Blok);
                    }
                }
            }
        }//выбр выставляемой детали в списке
        private void button1_Click(object sender, EventArgs e)
        {
            Rasp = this.textBox7.Text;
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                //ZapisSlov(this.textBox1.Text);
                SOZDlov("POM");
                ZapisSlov("POM", this.textBox1.Text);
                SOZDlov("VIS");
                ZapisSlov("VIS", this.textBox8.Text);
                string KEI = "796";
                string BlokKatal = "";
                string BlokType = "DB";

                CreateLayer("Насыщение");

                if (TipBLOK == "") { BlokKatal = @Rasp + @"\796DB\" + Blok + ".dwg"; KEI = "796"; }
                if (TipBLOK == "Детали поштучно (КЕИ=796)") { BlokKatal = @Rasp + @"\796DB\" + Blok + ".dwg"; KEI = "796"; }
                if (TipBLOK == "Детали метражем (КЕИ=006)") { BlokKatal = @Rasp + @"\006DB\" + Blok + ".dwg"; KEI = "006"; }
                if (TipBLOK == "Детали поштучно (КЕИ=796 статические блоки)") { BlokKatal = @Rasp + @"\796\" + Blok + ".dwg"; KEI = "006"; BlokType = "SB"; }
                //Application.ShowAlertDialog(Blok);
                InsBlockRef(BlokKatal, Blok, BlokType);
                SetDynamicBlkProperty(Blok, this.textBox8.Text, "", this.textBox1.Text, "Доизол.д", KEI, "", "", 0, "", "", 0, "Насыщение");
            }
            this.Show();
        }//Доизоляционная деталь
        private void button2_Click(object sender, EventArgs e)
        {
            Rasp = this.textBox7.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                this.Hide();
                SOZDlov("POM");
                ZapisSlov("POM", this.textBox1.Text);
                SOZDlov("VIS");
                ZapisSlov("VIS", this.textBox8.Text);
                string KEI = "796";
                string BlokKatal = "";
                string BlokType = "DB";
                CreateLayer("Насыщение");
                if (TipBLOK == "") { BlokKatal = @Rasp + @"\796DB\" + Blok + ".dwg"; KEI = "796"; }
                if (TipBLOK == "Детали поштучно (КЕИ=796)") { BlokKatal = @Rasp + @"\796DB\" + Blok + ".dwg"; KEI = "796"; }
                if (TipBLOK == "Детали метражем (КЕИ=006)") { BlokKatal = @Rasp + @"\006DB\" + Blok + ".dwg"; KEI = "006"; }
                if (TipBLOK == "Детали поштучно (КЕИ=796 статические блоки)") { BlokKatal = @Rasp + @"\796\" + Blok + ".dwg"; KEI = "006"; BlokType = "SB"; }
                //Application.ShowAlertDialog(Blok);
                InsBlockRef(BlokKatal, Blok, BlokType);
                SetDynamicBlkProperty(Blok, this.textBox8.Text, "", this.textBox1.Text, "Послеизол.д", KEI, "", "", 0, "", "", 0, "Насыщение");
            }
            this.Show();
        }//Послеизоляционная деталь
        private void button3_Click(object sender, EventArgs e)
        {
            Rasp = this.textBox7.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("POM");
                ZapisSlov("POM", this.textBox1.Text);
                SOZDlov("VIS");
                ZapisSlov("VIS", this.textBox8.Text);
                string KEI = "796";
                string BlokKatal = "";
                string BlokType = "DB";
                CreateLayer("Насыщение");
                if (TipBLOK == "") { BlokKatal = @Rasp + @"\796DB\" + Blok + ".dwg"; KEI = "796"; }
                if (TipBLOK == "Детали поштучно (КЕИ=796)") { BlokKatal = @Rasp + @"\796DB\" + Blok + ".dwg"; KEI = "796"; }
                if (TipBLOK == "Детали метражем (КЕИ=006)") { BlokKatal = @Rasp + @"\006DB\" + Blok + ".dwg"; KEI = "006"; }
                if (TipBLOK == "Детали поштучно (КЕИ=796 статические блоки)") { BlokKatal = @Rasp + @"\796\" + Blok + ".dwg"; KEI = "006"; BlokType = "SB"; }
                //Application.ShowAlertDialog(Blok);
                InsBlockRef(BlokKatal, Blok, BlokType);
                SetDynamicBlkProperty(Blok, this.textBox8.Text, "", this.textBox1.Text, "Тр пом", KEI, "", "", 0, "", "", 0, "Насыщение");
            }
        }//Деталь ТЗК
        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                List<POZIZIA> SpPOZ = new List<POZIZIA>();
                List<POZIZIA> SpPOZob = new List<POZIZIA>();
                //if (System.IO.File.Exists(@"C:\МАРШРУТ\SpIzNANOsPoz.txt")) HCteniPOZTXT(ref SpPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
                //if (System.IO.File.Exists(@"C:\МАРШРУТ\SpIzNANOsPozObor.txt")) HCteniPOZTXT(ref SpPOZob, @"C:\МАРШРУТ\SpIzNANOsPozObor.txt");
                Database db = doc.Database;
                Editor ed = doc.Editor;
                PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
                PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("");
                Point3d T1 = prEntResult.PickedPoint;
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = T1;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d T2 = pPtRes.Value;
                ProstPOZ(T1, T2, prEntResult.ObjectId, SpPOZ, SpPOZob);
            }
            this.Show();
        }//Указать позицию 
        private void button5_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                List<POZIZIA> SpPOZ = new List<POZIZIA>();
                List<POZIZIA> SpPOZsPOZ = new List<POZIZIA>();
                HCteniPOZTXT(ref SpPOZsPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
                SBORpl(ref SpPOZ, SpPOZsPOZ, ref SpCWAYcshert);
                SBORBl(ref SpPOZ, SpPOZsPOZ, ref SpPLOS);
                ZapolnTabl(SpPOZ, SpPOZsPOZ);
            }
        }//кнопка обновить
        private void button6_Click(object sender, EventArgs e)
        {
            //Int32 nStr=this.dataGridView1.CurrentCell.RowIndex;
            if (this.dataGridView1.CurrentRow.Cells[1].Value != null)
            {
                String FindK = this.dataGridView1.CurrentRow.Cells[7].Value.ToString();
                String FindR = this.dataGridView1.CurrentRow.Cells[6].Value.ToString();
                String FindP = this.dataGridView1.CurrentRow.Cells[4].Value.ToString();
                List<ObjectId> SpOBJID = new List<ObjectId>();
                //Application.ShowAlertDialog(Find);
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    Database db = doc.Database;
                    Editor ed = doc.Editor;
                    SBORBl_FIND(ref SpOBJID, FindK, FindR, FindP);
                    SBORpl_FIND(ref SpOBJID, FindK, FindR, FindP);
                    ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                    ed.SetImpliedSelection(idarrayEmpty);
                }
            }
        }//кнопка поиска деталей по омпоненту
        private void button17_Click(object sender, EventArgs e)
        {
            this.dataGridView6.Rows.Clear();
            List<string> StrSpLESTN = new List<string>();
            HCteniSP_LESTN(ref StrSpLESTN, this.textBox5.Text);
            SOZDlov("SPDETPOLY");
            ZapisSlovSistTR("SPDETPOLY", StrSpLESTN);
            //Application.ShowAlertDialog(StrSpLESTN.Count.ToString());
            SozSpLest(ref SpLest, StrSpLESTN);
            zapDGWLes(SpLest);
        }//загрузить спецификацию
        private void button14_Click(object sender, EventArgs e)
        {
            List<POZIZIA> SpPOZob = new List<POZIZIA>();
            HCteniPOZTXT(ref SpPOZob, @"C:\МАРШРУТ\SpIzNANOsPozObor.txt");
            SpNamenkl.Clear();
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; ZagrExcel(File); this.label9.Text = File; }
            else return;
            this.dataGridView5.Rows.Clear();
            foreach (POZIZIA tPoz in SpNamenkl)
                this.dataGridView5.Rows.Add(tPoz.NOMpoz, tPoz.Compon, tPoz.Hozain, tPoz.polnNAIM, tPoz.Ind, tPoz.Dobav, tPoz.Pom);
            OBNNpoz(ref SpNamenkl, ref SpPOZob);
        }//изменить расположение файла с номенклатурой
        private void button15_Click(object sender, EventArgs e)
        {
            List<POZIZIA> SpPOZob = new List<POZIZIA>();
            HCteniPOZTXT(ref SpPOZob, @"C:\МАРШРУТ\SpIzNANOsPozObor.txt");
            SpNamenkl.Clear();
            if (!System.IO.File.Exists(this.label9.Text)) return;
            ZagrExcel(this.label9.Text);
            this.dataGridView5.Rows.Clear();
            foreach (POZIZIA tPoz in SpNamenkl)
                this.dataGridView5.Rows.Add(tPoz.NOMpoz, tPoz.Compon, tPoz.Hozain, tPoz.polnNAIM, tPoz.Ind, tPoz.Dobav, tPoz.Pom);
            OBNNpoz(ref SpNamenkl, ref SpPOZob);
        }//загрузить наменклатуру
        private void button21_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("DERBL");
                ZapisSlov("DERBL", this.textBox7.Text);
            }
        }//Сохранить дерикторию с блоками 
        private void button22_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("DERIZ");
                ZapisSlov("DERIZ", this.textBox9.Text);
            }
        }//Сохранить дерикторию с изображениями
        private void button24_Click(object sender, EventArgs e)
        {
            string File = this.textBox10.Text;
            if (System.IO.File.Exists(@File) == true)
                ZagrExcelCW(File);
            else
                Application.ShowAlertDialog(File + "-  не найден");
        }//загрузить трассы из Ecxel
        private void button23_Click(object sender, EventArgs e)
        {
            if (SpPLOS.Count == 0) { Application.ShowAlertDialog("Нет ни одной плоскости построения"); return; }
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            double GlubinaSR = Convert.ToDouble(this.textBox21.Text);
            Point3d BP = new Point3d();
            Point3d PSKtl = new Point3d();
            Point3d MSKtl = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            List<Part> SpPartTCWAY = new List<Part>();
            List<Part> SpPartTCWAYSort = new List<Part>();
            List<TPosrL> SpTPostrL = new List<TPosrL>();
            string Room = this.textBox1.Text;
            string OSItl = "";
            string Name = "";
            string StartPoint = "";
            string EndPoint = "";
            string StartPointTr = "";
            string EndPointTr = "";
            string COMPON = "";
            string NAME = "";
            double Dles = 100;
            double Delt = 0, DeltX=0,DeltY=0;
            Point3d MAX = new Point3d();
            Point3d MIN = new Point3d();
            string[] COMPONm;
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
                            if (prop.PropertyName == "Расстояние1") { DeltX = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Расстояние2") { DeltY = Convert.ToDouble(prop.Value.ToString()); }
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
                    MIN = BP;
                    MAX = new Point3d(MIN.X + DeltX, MIN.Y + DeltY, 0);
                }
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
                foreach (CWAY tCWAY in SpCWAY)
                {
                    double CountStartPoint = 0;
                    double CountEndPoint = 0;
                    Name = tCWAY.CWName;
                    if(this.checkBox8.Checked) NAME = Name;
                    SpTPostrL.Clear();
                    SpPartTCWAYSort.Clear();
                    SpPartTCWAY = SpPart.FindAll(x => x.CWName == Name);
                    foreach (Part TPart in SpPartTCWAY)
                    {
                        StartPoint = TPart.StartPoint;
                        EndPoint = TPart.EndPoint;
                        List<Part> SpPartStartPoint = SpPartTCWAY.FindAll(x => x.StartPoint == StartPoint | x.EndPoint == StartPoint);
                        List<Part> SpPartEndPoint = SpPartTCWAY.FindAll(x => x.StartPoint == EndPoint | x.EndPoint == EndPoint);
                        if (SpPartStartPoint.Count == 1) { StartPointTr = StartPoint; CountStartPoint += 1; }
                        if (SpPartEndPoint.Count == 1) { EndPointTr = EndPoint; CountEndPoint += 1; }
                    }
                    string TPointTr = StartPointTr;
                    string Prodol = "Y";
                    if (CountStartPoint == 1 & CountEndPoint == 1)
                    {
                        while (TPointTr != EndPointTr | Prodol != "Y")
                        {
                            if (SpPartTCWAY.Exists(x => x.StartPoint == TPointTr))
                            {
                                Part TPart = SpPartTCWAY.Find(x => x.StartPoint == TPointTr);
                                TPointTr = TPart.EndPoint;
                                SpPartTCWAYSort.Add(TPart);
                            }
                            else
                                Prodol = "N";
                        }
                        foreach (Part TPart in SpPartTCWAYSort)
                        {
                            TPosrL TTpstrL = new TPosrL();
                            //Point3d TPostr = new Point3d();
                            string Visot = "";
                            //Application.ShowAlertDialog(TPart.CWName);
                            if (TPart.CompName.Contains("CTRSZ") == true)
                            {
                                StartPoint = TPart.StartPoint;
                                Point3d TPostr = new Point3d(Convert.ToDouble(StartPoint.Split(' ')[0]), Convert.ToDouble(StartPoint.Split(' ')[1]), 0);
                                COMPONm = TPart.CompName.Split('_');
                                if (COMPONm.Length > 1) { if (Double.TryParse(COMPONm[1], out Delt)) TTpstrL.NDelt(Convert.ToDouble(COMPONm[1])); }
                                Visot = StartPoint.Split(' ')[2];
                                TTpstrL.NTPoint(TPostr);
                                TTpstrL.NVisot(Visot);
                                SpTPostrL.Add(TTpstrL);
                            }
                            else
                            {
                                StartPoint = TPart.Coordinates;
                                Point3d TPostr = new Point3d(Convert.ToDouble(StartPoint.Split(' ')[0]), Convert.ToDouble(StartPoint.Split(' ')[1]), 0);
                                COMPONm = TPart.CompName.Split('_');
                                if (COMPONm.Length > 1) { if (Double.TryParse(COMPONm[1], out Delt)) TTpstrL.NDelt(Convert.ToDouble(COMPONm[1])); }
                                Visot = StartPoint.Split(' ')[2];
                                TTpstrL.NTPoint(TPostr);
                                TTpstrL.NVisot(Visot);
                                SpTPostrL.Add(TTpstrL);
                            }
                        }
                        if (SpPartTCWAYSort.Last().CompName.Contains("CTRSZ"))
                        {
                            TPosrL TTpstrL = new TPosrL();
                            StartPoint = SpPartTCWAYSort.Last().EndPoint;
                            Point3d TPostr = new Point3d(Convert.ToDouble(StartPoint.Split(' ')[0]), Convert.ToDouble(StartPoint.Split(' ')[1]), 0);
                            string Visot = StartPoint.Split(' ')[2];
                            TTpstrL.NTPoint(TPostr);
                            TTpstrL.NVisot(Visot);
                            SpTPostrL.Add(TTpstrL);
                        }
                        if (tCWAY.CompName.Contains("-100_")) { Dles = 108; COMPON = "#ЛестУсил100"; }
                        if (tCWAY.CompName.Contains("-200_")) { Dles = 208; COMPON = "#ЛестУсил200"; }
                        if (tCWAY.CompName.Contains("-300_")) { Dles = 308; COMPON = "#ЛестУсил300"; }
                        if (tCWAY.CompName.Contains("-400_")) { Dles = 408; COMPON = "#ЛестУсил400"; }
                        if (tCWAY.CompName.Contains("-500_")) { Dles = 508; COMPON = "#ЛестУсил500"; }
                        if (tCWAY.CompName.Contains("-600_")) { Dles = 608; COMPON = "#ЛестУсил600"; }
                        if (tCWAY.CompName.Contains("-700_")) { Dles = 708; COMPON = "#ЛестУсил700"; }
                        if (tCWAY.CompName.Contains("-800_")) { Dles = 808; COMPON = "#ЛестУсил800"; }
                        List<TPosrL> SpTPostrL_Per = new List<TPosrL>();
                        bool Prov = true;
                        foreach (TPosrL TT in SpTPostrL)
                        {
                            double mid = 0;
                            string[] COMPONm1 = tCWAY.CompName.Split('_');
                            if (COMPONm1.Length > 1)
                            { if (Double.TryParse(COMPONm1[1], out Delt)) mid = Convert.ToDouble(COMPONm1[1]) / 2; }
                            TPosrL NT = PerXYZ(TT, OSItl, MSKtl, PSKtl, mid,Convert.ToDouble(textBox29.Text));
                            if (Math.Abs(Convert.ToDouble(NT.Visot)) > GlubinaSR) Prov = false;
                            if (this.checkBox7.Checked)
                            {
                                if (NT.TPoint.X >= MIN.X & NT.TPoint.X <= MAX.X & NT.TPoint.Y >= MIN.Y & NT.TPoint.Y <= MAX.Y)
                                 SpTPostrL_Per.Add(NT);
                            }
                            else
                           SpTPostrL_Per.Add(NT); 
                        }
                        pruning_the_vertical_ends(ref SpTPostrL_Per,ref Prov);
                        if (Prov)
                        {
                            UdalToch(ref SpTPostrL_Per);
                            Lest_N_INtr_SP(SpTPostrL_Per, COMPON, Dles, "Лес", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 400 + Dles / 2, 1200, this.checkBox6.Checked ? tCWAY.ModuleName : this.textBox1.Text, NAME,checkBox1.Checked);
                        }
                    }
                }
            }
            this.Show();
        }//построить трассы из TRIBON
        private void button28_Click(object sender, EventArgs e)
        {
            string BlokType = "DB";
            string Dir = this.textBox14.Text;
            //string BlocCatal =Path.GetDirectoryName(@Catal);
            Blok = "Плоскость_пр";
            //Application.ShowAlertDialog();
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                //Application.ShowAlertDialog(@Catal);
                InsBlockRef($"{@Dir}Плоскость_пр.dwg", Blok, BlokType);
                SetDynamicBlkProperty(Blok, "", "", "", "", "", "", "", 0, "", "", 0, "Плоскости");
            }
            this.Show();
        }//плоскость построения
        #region вкладка детали полилиниями
        private void button16_Click(object sender, EventArgs e)
        {
            if (this.dataGridView6.CurrentRow.Cells[0].Value != null)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    SOZDlov("POM");
                    ZapisSlov("POM", this.textBox1.Text);
                }
                string Tipor = "";
                string Zagal = "";
                string Povorot = "";
                string Soed = "";
                string Xvost = "";
                string DopKR = "";
                string DopSoed = "";
                double Hsirin = 0;
                double HSagSoed = 0;
                double HSagKr = 0;
                if (this.dataGridView6.CurrentRow.Cells[0].Value != null) Tipor = this.dataGridView6.CurrentRow.Cells[0].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[1].Value != null) Hsirin = Convert.ToDouble(this.dataGridView6.CurrentRow.Cells[1].Value.ToString());
                if (this.dataGridView6.CurrentRow.Cells[2].Value != null) Zagal = this.dataGridView6.CurrentRow.Cells[2].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[3].Value != null) Povorot = this.dataGridView6.CurrentRow.Cells[3].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[4].Value != null) Soed = this.dataGridView6.CurrentRow.Cells[4].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[5].Value != null) Xvost = this.dataGridView6.CurrentRow.Cells[5].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[6].Value != null) DopKR = this.dataGridView6.CurrentRow.Cells[6].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[7].Value != null) DopSoed = this.dataGridView6.CurrentRow.Cells[7].Value.ToString();
                if (this.dataGridView6.CurrentRow.Cells[8].Value != null & this.dataGridView6.CurrentRow.Cells[8].Value != "") HSagKr = Convert.ToDouble(this.dataGridView6.CurrentRow.Cells[8].Value.ToString());
                if (this.dataGridView6.CurrentRow.Cells[9].Value != null & this.dataGridView6.CurrentRow.Cells[9].Value != "") HSagSoed = Convert.ToDouble(this.dataGridView6.CurrentRow.Cells[9].Value.ToString());
                //String Adr = this.textBox6.Text;
                if (this.textBox6.Text != "") DopKR = DopKR + "#" + this.textBox6.Text;
                if (this.textBox11.Text != "") DopSoed = DopSoed + "#" + this.textBox11.Text;
                if (Tipor != "" & Zagal == "Лес") Lest_INtr_SP(Tipor, Hsirin, Zagal, Povorot, Xvost, Soed, "", 400 + Hsirin / 2, HSagKr,checkBox1.Checked);
                if (Tipor != "" & Zagal == "Тр") TR_INtr_SP(Tipor, Hsirin, Zagal, Povorot, Xvost, Soed, "", 6 * Hsirin, HSagKr, DopSoed, DopKR, HSagSoed,checkBox1.Checked);
            }
        }//создать полилинию
        private void button25_Click(object sender, EventArgs e)
        {
            this.Hide();
            SpPOZ.Sort(delegate (POZIZIA x, POZIZIA y) { return x.NOMpoz.CompareTo(y.NOMpoz); });
            Tabl(ref SpPOZ);
            this.Show();
        }//таблица спецификации
        private void button26_Click(object sender, EventArgs e)
        {
            double B = 330;
            int max = 0;
            int k = 0;
            string SpToch = "";
            double Dist = 999999999999999999;
            //List<POZIZIA> SpPOZ_XVOST = new List<POZIZIA>();
            List<TPosrL> spTles = new List<TPosrL>();
            List<Lest> spLes = new List<Lest>();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Start a transaction
            this.Hide();
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                SOZD_Sp_TP_Lest(ref spTles);
                SOZD_Sp_TP_Lest_BL(ref spTles);
                SBORpl_Dla_T(ref spLes, spTles, SpNamenkl, SpPOZ);
                foreach (Lest Ind in spLes) { if (Ind.SpT.Count > max) max = Ind.SpT.Count; }
                Transaction tr = acCurDb.TransactionManager.StartTransaction();
                using (tr)
                {
                    PromptPointResult pPtRes;
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                    Point3d Toch1 = pPtRes.Value;
                    Point3d TochT = Toch1;
                    List<string> SpT = new List<string>();
                    SpT.Add("X0 Y0 Z0");
                    for (int i = 1; i < max; i++) SpT.Add("dX" + i.ToString() + " dY" + i.ToString() + " dZ" + i.ToString());
                    Strok_tabl_L("Базавая т", "Номер л.", "Позиция", "Хвостовик", "Планка", SpT, max, TochT, B);
                    TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                    B = 280;
                    foreach (Lest Ind in spLes)
                    {
                        PLOS Tplos = new PLOS();
                        Dist = 999999999999999999;
                        foreach (PLOS pl in SpPLOS) { if (Ind.SpT[0].Koord.DistanceTo(pl.Psk) < Dist) { Dist = Ind.SpT[0].Koord.DistanceTo(pl.Psk); Tplos = pl; } }
                        SpT.Clear();
                        Ind.SpT.Reverse();
                        k = 0;
                        foreach (TPodk TT in Ind.SpT)
                        {
                            if (k == 0)
                                SpT.Add((TT.Koord.X - Tplos.Psk.X).ToString("0.") + " " + (TT.Koord.Y - Tplos.Psk.Y).ToString("0.") + " " + TT.Sist);
                            else
                                SpT.Add((TT.Koord.X - Ind.SpT[0].Koord.X).ToString("0.") + " " + (TT.Koord.Y - Ind.SpT[0].Koord.Y).ToString("0.") + " " + (Convert.ToDouble(TT.Sist) - Convert.ToDouble(Ind.SpT[0].Sist)).ToString("0."));
                            k = k + 1;
                        }
                        Strok_tabl_L(Tplos.Vid, Ind.Nazv, Ind.poz, Ind.xvost, Ind.perem, SpT, max, TochT, B);
                        TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                    }
                    tr.Commit();
                }
            }
            this.Show();
        }//Таблица лестниц
        private void button27_Click(object sender, EventArgs e)
        {
            double B = 330;
            int max = 0;
            int k = 0;
            string SpToch = "";
            double Dist = 999999999999999999;
            List<TPosrL> spTles = new List<TPosrL>();
            List<Lest> spLes = new List<Lest>();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                SOZD_Sp_TP_Lest(ref spTles);
                SOZD_Sp_TP_Lest_BL(ref spTles);
                SBORpl_Dla_T(ref spLes, spTles, SpNamenkl, SpPOZ);
                foreach (Lest Ind in spLes) { if (Ind.SpT.Count > max) max = Ind.SpT.Count; }
                Transaction tr = acCurDb.TransactionManager.StartTransaction();
                using (tr)
                {
                    PromptPointResult pPtRes;
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    //PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                    //pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                    //Point3d Toch1 = pPtRes.Value;
                    //Point3d TochT = Toch1;
                    List<string> SpT = new List<string>();
                    //SpT.Add("X0 Y0 Z0");
                    //for (int i = 1; i < max; i++) SpT.Add("dX" + i.ToString() + " dY" + i.ToString() + " dZ" + i.ToString());
                    //TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                    //B = 280;
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\МАРШРУТ\Трассы.txt", false, Encoding.GetEncoding("Windows-1251")))
                    {
                        foreach (Lest Ind in spLes)
                        {
                            PLOS Tplos = new PLOS();
                            foreach (PLOS pl in SpPLOS) { if (Ind.SpT[0].Koord.DistanceTo(pl.Psk) < Dist) { Dist = Ind.SpT[0].Koord.DistanceTo(pl.Psk); Tplos = pl; } }
                            SpT.Clear();
                            Ind.SpT.Reverse();
                            k = 0;
                            SpToch = "";
                            foreach (TPodk TT in Ind.SpT)
                            {
                                //SpToch = SpToch + "," + (Tplos.Msk.X + (TT.Koord.X - Tplos.Psk.X)).ToString("0.") + " " + (Tplos.Msk.Y + (TT.Koord.Y - Tplos.Psk.Y)).ToString("0.") + " " + Tplos.Msk.Z.ToString("0.") ;
                                SpToch = SpToch + "," + (Tplos.Msk.X + (TT.Koord.X - Tplos.Psk.X)).ToString("0.") + " " + (Tplos.Msk.Y + (TT.Koord.Y - Tplos.Psk.Y)).ToString("0.") + " " + Tplos.Msk.Z.ToString("0.");
                                //if (k == 0)
                                //    SpT.Add((TT.Koord.X - Tplos.Psk.X).ToString("0.") + " " + (TT.Koord.Y - Tplos.Psk.Y).ToString("0.") + " " + TT.Sist);
                                //else
                                //    SpT.Add((TT.Koord.X - Ind.SpT[0].Koord.X).ToString("0.") + " " + (TT.Koord.Y - Ind.SpT[0].Koord.Y).ToString("0.") + " " + (Convert.ToDouble(TT.Sist) - Convert.ToDouble(Ind.SpT[0].Sist)).ToString("0."));
                                //k = k + 1;
                            }
                            file.WriteLine(Ind.Nazv + ":" + Ind.Comp + ":" + Ind.pom + ":" + SpToch + ":" + Tplos.Msk.ToString() + ":" + Tplos.Psk.ToString());
                        }
                        //Strok_tabl_L(Tplos.Vid, Ind.Nazv, Ind.poz, Ind.xvost, Ind.perem, SpT, max, TochT, B);
                        //TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                    }
                    tr.Commit();
                }
            }
        }//лестницы в TXT
        private void button20_Click(object sender, EventArgs e)
        {
            this.Hide();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Point3d T1 = new Point3d(0, 0, 0);
                //FPovorot(T1, "","",0,"",208);
                //    //Document doc = Application.DocumentManager.MdiActiveDocument;
                //    Database db = doc.Database;
                //    Editor ed = doc.Editor;
                //    PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");

                //    PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
                //        if (prEntResult.Status != PromptStatus.OK)
                //        {
                //         ed.WriteMessage("Ошибка...");
                //         return;
                //        }
                //    using (Transaction Tx = db.TransactionManager.StartTransaction())
                //   {
                //       BlockTableRecord btr = (BlockTableRecord)Tx.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                //       Line bref1 = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Line;
                //       Point3d stPoint = bref1.StartPoint;
                //       Point3d enPoint = bref1.EndPoint;
                //       Vector3d vek1 = stPoint.GetVectorTo(enPoint);
                //       Point3d T1 = polar1(new Point3d(0, 0, 0), vek1, stPoint.DistanceTo(enPoint));
                //       Line acLine1 = new Line();
                //       acLine1.SetDatabaseDefaults();
                //       acLine1.StartPoint = new Point3d(0, 0, 0);
                //       acLine1.EndPoint = T1;
                //       acLine1.ColorIndex = 1;
                //       btr.AppendEntity(acLine1);
                //       Tx.AddNewlyCreatedDBObject(acLine1, true);
                //       Tx.Commit();
                //   }
            }
            this.Show();
        }//проверка функции
        private void button18_Click(object sender, EventArgs e)
        {

        }//изменить расположение файла со спецификацией лесниц
        #endregion
        #region вкладка база

        private void button7_Click(object sender, EventArgs e)
        {
            string label = "";
            string doc_N = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("DOC");
                ZapisSlov("DOC", this.textBox4.Text);
                //LoadData();
                List<POZIZIA> SpPOZ_KK = new List<POZIZIA>();
                SBORBl_KK(ref SpPOZ_KK);
                SpPOZ_KK.Sort(delegate (POZIZIA x, POZIZIA y) { return x.Hozain.CompareTo(y.Hozain); });
                for (int i = 0; i < this.dataGridView2.RowCount - 1; i++)
                {
                    label = "";
                    doc_N = "";
                    if (this.dataGridView2.Rows[i].Cells[1].Value != null) label = this.dataGridView2.Rows[i].Cells[1].Value.ToString();
                    if (this.dataGridView2.Rows[i].Cells[2].Value != null) doc_N = this.dataGridView2.Rows[i].Cells[2].Value.ToString();
                    if (SpPOZ_KK.Exists(x => x.RAB == label & x.LINK == doc_N))
                    {
                        this.dataGridView2.Rows[i].Cells[0].Value = SpPOZ_KK.Find(x => x.RAB == label & x.LINK == doc_N).Hozain;
                    }
                }
                this.dataGridView3.Rows.Clear();
                double Nom;
                foreach (POZIZIA Poz in SpPOZ_KK)
                {
                    if (double.TryParse(Poz.HtoEto.TrimStart('C', 'С', 'В', 'К'), out Nom))
                        Nom = Convert.ToDouble(Poz.HtoEto.TrimStart('C', 'С', 'В', 'К'));
                    else
                        Nom = 0;
                    this.dataGridView3.Rows.Add(Nom, Poz.HtoEto, "", "");
                }
                //foreach (POZIZIA Poz in SpPOZ_KK) { this.dataGridView3.Rows.Add("", Poz.HtoEto, "", ""); }
            }
        }//найти по запросу
        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockTableRecord acBlkTblRec;
            BlockTable acBlkTbl;
            using (DocumentLock docLock = doc.LockDocument())
            {
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("Точка начала построения");
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d XYZ = pPtRes.Value;
                Point3d T1 = new Point3d();
                Point3d T2 = new Point3d();
                Point3d T3 = new Point3d();
                Point3d T4 = new Point3d();
                Point3d T5 = new Point3d();
                Point3dCollection TkoorPL = new Point3dCollection();
                string label = "";
                string doc_N = "";
                string label_Kor = "";
                double A = 165;
                double B = 500;
                double a = 115;
                double b = 220;
                StrokaTabl(T1, T2, T3, T4, T5, XYZ, "Позиция", "N черт. насыщения", "Поз. по черт. насыщения");
                XYZ = new Point3d(XYZ.X, XYZ.Y - 165, XYZ.Z);
                for (int i = 0; i < this.dataGridView3.RowCount - 1; i++)
                {
                    label = "";
                    doc_N = "";
                    label_Kor = "";
                    if (this.dataGridView3.Rows[i].Cells[1].Value != null) label = this.dataGridView3.Rows[i].Cells[1].Value.ToString();
                    if (this.dataGridView3.Rows[i].Cells[2].Value != null) doc_N = this.dataGridView3.Rows[i].Cells[2].Value.ToString();
                    if (this.dataGridView3.Rows[i].Cells[3].Value != null) label_Kor = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                    StrokaTabl(T1, T2, T3, T4, T5, XYZ, label, doc_N, label_Kor);
                    XYZ = new Point3d(XYZ.X, XYZ.Y - 165, XYZ.Z);
                }
            }
            this.Show();
        }//Вывести таблицу


        public void StrokaTabl(Point3d T1, Point3d T2, Point3d T3, Point3d T4, Point3d T5, Point3d XYZ, string label, string doc_N, string label_Kor)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockTableRecord acBlkTblRec;
            BlockTable acBlkTbl;
            Point3dCollection TkoorPL = new Point3dCollection();
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, XYZ, 500, 165, 115, 120);

                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                DBText Poz = new DBText();
                Poz.SetDatabaseDefaults();
                Poz.Position = T5;
                Poz.TextString = label;
                Poz.Height = 50;
                acBlkTblRec.AppendEntity(Poz);
                Tx.AddNewlyCreatedDBObject(Poz, true);
                FPoly2d(TkoorPL);

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, new Point3d(XYZ.X + 500, XYZ.Y, XYZ.Z), 1335, 165, 115, 120);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                DBText DOC_N = new DBText();
                DOC_N.SetDatabaseDefaults();
                DOC_N.Position = T5;
                DOC_N.TextString = doc_N;
                DOC_N.Height = 50;
                acBlkTblRec.AppendEntity(DOC_N);
                Tx.AddNewlyCreatedDBObject(DOC_N, true);
                FPoly2d(TkoorPL);

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, new Point3d(XYZ.X + 1835, XYZ.Y, XYZ.Z), 1250, 165, 115, 120);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                DBText Poz_Korp = new DBText();
                Poz_Korp.SetDatabaseDefaults();
                Poz_Korp.Position = T5;
                Poz_Korp.TextString = label_Kor;
                Poz_Korp.Height = 50;
                acBlkTblRec.AppendEntity(Poz_Korp);
                Tx.AddNewlyCreatedDBObject(Poz_Korp, true);
                FPoly2d(TkoorPL);

                XYZ = new Point3d(XYZ.X, XYZ.Y - 165, XYZ.Z);
                Tx.Commit();
            }
        }

        public void RasKoor(ref Point3d T1, ref Point3d T2, ref Point3d T3, ref Point3d T4, ref Point3d T5, Point3d XYZ, double A, double B, double a, double b)
        {
            T1 = new Point3d(XYZ.X, XYZ.Y, 0);
            T2 = new Point3d(XYZ.X + A, XYZ.Y, 0);
            T3 = new Point3d(XYZ.X + A, XYZ.Y - B, 0);
            T4 = new Point3d(XYZ.X, XYZ.Y - B, 0);
            T5 = new Point3d(XYZ.X + a, XYZ.Y - b, 0);
        }
        public void FPoly2d(Point3dCollection TkoorPL)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            // Append the point to the database
            using (tr1)
            {
                Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = 0;
                poly.Closed = true;
                //poly.Layer = "Kabeli";
                int i = 0;
                foreach (Point3d pt in TkoorPL)
                {
                    poly.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    i = i + 1;
                }
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);

                btr.Dispose();
                tr1.Commit();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (this.dataGridView3.CurrentRow.Cells[0].Value != null)
            {
                String FindLab = this.dataGridView3.CurrentRow.Cells[1].Value.ToString();
                String FindDoc = this.dataGridView3.CurrentRow.Cells[2].Value.ToString();
                List<ObjectId> SpOBJID = new List<ObjectId>();
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    Database db = doc.Database;
                    Editor ed = doc.Editor;
                    SBORBl_FIND_KK(ref SpOBJID, FindLab, FindDoc);
                    ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                    ed.SetImpliedSelection(idarrayEmpty);
                }
            }
        }//найти
        private void button10_Click(object sender, EventArgs e)
        {
            this.Hide();
            double dMas = Convert.ToDouble(Mas);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                string stLabel = "-";
                int Schet;
                Database db = doc.Database;
                Editor ed = doc.Editor;
                PromptEntityOptions prEntOptions = new PromptEntityOptions("Выберите вставку динамического блока...");
                PromptEntityResult prEntResult = ed.GetEntity(prEntOptions);
                if (prEntResult.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Ошибка...");
                    return;
                }
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    BlockTableRecord acBlkTblRec;
                    acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Entity bref = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 8) { stLabel = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    PromptPointResult pPtRes;
                    PromptPointOptions pPtOpts = new PromptPointOptions("");
                    Point3d ptStart = prEntResult.PickedPoint;
                    pPtOpts.UseBasePoint = true;
                    pPtOpts.BasePoint = ptStart;

                    pPtRes = doc.Editor.GetPoint(pPtOpts);
                    Point3d ZKr = pPtRes.Value;


                    Point3d ptEnd = new Point3d();
                    Point3d txt = new Point3d();
                    Point3d txtNP = new Point3d();
                    Point3d txtPP = new Point3d();
                    Point3d TPol1 = new Point3d();
                    Point3d TPol2 = new Point3d();
                    RaschKoordin(ptStart, ZKr, ref ptEnd, ref txt, ref txtNP, ref txtPP, ref TPol1, ref TPol2, "");


                    Line acLine = new Line(ptStart, ptEnd);
                    acLine.SetDatabaseDefaults();
                    acLine.ColorIndex = 1;
                    acLine.Layer = "Насыщение";
                    acBlkTblRec.AppendEntity(acLine);
                    Tx.AddNewlyCreatedDBObject(acLine, true);


                    Circle KrugPOZ = new Circle();
                    KrugPOZ.SetDatabaseDefaults();
                    KrugPOZ.Center = ZKr;
                    KrugPOZ.Radius = 100 * dMas;
                    KrugPOZ.ColorIndex = 1;
                    KrugPOZ.Layer = "Насыщение";
                    acBlkTblRec.AppendEntity(KrugPOZ);
                    Tx.AddNewlyCreatedDBObject(KrugPOZ, true);


                    Circle KrugPOZVn = new Circle();
                    KrugPOZVn.SetDatabaseDefaults();
                    KrugPOZVn.Center = ZKr;
                    KrugPOZVn.Radius = 80 * dMas;
                    KrugPOZVn.ColorIndex = 1;
                    KrugPOZVn.Layer = "Насыщение";
                    acBlkTblRec.AppendEntity(KrugPOZVn);
                    Tx.AddNewlyCreatedDBObject(KrugPOZVn, true);


                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = txt;
                    Poz.Height = 50 * dMas;
                    Poz.ColorIndex = 2;
                    Poz.TextString = stLabel;
                    Poz.Layer = "Насыщение";
                    acBlkTblRec.AppendEntity(Poz);
                    Tx.AddNewlyCreatedDBObject(Poz, true);
                    Tx.Commit();
                }
            }
            this.Show();
        }//указать позицию

        static public void SBORBl_KK(ref List<POZIZIA> SpPOZ_KK)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPOLnNaim = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей изображенных блоками...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stHOZ = "";
                    stRab = "";
                    sLINK = "";
                    dVisot = 0;
                    dDlin = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); stPOLnNaim = stKomp; }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Convert.ToDouble(prop.Value.ToString()) / 1000; }
                        }
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                    if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                    if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                    if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString) / 1000; }
                                }
                            }
                        }
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    else
                    {
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                                if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()) / 1000; }
                                if (Schet == 10) { stRab = value.Value.ToString(); }
                                if (Schet == 11) { sLINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    POZIZIA tPOZ = new POZIZIA();
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    string[] VirezM = stKompM.First().Split('%');
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    tPOZ.NpolnNAIM(bref.Handle.ToString());
                    tPOZ.NHozain(stHOZ);
                    tPOZ.NRAB(stRab);
                    tPOZ.NLINK(sLINK);
                    if (stHtoEto != "" && stHtoEto != "Лес" && stHtoEto != "тр" && stHtoEto != "хв")
                    {
                        SpPOZ_KK.Add(tPOZ);
                    }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками
        static public void SBORBl_FIND_KK(ref List<ObjectId> SpOBJID, string FIND_K, string FIND_D)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            string stPOLnNaim = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 3);
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
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stHOZ = "";
                    stRab = "";
                    sLINK = "";
                    stPOLnNaim = "";
                    dVisot = 0;
                    dDlin = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); stPOLnNaim = stKomp; }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Convert.ToDouble(prop.Value.ToString()) / 1000; }
                        }
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                    if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                    if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                    if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); }
                                }
                            }
                        }
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    else
                    {
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                                if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 10) { stRab = value.Value.ToString(); }
                                if (Schet == 11) { sLINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    if (FIND_K == stHOZ & FIND_D == sLINK) { SpOBJID.Add(acSSObj.ObjectId); }
                    SpDOP_POZ.Clear();
                }
                Tx.Commit();
            }
        }//поиск блоков по компоненту

        public static void SetConn()
        {
            connection = new SQLiteConnection("Data Source=D:/DBtest.db;Version=3;New=false;Compress=true");
        }
        public void LoadData()
        {
            SetConn();
            connection.Open();
            command = connection.CreateCommand();
            string ComTXT = "SELECT label as 'Позиция по черт. насыщения',doc_name  as 'Номер документа',position_description  as 'Наименование' FROM DOC WHERE position_description LIKE '%" + this.textBox4.Text + "%'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(ComTXT, connection);
            DS1.Reset();
            adapter.Fill(DS1);
            DT1 = DS1.Tables[0];
            this.dataGridView2.DataSource = DT1;
            connection.Close();
        }

        #endregion
        #region вкладка ..sp

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            this.dataGridView4.Rows.Clear();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                File = ofd.FileName;
                this.label8.Text = File;
                this.dataGridView4.Rows.Clear();
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    SOZDlov("sp");
                    ZapisSlov("sp", File);
                    XMLr(File);
                }
            }
            else
                return;
        }//считать sp по новому адрсу
        private void button12_Click(object sender, EventArgs e)
        {
            string File = this.label8.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            this.dataGridView4.Rows.Clear();
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("sp");
                ZapisSlov("sp", File);
                XMLr(File);
            }
        }//считать sp по сохраненому адресу
        private void button13_Click(object sender, EventArgs e)
        {
            string fileName = this.label8.Text;
            string caption_description_G2 = "";
            string label = "";
            string quantity = "";
            string quantitySP = "";
            XDocument doc = XDocument.Load(fileName);
            for (int i = 0; i < this.dataGridView4.RowCount - 1; i++)
            {
                caption_description_G2 = "";
                label = "";
                quantity = "";
                quantitySP = "";
                if (this.dataGridView4.Rows[i].Cells[0].Value != null) label = this.dataGridView4.Rows[i].Cells[0].Value.ToString();
                if (this.dataGridView4.Rows[i].Cells[5].Value != null) quantitySP = this.dataGridView4.Rows[i].Cells[5].Value.ToString();
                if (this.dataGridView4.Rows[i].Cells[6].Value != null) quantity = this.dataGridView4.Rows[i].Cells[6].Value.ToString();
                if (this.dataGridView4.Rows[i].Cells[9].Value != null) caption_description_G2 = this.dataGridView4.Rows[i].Cells[9].Value.ToString();
                IEnumerable<XElement> FindZ = from el in doc.Descendants("caption") where ((string)el.Attribute("description") == caption_description_G2) select el;
                if (FindZ.Count() != 0 & quantity != "")
                {
                    IEnumerable<XElement> Find = from el in FindZ.Descendants("position") where (string)el.Attribute("label") == label select el;
                    foreach (XElement tEl in Find) { tEl.SetAttributeValue("quantity", quantity); }
                }
            }
            doc.Save(fileName);
            this.dataGridView4.Rows.Clear();
            XMLr(fileName);
        }//сохранить изменения в SP


        public void XMLr(string path)
        {
            List<POZIZIA> SpPOZ_XML = new List<POZIZIA>();
            this.label7.Text = path;
            string doc_name = "";
            string caption_description = "";
            string caption_description_G2 = "";
            string caption_description_G3 = "";
            string label = "";
            string name = "";
            string position_description = "";
            string mass = "";
            string code = "";
            string weighting_code = "";
            string quantity = "";
            string placement = "";
            string order_code = "";
            string unit_code = "";
            string project = "";
            //задаем путь к нашему рабочему файлу XML
            //string fileName = "base.xml";
            //читаем данные из файла
            XDocument doc = XDocument.Load(path);
            //проходим по каждому элементу в найшей library
            //(этот элемент сразу доступен через свойство doc.Root)
            project = doc.Root.Element("document").Attribute("project").Value;
            this.label1.Text = project;
            foreach (XElement el in doc.Root.Elements())
            {
                doc_name = el.Attribute("name").Value.ToString();
                //выводим в цикле названия всех дочерних элементов и их значения
                foreach (XElement element in el.Elements())
                {
                    if (element.Name == "body")
                    {
                        foreach (XElement element1 in element.Elements())
                        {
                            foreach (XAttribute attr in element1.Attributes())
                            {
                                if (attr.Name == "description") caption_description = attr.Value;
                                if (attr.Name == "label") label = attr.Value;
                            }
                            if (label == "Г2") { caption_description_G2 = caption_description; caption_description_G3 = ""; }
                            if (label == "Г3") caption_description_G3 = caption_description;
                            //this.dataGridView1.Rows.Add(label, caption_description);
                            foreach (XElement element2 in element1.Elements())
                            {
                                label = "";
                                name = "";
                                position_description = "";
                                mass = "";
                                code = "";
                                weighting_code = "";
                                quantity = "";
                                placement = "";
                                order_code = "";
                                unit_code = "";
                                foreach (XAttribute attr in element2.Attributes())
                                {
                                    if (attr.Name == "description") position_description = attr.Value;
                                    if (attr.Name == "label") label = attr.Value;
                                    if (attr.Name == "mass") mass = attr.Value;
                                    if (attr.Name == "code") code = attr.Value;
                                    if (attr.Name == "name") name = attr.Value;
                                    if (attr.Name == "code") code = attr.Value;
                                    if (attr.Name == "weighting_code") weighting_code = attr.Value;
                                    if (attr.Name == "quantity") quantity = attr.Value;
                                    if (attr.Name == "order_code") order_code = attr.Value;
                                    if (attr.Name == "unit_code") unit_code = attr.Value;
                                }
                                POZIZIA NPoz = new POZIZIA();
                                NPoz.NNOMpoz(label);
                                NPoz.NCompon(name);
                                NPoz.NpolnNAIM(position_description);
                                NPoz.NHtoEto(code);
                                NPoz.NLINK(quantity);
                                NPoz.NKEI(order_code);
                                NPoz.NHozain(caption_description_G2);
                                SpPOZ_XML.Add(NPoz);
                            }
                        }
                    }
                }
            }
            string tKol = "";
            string tComp = "";
            int i = -1;
            foreach (POZIZIA Tpoz in SpPOZ_XML)
            {
                i = i + 1;
                if (SpPOZ.Exists(x => x.NOMpoz == Tpoz.NOMpoz))
                {
                    tKol = SpPOZ.Find(x => x.NOMpoz == Tpoz.NOMpoz).Kol.ToString("0.###");
                    tComp = SpPOZ.Find(x => x.NOMpoz == Tpoz.NOMpoz).Compon;
                    if (tKol == Tpoz.LINK)
                    {
                        this.dataGridView4.Rows.Add(Tpoz.NOMpoz, Tpoz.Compon, tComp, Tpoz.polnNAIM, Tpoz.HtoEto, Tpoz.LINK, tKol, "", Tpoz.KEI, Tpoz.Hozain);
                        this.dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    else
                    {
                        this.dataGridView4.Rows.Add(Tpoz.NOMpoz, Tpoz.Compon, tComp, Tpoz.polnNAIM, Tpoz.HtoEto, Tpoz.LINK, tKol, "", Tpoz.KEI, Tpoz.Hozain);
                        this.dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                    }
                    if (Tpoz.Compon != tComp & Tpoz.HtoEto != tComp)
                    {
                        this.dataGridView4.Rows[i].Cells[1].Style.BackColor = Color.LightSeaGreen;
                        this.dataGridView4.Rows[i].Cells[2].Style.BackColor = Color.LightSeaGreen;
                    }
                }
                else
                    this.dataGridView4.Rows.Add(Tpoz.NOMpoz, Tpoz.Compon, "", Tpoz.polnNAIM, Tpoz.HtoEto, Tpoz.LINK, "", "", Tpoz.KEI, Tpoz.Hozain);
            }
        }//чтение XML


        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "2.5";
            ZapisSlov("MAS#", "2.5");

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "1";
            ZapisSlov("MAS#", "1");
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "0.5";
            ZapisSlov("MAS#", "0.5");
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "1.25";
            ZapisSlov("MAS#", "1.25");
        }





        #endregion
        private void button19_Click(object sender, EventArgs e)
        {
            this.Hide();
            string Rez = "N";
            string Rad = this.textBox13.Text;
            if (this.checkBox3.Checked) Rez = "Y";
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
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    BlockTableRecord acBlkTblRec;
                    // Open Model space for write
                    acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Line acLine = new Line(T1, T2);
                    acLine.SetDatabaseDefaults();
                    if (this.checkBox3.Checked) acLine.ColorIndex = 4; else acLine.ColorIndex = 3;
                    acLine.Layer = "Насыщение";
                    acLine.XData = new ResultBuffer
                        (
                        new TypedValue(1001, "LAUNCH01"),
                        new TypedValue(1000, "Препядствие"),
                        new TypedValue(1000, this.textBox12.Text),
                        new TypedValue(1000, this.textBox13.Text),
                        new TypedValue(1000, Rez),
                        new TypedValue(1000, this.textBox30.Text),
                        new TypedValue(1000, this.textBox31.Text),
                        new TypedValue(1000, this.textBox32.Text),
                        new TypedValue(1000, this.textBox33.Text),
                        new TypedValue(1000, this.textBox34.Text)
                        );
                    acBlkTblRec.AppendEntity(acLine);
                    Tx.AddNewlyCreatedDBObject(acLine, true);
                    Tx.Commit();
                }
            }
            this.Show();
        }//построить препядствие
        private void button29_Click(object sender, EventArgs e)
        {
            double A = Convert.ToDouble(this.textBox15.Text);
            double B = Convert.ToDouble(this.textBox16.Text);
            double C = Convert.ToDouble(this.textBox17.Text);
            double D = Convert.ToDouble(this.textBox26.Text);
            bool OutPlane = true;
            List<Point3d> ListPoint1 = new List<Point3d>();
            List<TPosrL> SpGibov = new List<TPosrL>();
            List<TPosrL> SpMetok = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {    
                SBORBl_Gibov(ref SpGibov, ref SpMetok);
                if (this.checkBox4.Checked == true)
                {
                    Transaction tr1 = db.TransactionManager.StartTransaction();
                    using (tr1)
                    {
                        foreach (TPosrL ID in SpMetok) { Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity; Obj.Erase(); }
                        SpMetok.Clear();
                        tr1.Commit();
                    }
                }
                foreach (TPosrL TT in SpGibov)
                {
                OutPlane = true;
                    if (!(SpMetok.Exists(x => x.Komp == TT.Komp))) {
                        Point3d T0 = TT.TPoint;
                        Point3d T1 = T0;
                        Point3d T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                        Point3d T3 = new Point3d(T1.X, T1.Y + 5, 0);
                        foreach (PLOS TP in SpPLOS)
                        {
                            if (TT.TPoint.X > TP.min.X & TT.TPoint.X < TP.max.X & TT.TPoint.Y > TP.min.Y & TT.TPoint.Y < TP.max.Y)
                            {
                                if (TT.TPoint1.DistanceTo(TP.max) < TT.TPoint.DistanceTo(TP.max)) T0 = TT.TPoint1;
                                double DeltX = TP.min.X + Math.Abs(TP.max.X - TP.min.X) / 2;
                                double DeltY = TP.min.Y + Math.Abs(TP.max.Y - TP.min.Y) / 2;
                                OutPlane = false;
                                Point3d Zentr = new Point3d(DeltX, DeltY, 0);
                                if (T0.X > Zentr.X & T0.Y > Zentr.Y)
                                {
                                    if (TT.TPoint1.DistanceTo(TP.min_v) < TT.TPoint.DistanceTo(TP.min_v)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.RigthUp);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                                    T3 = new Point3d(T1.X, T1.Y + 5, 0);
                                }
                                if (T0.X < Zentr.X & T0.Y > Zentr.Y)
                                {
                                    if (TT.TPoint1.DistanceTo(TP.min_v) < TT.TPoint.DistanceTo(TP.min_v)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.LeftUp);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X - C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y , 0);
                                    T3 = new Point3d(T2.X, T1.Y + 5, 0); }
                                if (T0.X < Zentr.X & T0.Y < Zentr.Y)
                                {
                                    T0 = TT.TPoint;
                                    if (TT.TPoint1.DistanceTo(TP.min) < TT.TPoint.DistanceTo(TP.min))  { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.LeftBottom);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X - C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y , 0);
                                    T3 = new Point3d(T1.X - C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y + 5, 0);
                                }
                                if (T0.X > Zentr.X & T0.Y < Zentr.Y)
                                {
                                    T0 = TT.TPoint;
                                    if (TT.TPoint1.DistanceTo(TP.max_n) < TT.TPoint.DistanceTo(TP.max_n)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.RigthBottom);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                                    T3 = new Point3d(T1.X, T1.Y + 5, 0);
                                }
                            }
                        }
                        if (OutPlane) 
                        {
                         T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.LeftUp);
                         T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                         T3 = new Point3d(T1.X, T1.Y + 5, 0);
                        }
                        Point3dCollection spTpostrMetk = new Point3dCollection() { T0, T1, T2 };
                        Metka(spTpostrMetk, TT.LINK, "H=" + TT.Visot, Convert.ToDouble(Mas), T3, TT.Komp); }
                }
            }
        }//Подписать высоту гибов
        private void button30_Click(object sender, EventArgs e)
        {
            List<Point3d> ListPoint1 = new List<Point3d>();
            bool OutPlane = true;
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            List<POZIZIA> SpPOZob = new List<POZIZIA>();
            //if (Directory.Exists(@"C:\МАРШРУТ\SpIzNANOsPoz.txt") == true) HCteniPOZTXT(ref SpPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
            //if (Directory.Exists(@"C:\МАРШРУТ\SpIzNANOsPozObor.txt") == true) HCteniPOZTXT(ref SpPOZob, @"C:\МАРШРУТ\SpIzNANOsPozObor.txt");
            double A = Convert.ToDouble(this.textBox18.Text);
            double B = Convert.ToDouble(this.textBox20.Text);
            double C = Convert.ToDouble(this.textBox19.Text);
            List<TPosrL> SpPoz = new List<TPosrL>();
            List<TPosrL> SpPoZProstavl = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SBORBl_Poz_APr(ref SpPoz, ref SpPoZProstavl,ref ListPoint1);
                foreach(Point3d Point in ListPoint1) this.listBox2.Items.Add(Point);
                if (this.checkBox5.Checked == true)
                {
                    Transaction tr1 = db.TransactionManager.StartTransaction();
                    using (tr1)
                    {
                        foreach (TPosrL ID in SpPoZProstavl) { Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity; Obj.Erase(); }
                        SpPoZProstavl.Clear();
                        tr1.Commit();
                    }
                }
                foreach (TPosrL TT in SpPoz)
                {
                    if (!(SpPoZProstavl.Exists(x => x.Komp == TT.Komp)))
                    {
                        Point3d T0 = TT.TPoint;
                        Point3d T1 = TT.TPoint;
                        OutPlane = true;
                        foreach (PLOS TP in SpPLOS)
                        {
                            if (TT.TPoint.X > TP.min.X & TT.TPoint.X < TP.max.X & TT.TPoint.Y > TP.min.Y & TT.TPoint.Y < TP.max.Y)
                            {
                                OutPlane = false;
                                T0 = TP.FFindPoint(TT.SpKoor, out Side TSide);
                                T1 = T0;
                                T1 = FindAcceptablePoint(ref T1, ListPoint1, C, A, B, TSide);
                                ListPoint1.Add(T1);
                            }
                        }
                        if (OutPlane) { T1 = FindAcceptablePoint(ref T1, ListPoint1, C, A, B, Side.RigthUp); ListPoint1.Add(T1); }
                        ProstPOZ(T0, T1, TT.OId, SpPOZ, SpPOZob);
                    }
                }
            }
        }//Подписать позиции
        private void button31_Click(object sender, EventArgs e)
        {
            if (this.radioButton6.Checked) DimOfОbstacles();
            if (this.radioButton5.Checked) DimOfGird();
        }//проставить размеры
        private void button32_Click(object sender, EventArgs e)
        {
            if (SpPLOS.Count == 0) { Application.ShowAlertDialog("Нет ни одной плоскости построения"); return; }
            this.Hide();
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMax = 0;
            double yMax = 0;
            double xMin = 0;
            double yMin = 0;
            Point3d BP = new Point3d();
            Point3d PSKtl = new Point3d();
            Point3d MSKtl = new Point3d();
            double stepGor = Convert.ToDouble(this.textBox22.Text);
            double stepVert = Convert.ToDouble(this.textBox23.Text);
            double relativeСoordinateX = Convert.ToDouble(this.textBox25.Text);
            double relativeСoordinateY = Convert.ToDouble(this.textBox24.Text);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
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
                        }
                    }
                    PSKtl = new Point3d(BP.X + x1, BP.Y + y1, 0);
                    CreatGird(stepGor, stepVert, PSKtl, bref.GeometricExtents.MaxPoint, bref.GeometricExtents.MinPoint, relativeСoordinateX, relativeСoordinateY);
                    tr.Commit();    
                }
            }
            this.Show();
        }//сделать сетку для вида
        private void button33_Click(object sender, EventArgs e)
        {
            int i = 0;
            string File = this.textBox28.Text;
            if (System.IO.File.Exists(@File) == true)
                { 
                LoadCWinTXT(File,ref SpPart,ref SpCWAY);
                this.dataGridView7.Rows.Clear();
                foreach (CWAY tCW in SpCWAY) 
                { 
                    this.dataGridView7.Rows.Add(tCW.CWName, tCW.CompName, tCW.ModuleName);
                    if (SpCWAYcshert.Exists(x=>x.CWName==tCW.CWName)) this.dataGridView7.Rows[i].DefaultCellStyle.BackColor = Color.Aqua;
                    i++;
                }
            }
            else
                Application.ShowAlertDialog(File + "-  не найден");
        }//загрузка трасс из текстового файла
        private void button34_Click(object sender, EventArgs e)
        {
            
            
        }//применить препядствия
        private void button39_Click(object sender, EventArgs e)
        {
            double DistPen = Convert.ToDouble(this.textBox37.Text);
            double RadPen = Convert.ToDouble(this.textBox36.Text);
            List<BLOK.Form3.TPosrL> spTles = new List<BLOK.Form3.TPosrL>();
            List<Point3d> ListPoint1 = new List<Point3d>();
            List<BLOK.Form3.TPosrL> SpPoz = new List<BLOK.Form3.TPosrL>();
            List<TPosrL> SpPoZProstavl = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SBORBl_Poz_APr(ref SpPoz, ref SpPoZProstavl, ref ListPoint1);
                SOZD_Sp_TP_Lest_BL(ref spTles);
                List<BLOK.Form3.TPosrL> ListCWEY = SpPoz.FindAll(x => x.Tip == "Лес");
                List<BLOK.Form3.TPosrL> ListPen = SpPoz.FindAll(x => x.Komp.Contains("Вырез"));
                listBox2.Items.Clear();
                foreach (BLOK.Form3.TPosrL TCW in ListCWEY) 
                    {
                    double tID = 0;
                    bool changes = false;
                    List<BLOK.Form3.TPosrL> CWEY = spTles.FindAll(x=>x.LINK== TCW.NameCWAY);
                    CWEY.Sort(delegate (BLOK.Form3.TPosrL x, BLOK.Form3.TPosrL y) { return x.NomT.CompareTo(y.NomT);});
                    Vector3d vekCWEY_0_1 = CWEY[0].TPoint.GetVectorTo(CWEY[1].TPoint);
                    Vector3d vekCWEYP_n_nMin1 = CWEY.Last().TPoint.GetVectorTo(CWEY[CWEY.Count-2].TPoint);
                    List<BLOK.Form3.TPosrL> FindPen0 = ListPen.FindAll(x=> x.MaxPoint.DistanceTo(CWEY[0].TPoint)< RadPen | x.MinPoint.DistanceTo(CWEY[0].TPoint) < RadPen);
                    List<BLOK.Form3.TPosrL> FindPenN = ListPen.FindAll(x => x.MaxPoint.DistanceTo(CWEY.Last().TPoint) < RadPen | x.MinPoint.DistanceTo(CWEY.Last().TPoint) < RadPen);
                        if (FindPen0.Count > 0) 
                        {
                        Vector3d vekCWEY_0_Pen_Min = FindPen0[0].MinPoint.GetVectorTo(CWEY[0].TPoint);
                        Vector3d vekCWEY_0_Pen_Max = FindPen0[0].MaxPoint.GetVectorTo(CWEY[0].TPoint);
                        Vector3d Vek_CWEY = CWEY[0].TPoint.GetVectorTo(CWEY[1].TPoint);
                        double Ofset = 0;
                            if (Math.Abs(Vek_CWEY.X) < 0.1)
                            { if (Vek_CWEY.Y > 0) Ofset = DistPen - Math.Min(vekCWEY_0_Pen_Max.Y,vekCWEY_0_Pen_Min.Y);else Ofset = DistPen + Math.Max(vekCWEY_0_Pen_Max.Y, vekCWEY_0_Pen_Min.Y); }
                            if (Math.Abs(Vek_CWEY.Y) < 0.1) 
                            { if (Vek_CWEY.X > 0) Ofset = DistPen - Math.Min(vekCWEY_0_Pen_Max.X,vekCWEY_0_Pen_Min.X); else Ofset = DistPen + Math.Max(vekCWEY_0_Pen_Max.X, vekCWEY_0_Pen_Min.X); }
                            if (Ofset > 0) 
                            {
                            changes = true;
                            Point3d PoslT = polar1(CWEY[0].TPoint, Vek_CWEY, Ofset);
                            BLOK.Form3.TPosrL NPointBild = new BLOK.Form3.TPosrL();
                            NPointBild.NLINK(CWEY[0].LINK);
                            NPointBild.NNomT(CWEY[0].NomT);
                            NPointBild.NVisot(CWEY[0].Visot);
                            NPointBild.NTPoint(PoslT);
                            CWEY.Remove(CWEY[0]);
                            CWEY.Insert(0, NPointBild);
                            }
                        listBox2.Items.Add("FirstPointCWAY-" +CWEY[0].LINK + "|CWEY[0]=" + CWEY[0].TPoint);
                        }
                        if (FindPenN.Count > 0)
                        {
                        Vector3d vekCWEY_N_Pen_Min = FindPenN[0].MinPoint.GetVectorTo(CWEY.Last().TPoint);
                        Vector3d vekCWEY_N_Pen_Max = FindPenN[0].MaxPoint.GetVectorTo(CWEY.Last().TPoint);
                        Vector3d Vek_CWEY = CWEY.Last().TPoint.GetVectorTo(CWEY[CWEY.Count-2].TPoint);
                        double Ofset = 0;
                            if (Math.Abs(Vek_CWEY.X) < 0.1)
                            { if (Vek_CWEY.Y > 0) Ofset = DistPen - Math.Min(vekCWEY_N_Pen_Max.Y,vekCWEY_N_Pen_Min.Y);else Ofset = DistPen + Math.Max(vekCWEY_N_Pen_Max.Y, vekCWEY_N_Pen_Min.Y); }
                            if (Math.Abs(Vek_CWEY.Y) < 0.1)
                            { if (Vek_CWEY.X > 0) Ofset = DistPen - Math.Min(vekCWEY_N_Pen_Max.X, vekCWEY_N_Pen_Min.X); else Ofset = DistPen + Math.Max(vekCWEY_N_Pen_Max.X, vekCWEY_N_Pen_Min.X); }
                            if (Ofset > 0)
                            {
                            changes = true;
                            listBox2.Items.Add("Ofset-" + Ofset + "Math.Min(vekCWEY_N_Pen_Max.Y,vekCWEY_N_Pen_Min.Y)-" + Math.Min(vekCWEY_N_Pen_Max.Y, vekCWEY_N_Pen_Min.Y)  + "Math.Min(vekCWEY_N_Pen_Max.X,vekCWEY_N_Pen_Min.X)-" + Math.Min(vekCWEY_N_Pen_Max.X, vekCWEY_N_Pen_Min.X));
                            Point3d PoslT = polar1(CWEY.Last().TPoint, Vek_CWEY, Ofset);
                            BLOK.Form3.TPosrL NPointBild = new BLOK.Form3.TPosrL();
                            NPointBild.NLINK(CWEY.Last().LINK);
                            NPointBild.NNomT(CWEY.Last().NomT);
                            NPointBild.NVisot(CWEY.Last().Visot);
                            NPointBild.NTPoint(PoslT);
                            CWEY.Remove(CWEY.Last());
                            CWEY.Add(NPointBild);
                            }
                        listBox2.Items.Add("LastPointCWAY-" + CWEY[0].LINK + "|CWEY.Last()=" + CWEY.Last().TPoint + "| Vek_CWEY-" + Vek_CWEY + "| vekCWEY_N_Pen_Min-" + vekCWEY_N_Pen_Min + "| vekCWEY_N_Pen_Max-" + vekCWEY_N_Pen_Max);
                        }
                    if (changes)
                    {
                        List<BLOK.Form3.TPosrL> SpPoint1 = new List<BLOK.Form3.TPosrL>();
                        List<ObjectId> SpOBJID = new List<ObjectId>();
                        SpOBJID.Clear();
                        Form3.SBORBl_FIND(TCW.NameCWAY, ref SpOBJID, ref SpPoint1);
                        Transaction tr1 = db.TransactionManager.StartTransaction();
                        using (tr1)
                        {
                            //if (this.checkBox1.Checked) SpPoint_poln.RemoveAll(x => x.Tip == "хв");
                            foreach (ObjectId tID_CW in SpOBJID)
                            { Entity Obj = tr1.GetObject(tID_CW, OpenMode.ForWrite) as Entity; Obj.Erase(); }
                            tr1.Commit();
                        }
                        Lestn TL = ListLader.Find(x => x.Name == TCW.Komp);
                        string ID = TCW.LINK.Split('-')[0];
                        Form3.Lest_N_INtr_SP(CWEY, TCW.Komp, TL.Hsir, TL.RazdelSP, TL.Povorot, TL.Hvost, TL.Soed, "", 400 + TL.Hsir / 2, Convert.ToDouble(TL.StepHvost), ID, ref tID, true, "", false);
                    }
                }
            }
        }//Сместить от проходок
        private void button35_Click(object sender, EventArgs e)
        {
            if (SpPLOS.Count == 0) { Application.ShowAlertDialog("Нет ни одной плоскости построения"); return; }
            if (this.dataGridView7.CurrentRow.Cells[0].Value == null) return;
            string IND = this.dataGridView7.CurrentRow.Cells[0].Value.ToString();
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            double GlubinaSR = Convert.ToDouble(this.textBox21.Text);
            double DeltH = Convert.ToDouble(this.textBox29.Text);
            Point3d BP = new Point3d();
            Point3d PSKtl = new Point3d();
            Point3d MSKtl = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            List<Part> SpPartTCWAY = new List<Part>();
            List<Part> SpPartTCWAYSort = new List<Part>();
            List<TPosrL> SpTPostrL = new List<TPosrL>();
            PLOS TPlane = new PLOS();
            string Room = this.textBox1.Text;
            string OSItl = "";
            string Name = "";
            string StartPoint = "";
            string EndPoint = "";
            string StartPointTr = "";
            string EndPointTr = "";
            string COMPON = "";
            string NAME = "";
            double Dles = 100;
            double Delt = 0, DeltX = 0, DeltY = 0;
            Point3d MAX = new Point3d();
            Point3d MIN = new Point3d();
            string[] COMPONm;
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
                            if (prop.PropertyName == "Расстояние1") { DeltX = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Расстояние2") { DeltY = Convert.ToDouble(prop.Value.ToString()); }
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
                    MIN = BP;
                    MAX = new Point3d(MIN.X + DeltX, MIN.Y + DeltY, 0);
                    TPlane.NPsk(PSKtl);
                    TPlane.NMsk(MSKtl);
                    TPlane.Nmin(MIN);
                    TPlane.Nmax(MAX);
                    TPlane.NOsi(OSItl);
                }
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
                CWAY tCWAY = SpCWAY.Find(x=>x.CWName== IND);
                if (this.checkBox8.Checked) NAME = IND;
                int TipeHigt = 3;
                if(radioButton7.Checked) TipeHigt = 1;
                if (radioButton8.Checked) TipeHigt = 2;
                BuildSelectCWEY(SpPart, tCWAY, GlubinaSR, DeltH,checkBox7.Checked, checkBox6.Checked,textBox1.Text, checkBox1.Checked, TipeHigt, TPlane, NAME);
            }
            this.Show();
        }//построить выделеную трассу
        private void button36_Click(object sender, EventArgs e)
        {
            if (this.dataGridView9.CurrentRow.Cells[0].Value == null) return;
            String IND = this.dataGridView9.CurrentRow.Cells[0].Value.ToString();
            List<ObjectId> SpOBJID = new List<ObjectId>();
            List<BLOK.Form3.TPosrL> SpPoint1 = new List<BLOK.Form3.TPosrL>();
            List<TPosrL> SpPoint_poln = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Database db = doc.Database;
                Editor ed = doc.Editor;
                SpPoint1.Clear();
                Form3.SBORBl_FIND(IND, ref SpOBJID,ref SpPoint1);
                ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                ed.SetImpliedSelection(idarrayEmpty);
            }
        }//найти трассу
        private void button37_Click(object sender, EventArgs e)
        {
            string Name = "";
            string Zagal = "";
            string Povorot = "";
            string Soed = "";
            string Xvost = "";
            string DopKR = "";
            string DopSoed = "";
            double Hsirin = 0;
            double HSagSoed = 0;
            double HSagKr = 0;
            double tID = 0;
            Int32 selectedRowCount =dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected);
            if (selectedRowCount == 0) return;
            List<ObjectId> SpOBJID = new List<ObjectId>();
            List<BLOK.Form3.TPosrL> SpPoint1 = new List<BLOK.Form3.TPosrL>();
            List<TPosrL> SpPoint_poln = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Database db = doc.Database;
                Editor ed = doc.Editor;
                for (int i = 0; i < selectedRowCount; i++)
                {
                    string IND = this.dataGridView9.Rows[dataGridView9.SelectedRows[i].Index].Cells[0].Value.ToString();
                    string Tipor = this.dataGridView9.Rows[dataGridView9.SelectedRows[i].Index].Cells[1].Value.ToString();
                    SpPoint1.Clear();
                    SpOBJID.Clear();
                    Form3.SBORBl_FIND(IND, ref SpOBJID, ref SpPoint1);
                    Lestn TL = ListLader.Find(x => x.Name == Tipor);
                    Form3.Lest_N_INtr_SP(SpPoint1, Tipor, TL.Hsir, TL.RazdelSP, TL.Povorot, TL.Hvost, TL.Soed, "", 400 + TL.Hsir / 2, Convert.ToDouble(TL.StepHvost), IND, ref tID, true, "", false);
                    Transaction tr1 = db.TransactionManager.StartTransaction();
                    using (tr1)
                    {
                        if (this.checkBox1.Checked) SpPoint_poln.RemoveAll(x => x.Tip == "хв");
                        foreach (ObjectId ID in SpOBJID)
                        { Entity Obj = tr1.GetObject(ID, OpenMode.ForWrite) as Entity; Obj.Erase(); }
                        tr1.Commit();
                    }
                }
            }
        }//перестроить выделеные лестницы
        private void button38_Click(object sender, EventArgs e)
        {
            double A = Convert.ToDouble(this.textBox15.Text);
            double B = Convert.ToDouble(this.textBox16.Text);
            double C = Convert.ToDouble(this.textBox17.Text);
            double D = Convert.ToDouble(this.textBox26.Text);
            bool OutPlane = true;
            List<Point3d> ListPoint1 = new List<Point3d>();
            List<TPosrL> SpGibovGlob = new List<TPosrL>();
            List<TPosrL> SpMetokGlob = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                string NameCWEY = FindNameSelectCWEY();
                SBORBl_Gibov(ref SpGibovGlob, ref SpMetokGlob);
                List<TPosrL> SpGibov = SpGibovGlob.FindAll(x=>x.LINK.Split('-')[0]== NameCWEY);
                List<TPosrL> SpMetok = SpMetokGlob.FindAll(x => x.LINK.Split('-')[0] == NameCWEY);
                if (this.checkBox4.Checked == true)
                {
                    Transaction tr1 = db.TransactionManager.StartTransaction();
                    using (tr1)
                    {
                        foreach (TPosrL ID in SpMetok) { Entity Obj = tr1.GetObject(ID.OId, OpenMode.ForWrite) as Entity; Obj.Erase(); }
                        SpMetok.Clear();
                        tr1.Commit();
                    }
                }
                foreach (TPosrL TT in SpGibov)
                {
                    OutPlane = true;
                    if (!(SpMetok.Exists(x => x.Komp == TT.Komp)))
                    {
                        Point3d T0 = TT.TPoint;
                        Point3d T1 = T0;
                        Point3d T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                        Point3d T3 = new Point3d(T1.X, T1.Y + 5, 0);
                        foreach (PLOS TP in SpPLOS)
                        {
                            if (TT.TPoint.X > TP.min.X & TT.TPoint.X < TP.max.X & TT.TPoint.Y > TP.min.Y & TT.TPoint.Y < TP.max.Y)
                            {
                                if (TT.TPoint1.DistanceTo(TP.max) < TT.TPoint.DistanceTo(TP.max)) T0 = TT.TPoint1;
                                double DeltX = TP.min.X + Math.Abs(TP.max.X - TP.min.X) / 2;
                                double DeltY = TP.min.Y + Math.Abs(TP.max.Y - TP.min.Y) / 2;
                                OutPlane = false;
                                Point3d Zentr = new Point3d(DeltX, DeltY, 0);
                                if (T0.X > Zentr.X & T0.Y > Zentr.Y)
                                {
                                    if (TT.TPoint1.DistanceTo(TP.min_v) < TT.TPoint.DistanceTo(TP.min_v)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.RigthUp);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                                    T3 = new Point3d(T1.X, T1.Y + 5, 0);
                                }
                                if (T0.X < Zentr.X & T0.Y > Zentr.Y)
                                {
                                    if (TT.TPoint1.DistanceTo(TP.min_v) < TT.TPoint.DistanceTo(TP.min_v)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.LeftUp);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X - C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                                    T3 = new Point3d(T2.X, T1.Y + 5, 0);
                                }
                                if (T0.X < Zentr.X & T0.Y < Zentr.Y)
                                {
                                    T0 = TT.TPoint;
                                    if (TT.TPoint1.DistanceTo(TP.min) < TT.TPoint.DistanceTo(TP.min)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.LeftBottom);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X - C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                                    T3 = new Point3d(T1.X - C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y + 5, 0);
                                }
                                if (T0.X > Zentr.X & T0.Y < Zentr.Y)
                                {
                                    T0 = TT.TPoint;
                                    if (TT.TPoint1.DistanceTo(TP.max_n) < TT.TPoint.DistanceTo(TP.max_n)) { T0 = TT.TPoint1; T1 = T0; }
                                    T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.RigthBottom);
                                    ListPoint1.Add(T1);
                                    T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                                    T3 = new Point3d(T1.X, T1.Y + 5, 0);
                                }
                            }
                        }
                        if (OutPlane)
                        {
                            T1 = FindAcceptablePoint(ref T1, ListPoint1, D, A, B, Side.LeftUp);
                            T2 = new Point3d(T1.X + C * (TT.Visot.Length) * Convert.ToDouble(Mas), T1.Y, 0);
                            T3 = new Point3d(T1.X, T1.Y + 5, 0);
                        }
                        Point3dCollection spTpostrMetk = new Point3dCollection() { T0, T1, T2 };
                        Metka(spTpostrMetk, TT.LINK, "H=" + TT.Visot, Convert.ToDouble(Mas), T3, TT.Komp);
                    }
                }
            }
        }//подписать высоту гибов отдельной трассы 
        public static void BuildSelectCWEY(List<Part> SpPart, CWAY tCWAY,double GlubinaSR, double DeltH, bool TrimViev, bool NameModul,string Room,bool BuildHv,int TipeHigt, PLOS Plane, string NAME)
        {
            Point3d PSKtl = Plane.Psk;
            Point3d MSKtl = Plane.Msk;
            Point3d MAX = Plane.max;
            Point3d MIN = Plane.min;
            string OSItl = Plane.Osi;
            string Name = "";
            string StartPoint = "";
            string EndPoint = "";
            string StartPointTr = "";
            string EndPointTr = "";
            string COMPON = "";
            //string NAME = "";
            double Dles = 100;
            double Delt = 0, DeltX = 0, DeltY = 0;
            List<Part> SpPartTCWAY = new List<Part>();
            List<Part> SpPartTCWAYSort = new List<Part>();
            List<TPosrL> SpTPostrL = new List<TPosrL>();
            string Prodol = "Y";
            string[] COMPONm;
            SpTPostrL.Clear();
            SpPartTCWAYSort.Clear();
            double CountStartPoint = 0;
            double CountEndPoint = 0;
            SpPartTCWAY = SpPart.FindAll(x => x.CWName == tCWAY.CWName);
            foreach (Part TPart in SpPartTCWAY)
            {
                StartPoint = TPart.StartPoint;
                EndPoint = TPart.EndPoint;
                List<Part> SpPartStartPoint = SpPartTCWAY.FindAll(x => x.StartPoint == StartPoint | x.EndPoint == StartPoint);
                List<Part> SpPartEndPoint = SpPartTCWAY.FindAll(x => x.StartPoint == EndPoint | x.EndPoint == EndPoint);
                if (SpPartStartPoint.Count == 1) { StartPointTr = StartPoint; CountStartPoint += 1; }
                if (SpPartEndPoint.Count == 1) { EndPointTr = EndPoint; CountEndPoint += 1; }
            }
            string TPointTr = StartPointTr;
            //Application.ShowAlertDialog(tCWAY.CWName + "|SpPartTCWAY.Count-" + SpPartTCWAY.Count + "|CountStartPoint-" + CountStartPoint + "|CountEndPoint-" + CountEndPoint + "|TPointTr-" + TPointTr);
            //Application.ShowAlertDialog("Plane.Psk-" + Plane.Psk + "|Plane.Msk-" + Plane.Msk + "|Plane.max-" + Plane.max + "|Plane.min-" + Plane.min + "|Plane.Osi-" + Plane.Osi);
            if (CountStartPoint == 1 & CountEndPoint == 1)
            {
                while (TPointTr != EndPointTr | Prodol != "Y")
                {
                    if (SpPartTCWAY.Exists(x => x.StartPoint == TPointTr))
                    {
                        Part TPart = SpPartTCWAY.Find(x => x.StartPoint == TPointTr);
                        TPointTr = TPart.EndPoint;
                        SpPartTCWAYSort.Add(TPart);
                    }
                    else
                        Prodol = "N";
                }
                foreach (Part TPart in SpPartTCWAYSort)
                {
                    TPosrL TTpstrL = new TPosrL();
                    //Point3d TPostr = new Point3d();
                    string Visot = "";
                    //Application.ShowAlertDialog(TPart.CWName);
                    if (TPart.CompName.Contains("CTRSZ") == true)
                    {
                        StartPoint = TPart.StartPoint;
                        Point3d TPostr = new Point3d(Convert.ToDouble(StartPoint.Split(' ')[0]), Convert.ToDouble(StartPoint.Split(' ')[1]), 0);
                        COMPONm = TPart.CompName.Split('_');
                        if (COMPONm.Length > 1) { if (Double.TryParse(COMPONm[1], out Delt)) TTpstrL.NDelt(Convert.ToDouble(COMPONm[1])); }
                        Visot = StartPoint.Split(' ')[2];
                        TTpstrL.NTPoint(TPostr);
                        TTpstrL.NVisot(Visot);
                        SpTPostrL.Add(TTpstrL);
                    }
                    else
                    {
                        StartPoint = TPart.Coordinates;
                        Point3d TPostr = new Point3d(Convert.ToDouble(StartPoint.Split(' ')[0]), Convert.ToDouble(StartPoint.Split(' ')[1]), 0);
                        COMPONm = TPart.CompName.Split('_');
                        if (COMPONm.Length > 1) { if (Double.TryParse(COMPONm[1], out Delt)) TTpstrL.NDelt(Convert.ToDouble(COMPONm[1])); }
                        Visot = StartPoint.Split(' ')[2];
                        TTpstrL.NTPoint(TPostr);
                        TTpstrL.NVisot(Visot);
                        SpTPostrL.Add(TTpstrL);
                    }
                }
                if (SpPartTCWAYSort.Last().CompName.Contains("CTRSZ"))
                {
                    TPosrL TTpstrL = new TPosrL();
                    StartPoint = SpPartTCWAYSort.Last().EndPoint;
                    Point3d TPostr = new Point3d(Convert.ToDouble(StartPoint.Split(' ')[0]), Convert.ToDouble(StartPoint.Split(' ')[1]), 0);
                    string Visot = StartPoint.Split(' ')[2];
                    TTpstrL.NTPoint(TPostr);
                    TTpstrL.NVisot(Visot);
                    SpTPostrL.Add(TTpstrL);
                }
                if (tCWAY.CompName.Contains("-100_")) { Dles = 108; COMPON = "#ЛестУсил100"; }
                if (tCWAY.CompName.Contains("-200_")) { Dles = 208; COMPON = "#ЛестУсил200"; }
                if (tCWAY.CompName.Contains("-300_")) { Dles = 308; COMPON = "#ЛестУсил300"; }
                if (tCWAY.CompName.Contains("-400_")) { Dles = 408; COMPON = "#ЛестУсил400"; }
                if (tCWAY.CompName.Contains("-500_")) { Dles = 508; COMPON = "#ЛестУсил500"; }
                if (tCWAY.CompName.Contains("-600_")) { Dles = 608; COMPON = "#ЛестУсил600"; }
                if (tCWAY.CompName.Contains("-700_")) { Dles = 708; COMPON = "#ЛестУсил700"; }
                if (tCWAY.CompName.Contains("-800_")) { Dles = 808; COMPON = "#ЛестУсил800"; }
                List<TPosrL> SpTPostrL_Per = new List<TPosrL>();
                bool Prov = true;
                foreach (TPosrL TT in SpTPostrL)
                {
                    double mid = 0;
                    string[] COMPONm1 = tCWAY.CompName.Split('_');
                    if (COMPONm1.Length > 1)
                    { if (Double.TryParse(COMPONm1[1], out Delt)) mid = Convert.ToDouble(COMPONm1[1]) / 2; }
                    if (TipeHigt == 1) mid = mid * (-1);
                    if (TipeHigt == 2) mid = 0;
                    TPosrL NT = PerXYZ(TT, OSItl, MSKtl, PSKtl, mid, DeltH);
                    if (Math.Abs(Convert.ToDouble(NT.Visot)) > GlubinaSR) Prov = false;
                    if (TrimViev)
                    {
                        if (NT.TPoint.X >= MIN.X & NT.TPoint.X <= MAX.X & NT.TPoint.Y >= MIN.Y & NT.TPoint.Y <= MAX.Y)
                            SpTPostrL_Per.Add(NT);
                    }
                    else
                        SpTPostrL_Per.Add(NT);
                }
                pruning_the_vertical_ends(ref SpTPostrL_Per,ref Prov);
                if (Prov)
                {
                    UdalToch(ref SpTPostrL_Per);
                    Lest_N_INtr_SP(SpTPostrL_Per, COMPON, Dles, "Лес", "", "40*40*4#Уголок40х40х4#00311742137", "", "", 400 + Dles / 2, 1200, NameModul ? tCWAY.ModuleName : Room, NAME, BuildHv);
                }
            }
        }
        public static void pruning_the_vertical_ends(ref List<TPosrL> SpTPostrL_Per,ref bool Prov)
        {
            if (SpTPostrL_Per.Count > 2)
                for (int i= 0; i < SpTPostrL_Per.Count-2; i++)
                {
                TPosrL Pointi = SpTPostrL_Per[i];
                TPosrL Point1pl1 = SpTPostrL_Per[i+1];
                TPosrL Point1pl2 = SpTPostrL_Per[i + 2];
                if (Pointi.TPoint.DistanceTo(Point1pl1.TPoint) < 1) 
                {
                    Vector3d Vekt1 = Point1pl1.TPoint.GetVectorTo(Point1pl2.TPoint)/ Point1pl1.TPoint.DistanceTo(Point1pl2.TPoint);
                    TPosrL NevPoint = new TPosrL();
                    NevPoint.NVisot(Point1pl1.Visot);
                    NevPoint.NTPoint(Point1pl1.TPoint + Vekt1);
                    SpTPostrL_Per.Remove(Point1pl1);
                    SpTPostrL_Per.Insert(i + 1, NevPoint);
                }
                }
            if (SpTPostrL_Per.Count == 2) 
            {
                TPosrL Point0 = SpTPostrL_Per[0];
                TPosrL Point1 = SpTPostrL_Per[1];
                if (Point0.TPoint.DistanceTo(Point1.TPoint) < 1) Prov = false;
            }
        }
        public string FindNameSelectCWEY() 
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            int Schet = 0;
            string stRab = "";
            Transaction tr = db.TransactionManager.StartTransaction();
            this.Hide();
            using (tr)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                try
                {
                    // Просим пользователя выбрать примитив
                    PromptEntityResult ers = ed.GetEntity("Укажите примитив ");
                    // Открываем выбранный примитив
                    Entity ent = (Entity)tr.GetObject(ers.ObjectId, OpenMode.ForWrite);
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
                            if (Schet == 10) { stRab = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }
            this.Show();
            return stRab;
        }//узнать имя кабельной лестницы
        public void DimOfОbstacles() 
        {
            double MinDistGor = 99999999;
            double MinDistVert = 99999999;
            double Corner = 0;
            double AngeGmin = 0;
            double Dim = 0;
            bool FindObstclVert = false;
            bool FindObstclGor = false;
            bool FindColl = false;
            string DeltAng = "";
            double CountVer = 0;
            orientation OrCwPart = orientation.gor;
            orientation OrObst = orientation.gor;
            List<double> listDim = new List<double>();
            List<TPosrL> Оbstacles = new List<TPosrL>();
            List<TPosrL> CWEY = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            this.listBox2.Items.Clear();
            using (DocumentLock docLock = doc.LockDocument())
            {
                SBORBl_Poz_Dim(ref CWEY, ref Оbstacles);
                foreach (TPosrL TT in CWEY)
                {
                    listDim.Clear();
                    CountVer = TT.SpKoor.Count - 1;
                    this.listBox2.Items.Add(TT.SpKoor.Count);
                    for (int j = 0; j < (TT.SpKoor.Count - 1) / 2; j++)
                    {
                        Point3d Point1 = TT.SpKoor[j];
                        Point3d Point2Gor = TT.SpKoor[j];
                        Point3d Point2Vert = TT.SpKoor[j];
                        double AngeG = Point1.GetVectorTo(TT.SpKoor[j + 1]).GetAngleTo(new Vector3d(1, 0, 0));
                        double AngeV = Point1.GetVectorTo(TT.SpKoor[j + 1]).GetAngleTo(new Vector3d(0, 1, 0));
                        OrCwPart = FindOri(AngeG, AngeV);
                        if (j > 0)
                        { AngeGmin = Point1.GetVectorTo(TT.SpKoor[j - 1]).GetAngleTo(new Vector3d(1, 0, 0)); }
                        MinDistGor = 99999999;
                        MinDistVert = 99999999;
                        FindObstclVert = false;
                        FindObstclGor = false;
                        foreach (TPosrL CurObst in Оbstacles)
                        {
                            FindColl = false;
                            OrObst = FindOri(CurObst.AngelG, CurObst.AngelV);
                            FindCollision(Point1, CurObst.TPointZ, TT.SpKoor, ref FindColl);
                            this.listBox2.Items.Add(j + "| OrCwPart=" + OrCwPart + " | OrObst=" + OrObst + "| FindColl=" + FindColl);
                            if (CurObst.TPointZ.DistanceTo(Point1) < MinDistGor & OrCwPart == orientation.gor & OrObst == orientation.gor & !FindColl)
                            {
                                MinDistGor = CurObst.TPointZ.DistanceTo(Point1);
                                Point2Gor = CurObst.TPointZ; Corner = Math.PI / 2;
                                FindObstclGor = true;
                                Dim = Math.Abs(Point1.Y - Point2Gor.Y);
                            }
                            if (CurObst.TPointZ.DistanceTo(Point1) < MinDistVert & OrCwPart == orientation.vert & OrObst == orientation.vert & !FindColl)
                            {
                                MinDistVert = CurObst.TPointZ.DistanceTo(Point1);
                                Point2Vert = CurObst.TPointZ; Corner = 0;
                                FindObstclVert = true;
                                Dim = Math.Abs(Point1.X - Point2Vert.X);
                            }
                        }
                        if (j == (CountVer - 1) | j == ((CountVer / 2) - 1)) { FindObstclVert = false; FindObstclGor = false; }
                        if (FindObstclVert & !(listDim.Exists(x => x == Dim)))
                        {
                            AutoDim(Point1, Point2Vert, Corner);
                            listDim.Add(Dim);
                        }
                        if (FindObstclGor & !(listDim.Exists(x => x == Dim)))
                        {
                            AutoDim(Point1, Point2Gor, Corner);
                            listDim.Add(Dim);
                        }
                    }
                }
            }
        }//Размеры от препядствий
        public void DimOfGird()
        {
            double MinDistGor = 99999999;
            double MinDistVert = 99999999;
            double Corner = 0;
            double AngeGmin = 0;
            double Dim = 0;
            bool FindObstclVert = false;
            bool FindObstclGor = false;
            bool FindColl = false;
            double CountVer = 0;
            double RadiusFindNode = Convert.ToDouble(this.textBox27.Text);
            orientation OrCwPart = orientation.gor;
            List<double> listDim = new List<double>();
            List<TPosrL> Оbstacles = new List<TPosrL>();
            List<TPosrL> CWEY = new List<TPosrL>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            this.listBox2.Items.Clear();
            using (DocumentLock docLock = doc.LockDocument())
            {
                List<Point3d> ListPointGird = SBORBl_Poz_Dim_Gird(ref CWEY);
                foreach (Point3d CurPoi in ListPointGird) this.listBox2.Items.Add(CurPoi); ;
                this.listBox2.Items.Add(ListPointGird.Count + " " + CWEY.Count);
                foreach (TPosrL TT in CWEY)
                {
                    listDim.Clear();
                    CountVer = TT.SpKoor.Count - 1;
                    this.listBox2.Items.Add(TT.SpKoor.Count);
                    List<Point3d> FindListPointGirt = ListPointGird.FindAll(x => x.X > TT.MinPoint.X - RadiusFindNode &
                                                                                x.Y > TT.MinPoint.Y - RadiusFindNode &
                                                                                x.X < TT.MaxPoint.X + RadiusFindNode &
                                                                                x.Y < TT.MaxPoint.Y + RadiusFindNode);
                    for (int j = 0; j < (TT.SpKoor.Count - 1) / 2; j++)
                    {
                        Point3d Point1 = TT.SpKoor[j];
                        Point3d Point2Gor = TT.SpKoor[j];
                        Point3d Point2Vert = TT.SpKoor[j];
                        double AngeG = Point1.GetVectorTo(TT.SpKoor[j + 1]).GetAngleTo(new Vector3d(1, 0, 0));
                        double AngeV = Point1.GetVectorTo(TT.SpKoor[j + 1]).GetAngleTo(new Vector3d(0, 1, 0));
                        OrCwPart = FindOri(AngeG, AngeV);
                        if (j > 0)
                        { AngeGmin = Point1.GetVectorTo(TT.SpKoor[j - 1]).GetAngleTo(new Vector3d(1, 0, 0)); }
                        MinDistGor = 99999999;
                        MinDistVert = 99999999;
                        FindObstclVert = false;
                        FindObstclGor = false;
                        foreach (Point3d CurPoint in FindListPointGirt)
                        {
                            FindColl = false;
                            FindCollision(Point1, CurPoint, TT.SpKoor, ref FindColl);
                            this.listBox2.Items.Add(j + "| OrCwPart=" + OrCwPart +  "| FindColl=" + FindColl);
                            if (CurPoint.DistanceTo(Point1) < MinDistGor & OrCwPart == orientation.gor & !FindColl)
                            {
                                MinDistGor = CurPoint.DistanceTo(Point1);
                                Point2Gor = CurPoint;
                                Corner = Math.PI / 2;
                                FindObstclGor = true;
                                Dim = Math.Abs(Point1.Y - Point2Gor.Y);
                            }
                            if (CurPoint.DistanceTo(Point1) < MinDistVert & OrCwPart == orientation.vert & !FindColl)
                            {
                                MinDistVert = CurPoint.DistanceTo(Point1);
                                Point2Vert = CurPoint;
                                Corner = 0;
                                FindObstclVert = true;
                                Dim = Math.Abs(Point1.X - Point2Vert.X);
                            }
                        }
                        if (j == (CountVer - 1) | j == ((CountVer / 2) - 1)) { FindObstclVert = false; FindObstclGor = false; }
                        if (FindObstclVert & !(listDim.Exists(x => x == Dim)))
                        {
                            AutoDim(Point1, Point2Vert, Corner);
                            listDim.Add(Dim);
                        }
                        if (FindObstclGor & !(listDim.Exists(x => x == Dim)))
                        {
                            AutoDim(Point1, Point2Gor, Corner);
                            listDim.Add(Dim);
                        }
                    }
                }
            }
        }//Размеры от сетки
        public void CreatGird(double stepGor, double stepVert, Point3d BPoint, Point3d MaxPoint, Point3d MinPoint, double relativeСoordinateX, double relativeСoordinateY)
        {
            Point3d Point1 = new Point3d();
            Point3d Point2 = new Point3d();
            double Ngor0 = Math.Truncate(relativeСoordinateX);
            double NVert0 = Math.Truncate(relativeСoordinateY);
            double NgorT = Ngor0;
            double NVertT = NVert0;

            double X0 = BPoint.X - stepGor* (relativeСoordinateX - Math.Truncate(relativeСoordinateX));
            double Y0 = BPoint.Y - stepVert * (relativeСoordinateY- Math.Truncate(relativeСoordinateY));
            double Xt = X0;
            double Yt = Y0;

            double Xl = MinPoint.X;
            double Xr = MaxPoint.X;

            double Yu = MaxPoint.Y;
            double Yd = MinPoint.Y;

            while (Xt < MaxPoint.X) 
            {
                Point1 = new Point3d(Xt, Yu,0);
                Point2 = new Point3d(Xt, Yd, 0);
                LineGird(Point1, Point2, "Vert", NgorT.ToString("0."));
                Xt += stepGor;
                NgorT += 1;
            }
            Xt = X0;
            NgorT = Ngor0;
            while (Xt > MinPoint.X)
            {
                Point1 = new Point3d(Xt, Yu, 0);
                Point2 = new Point3d(Xt, Yd, 0);
                LineGird(Point1, Point2, "Vert", NgorT.ToString("0."));
                Xt -= stepGor;
                NgorT -= 1;
            }
            while (Yt < MaxPoint.Y)
            {
                Point1 = new Point3d(Xr, Yt, 0);
                Point2 = new Point3d(Xl, Yt, 0);
                LineGird(Point1, Point2, "Gor", NVertT.ToString("0."));
                Yt += stepVert;
                NVertT += 1;
            }
            Yt = Y0;
            NVertT = NVert0;
            while (Yt > MinPoint.Y)
            {
                Point1 = new Point3d(Xr, Yt, 0);
                Point2 = new Point3d(Xl, Yt, 0);
                LineGird(Point1, Point2, "Gor", NVertT.ToString("0."));
                Yt -= stepVert;
                NVertT -= 1;
            }
        }
        public void LineGird(Point3d point1, Point3d point2, string tip, string Nom)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            // Append the point to the database
            using (tr1)
            {
                Line LineGird = new Line();
                LineGird.StartPoint = point1;
                LineGird.EndPoint = point2;
                LineGird.ColorIndex = 164;
                LineGird.Layer = "Плоскости";
                LineGird.XData = new ResultBuffer
                (
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, tip)
                );
                btr.AppendEntity(LineGird);
                tr1.AddNewlyCreatedDBObject(LineGird, true);

                DBText Poz = new DBText();
                Poz.SetDatabaseDefaults();
                Poz.Position = point1;
                Poz.Height = 50 ;
                Poz.TextString = Nom;
                Poz.Layer = "Плоскости";
                btr.AppendEntity(Poz);
                tr1.AddNewlyCreatedDBObject(Poz, true);

                //btr.Dispose();
                tr1.Commit();
            }

        }
        public orientation FindOri(double AngG,double AngV) 
        {
            orientation FindOr = orientation.notDef;
            if (Math.Abs(AngG - 0) < 0.05) FindOr = orientation.gor;
            if (Math.Abs(AngG - Math.PI) < 0.05) FindOr = orientation.gor;
            if (Math.Abs(AngV - 0) < 0.05) FindOr = orientation.vert;
            if (Math.Abs(AngV - Math.PI) < 0.05) FindOr = orientation.vert;
            return FindOr;
        }
        public void FindCollision(Point3d point1 , Point3d point2, Point3dCollection ListPoint,ref bool FindColl)
        {
            List<Point3d> SpToch = new List<Point3d>();
            for (int i = 0; i < ListPoint.Count - 1; i++)
            {
                Point3d TOtr1 = ListPoint[i];
                Point3d TOtr2 = ListPoint[i + 1];
                Point3d TperK = TPer1(TOtr1, TOtr2, point1, point2);
                if (TperK != new Point3d() & SpToch.Exists(X => X == TperK) == false) { SpToch.Add(TperK); }
            }
            if (SpToch.Count > 1 ) FindColl = true;
        }
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
        public Point3d FindAcceptablePoint(ref Point3d VerifiablePoint, List<Point3d> ListPoint1,double AcceptableDist,double A,double B, Side TSide) 
        {
            bool Error = false;
            if (TSide == Side.RigthUp) VerifiablePoint = new Point3d(VerifiablePoint.X + A, VerifiablePoint.Y + B, 0);
            if (TSide == Side.RigthBottom) VerifiablePoint = new Point3d(VerifiablePoint.X + A, VerifiablePoint.Y - B, 0);
            if (TSide == Side.LeftBottom) VerifiablePoint = new Point3d(VerifiablePoint.X - A, VerifiablePoint.Y - B, 0);
            if (TSide == Side.LeftUp) VerifiablePoint = new Point3d(VerifiablePoint.X - A, VerifiablePoint.Y + B, 0);
            foreach (Point3d TP in ListPoint1) { if (TP.DistanceTo(VerifiablePoint) < AcceptableDist) Error = true; }
            if (Error) FindAcceptablePoint(ref VerifiablePoint, ListPoint1, AcceptableDist, A, B, TSide);
            return VerifiablePoint;
        }
        public static TPosrL PerXYZ(TPosrL TT, string Osi, Point3d MSKtl, Point3d PSKtl,double mid, double DeltHi) 
        {
            TPosrL tPer = new TPosrL();
            double delX = TT.TPoint.X - MSKtl.X;
            double delY = TT.TPoint.Y - MSKtl.Y;
            double delZ = Convert.ToDouble(TT.Visot) - MSKtl.Z;

            double Xnow = PSKtl.X;
            double Ynow = PSKtl.Y;
            double Znow = PSKtl.Z;

            double Visot = 0;
            //double DeltHi = Convert.ToDouble(this.textBox29.Text);

            if (Osi == "ZY") { Xnow = Xnow - delY; Ynow = Ynow + delZ; }
            if (Osi == "YZ") { Xnow = Xnow + delY; Ynow = Ynow + delZ; }

            if (Osi == "XZ") { Xnow = Xnow + delX; Ynow = Ynow + delZ; }
            if (Osi == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; }

            if (Osi == "XY") 
                            { 
                            Xnow = Xnow + delX; 
                            Ynow = Ynow + delY; 
                            Visot = -1* delZ - DeltHi;
                            Visot = Math.Round(Visot / 10, 0); Visot = Visot * 10;
                            Visot += mid;
                            //if (this.radioButton7.Checked) Visot -= mid;
                            //if (this.radioButton9.Checked) Visot += mid;
                            //this.listBox2.Items.Add(TT.Delt);
            }
            Point3d IskTP = new Point3d(Xnow, Ynow, 0);

            tPer.NTPoint(IskTP);
            tPer.NVisot(Visot.ToString());

            return tPer;
        }//пересчет координат
        public static void UdalToch(ref List<TPosrL> spTpostr)
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
        }
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
        public void InsBlockRef(string BlockPath, string BlockName, string BlockType)
        {
            Application.ShowAlertDialog(BlockPath);
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
                        db1.ReadDwgFile(BlockPath, System.IO.FileShare.Read, true, "");
                        // Получаем ID нового блока
                        DbS.ObjectId BlkId = db.Insert(BlockName, db1, false);
                        DbS.BlockReference bref = new DbS.BlockReference(ptStart, BlkId);
                        // Дефолтные свойства блока (слой, цвет и пр.)
                        bref.SetDatabaseDefaults();
                        // Добавляем блок в модель
                        model.AppendEntity(bref);
                        // Добавляем блок в транзакцию
                        tr.AddNewlyCreatedDBObject(bref, true);
                        // Расчленяем блок
                        if (BlockType == "DB")
                        {
                            bref.ExplodeToOwnerSpace();
                            bref.Erase();
                        }
                        else
                            ADDAtr_Cr_Bl(BlockName);
                        // Закрываем транзакцию
                        tr.Commit();
                    }
                }
            //}
        }//Блок из файла
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
                                if (atrRef.Tag == "ДопДетали") { atrRef.TextString = StrDobav; }
                            }
                        }
                    }
                }
                Tx.Commit();
            }
        }//расширеные данные
        public void RaschKoordin(Point3d ptStart, Point3d ZKr, ref Point3d ptEnd, ref Point3d txt, ref Point3d txtNP, ref Point3d txtPP, ref Point3d TPol1, ref Point3d TPol2, string stNP) 
        {

            double dMas = Convert.ToDouble(Mas);
            double DlinT = stNP.Length * 50 * dMas;
            //double Xkr = ptEnd.X + 100;
            double Xtxt = ZKr.X - 60 * dMas;
            double Ytxt = ZKr.Y - 30 * dMas;

            double Xp1 = ZKr.X + 100 * dMas;
            double Xp2 = Xp1 + DlinT;
            double Xnp = ZKr.X + 110 * dMas;
            double Xpp = ZKr.X + 110 * dMas;
            double Ynp = ZKr.Y + 10 * dMas;
            double Ypp = ZKr.Y - 60 * dMas;
            double Gipot = ptStart.DistanceTo(ZKr);
            double BProKa = ptStart.Y - ZKr.Y;
            double BPriKa = ptStart.X - ZKr.X;

            double COS = BPriKa / Gipot;
            double SIN = BProKa / Gipot;

            double DeltX = COS * 100 * dMas;
            double DeltY = SIN * 100 * dMas;

            if (ZKr.X < ptStart.X)
            {
                //Xkr = ptEnd.X - 100;
                Xtxt = ZKr.X - 60 * dMas;
                Xp1 = ZKr.X - 100 * dMas;
                Xp2 = Xp1 - DlinT;
                Xnp = Xp2 + 8 * dMas;
                Xpp = Xp2 + 8 * dMas;
            }

             //ptEnd = new Point3d(Xkr, ptEnd.Y, ptEnd.Z);
            //Vector3d Ugol = ptStart.GetVectorTo(ZKr);
            //double 2DUg= Ugol.;
             //ptEnd.RotateBy(,Ugol,ZKr)
            
             //ptEnd = ZKr;
             ptEnd = new Point3d(ZKr.X + DeltX, ZKr.Y + DeltY, ZKr.Z);
             txt = new Point3d(Xtxt, Ytxt, ZKr.Z);
             txtNP = new Point3d(Xnp, Ynp, ZKr.Z);
             txtPP = new Point3d(Xpp, Ypp, ZKr.Z);
             TPol1 = new Point3d(Xp1, ZKr.Y, ZKr.Z);
             TPol2 = new Point3d(Xp2, ZKr.Y, ZKr.Z);

        }//расчет координат
        static public void TextPP_NP_POZ(ref List<POZIZIA> SpPOZ, string stKomp, string stPom, string stKEI, string[] VirezM, ref string strTextNP, ref string strTextPP, ref string NOMpoz, double dDlin, double dVisot, string stHtoEto, double kol, ref List<POZIZIA> SpPOZob)
        {
            if (VirezM.Length > 1) { strTextNP = VirezM[1]; }
            if (dVisot == 0 & stKEI == "006") { strTextNP = "L=" + dDlin.ToString(); }
            if (dVisot > 0 & stKEI == "006") { strTextNP = "L=" + Convert.ToString(dDlin + 2 * dVisot); strTextPP = "2фл." + dVisot.ToString(); }
            if (dVisot >= 250 & stKEI == "006" & stKomp == "688-78.1611") { strTextNP = "L=" + Convert.ToString(dDlin + 40); strTextPP = "2фл.20"; }
            if (dVisot > 0 & stKEI == "006" & stHtoEto == "хв") { strTextNP = "L=" + Convert.ToString(dVisot); if (kol > 1) { strTextPP = kol.ToString() + " шт"; } else { strTextPP = ""; } }
            if (dVisot > 0 & stKEI == "796") { strTextPP = "H=" + dVisot.ToString(); }
            if (kol > 1 & stKEI == "796") { strTextPP = kol.ToString() + " шт"; }
            if (SpPOZ.Exists(x => x.Compon == stKomp & x.Pom == stPom) == true) { NOMpoz = SpPOZ.Find(x => x.Compon == stKomp & x.Pom == stPom).NOMpoz; }
            if (SpPOZob.Exists(x => x.Ind == stKomp) == true) { NOMpoz = SpPOZob.Find(x => x.Ind == stKomp).NOMpoz; strTextNP = stKomp; strTextPP = SpPOZob.Find(x => x.Ind == stKomp).shabl; }
            if (strTextNP == "" & (strTextPP == "") == false) { strTextNP = strTextPP; strTextPP = ""; }
        }//текст позиций текст под полкой и над полкой
        public void Postpoen(Point3d ZKr, Point3d txt, Point3d txtNP, Point3d txtPP, Point3d TPol1, Point3d TPol2, string stPoz, string stNP, string stPP, string Handl_Det, string COMP_HAND, string LINK, string RAZDEL, string POM, string Kompon, string AP_LINK)
        {
            double dMas = Convert.ToDouble(Mas);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                acBlkTbl = Tx.GetObject(db.BlockTableId,OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;

                Circle KrugPOZ = new Circle();
                KrugPOZ.SetDatabaseDefaults();
                KrugPOZ.Center = ZKr;
                KrugPOZ.Radius = 100 * dMas;
                KrugPOZ.ColorIndex=1;
                KrugPOZ.Layer = "Насыщение";
                KrugPOZ.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, ""),
                new TypedValue(1040, 0),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, AP_LINK),
                new TypedValue(1000, "Позиция"));
                acBlkTblRec.AppendEntity(KrugPOZ);
                Tx.AddNewlyCreatedDBObject(KrugPOZ, true);

                if (RAZDEL == "Доизол.д" || RAZDEL == "Доиз.Детали")
                {
                    Circle KrugPOZVn = new Circle();
                    KrugPOZVn.SetDatabaseDefaults();
                    KrugPOZVn.Center = ZKr;
                    KrugPOZVn.Radius = 80 * dMas;
                    KrugPOZVn.ColorIndex = 1;
                    KrugPOZVn.Layer = "Насыщение";
                    KrugPOZVn.XData = new ResultBuffer(
                                new TypedValue(1001, "LAUNCH01"),
                                new TypedValue(1000, ""),
                                new TypedValue(1040, 0),
                                new TypedValue(1000, ""),
                                new TypedValue(1000, ""),
                                new TypedValue(1000, ""),
                                new TypedValue(1000, AP_LINK),
                                new TypedValue(1000, "Позиция"));
                    acBlkTblRec.AppendEntity(KrugPOZVn);
                    Tx.AddNewlyCreatedDBObject(KrugPOZVn, true);
                }

                DBText Poz = new DBText();
                Poz.SetDatabaseDefaults();
                Poz.Position = txt;
                Poz.Height = 50 * dMas;
                if (RAZDEL == "не важно") { Poz.ColorIndex = 2; } else { Poz.ColorIndex = 3; }
                Poz.TextString = stPoz;
                Poz.Layer = "Насыщение";
                Poz.XData = new ResultBuffer(
                    new TypedValue(1001, "LAUNCH01"),
                    new TypedValue(1000, Handl_Det),
                    new TypedValue(1040, 0),
                    new TypedValue(1000, "POZIZIA"),
                    new TypedValue(1000, POM),
                    new TypedValue(1000, Kompon),
                    new TypedValue(1000, AP_LINK),
                    new TypedValue(1000, "Позиция"));
                //if (COMP_HAND == "HANDL")
                //{ Poz.XData = new ResultBuffer(new TypedValue(1001, "LAUNCH01"), new TypedValue(1000, Handl_Det), new TypedValue(1040, 0), new TypedValue(1000, "POZIZIA"), new TypedValue(1000, POM), new TypedValue(1000, Kompon)); }
                //else
                //{Poz.XData = new ResultBuffer(new TypedValue(1001, "LAUNCH01"), new TypedValue(1000, Handl_Det), new TypedValue(1040, 0), new TypedValue(1000, ""), new TypedValue(1000, LINK)); }
                acBlkTblRec.AppendEntity(Poz);
                Tx.AddNewlyCreatedDBObject(Poz, true);

                DBText TNPol = new DBText();
                TNPol.SetDatabaseDefaults();
                TNPol.Position = txtNP;
                TNPol.Height = 50 * dMas;
                if (RAZDEL == "не важно") { TNPol.ColorIndex = 2; } else { TNPol.ColorIndex = 3; }
                TNPol.TextString = stNP;
                TNPol.XData = new ResultBuffer(
                    new TypedValue(1001, "LAUNCH01"),
                    new TypedValue(1000, Handl_Det),
                    new TypedValue(1040, 0),
                    new TypedValue(1000, "T_NAD_POL"),
                    new TypedValue(1000, LINK),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, AP_LINK),
                    new TypedValue(1000, "Позиция"));
                TNPol.Layer = "Насыщение";
                acBlkTblRec.AppendEntity(TNPol);
                Tx.AddNewlyCreatedDBObject(TNPol, true);

                if ((stPP == "") == false)
                {
                    DBText TPPol = new DBText();
                    TPPol.SetDatabaseDefaults();
                    TPPol.Position = txtPP;
                    TPPol.Height = 50 * dMas;
                    TPPol.ColorIndex = 3;
                    TPPol.TextString = stPP;
                    TPPol.XData = new ResultBuffer(
                        new TypedValue(1001, "LAUNCH01"), 
                        new TypedValue(1000, Handl_Det), 
                        new TypedValue(1040, 0), 
                        new TypedValue(1000, "T_POD_POL"), 
                        new TypedValue(1000, LINK),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, AP_LINK),
                        new TypedValue(1000, "Позиция"));
                    TPPol.Layer = "Насыщение";
                    acBlkTblRec.AppendEntity(TPPol);
                    Tx.AddNewlyCreatedDBObject(TPPol, true);
                }

                Line acLine1 = new Line();
                acLine1.SetDatabaseDefaults();
                acLine1.StartPoint = TPol1;
                acLine1.EndPoint = TPol2;
                acLine1.ColorIndex = 1;
                acLine1.Layer = "Насыщение";
                acLine1.XData = new ResultBuffer(
                    new TypedValue(1001, "LAUNCH01"),
                    new TypedValue(1000, ""),
                    new TypedValue(1040, 0),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, AP_LINK),
                    new TypedValue(1000, "Позиция"));
                // Add the line to the drawing
                acBlkTblRec.AppendEntity(acLine1);
                Tx.AddNewlyCreatedDBObject(acLine1, true);
                Tx.Commit();
            }

        } //построение меток с номерами позиций  
        public void CreateLayer(string SloiName)
        {
            ObjectId layerID;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                LayerTable LT = (LayerTable)Trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                if (LT.Has(SloiName))
                { layerID = LT[SloiName]; }
                else
                {
                    LayerTableRecord LTR = new LayerTableRecord();
                    LTR.Name = SloiName;
                    layerID = LT.Add(LTR);
                    Trans.AddNewlyCreatedDBObject(LTR, true);
                }
                Trans.Commit();
            }
        }//Создание слоя
        static public void HCteniPOZTXT(ref List<POZIZIA> SpPOZ, string Adres)
        {
            string[] lines = System.IO.File.ReadAllLines(Adres, Encoding.GetEncoding("Windows-1251"));
            foreach (string Strok in lines)
            {
            POZIZIA TPOZ = new POZIZIA();
            string[] strRazd = Strok.Split(':');
            TPOZ.NNOMpoz(strRazd[0]);
            TPOZ.NCompon(strRazd[1]);
            TPOZ.NPom(strRazd[3]);
            TPOZ.NRazdelSp(strRazd[5]);
                if (strRazd.Count() >6) 
                {
                    TPOZ.NInd(strRazd[1]);
                    TPOZ.NDobav(strRazd[2]);
                    TPOZ.Nshabl(strRazd[6]);
                    TPOZ.NTiPObor(strRazd[7]);
                    TPOZ.NrNAS(strRazd[8]);
                }
            SpPOZ.Add(TPOZ);
            }
        }//чтение файла с номерами позиций
        static public void HCteniSP_LESTN(ref List<string> SpPOZ, string Adres)
        {
            string[] lines = System.IO.File.ReadAllLines(Adres, Encoding.GetEncoding("Windows-1251"));
            foreach (string Strok in lines)SpPOZ.Add(Strok);
        }//чтение файла с номерами позиций
        static public void DOPpoz(ref List<POZIZIA> SpDOP_POZ, ref List<POZIZIA> SpPOZ, string compon, string pom, double Dlin, double Vis, List<POZIZIA> SpPOZsPOZ)
        {
            if (compon == "688-78.1611")
            {
                string comp = "";
                POZIZIA TPOZ = new POZIZIA();
                double Kol = Dlin / 700;
                if (Kol >= 1)
                {
                    Kol= Math.Truncate(Dlin / 700);
                    TPOZ.NNOMpoz("Без поз");
                    TPOZ.NPom(pom);
                    TPOZ.NRazdelSp("Доизол.д");
                    if (Vis < 250)
                    {
                        TPOZ.NCompon("#Проставыш");
                        if (SpPOZsPOZ.Exists(x => x.Compon == "#Проставыш" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "#Проставыш" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                        TPOZ.NpolnNAIM("#Проставыш");
                        TPOZ.NKol(Kol);
                        TPOZ.NKolF(Convert.ToInt16(Kol));
                        TPOZ.NVisot(Vis);
                        TPOZ.NKEI("006");
                        TPOZ.NHtoEto("хв");
                    }
                    else
                    {
                        //if (Vis <= 210 & Vis >= 200) { TPOZ.NCompon("#КЛИГ.746145.001"); }
                        //if (Vis <= 220 & Vis > 210) { TPOZ.NCompon("#КЛИГ.746145.001-01"); }
                        //if (Vis <= 230 & Vis > 220) { TPOZ.NCompon("#КЛИГ.746145.001-02"); }
                        //if (Vis <= 240 & Vis > 230) { TPOZ.NCompon("#КЛИГ.746145.001-03"); }
                        //if (Vis <= 250 & Vis > 240) { TPOZ.NCompon("#КЛИГ.746145.001-04"); }
                        if (Vis <= 260 & Vis > 250) {  comp = "#КЛИГ.746145.001-05"; }
                        if (Vis <= 270 & Vis > 260) {  comp = "#КЛИГ.746145.001-06"; }
                        if (Vis <= 280 & Vis > 270) {  comp = "#КЛИГ.746145.001-07"; }
                        if (Vis <= 290 & Vis > 280) {  comp = "#КЛИГ.746145.001-08"; }
                        if (Vis <= 300 & Vis > 290) {  comp = "#КЛИГ.746145.001-09"; }
                        if (Vis <= 310 & Vis > 300) {  comp = "#КЛИГ.746145.001-10"; }
                        if (Vis <= 320 & Vis > 310) {  comp = "#КЛИГ.746145.001-11"; }
                        if (Vis <= 330 & Vis > 320) {  comp = "#КЛИГ.746145.001-12"; }
                        if (Vis <= 340 & Vis > 330) {  comp = "#КЛИГ.746145.001-13"; }
                        if (Vis <= 350 & Vis > 340) {  comp = "#КЛИГ.746145.001-15"; }
                        if (Vis <= 380 & Vis > 350) {  comp = "#КЛИГ.746145.001-16"; }
                        if (Vis > 380) { comp = "#Самый большой проставыш"; }
                        TPOZ.NCompon(comp);
                        TPOZ.NpolnNAIM(comp);
                        if (SpPOZsPOZ.Exists(x => x.Compon == comp && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == comp && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                        TPOZ.NKol(Kol);
                        TPOZ.NKolF(Convert.ToInt16(Kol));
                        TPOZ.NVisot(0);
                        TPOZ.NKEI("796");
                        TPOZ.NHtoEto("");
                    }
                    SpDOP_POZ.Add(TPOZ);
                }
                Kol = Dlin;
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("УзкаяЛентаНерж");
                TPOZ.NpolnNAIM("УзкаяЛентаНерж");
                TPOZ.NPom(pom);
                TPOZ.NKol(Kol);
                TPOZ.NKolF(1);
                TPOZ.NVisot(Vis);
                TPOZ.NDlin(Dlin);
                TPOZ.NRazdelSp("КрЛот/Лес");
                TPOZ.NKEI("006");
                if (SpPOZsPOZ.Exists(x => x.Compon == "УзкаяЛентаНерж" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "УзкаяЛентаНерж" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                SpDOP_POZ.Add(TPOZ);

                Kol = Math.Round(Dlin/200);
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("УзкийЗамокНерж");
                TPOZ.NpolnNAIM("УзкийЗамокНерж");
                TPOZ.NPom(pom);
                TPOZ.NKol(Kol);
                TPOZ.NKolF(Convert.ToInt16(Kol));
                TPOZ.NRazdelSp("КрЛот/Лес");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);
            }
            if (compon.Contains("305142.002") == true)
            {
                POZIZIA TPOZ = new POZIZIA();
                double Kol = Math.Ceiling((Dlin) / 2000);
                double Kolkr = Math.Ceiling(Dlin/1000);

                if (SpPOZsPOZ.Exists(x => x.Compon == "КЛИГ.363613.006-008" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "КЛИГ.363613.006-008" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                TPOZ.NCompon("КЛИГ.363613.006-008");
                TPOZ.NpolnNAIM("КЛИГ.363613.006-008");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kolkr < 1) TPOZ.NKol(4); else TPOZ.NKol(Kolkr * 4);
                if (Kolkr < 1) TPOZ.NKolF(4); else TPOZ.NKolF(Convert.ToInt16(Kolkr) * 4);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);


                if (SpPOZsPOZ.Exists(x => x.Compon == "ТЛИШ.685614.002-13" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "ТЛИШ.685614.002-13" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                TPOZ.NCompon("ТЛИШ.685614.002-13");
                TPOZ.NpolnNAIM("ТЛИШ.685614.002-13");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(1); else TPOZ.NKol(Kol);
                if (Kol < 1) TPOZ.NKolF(1); else TPOZ.NKolF(Convert.ToInt16(Kol));
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);

                double KolPl = Kol * 2;
                if (SpPOZsPOZ.Exists(x => x.Compon == "ТЛИШ.741124.011" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "ТЛИШ.741124.011" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                TPOZ.NCompon("ТЛИШ.741124.011");
                TPOZ.NpolnNAIM("ТЛИШ.741124.011");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(2); else TPOZ.NKol(KolPl);
                if (Kol < 1) TPOZ.NKolF(2); else TPOZ.NKolF(Convert.ToInt16(KolPl));
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);


                double KolBalt = Kol * 2;
                if (SpPOZsPOZ.Exists(x => x.Compon == "01851363247" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "01851363247" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                TPOZ.NCompon("01851363247");
                TPOZ.NpolnNAIM("01851363247");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(2); else TPOZ.NKol(KolBalt);
                if (Kol < 1) TPOZ.NKolF(2); else TPOZ.NKolF(Convert.ToInt16(KolBalt));
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);


                double KolG = Kol * 4;
                if (SpPOZsPOZ.Exists(x => x.Compon == "01880161010" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "01880161010" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                TPOZ.NCompon("01880161010");
                TPOZ.NpolnNAIM("01880161010");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(4); else TPOZ.NKol(KolG);
                if (Kol < 1) TPOZ.NKolF(4); else TPOZ.NKolF(Convert.ToInt16(KolG));
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);

                double KolH = Kol * 2;
                if (SpPOZsPOZ.Exists(x => x.Compon == "01990108060" && x.Pom == pom)) TPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == "01990108060" && x.Pom == pom).NOMpoz); else TPOZ.NNOMpoz("без поз");
                TPOZ.NCompon("01990108060");
                TPOZ.NpolnNAIM("01990108060");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(2); else TPOZ.NKol(KolH);
                if (Kol < 1) TPOZ.NKolF(2); else TPOZ.NKolF(Convert.ToInt16(KolH));
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);
            }
        }//Дополнительные позиции для перфоволосы
        static public void SBORBl_RAB(ref List<POZIZIA> SpPOZ, string RAB, List<POZIZIA> SpPOZsPOZ, string LINK)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            string sDopDet = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 3);
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
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    sDopDet = "";
                    stHOZ = "";
                    stRab = "";
                    sLINK = "";
                    dVisot = 0;
                    dDlin = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Convert.ToDouble(prop.Value.ToString()) / 1000; }
                        }
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                    if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                    if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                    if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out dDlin)) dVisot = Convert.ToDouble(atrRef.TextString); }
                                }
                            }
                        }
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    else
                    {
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 10) { stRab = value.Value.ToString(); }
                                if (Schet == 11) { sLINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out dDlin)) dVisot = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Ссылка") { stHOZ = atrRef.TextString; }
                                if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                                if (atrRef.Tag == "ДопДетали") { stDobav = atrRef.TextString; }
                            }
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    POZIZIA tPOZ = new POZIZIA();
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    string[] VirezM = stKompM.First().Split('%');
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    tPOZ.NHozain(stHOZ);
                    tPOZ.NRAB(stRab);
                    tPOZ.NRAB(sLINK);
                    if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot); }
                    if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot); }
                    if (stKEI == "796")
                    {
                        //Application.ShowAlertDialog(VirezM.Last().Contains("шт").ToString());
                        if (VirezM.Last().Contains("шт") == true)
                        {
                            string[] KolM = VirezM.Last().Split('ш');
                            int kol = Convert.ToInt32(KolM.First());
                            double dkol = Convert.ToDouble(KolM.First());
                            tPOZ.NKolF(kol);
                            tPOZ.NKol(dkol);
                        }
                        else
                        {
                            tPOZ.NKolF(1);
                            tPOZ.NKol(1);
                        }
                    }
                    //DOBOVLpoz(ref SpPOZ, tPOZ);
                    if ((stKomp == "") == false & (RAB == tPOZ.Hozain | LINK == tPOZ.Hozain))
                    {
                        //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                        if (stDobav != "") { RASKR_DOB(ref SpPOZ, stDobav, stPom, SpPOZsPOZ); }
                        DOBOVLpoz(ref SpPOZ, tPOZ);
                    }
                    SpDOP_POZ.Clear();
                    //strPoz.Add(stKomp + ":" + stPom + ":" + stRazd + ":" + stKEI);
                }
                Tx.Commit();
            }
        }//поиск связаных деталей среди блоков
        static public void SBORpl_RAB(ref List<POZIZIA> SpPOZ, string RAB, List<POZIZIA> SpPOZsPOZ, string LINK,string stRazdGP ,ref double Dlin)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            if (stRazdGP == "Тр") Dlin = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
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
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stHOZ = "";
                    stRab = "";
                    sLINK = "";
                    dVisot = 0;
                    dDlin = 0;
                    Entity bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out dVisot)) dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            if (Schet == 3) { stDobav = value.Value.ToString(); }
                            if (Schet == 4) { stPom = value.Value.ToString(); }
                            if (Schet == 5) { stRazd = value.Value.ToString(); }
                            if (Schet == 6) { stKEI = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 8) { stHOZ = value.Value.ToString(); }
                            if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Convert.ToDouble(value.Value.ToString()); dDlin = Math.Round(dDlin / 10, 0); dDlin = dDlin * 10; }
                            //if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            if (Schet == 10) { stRab = value.Value.ToString(); }
                            if (Schet == 11) { sLINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    POZIZIA tPOZ = new POZIZIA();
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot); }
                    if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot); }
                    if (stKEI == "006" & stHtoEto == "Лес") { tPOZ.NKol(dDlin); }
                    if (stKEI == "006" & stHtoEto == "Тр") { tPOZ.NKol(dDlin); }
                    if (stKEI == "796") { tPOZ.NKol(1); }
                    if ((stKomp != "") & (RAB == sLINK | LINK == sLINK)) 
                    {
                        if (stRazdGP == "Тр") Dlin += dDlin;
                        DOBOVLpoz(ref SpPOZ, tPOZ); 
                        if (stDobav != "")RASKR_DOB(ref SpPOZ, stDobav, stPom, SpPOZsPOZ); 
                    }
                    //strPoz.Add(stKomp + ":" + stPom + ":" + stRazd + ":" + stKEI);
                }
                Tx.Commit();
            }
        }//поиск связаных деталей среди полилиний
        static public void RASKR_DOB(ref List<POZIZIA> SpPOZ, string DOBAVKA, string POM, List<POZIZIA> SpPOZsPOZ)
        {
            string stKomp = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            double dDlin = 0;
            double dVisot = 0;
            string[] DOBAVm = DOBAVKA.Split('#');
            foreach (string TDetal in DOBAVm)
                {
                    if (TDetal != "")
                    {
                        string[] TdetM = TDetal.Split('*');
                        stKomp = TdetM[1];
                        stRazd = TdetM[0];

                        if (stRazd == "ДК" | stRazd == "ДЗ") { stRazd = "Доизол.д"; }
                        if (TdetM.Count() > 3) { stKEI = TdetM[3]; } else { stKEI = "796"; }
                        stDobav = "";
                        stHtoEto = "";
                        stHOZ = "";
                        stRab = "";
                        sLINK = "";
                        dVisot = 0;
                    if (double.TryParse(TdetM[2], out dDlin)) dDlin = Convert.ToDouble(TdetM[2]) / 1000; else dDlin = 0;
                        POZIZIA tPOZ = new POZIZIA();
                        string[] stKompM = stKomp.Split('#');
                        stKomp = stKompM.Last().Replace('@', '.');
                        string[] VirezM = stKompM.First().Split('%');
                        tPOZ.NCompon(stKomp);
                        tPOZ.NVisot(dVisot);
                        tPOZ.NDobav(stDobav);
                        tPOZ.NPom(POM);
                        tPOZ.NRazdelSp(stRazd);
                        tPOZ.NKEI(stKEI);
                        tPOZ.NDlin(dDlin);
                        tPOZ.NHtoEto(stHtoEto);
                        tPOZ.NHozain(stHOZ);
                        tPOZ.NRAB(stRab);
                        tPOZ.NRAB(sLINK);
                        tPOZ.NpolnNAIM(stKomp);
                        if (SpPOZsPOZ.Exists(x => x.Compon == stKomp && x.Pom == POM)) tPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == stKomp && x.Pom == POM).NOMpoz); else tPOZ.NNOMpoz("без поз");
                        if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot); tPOZ.NKolF(dDlin); }
                        if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot); tPOZ.NKolF(dDlin); }
                        if (stKEI == "796")
                        {
                        if (double.TryParse(TdetM[2], out dDlin)) tPOZ.NKolF(Convert.ToInt16(TdetM[2]));
                        if (double.TryParse(TdetM[2], out dDlin)) tPOZ.NKol(Convert.ToDouble(TdetM[2]));
                        }
                        DOBOVLpoz(ref SpPOZ, tPOZ);
                    }
            }
        }//Создание списков деталей из детай в добавке
        static void DOBOVLpoz(ref List<POZIZIA> SpPOZ, POZIZIA tPOZ)
        {
            if (SpPOZ.Exists(x => x.Compon == tPOZ.Compon & x.Pom == tPOZ.Pom & x.RazdelSp == tPOZ.RazdelSp &  x.Visot== tPOZ.Visot) == true)
            {
                POZIZIA izmPOZ = SpPOZ.Find(x => x.Compon == tPOZ.Compon & x.Pom == tPOZ.Pom & x.Visot == tPOZ.Visot);
                SpPOZ.Remove(izmPOZ);
                izmPOZ.NKolF(izmPOZ.KolF + tPOZ.KolF);
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin + 2 * tPOZ.Visot); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "" & izmPOZ.Compon == "688-78.1611" & tPOZ.Visot >= 0.25) { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin + 0.04); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "хв") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Visot); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "Лес") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "Тр") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin); }
                if (izmPOZ.KEI == "796") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Kol); }
                izmPOZ.NDobav(izmPOZ.Dobav + "!" + tPOZ.Dobav);
                SpPOZ.Add(izmPOZ);
            }
            else { SpPOZ.Add(tPOZ); }
        }//добовление позиций в список позиций для проставления 
        static void DOBOVLpozSP(ref List<POZIZIA> SpPOZ, POZIZIA tPOZ)
        {
            if (SpPOZ.Exists(x => x.Compon == tPOZ.Compon & x.Pom == tPOZ.Pom & x.RazdelSp == tPOZ.RazdelSp ) == true)
            {
                POZIZIA izmPOZ = SpPOZ.Find(x => x.Compon == tPOZ.Compon & x.Pom == tPOZ.Pom );
                SpPOZ.Remove(izmPOZ);
                izmPOZ.NKolF(izmPOZ.KolF + tPOZ.KolF);
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin + 2 * tPOZ.Visot); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "" & izmPOZ.Compon == "688-78.1611" & tPOZ.Visot >= 0.25) { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin + 0.04); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "хв") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Visot); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "Лес") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin); }
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "Тр") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin); }
                if (izmPOZ.KEI == "796") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Kol); }
                izmPOZ.NDobav(izmPOZ.Dobav + "!" + tPOZ.Dobav);
                SpPOZ.Add(izmPOZ);
            }
            else { SpPOZ.Add(tPOZ); }
        }//добовление позиций в список позиций для спецификации
        static string HCtenSlov(string Slov,string PoUmol) 
        {
            string ZNACH = PoUmol;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()) 
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForRead);
                            foreach (DBDictionaryEntry Pom in PomD)
                            {
                                Xrecord xRec = (Xrecord)tr.GetObject(Pom.Value, OpenMode.ForRead, false);
                                TypedValue[] rez = xRec.Data.AsArray();
                                ZNACH = rez[0].Value.ToString();
                            }
                        }
                    }

                }
            }
            return ZNACH;
        }//чтение словоря
        static void SOZDlov(string Slov) 
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            DBDictionary NewDict=new DBDictionary();
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                if (!nod.Contains(Slov))
                {
                    nod.SetAt(Slov, NewDict);
                    tr.AddNewlyCreatedDBObject(NewDict, true);
                }
                tr.Commit();
            }
        }//создание словоря если его нет
        static void ZapisSlov(string Slov, string ZNACH)
        {
                        Document doc = Application.DocumentManager.MdiActiveDocument;
                        using (DocumentLock docLock = doc.LockDocument())
                        {
                            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
                            {
                                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                                if (nod.Contains(Slov))
                                {
                                    foreach (DBDictionaryEntry de in nod)
                                    {
                                        if (de.Key == Slov)
                                        {
                                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForWrite);
                                            ResultBuffer reZ = new ResultBuffer(new TypedValue(1000, ZNACH));
                                            Xrecord xRec = new Xrecord();
                                            xRec.Data = reZ;
                                            PomD.SetAt(Slov, xRec);
                                            tr.AddNewlyCreatedDBObject(xRec, true);
                                        }
                                    }

                                }
                                tr.Commit();
                            }
                        }
        }//запись данных в словарь
        //private void radioButton1_CheckedChanged(object sender, EventArgs e)
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    using (DocumentLock docLock = doc.LockDocument())
        //    {
        //        SOZDlov("MAS");
        //        ZapisSlov("MAS", "0.8");
        //        Mas = "0.8";
        //    }
        //}//изменение масштаба 1к25
        //private void radioButton2_CheckedChanged(object sender, EventArgs e)
        //{
        //     Document doc = Application.DocumentManager.MdiActiveDocument;
        //     using (DocumentLock docLock = doc.LockDocument())
        //     {
        //          SOZDlov("MAS");
        //          ZapisSlov("MAS", "1");
        //          Mas = "1";
        //     }
        //}//изменение масштаба 1к20
        //private void radioButton3_CheckedChanged(object sender, EventArgs e)
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    using (DocumentLock docLock = doc.LockDocument())
        //    {
        //        SOZDlov("MAS");
        //        ZapisSlov("MAS", "2");
        //        Mas = "2";
        //    }
        //}//изменение масштаба 1к10
        public double angele(Point3d stPoint, Point3d enPoint) 
        {
            double angele_Is=0;
            Vector3d vek = stPoint.GetVectorTo(enPoint);
            double alfX = vek.GetAngleTo(new Vector3d(1, 0, 0));
            double alfY = vek.GetAngleTo(new Vector3d(0, 1, 0));
            //double tan = (enPoint.Y - stPoint.Y) / (enPoint.X - stPoint.X);
            //double alf = Math.Atan(tan);
            Application.ShowAlertDialog("alfX=" + alfX.ToString() + " alfY=" + alfY.ToString() + " Vek=" + vek.ToString());
            return angele_Is;
        }
        static public void SBORBl(ref List<POZIZIA> SpPOZ, List<POZIZIA> SpPOZsPOZ, ref List<PLOS> SpPLOS)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPOLnNaim = "";
            string LINK = "";
            double dDlin = 0;
            double dVisot = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            string Vid = "", Mash = "", Sprav_plos = "", List = "", Osi = "", Nom = "", KrNaim = "";
            int Schet = 0;
            Point3d BP = new Point3d();
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[6];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 3);
            acTypValAr.SetValue(new TypedValue(8, "Плоскости"), 4);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 5);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей изображенных блоками...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    LINK = "";
                    //для плоскостей
                    Vid = "";
                    Mash = "";
                    Sprav_plos = "";
                    List = "";
                    Osi = "";
                    x1 = 0;
                    y1 = 0;
                    x2 = 0;
                    y2 = 0;
                    xMir = 0;
                    yMir = 0;
                    zMir = 0;
                    Nom = "";
                    KrNaim = "";
                    dVisot = 0;
                    dDlin = 0;
                    double Out = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    BP = bref.Position;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); stPOLnNaim = stKomp; }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Math.Ceiling(Convert.ToDouble(prop.Value.ToString()) / 100) / 10; }
                            if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                        }
                    }
                    else
                    {
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                                if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out Out)) dVisot = Math.Ceiling(Convert.ToDouble(value.Value.ToString()) / 100) / 10; }
                                //if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out Out))dDlin = Math.Ceiling(Convert.ToDouble(value.Value.ToString()) / 100) / 10; }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; stPOLnNaim = stKomp; }
                                if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                                if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out Out)) { dVisot = Math.Ceiling(Convert.ToDouble(atrRef.TextString) / 100) / 10; if (stHtoEto == "Тр") { dDlin = dVisot; dVisot = 0; } } }
                                //if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out Out)) { dVisot = Math.Ceiling(Convert.ToDouble(atrRef.TextString) / 100) / 10; }}
                                if (atrRef.Tag == "Ссылка") { LINK = atrRef.TextString; }
                                if (atrRef.Tag == "ДопДетали") { stDobav = atrRef.TextString; }
                                //плоскости
                                if (atrRef.Tag == "X") { xMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Y") { yMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Z") { zMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "ВИД") { Vid = atrRef.TextString; }
                                if (atrRef.Tag == "МАСШТАБ") { Mash = atrRef.TextString; }
                                if (atrRef.Tag == "ПРИМЕЧАНИЕ") { Sprav_plos = atrRef.TextString; }
                                if (atrRef.Tag == "ЛИСТ") { List = atrRef.TextString; }
                                if (atrRef.Tag == "НОМЕР") { Nom = atrRef.TextString; }
                                if (atrRef.Tag == "КРАДКОЕ_НАИМЕНОВАНИЕ") { KrNaim = atrRef.TextString; }
                            }
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    POZIZIA tPOZ = new POZIZIA();
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    string[] VirezM = stKompM.First().Split('%');
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    tPOZ.NpolnNAIM(stPOLnNaim);
                    tPOZ.NLINK(LINK);
                    if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot); }
                    if (stKEI == "006" & stHtoEto == "Тр") { tPOZ.NKol(dDlin); }
                    if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot); }
                    if (stKEI == "796")
                    {
                        if (VirezM.Last().Contains("шт") == true)
                        {
                            string[] KolM = VirezM.Last().Split('ш');
                            int kol = Convert.ToInt32(KolM.First());
                            double dkol = Convert.ToDouble(KolM.First());
                            tPOZ.NKolF(kol);
                            tPOZ.NKol(dkol);
                        }
                        else
                        {
                            tPOZ.NKolF(1);
                            tPOZ.NKol(1);
                        }
                    }
                    if ((stKomp == "") == false && (stKEI == "") == false)
                    {
                        if (SpPOZsPOZ.Exists(x => x.Compon == stKomp && x.Pom == stPom)) tPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == stKomp && x.Pom == stPom).NOMpoz); else tPOZ.NNOMpoz("без поз");
                        DOBOVLpozSP(ref SpPOZ, tPOZ);
                        DOPpoz(ref SpDOP_POZ, ref SpPOZ, stKomp, stPom, dDlin, dVisot, SpPOZsPOZ);
                            foreach (POZIZIA POZ in SpDOP_POZ) { DOBOVLpozSP(ref SpPOZ, POZ); }
                        SpDOP_POZ.Clear();
                        if (stDobav != "" & stDobav != "не важно") { RASKR_DOB(ref SpPOZ, stDobav, stPom, SpPOZsPOZ); }
                    }
                    if (Vid != "")
                    {
                        Point3d Psk = new Point3d(BP.X + x1, BP.Y + y1, 0);
                        Point3d Msk = new Point3d(xMir, yMir, zMir);
                        PLOS tPL = new PLOS();
                        tPL.Nmax(bref.GeometricExtents.MaxPoint);
                        tPL.Nmin(bref.GeometricExtents.MinPoint);
                        tPL.Nmin_max();
                        tPL.NPsk(Psk);
                        tPL.NMsk(Msk);
                        tPL.NVid(Vid);
                        tPL.NMas(Mash);
                        tPL.NList(List);
                        tPL.NOsi(Osi);
                        tPL.NSpravka(Sprav_plos);
                        tPL.NNomer(Nom);
                        tPL.NKrad(KrNaim);
                        SpPLOS.Add(tPL);
                    }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками
        static public void SBORBl( ref List<PLOS> SpPLOS)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPOLnNaim = "";
            string LINK = "";
            double dDlin = 0;
            double dVisot = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            string Vid = "", Mash = "", Sprav_plos = "", List = "", Osi = "", Nom = "", KrNaim = "";
            int Schet = 0;
            Point3d BP = new Point3d();
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[6];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 3);
            acTypValAr.SetValue(new TypedValue(8, "Плоскости"), 4);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 5);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей изображенных блоками...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    LINK = "";
                    //для плоскостей
                    Vid = "";
                    Mash = "";
                    Sprav_plos = "";
                    List = "";
                    Osi = "";
                    x1 = 0;
                    y1 = 0;
                    x2 = 0;
                    y2 = 0;
                    xMir = 0;
                    yMir = 0;
                    zMir = 0;
                    Nom = "";
                    KrNaim = "";
                    dVisot = 0;
                    dDlin = 0;
                    double Out = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    BP = bref.Position;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); stPOLnNaim = stKomp; }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Math.Ceiling(Convert.ToDouble(prop.Value.ToString()) / 100) / 10; }
                            if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Видимость1") { Osi = prop.Value.ToString(); }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; stPOLnNaim = stKomp; }
                                if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                                if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out Out)) { dVisot = Math.Ceiling(Convert.ToDouble(atrRef.TextString) / 100) / 10; if (stHtoEto == "Тр") { dDlin = dVisot; dVisot = 0; } } }
                                //if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out Out)) { dVisot = Math.Ceiling(Convert.ToDouble(atrRef.TextString) / 100) / 10; }}
                                if (atrRef.Tag == "Ссылка") { LINK = atrRef.TextString; }
                                if (atrRef.Tag == "ДопДетали") { stDobav = atrRef.TextString; }
                                //плоскости
                                if (atrRef.Tag == "X") { xMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Y") { yMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Z") { zMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "ВИД") { Vid = atrRef.TextString; }
                                if (atrRef.Tag == "МАСШТАБ") { Mash = atrRef.TextString; }
                                if (atrRef.Tag == "ПРИМЕЧАНИЕ") { Sprav_plos = atrRef.TextString; }
                                if (atrRef.Tag == "ЛИСТ") { List = atrRef.TextString; }
                                if (atrRef.Tag == "НОМЕР") { Nom = atrRef.TextString; }
                                if (atrRef.Tag == "КРАДКОЕ_НАИМЕНОВАНИЕ") { KrNaim = atrRef.TextString; }
                            }
                        }
                    }
                    if (Vid != "")
                    {
                        Point3d Psk = new Point3d(BP.X + x1, BP.Y + y1, 0);
                        Point3d Msk = new Point3d(xMir, yMir, zMir);
                        PLOS tPL = new PLOS();
                        tPL.Nmax(bref.GeometricExtents.MaxPoint);
                        tPL.Nmin(bref.GeometricExtents.MinPoint);
                        tPL.Nmin_max();
                        tPL.NPsk(Psk);
                        tPL.NMsk(Msk);
                        tPL.NVid(Vid);
                        tPL.NMas(Mash);
                        tPL.NList(List);
                        tPL.NOsi(Osi);
                        tPL.NSpravka(Sprav_plos);
                        tPL.NNomer(Nom);
                        tPL.NKrad(KrNaim);
                        SpPLOS.Add(tPL);
                    }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками
        static public void SBORpl(ref List<POZIZIA> SpPOZ, List<POZIZIA> SpPOZsPOZ,ref List<CWAY> SpCWAYcshert)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPOLnNaim = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
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
                    stKomp = "";
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stHOZ = "";
                    dVisot = 0;
                    dDlin = 0;
                    Entity bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                            if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out dVisot)) dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            //if (Schet == 3) { stDobav = value.Value.ToString(); }
                            if (Schet == 4) { stPom = value.Value.ToString(); }
                            if (Schet == 5) { stRazd = value.Value.ToString(); }
                            if (Schet == 6) { stKEI = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 10) { stHOZ = value.Value.ToString(); }
                            if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Math.Ceiling(Convert.ToDouble(value.Value.ToString()) / 100) / 10; }
                            Schet = Schet + 1;
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    CWAY nCWAY = new CWAY();
                    nCWAY.NCompName(stKomp);
                    nCWAY.NCWName(stHOZ);
                    nCWAY.Nlength(dDlin.ToString());
                    nCWAY.NModuleName(stPom);
                    POZIZIA tPOZ = new POZIZIA();
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NpolnNAIM(stPOLnNaim);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot); }
                    if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot); }
                    if (stKEI == "006" & stHtoEto == "Лес") { tPOZ.NKol(dDlin); }
                    if (stKEI == "006" & stHtoEto == "тр") { tPOZ.NKol(dDlin); }
                    if (stKEI == "796") { tPOZ.NKol(1); }
                    if (stKomp != "" & stHtoEto!="Метка") 
                    {
                        if (SpPOZsPOZ.Exists(x => x.Compon == stKomp && x.Pom == stPom)) tPOZ.NNOMpoz(SpPOZsPOZ.Find(x => x.Compon == stKomp && x.Pom == stPom).NOMpoz); else tPOZ.NNOMpoz("без поз");
                        DOBOVLpozSP(ref SpPOZ, tPOZ);
                        if (stDobav != "" & stDobav != "не важно") { RASKR_DOB(ref SpPOZ, stDobav, stPom, SpPOZsPOZ); }
                        SpCWAYcshert.Add(nCWAY);
                    }
                    //strPoz.Add(stKomp + ":" + stPom + ":" + stRazd + ":" + stKEI);
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных полилиниями
        static public void SBORpl_Dla_T(ref List<Lest> SpLes, List<TPosrL> SpNOD, List<POZIZIA> SpPOZsPOZ, List<POZIZIA> SpPOZ)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPOLnNaim = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            //List<POZIZIA> SpPOZ_XVOST = new List<POZIZIA>();
            //POZIZIA Xvost = new POZIZIA();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
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
                    stKomp = "";
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stHOZ = "";
                    dVisot = 0;
                    dDlin = 0;
                    Entity bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                            //if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            if (Schet == 3) { stDobav = value.Value.ToString(); }
                            if (Schet == 4) { stPom = value.Value.ToString(); }
                            if (Schet == 5) { stRazd = value.Value.ToString(); }
                            if (Schet == 6) { stKEI = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 10) { stHOZ = value.Value.ToString(); }
                            if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            Schet = Schet + 1;
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    Lest tPOZ = new Lest();
                    POZIZIA Xvost = new POZIZIA();
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    tPOZ.NBaza("МСК");
                    tPOZ.Nperem("-");
                    //tPOZ.Nxvost("-");
                    Xvost = SpPOZ.Find(x=>x.LINK== stHOZ);
                    if (SpPOZsPOZ.Exists(x => x.Compon == Xvost.Compon && x.Pom == Xvost.Pom)) tPOZ.Nxvost(SpPOZsPOZ.Find(x => x.Compon == Xvost.Compon && x.Pom == Xvost.Pom).NOMpoz); else tPOZ.Nxvost("-");
                    if (SpPOZsPOZ.Exists(x => x.Compon == stKomp && x.Pom == stPom)) tPOZ.Npoz(SpPOZsPOZ.Find(x => x.Compon == stKomp && x.Pom == stPom).NOMpoz); else tPOZ.Npoz("-");
                    //tPOZ.Npoz("без поз.");
                    tPOZ.Ndobav(stDobav);
                    tPOZ.NNazv(stHOZ);
                    tPOZ.NTip_Pom(stKomp + "_" + stPom);
                    tPOZ.Npom(stPom);
                    tPOZ.NComp(stKomp);
                    tPOZ.NSpT(SpNOD.FindAll(x => x.LINK == stHOZ));
                    if ((stKomp == "") == false) { SpLes.Add(tPOZ); }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных полилиниями
        static public void SBORpl_ID( ref List<string> SpCWAY_ID)
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
        public void ZapolnTabl(List<POZIZIA> strPoz, List<POZIZIA> strPozsPoz)
        {
            this.dataGridView1.Rows.Clear();
            string POlnNaim;
            string Naim="";
            int i = -1;
            Color Zvet=new Color();
                foreach (POZIZIA line in strPoz)
                {
                    Naim = "";
                    if (line.polnNAIM != "") { POlnNaim = line.polnNAIM; } else { POlnNaim = line.Compon; }
                    if (SpNamenkl.Exists(x => x.Compon == line.Compon)) Naim = SpNamenkl.Find(x => x.Compon == line.Compon).Hozain;
                        i = i + 1;
                    this.dataGridView1.Rows.Add(line.NOMpoz, line.Compon, Naim, line.Kol.ToString(), line.Pom, line.KEI, line.RazdelSp, POlnNaim);                  
                }
            ADDlistLader("#ЛестУсил100", "108", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "",ref ListLader);
            ADDlistLader("#ЛестУсил200", "208", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("#ЛестУсил300", "308", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("#ЛестУсил400", "408", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("#ЛестУсил500", "508", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("#ЛестУсил600", "608", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("#ЛестУсил700", "708", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("#ЛестУсил800", "808", "Лес", "", "", "40*40*4#Уголок40х40х4#00311742137", "", "", "1200", "", ref ListLader);
            ADDlistLader("ТрубаЦ25х3,2#01420986011", "34", "Тр", "", "42*60#СоедТрМуфта25#01001902005", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-06*1*796#Тр*КЛИГ.754141.007-16*1*796", "Тр*01002702005*1*796", "1500", "3000", ref ListLader);
            ADDlistLader("Труба32х4ГОСТ8734-75#01362760346", "32", "Тр", "", "45*38#СоедТрШтуцер#ИТШЛ.302615.086-05", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-06*1*796#Тр*КЛИГ.754141.007-16*1*796", "Тр*ИТШЛ.753081.012-06*1*796", "1500", "3000", ref ListLader);
            ADDlistLader("Труба34х4ГОСТ8734-75#01270807175", "34", "Тр", "", "42*60#СоедТрМуфта25#01001902005", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-06*1*796#Тр*КЛИГ.754141.007-16*1*796", "Тр*01002702005*1*796", "1500", "3000", ref ListLader);
            ADDlistLader("ТрубаЦ50х3,5#01420986017", "60", "Тр", "", "66*80#СоедТрМуфта50#01001902008", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-12*1*796#ТЗК*КЛИГ.754141.007-33*1*796", "Тр*01002702008*1*796", "1500", "3000", ref ListLader);
            ADDlistLader("Труба89х4#01450310202", "89", "Тр", "", "20*120#СоедТрФланец#ИТШЛ.712402.005", "25*25*4#Уголок25х25х4#00321808050", "Тр*КЛИГ.301525.015-17*1*796#Тр*КЛИГ.754141.007-46*1*796", "Тр*КЛИГ.741314.044-06*1*796#Тр*01851376479*3*796#Тр*01880174014*3*796", "1500", "3000", ref ListLader);
            foreach (Lestn TL in ListLader) { this.dataGridView6.Rows.Add(TL.Name, TL.Hsir, TL.RazdelSP, TL.Povorot, TL.Soed, TL.Hvost, TL.AddidDetHvpst, TL.AddidDetSoed, TL.StepHvost, TL.StepSoed); }
            this.dataGridView9.Rows.Clear();
            foreach (CWAY tCw in SpCWAYcshert) this.dataGridView9.Rows.Add(tCw.CWName, tCw.CompName, tCw.ModuleName, tCw.length);
        }//заполнение таблицы с деталями
        public void ZapolnTabl_Nam()
        {
            List<string> Namen = HCtenSlovNod("Namen");
            this.dataGridView5.Rows.Clear();
            foreach (string line in Namen)
            {
                string[] NamM = (line + ":::").Split(':');
                POZIZIA Npoz = new POZIZIA();
                Npoz.NCompon(NamM[0]);
                Npoz.NpolnNAIM(NamM[1]);
                Npoz.NHozain(NamM[2]);
                Npoz.NDobav(NamM[4]);
                Npoz.NInd(NamM[3]);
                Npoz.NNOMpoz(NamM[5]);
                Npoz.NPom(NamM[6]);
                SpNamenkl.Add(Npoz);
                this.dataGridView5.Rows.Add(NamM[5], NamM[0], NamM[1], NamM[2], NamM[4], NamM[3], NamM[6]);
            }
        }//чтение словоря с наменклатурой наименований деталей
        public void ZapolnTabl_Sp_Les()
        {
            List<string> Namen = HCtenSlovNod("SPDETPOLY");
            this.dataGridView6.Rows.Clear();
            //Application.ShowAlertDialog(Namen.Count.ToString());
            foreach (string line in Namen)
            {
                string[] strTles = line.Split(':');
                Lestn tLest = new Lestn();
                tLest.NName(strTles[0]);
                tLest.NHsir(Convert.ToDouble(strTles[1]));
                tLest.NRazdelSP(strTles[2]);
                tLest.NPovorot(strTles[3]);
                tLest.NSoed(strTles[4]);
                tLest.NHvost(strTles[5]);
                SpLest.Add(tLest);
                this.dataGridView6.Rows.Add(tLest.Name, tLest.Hsir, tLest.RazdelSP, tLest.Povorot, tLest.Soed, tLest.Hvost);
            }
        }//чтение словоря с наменклатурой наименований деталей
       
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            List<POZIZIA> FSpPOZ = new List<POZIZIA>();
            if (this.textBox3.Text != "")
            {
                FSpPOZ = SpPOZ.FindAll(x => x.polnNAIM.Contains(this.textBox3.Text));
                //Application.ShowAlertDialog(FSpPOZ.Count.ToString());
                ZapolnTabl(FSpPOZ, SpPOZsPOZ);
            }
            else
            {
                ZapolnTabl(SpPOZ, SpPOZsPOZ);
            }
        }//изменение содержимого в текстовом поле
        private void textBox35_TextChanged(object sender, EventArgs e)
        {
            int i = 0;
            List<CWAY> FSpPOZ = new List<CWAY>();
            if (this.textBox35.Text != "")
            {
                FSpPOZ = SpCWAY.FindAll(x => x.CWName.Contains(this.textBox35.Text));
                this.dataGridView7.Rows.Clear();
                foreach (CWAY tCW in FSpPOZ) 
                { 
                    this.dataGridView7.Rows.Add(tCW.CWName, tCW.CompName, tCW.ModuleName);
                    if (SpCWAYcshert.Exists(x => x.CWName == tCW.CWName)) this.dataGridView7.Rows[i].DefaultCellStyle.BackColor = Color.Aqua;
                    i++;
                }
            }
            else
            {
                this.dataGridView7.Rows.Clear();
                foreach (CWAY tCW in SpCWAY) 
                { 
                    this.dataGridView7.Rows.Add(tCW.CWName, tCW.CompName, tCW.ModuleName);
                    if (SpCWAYcshert.Exists(x => x.CWName == tCW.CWName)) this.dataGridView7.Rows[i].DefaultCellStyle.BackColor = Color.Aqua;
                    i++;
                }
            }
        }//Фильтр трасс из ТРАЙБОНА
        public void ADDlistLader(string name, string hir, string razdelSP, string povorot, string soed, string hvost, string addidDetHvpst, string addidDetSoed, string stepHvost, string stepSoed,ref List<Lestn> ListLader) 
        {
            Lestn tLest = new Lestn();
            tLest.Creat( name, Convert.ToDouble(hir),  razdelSP,  povorot,  soed,  hvost,  addidDetHvpst,  addidDetSoed,  stepHvost,  stepSoed);
            ListLader.Add(tLest);
        }
        static public void SBORBl_FIND(ref List<ObjectId> SpOBJID, string FIND_K, string FIND_R, string FIND_P)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            string stPOLnNaim = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 3);
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
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stHOZ = "";
                    stRab = "";
                    sLINK = "";
                    stPOLnNaim = "";
                    dVisot = 0;
                    dDlin = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    stKomp = bref.Name;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); stPOLnNaim = stKomp; }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Convert.ToDouble(prop.Value.ToString()) / 1000; }
                        }
                        //foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        //{
                        //    using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        //    {
                        //        if (atrRef != null)
                        //        {
                        //            if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; ; stPOLnNaim = stKomp; }
                        //            if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                        //            if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                        //            if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                        //            if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); }
                        //        }
                        //    }
                        //}
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    else
                    {
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString() ; stPOLnNaim = stKomp; }
                                if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 10) { stRab = value.Value.ToString(); }
                                if (Schet == 11) { sLINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; ; stPOLnNaim = stKomp; }
                                if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                //if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); }
                            }
                        }
                    }
                    if (stPOLnNaim == FIND_K & stRazd == FIND_R & stPom == FIND_P) { SpOBJID.Add(acSSObj.ObjectId); }
                    SpDOP_POZ.Clear();
                }
                Tx.Commit();
            }

        }//поиск блоков по компоненту
        static public void SBORBl_Gibov( ref List<TPosrL> SpOBJID_Poln, ref List<TPosrL> SpMetok_Poln)
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
            acTypValAr.SetValue(new TypedValue(0, "TEXT"), 2);
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
                    ResultBuffer buffer = Obj.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            foreach (TypedValue value in buffer)
                            {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 4) { stRez = value.Value.ToString(); }
                            if (Schet == 5) { stRad = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 9) { dVisot = value.Value.ToString(); }
                            if (Schet == 11) { LINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                            }
                        }
                    if (Obj.GetType() == typeof(Line))
                    {
                        Line bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Line;
                        if (LINK.Contains( "-"))
                        {
                            TPosrL TPostr = new TPosrL();
                            TPostr.NKomp(stKomp);
                            TPostr.NNomT(0);
                            TPostr.NVisot(dVisot);
                            TPostr.NTPoint(bref.StartPoint);
                            TPostr.NTPoint1(bref.EndPoint);
                            TPostr.NLINK(LINK);
                            TPostr.NTip("Линия гиба");
                            TPostr.NOId(acSSObj.ObjectId);
                            SpOBJID_Poln.Add(TPostr);
                        }
                    }
                    if (stHtoEto== "Метка")
                        {
                            TPosrL TPostr = new TPosrL();
                            TPostr.NKomp(stKomp);
                            TPostr.NNomT(0);
                            TPostr.NVisot(dVisot);  
                            TPostr.NLINK(LINK);
                                if (Obj.GetType() == typeof(Polyline))
                                TPostr.NTip("Линия");
                                else
                                TPostr.NTip("Подпись");
                            TPostr.NOId(acSSObj.ObjectId);
                            SpMetok_Poln.Add(TPostr);
                        }
                }
                Tx.Commit();
            }
        }//создание списков для проставления высоты гибов
        static public void SBORBl_Poz_APr(ref List<TPosrL> SpOBJID_Poln, ref List<TPosrL> SpMetok_Poln,ref List<Point3d> ListPoint1)
        {
            Point3d BazePoint = new Point3d();
            int Schet = 0;
            string stKomp = "";
            string stRez = "";
            string stKompPoz = "";
            string stHtoEto = "";
            string LINK = "";
            string stTip = "";
            string strTipSvOb = "";
            string stPOLnNaim = "";
            string stDobav = "";
            string dVisot = "";
            string NameCWAY = "";
            Point3d PointPostr = new Point3d();
            double NomT = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[11];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 1);
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 2);
            acTypValAr.SetValue(new TypedValue(0, "Circle"), 3);
            acTypValAr.SetValue(new TypedValue(0, "LINE"), 4);
            acTypValAr.SetValue(new TypedValue(0, "TEXT"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 7);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 8);
            //acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 9);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 9);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 10);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
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
                    stKompPoz = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stDobav = "";
                    stTip = "";
                    LINK = "";
                    dVisot = "0";
                    NameCWAY = "";
                    NomT = 0;
                    Schet = 0;
                    Entity Obj = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = Obj.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 4) { stRez = value.Value.ToString(); }
                            if (Schet == 5) { stTip = value.Value.ToString(); }
                            if (Schet == 6) { stKompPoz = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 9) { dVisot = value.Value.ToString(); }
                            if (Schet == 10) { NameCWAY = value.Value.ToString(); }
                            if (Schet == 11) { LINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    if (Obj.GetType() == typeof(BlockReference))
                    {
                        BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        BazePoint = bref.Position;
                        if (bref.IsDynamicBlock)
                        {
                            DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); }
                            }
                        }
                         foreach (ObjectId idAtrRef in bref.AttributeCollection)
                                {using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                    if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                    if (atrRef.Tag == "Ссылка") { LINK = atrRef.TextString; }
                                    if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                                    }
                                }}
                    }
                    if (Obj.GetType() == typeof(Circle)) 
                    {
                        Circle bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Circle;
                        BazePoint = bref.Center;
                    }
                    if (stKomp!="" & stHtoEto != "Метка" & stHtoEto != "Позиция" & stHtoEto != "хв" & Obj.GetType() != typeof(DBText) & stKomp != "Препядствие" & Obj.GetType() != typeof(Line))
                        { 
                            TPosrL TPostr = new TPosrL();
                            TPostr.NKomp(stKomp);
                            TPostr.NNomT(0);
                            TPostr.NVisot(dVisot);
                            TPostr.NTPoint(BazePoint);
                            TPostr.NTPoint1(BazePoint);
                        if (Obj.GetType() == typeof(Polyline))
                        {
                            Point3dCollection SpPoint = new Point3dCollection();
                            Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                            TPostr.NTPoint(bref.StartPoint);
                            TPostr.NTPoint1(bref.GetPointAtParameter(bref.EndParam - 1));
                            for (int i = 0; i <= bref.EndParam; i++) { SpPoint.Add(bref.GetPointAtParameter(i)); }
                            TPostr.NSpKoor(SpPoint);
                        }
                        else
                        {
                            Point3dCollection SpPoint = new Point3dCollection();
                            SpPoint.Add(BazePoint);
                            TPostr.NSpKoor(SpPoint); 
                        }
                            TPostr.NLINK(LINK);
                            TPostr.NNameCWAY(NameCWAY);
                            TPostr.NTip(stHtoEto);
                            if(stTip=="Лес") TPostr.NTip(stTip);
                            TPostr.Dimensions(Obj.GeometricExtents.MaxPoint, Obj.GeometricExtents.MinPoint);
                            TPostr.NOId(acSSObj.ObjectId);
                            SpOBJID_Poln.Add(TPostr);
                    }
                    if (stHtoEto == "Метка" & Obj.GetType() == typeof(Polyline)) 
                        {
                        Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                        ListPoint1.Add(bref.GetPointAtParameter(1));
                        }
                    if (stHtoEto == "Позиция")
                    {
                        TPosrL TPostr = new TPosrL();
                        TPostr.NKomp(stKompPoz);
                        TPostr.NNomT(0);
                        TPostr.NVisot(dVisot);
                        TPostr.NLINK(LINK);
                        if (Obj.GetType() == typeof(Polyline))
                            TPostr.NTip("Линия");
                        else
                            TPostr.NTip("Подпись");
                        TPostr.NOId(acSSObj.ObjectId);
                        SpMetok_Poln.Add(TPostr);
                    }
                }
                Tx.Commit();
            }
        }//создание списков для проставленя позиций
        static public void SBORBl_Poz_APr(ref List<BLOK.Form3.TPosrL> SpOBJID_Poln, ref List<TPosrL> SpMetok_Poln, ref List<Point3d> ListPoint1)
        {
            Point3d BazePoint = new Point3d();
            int Schet = 0;
            string stKomp = "";
            string stRez = "";
            string stKompPoz = "";
            string stHtoEto = "";
            string LINK = "";
            string stTip = "";
            string strTipSvOb = "";
            string stPOLnNaim = "";
            string stDobav = "";
            string dVisot = "";
            string NameCWAY = "";
            Point3d PointPostr = new Point3d();
            double NomT = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[11];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 1);
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 2);
            acTypValAr.SetValue(new TypedValue(0, "Circle"), 3);
            acTypValAr.SetValue(new TypedValue(0, "LINE"), 4);
            acTypValAr.SetValue(new TypedValue(0, "TEXT"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 7);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 8);
            //acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 9);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 9);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 10);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
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
                    stKompPoz = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stDobav = "";
                    stTip = "";
                    LINK = "";
                    dVisot = "0";
                    NameCWAY = "";
                    NomT = 0;
                    Schet = 0;
                    Entity Obj = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = Obj.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 4) { stRez = value.Value.ToString(); }
                            if (Schet == 5) { stTip = value.Value.ToString(); }
                            if (Schet == 6) { stKompPoz = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 9) { dVisot = value.Value.ToString(); }
                            if (Schet == 10) { NameCWAY = value.Value.ToString(); }
                            if (Schet == 11) { LINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    if (Obj.GetType() == typeof(BlockReference))
                    {
                        BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        BazePoint = bref.Position;
                        if (bref.IsDynamicBlock)
                        {
                            DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); }
                            }
                        }
                        foreach (ObjectId idAtrRef in bref.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                    if (atrRef.Tag == "Ссылка") { LINK = atrRef.TextString; }
                                    if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                                }
                            }
                        }
                    }
                    if (Obj.GetType() == typeof(Circle))
                    {
                        Circle bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Circle;
                        BazePoint = bref.Center;
                    }
                    if (stKomp != "" & stHtoEto != "Метка" & stHtoEto != "Позиция" & stHtoEto != "хв" & Obj.GetType() != typeof(DBText) & stKomp != "Препядствие" & Obj.GetType() != typeof(Line))
                    {
                        BLOK.Form3.TPosrL TPostr = new BLOK.Form3.TPosrL();
                        TPostr.NKomp(stKomp);
                        TPostr.NNomT(0);
                        TPostr.NVisot(dVisot);
                        TPostr.NTPoint(BazePoint);
                        TPostr.NTPoint1(BazePoint);
                        if (Obj.GetType() == typeof(Polyline))
                        {
                            Point3dCollection SpPoint = new Point3dCollection();
                            Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                            TPostr.NTPoint(bref.StartPoint);
                            TPostr.NTPoint1(bref.GetPointAtParameter(bref.EndParam - 1));
                            for (int i = 0; i <= bref.EndParam; i++) { SpPoint.Add(bref.GetPointAtParameter(i)); }
                            TPostr.NSpKoor(SpPoint);
                        }
                        else
                        {
                            Point3dCollection SpPoint = new Point3dCollection();
                            SpPoint.Add(BazePoint);
                            TPostr.NSpKoor(SpPoint);
                        }
                        TPostr.NLINK(LINK);
                        TPostr.NNameCWAY(NameCWAY);
                        TPostr.NTip(stHtoEto);
                        if (stTip == "Лес") TPostr.NTip(stTip);
                        TPostr.Dimensions(Obj.GeometricExtents.MaxPoint, Obj.GeometricExtents.MinPoint);
                        TPostr.NOId(acSSObj.ObjectId);
                        SpOBJID_Poln.Add(TPostr);
                    }
                    if (stHtoEto == "Метка" & Obj.GetType() == typeof(Polyline))
                    {
                        Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                        ListPoint1.Add(bref.GetPointAtParameter(1));
                    }
                    if (stHtoEto == "Позиция")
                    {
                        TPosrL TPostr = new TPosrL();
                        TPostr.NKomp(stKompPoz);
                        TPostr.NNomT(0);
                        TPostr.NVisot(dVisot);
                        TPostr.NLINK(LINK);
                        if (Obj.GetType() == typeof(Polyline))
                            TPostr.NTip("Линия");
                        else
                            TPostr.NTip("Подпись");
                        TPostr.NOId(acSSObj.ObjectId);
                        SpMetok_Poln.Add(TPostr);
                    }
                }
                Tx.Commit();
            }
        }//создание списков для проставленя позиций
        static public void SBORBl_Poz_Dim(ref List<TPosrL> SpOBJID_Poln, ref List<TPosrL> SpMetok_Poln)
        {
            int Schet = 0;
            string stKomp = "";
            string stRez = "";
            string stKompPoz = "";
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
            TypedValue[] acTypValAr = new TypedValue[12];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 1);
            acTypValAr.SetValue(new TypedValue(0, "TEXT"), 2);
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 3);
            acTypValAr.SetValue(new TypedValue(0, "LINE"), 4);
            acTypValAr.SetValue(new TypedValue(0, "Circle"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 7);
            acTypValAr.SetValue(new TypedValue(8, "CWAYPoint"), 8);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 9);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 10);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 11);
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
                    stKompPoz = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stDobav = "";
                    stTip = "";
                    LINK = "";
                    dVisot = "0";
                    NomT = 0;
                    Schet = 0;
                    Entity Obj = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = Obj.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 4) { stRez = value.Value.ToString(); }
                            if (Schet == 6) { stKompPoz = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 9) { dVisot = value.Value.ToString(); }
                            if (Schet == 11) { LINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    if (Obj.GetType() == typeof(Polyline))
                    {
                        Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                        if (stKomp != "" & stHtoEto != "Метка")
                        {
                            Point3dCollection SpPoint = new Point3dCollection();
                            TPosrL TPostr = new TPosrL();
                            TPostr.NKomp(stKomp);
                            TPostr.NNomT(0);
                            TPostr.NVisot(dVisot);
                            TPostr.NTPoint(bref.StartPoint);
                            TPostr.NTPoint1(bref.GetPointAtParameter(bref.EndParam - 1));
                            for (int i = 0; i <= bref.EndParam; i++) { SpPoint.Add(bref.GetPointAtParameter(i)); }
                            //Application.ShowAlertDialog(bref.EndParam.ToString());
                            TPostr.NSpKoor(SpPoint);
                            TPostr.NLINK(LINK);
                            TPostr.NTip(stHtoEto);
                            TPostr.NOId(acSSObj.ObjectId);
                            if (!SpOBJID_Poln.Exists(x => x.LINK == LINK)) SpOBJID_Poln.Add(TPostr);
                        }
                    }
                    if (stKomp == "Препядствие")
                    {
                        Line bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Line;
                        TPosrL TPostr = new TPosrL();
                        TPostr.NKomp(stKomp);
                        TPostr.NTPoint(bref.StartPoint);
                        TPostr.NTPoint1(bref.EndPoint);
                        TPostr.NTPointZ();
                        TPostr.getAngel();
                        TPostr.NOId(acSSObj.ObjectId);
                        SpMetok_Poln.Add(TPostr);
                        //Application.ShowAlertDialog(TPostr.OId + ":Gorizont-(" + TPostr.AngelG + ")Vertikal(" + TPostr.AngelV + ")");
                    }
                }
                Tx.Commit();
            }
        }//создание списков для проставленя размеров от препядствий
        static public List<Point3d> SBORBl_Poz_Dim_Gird(ref List<TPosrL> SpOBJID_Poln)
        {
            List<Point3d> ListPointGirt = new List<Point3d>();
            List<TPosrL> SpGor=new List<TPosrL>();
            List<TPosrL> SpVert = new List<TPosrL>();
            int Schet = 0;
            string stKomp = "";
            string stRez = "";
            string stKompPoz = "";
            string stHtoEto = "";
            string LINK = "";
            string stTip = "";
            string strTipSvOb = "";
            string stPOLnNaim = "";
            string stDobav = "";
            string dVisot = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[8];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "LINE"), 1); 
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 2);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 4);
            acTypValAr.SetValue(new TypedValue(8, "Плоскости"), 5);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 6);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 7);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей...");
                return ListPointGirt;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stRez = "";
                    stKompPoz = "";
                    stHtoEto = "";
                    stPOLnNaim = "";
                    stDobav = "";
                    stTip = "";
                    LINK = "";
                    dVisot = "0";
                    Schet = 0;
                    Entity Obj = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = Obj.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            if (Schet == 4) { stRez = value.Value.ToString(); }
                            if (Schet == 6) { stKompPoz = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 9) { dVisot = value.Value.ToString(); }
                            if (Schet == 11) { LINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    if (Obj.GetType() == typeof(Polyline))
                    {
                        Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                        if (stKomp != "" & stHtoEto != "Метка")
                        {
                            Point3dCollection SpPoint = new Point3dCollection();
                            TPosrL TPostr = new TPosrL();
                            TPostr.NKomp(stKomp);
                            TPostr.NNomT(0);
                            TPostr.NVisot(dVisot);
                            TPostr.NTPoint(bref.StartPoint);
                            TPostr.NTPoint1(bref.GetPointAtParameter(bref.EndParam - 1));
                            for (int i = 0; i <= bref.EndParam; i++) { SpPoint.Add(bref.GetPointAtParameter(i)); }
                            TPostr.Dimensions(bref.GeometricExtents.MaxPoint, bref.GeometricExtents.MinPoint);
                            TPostr.NSpKoor(SpPoint);
                            TPostr.NLINK(LINK);
                            TPostr.NTip(stHtoEto);
                            TPostr.NOId(acSSObj.ObjectId);
                            if (!SpOBJID_Poln.Exists(x => x.LINK == LINK)) SpOBJID_Poln.Add(TPostr);
                        }
                    }
                    if (stKomp == "Gor")
                    {
                        Line bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Line;
                        TPosrL TPostr = new TPosrL();
                        TPostr.NTPoint(bref.StartPoint);
                        TPostr.NTPoint1(bref.EndPoint);
                        SpGor.Add(TPostr);
                    }
                    if (stKomp == "Vert")
                    {
                        Line bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Line;
                        TPosrL TPostr = new TPosrL();
                        TPostr.NTPoint(bref.StartPoint);
                        TPostr.NTPoint1(bref.EndPoint);
                        SpVert.Add(TPostr);
                    }
                }
                Tx.Commit();
            }
            foreach (TPosrL CurLinGor in SpGor) 
            {
                foreach (TPosrL CurLinVert in SpVert) 
                {
                    Point3d Point = TPer1(CurLinGor.TPoint, CurLinGor.TPoint1, CurLinVert.TPoint1, CurLinVert.TPoint);
                    if (Point != new Point3d()) { ListPointGirt.Add(Point); }
                }
            }
            return ListPointGirt;
        }//создание списков для проставленя размеров
        static public void SBORpl_FIND(ref List<ObjectId> SpOBJID, string FIND_K, string FIND_R, string FIND_P)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stHOZ = "";
            string stRab = "";
            string sLINK = "";
            string stPOLnNaim = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
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
                    stPom = "";
                    stRazd = "";
                    stKEI = "";
                    stDobav = "";
                    stHtoEto = "";
                    stHOZ = "";
                    stRab = "";
                    sLINK = "";
                    stPOLnNaim = "";
                    dVisot = 0;
                    dDlin = 0;
                    Entity bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); stPOLnNaim = stKomp; }
                            if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out dVisot)) dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            if (Schet == 3) { stDobav = value.Value.ToString(); }
                            if (Schet == 4) { stPom = value.Value.ToString(); }
                            if (Schet == 5) { stRazd = value.Value.ToString(); }
                            if (Schet == 6) { stKEI = value.Value.ToString(); }
                            if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                            if (Schet == 8) { stHOZ = value.Value.ToString(); }
                            if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Convert.ToDouble(value.Value.ToString()) / 1000; }
                            if (Schet == 10) { stRab = value.Value.ToString(); }
                            if (Schet == 11) { sLINK = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    if (stKomp == FIND_K & stRazd == FIND_R & stPom == FIND_P) { SpOBJID.Add(acSSObj.ObjectId); }
                }
                ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                ed.SetImpliedSelection(idarrayEmpty);
                Tx.Commit();
            }
        }//поиск полилиний по компоненту

        public void SozSpLest(ref List<Lestn> SpLest, List<string> StrSpLESTN) 
        {
            foreach (string str in StrSpLESTN) 
            {
                string[] strTles = str.Split(':');
                Lestn tLest = new Lestn();
                tLest.NName(strTles[0]);
                tLest.NHsir(Convert.ToDouble(strTles[1]));
                tLest.NRazdelSP(strTles[2]);
                tLest.NPovorot(strTles[3]);
                tLest.NSoed(strTles[4]);
                tLest.NHvost(strTles[5]);
                SpLest.Add(tLest);
            }
        }//создание списка лестниц

        public void SOZD_Sp_TP_Lest(ref List<TPosrL> SpNOD)
        {
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[3];
            acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 0);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 1);
            acTypValAr.SetValue(new TypedValue(40, 10.0), 2);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    TPosrL TNOD = new TPosrL();
                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
                        if (ln != null)
                        {
                            TNOD.NTPoint(ln.Center);
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { if(double.TryParse(value.Value.ToString(),out double rez)) TNOD.NNomT(Convert.ToDouble(value.Value.ToString())); }
                                    if (Schet == 2) { TNOD.NVisot(value.Value.ToString()); }
                                    if (Schet == 8) { TNOD.NLINK(value.Value.ToString()); }
                                    Schet = Schet + 1;
                                    //if (Schet > 2) { break; }
                                }
                            }
                            TNOD.NTPoint(ln.Center);
                            //Application.ShowAlertDialog(TNOD.IND + "," + TNOD.Sist + "," + TNOD.BlNOD);
                            SpNOD.Add(TNOD);
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка точек построения лестниц кругами
        public void SOZD_Sp_TP_Lest_BL(ref List<TPosrL> SpNOD)
        {
            int Schet = 0;
            string NameT, Visot,Hoz;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(8, "CWAYPoint"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    
                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        NameT = "";
                        Visot = "";
                        Hoz="";
                        BlockReference ln = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                        if (ln != null)
                        {
                            
                            foreach (ObjectId idAtrRef in ln.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Исполнение") { NameT = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") { Visot = atrRef.TextString; }
                                        if (atrRef.Tag == "Ссылка") { Hoz = atrRef.TextString; }
                                        //if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); }
                                    }
                                }
                            }
                            if (NameT != "") 
                            {
                                TPosrL TNOD = new TPosrL();
                                TNOD.NTPoint(ln.Position);
                                if(double.TryParse(NameT,out double rez)) TNOD.NNomT(Convert.ToDouble(NameT));
                                TNOD.NVisot(Visot);
                                TNOD.NLINK(Hoz);
                                SpNOD.Add(TNOD);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка точек построения лестниц блоками
        public void SOZD_Sp_TP_Lest_BL(ref List<BLOK.Form3.TPosrL> SpNOD)
        {
            int Schet = 0;
            string NameT, Visot, Hoz;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(8, "CWAYPoint"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;

                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        NameT = "";
                        Visot = "";
                        Hoz = "";
                        BlockReference ln = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                        if (ln != null)
                        {

                            foreach (ObjectId idAtrRef in ln.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Исполнение") { NameT = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") { Visot = atrRef.TextString; }
                                        if (atrRef.Tag == "Ссылка") { Hoz = atrRef.TextString; }
                                        //if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); }
                                    }
                                }
                            }
                            if (NameT != "")
                            {
                                BLOK.Form3.TPosrL TNOD = new BLOK.Form3.TPosrL();
                                TNOD.NTPoint(ln.Position);
                                if (double.TryParse(NameT, out double rez)) TNOD.NNomT(Convert.ToDouble(NameT));
                                TNOD.NVisot(Visot);
                                TNOD.NLINK(Hoz);
                                SpNOD.Add(TNOD);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка точек построения лестниц блоками
        public void zapDGWLes(List<Lestn> SpLest) 
        {
            foreach (Lestn tLes in SpLest) this.dataGridView6.Rows.Add(tLes.Name, tLes.Hsir, tLes.RazdelSP, tLes.Povorot, tLes.Soed, tLes.Hvost);
        }

        public void ZagrExcel(string File)
        {
            List<string> VKtxt = new List<string>();
            List<string> VKEcxel = new List<string>();
            List<Zag> SpZag = new List<Zag>();
            List<Zag> Hapka = new List<Zag>();
            SpisZag(ref SpZag, "Пом.");
            SpisZag(ref SpZag, "Поз.");
            SpisZag(ref SpZag, "Условное наименование");
            SpisZag(ref SpZag, "Обозначение");
            SpisZag(ref SpZag, "Наименование");
            SpisZag(ref SpZag, "Код ОКП");
            SpisZag(ref SpZag, "Масса ед");
            //считываем данные из Excel файла в двумерный массив
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга              
            Excel.Worksheet xlShtVK; //лист Excel 
            //Excel.Worksheets SPxlShtVK; //лист Excel 
            //Excel.Worksheet xlShtTR; //лист Excel
            //this.label33.Text = File;
            xlWB = xlApp.Workbooks.Open(@File); //название файла Excel                                             
            //xlSht = xlWB.Worksheets["Кабельный"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            //xlShtVK = xlWB.Worksheets;
            xlShtVK = xlWB.Worksheets[4]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            //xlShtVK = xlWB.Worksheets.get_Item(2);
            //xlShtTR = xlWB.Worksheets[14]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            ZagrLisExcel(xlShtVK, ref VKtxt, SpZag, Hapka, ref  VKEcxel);
            //ZagrLisExcel(xlShtTR, ref VKtxt, SpZag, Hapka, ref  VKEcxel);
            //int iLastRow = xlSht.Cells[xlSht.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
            //var arrData = (object[,])xlSht.Range["A1:AA" + iLastRow].Value; //берём данные с листа Excel
            ////xlApp.Visible = true; //отображаем Excel     
            xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel
            //настройка DataGridView
            this.dataGridView2.Rows.Clear();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                foreach (string Strok in VKEcxel)
                {
                    VKtxt.Add(Strok);
                }
                SOZDlov("Namen");
                SOZDlov("IstokNm");
                ZapisSlovSistTR("Namen", VKtxt);
                ZapisSlov("IstokNm", File);
            }
        }
        public void ZagrExcelCW(string File)
        {
            string Zagal = "";
            string CWName = "";
            string CompName = "";
            string StartPoint = "";
            string EndPoint = "";
            string Coordinates = "";
            string ModuleName = "";
            string Vibor = "";
            int Nom = 0;
            List<string> VKtxt = new List<string>();
            List<string> VKEcxel = new List<string>();
            List<Zag> SpZag = new List<Zag>();
            List<Zag> Hapka = new List<Zag>();
            SpisZag(ref SpZag, "CW Name");
            SpisZag(ref SpZag, "Comp Name");
            SpisZag(ref SpZag, "Start point");
            SpisZag(ref SpZag, "End point");
            SpisZag(ref SpZag, "Coordinates");
            SpisZag(ref SpZag, "Module name");
            SpisZag(ref SpZag, "Выбор");
            //считываем данные из Excel файла в двумерный массив
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга              
            Excel.Worksheet xlSht; //лист Excel   
            //this.label33.Text = File;
            xlWB = xlApp.Workbooks.Open(@File); //название файла Excel                                             
            xlSht = xlWB.Worksheets["трассы"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
            var arrData = (object[,])xlSht.Range["A1:J" + iLastRow].Value; //берём данные с листа Excel
            //xlApp.Visible = true; //отображаем Excel     
            xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel
            //настройка DataGridView
            this.dataGridView2.Rows.Clear();
            int RowsCount = arrData.GetUpperBound(0);
            int ColumnsCount = arrData.GetUpperBound(1);
            //dataGridView2.RowCount = RowsCount; //кол-во строк в DGV
            //dataGridView2.ColumnCount = ColumnsCount; //кол-во столбцов в DGV
            //заполняем DataGridView данными из массива
            int i, j;
            for (j = 1; j <= ColumnsCount; j++)
            {
                if (arrData[1, j] != null)
                {
                    Zagal = arrData[1, j].ToString();
                    foreach (Zag tZag in SpZag)
                    {
                        if (tZag.Zagal == Zagal)
                        {
                            tZag.NomGraf(j); Hapka.Add(tZag);
                        }
                    }
                }
            }
            for (i = 2; i <= RowsCount; i++)
            {
                CWName = "";
                CompName = "";
                StartPoint = "";
                EndPoint = "";
                Coordinates = "";
                ModuleName = "";
                Vibor = "";
                for (j = 1; j <= ColumnsCount; j++)
                {
                    if (arrData[i, j] != null)
                    {
                        Zagal = arrData[i, j].ToString();
                        foreach (Zag tZag in Hapka)
                        {
                            if (tZag.Zagal == "CW Name" & tZag.Graf == j) CWName = Zagal;
                            if (tZag.Zagal == "Comp Name" & tZag.Graf == j) CompName = Zagal;
                            if (tZag.Zagal == "Start point" & tZag.Graf == j) StartPoint = Zagal;
                            if (tZag.Zagal == "End point" & tZag.Graf == j) EndPoint = Zagal;
                            if (tZag.Zagal == "Coordinates" & tZag.Graf == j) Coordinates = Zagal;
                            if (tZag.Zagal == "Module name" & tZag.Graf == j) ModuleName = Zagal;
                            if (tZag.Zagal == "Выбор" & tZag.Graf == j) Vibor = Zagal;
                        }
                    }

                }
                if (Vibor != "")
                {
                    Part TPatr = new Part();
                    TPatr.NCWName(CWName);
                    TPatr.NCompName(CompName);
                    TPatr.NStartPoint(StartPoint);
                    TPatr.NEndPoint(EndPoint);
                    TPatr.NCoordinates(Coordinates);
                    TPatr.NModuleName(ModuleName);
                    //Application.ShowAlertDialog(CWName + ":" + CompName + ":" + ModuleName + ":" + StartPoint + ":" + EndPoint + ":" + Coordinates);
                    SpPart.Add(TPatr);
                    if (SpCWAY.Exists(x => x.CWName == CWName) == false)
                    {
                        //Application.ShowAlertDialog(CWName + ":" + CompName + ":" + ModuleName);
                        CWAY TCWAY = new CWAY();
                        TCWAY.NCWName(CWName);
                        TCWAY.NCompName(CompName);
                        TCWAY.NModuleName(ModuleName);
                        SpCWAY.Add(TCWAY);
                    }
                }
            }
            this.dataGridView7.Rows.Clear();
            foreach (CWAY tCW in SpCWAY) this.dataGridView7.Rows.Add(tCW.CWName, tCW.CompName, tCW.ModuleName);
        }
         public static void LoadCWinTXT(string Adres,ref List<Part> SpPart,ref List<CWAY> SpCWAY)
        {
            string[] lines = System.IO.File.ReadAllLines(Adres, Encoding.GetEncoding("Windows-1251"));
            foreach (string Strok in lines)
            {
                Part TPatr = new Part();
                string[] strRazd = Strok.Split(':');
                TPatr.NCWName(strRazd[0]);
                TPatr.NCompName(strRazd[2]);
                TPatr.NStartPoint(strRazd[3]);
                TPatr.NEndPoint(strRazd[4]);
                TPatr.NCoordinates(strRazd[5]);
                TPatr.NModuleName(strRazd[6]);
                if(strRazd[2]!="") SpPart.Add(TPatr);
                if (SpCWAY.Exists(x => x.CWName == strRazd[0]) == false & strRazd[2] != "")
                {
                    CWAY TCWAY = new CWAY();
                    TCWAY.NCWName(strRazd[0]);
                    TCWAY.NCompName(strRazd[2]);
                    TCWAY.NModuleName(strRazd[6]);
                    SpCWAY.Add(TCWAY);
                }
            }
        }//чтение файла с трассами
        public void ZagrLisExcel(Excel.Worksheet xlSht, ref List<string> VKtxt, List<Zag> SpZag, List<Zag> Hapka, ref List<string> VKEcxel)
        {
            string Pom = "";
            string Poz = "";
            string UslN = "";
            string Obozn = "";
            string Naimen = "";
            string Mass = "";
            string Zagal = "";
            string Code = "";
            //xlSht = xlWB.Worksheets[4]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "C"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
            var arrData = (object[,])xlSht.Range["A1:M" + iLastRow].Value; //берём данные с листа Excel
            //xlApp.Visible = true; //отображаем Excel     
            //xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            //xlApp.Quit(); //закрываем Excel
            //настройка DataGridView
            //this.dataGridView2.Rows.Clear();
            int RowsCount = arrData.GetUpperBound(0);
            int ColumnsCount = arrData.GetUpperBound(1);
            //dataGridView2.RowCount = RowsCount; //кол-во строк в DGV
            //dataGridView2.ColumnCount = ColumnsCount; //кол-во столбцов в DGV
            //заполняем DataGridView данными из массива
            int i, j;
            Hapka.Clear();
            for (j = 1; j <= ColumnsCount; j++)
            {
                if (arrData[1, j] != null)
                {
                    Zagal = arrData[1, j].ToString();
                    foreach (Zag tZag in SpZag)
                    {
                        if (tZag.Zagal == Zagal)
                        {
                            tZag.NomGraf(j); Hapka.Add(tZag);
                        }
                    }
                }
            }
            for (i = 2; i <= RowsCount; i++)
            {
                 Pom = "";
                 Poz = "";
                 UslN = "";
                 Obozn = "";
                 Naimen = "";
                 Mass = "";
                 Code = "";
                for (j = 1; j <= ColumnsCount; j++)
                {
                    if (arrData[i, j] != null)
                    {
                        if (arrData[i, j] != null)
                            Zagal = arrData[i, j].ToString();
                        else
                            Zagal = "";
                        foreach (Zag tZag in Hapka)
                        {
                            if (tZag.Zagal == "Условное наименование" & tZag.Graf == j) UslN = Zagal;
                            if (tZag.Zagal == "Обозначение" & tZag.Graf == j) Obozn = Zagal;
                            if (tZag.Zagal == "Наименование" & tZag.Graf == j) Naimen = Zagal;
                            if (tZag.Zagal == "Масса ед" & tZag.Graf == j) Mass = Zagal;
                            if (tZag.Zagal == "Код ОКП" & tZag.Graf == j) Code = Zagal;
                            if (tZag.Zagal == "Поз." & tZag.Graf == j) Poz = Zagal;
                            if (tZag.Zagal == "Пом." & tZag.Graf == j) Pom = Zagal;
                        }
                    }
                }
                if (Naimen != "")
                {
                    POZIZIA tPoz = new POZIZIA();
                    tPoz.NCompon(UslN);
                    tPoz.NpolnNAIM(Naimen);
                    tPoz.NHozain(Obozn);
                    tPoz.NDobav(Mass);
                    tPoz.NInd(Code);
                    tPoz.NNOMpoz(Poz);
                    tPoz.NPom(Pom);
                    SpNamenkl.Add(tPoz);
                    VKEcxel.Add(UslN + ":" + Obozn + ":" + Naimen + ":" + Mass + ":" + Code + ":" + Poz + ":" + Pom);
                }
            }
        }
        public void SpisZag(ref List<Zag> spZag, string tZag)
        {
            Zag nZag = new Zag();
            nZag.NomZagal(tZag);
            spZag.Add(nZag);
        }//пополнение списка загаловков
        static void ZapisSlovSistTR(string Slov, List<string> ZNACH)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForWrite);
                            ResultBuffer reZ = new ResultBuffer();
                            foreach (string TNod in ZNACH) reZ.Add(new TypedValue(1000, TNod));
                            Xrecord xRec = new Xrecord();
                            xRec.Data = reZ;
                            PomD.SetAt(Slov, xRec);
                            tr.AddNewlyCreatedDBObject(xRec, true);
                        }
                    }

                }
                tr.Commit();
            }
        }//запись данных в словарь список данных
        static List<string> HCtenSlovNod(string Slov)
        {
            List<string> ZNACH = new List<string>();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForRead);
                            foreach (DBDictionaryEntry Pom in PomD)
                            {
                                Xrecord xRec = (Xrecord)tr.GetObject(Pom.Value, OpenMode.ForRead, false);
                                TypedValue[] rez = xRec.Data.AsArray();
                                foreach (TypedValue valSl in rez)
                                {
                                    ZNACH.Add(valSl.Value.ToString());
                                }
                            }
                        }
                    }

                }
            }
            return ZNACH;
        }
        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            //DataGridViewCellStyle DefaultCellStyle = new DataGridViewCellStyle { Font = new Font("Tahoma", 9.75F, FontStyle.Bold), ForeColor = Color.Black };
            //this.dataGridView1.CurrentCell.Style = DefaultCellStyle;
        }
        public static void Lest_N_INtr_SP(List<TPosrL> SpPointVH, string Tipor, double Hsirin, string Zagal, string Povorot, string Xvost, string Soed, string Adr, double Otstup, double HSag_Hvost, string Pom, string Name, bool BildDetFast)
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
                string ID_Les = "0";
                if (Name == "")
                {
                    SBORpl_ID(ref SpID);
                    while (SpID.Exists(x => x == ID_Les)){ID_Les = (Convert.ToDouble(ID_Les) + 1).ToString(); }
                }
                else
                    ID_Les = Name;
                double ID = 0;
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
                            if (Math.Abs(alf - Math.PI / 2) < 0.1 | Math.Abs(alf - 3 * Math.PI / 2) < 0.1)
                            {
                                if (SpPoint.Last().TPoint.DistanceTo(SpPoint[SpPoint.Count - 2].TPoint) > Otstup & SpPoint.Last().TPoint.DistanceTo(TPoint.TPoint) > Otstup)
                                {
                                    Vector3d Vek1 = SpPoint.Last().TPoint.GetVectorTo(SpPoint[SpPoint.Count - 2].TPoint);
                                    Vector3d Vek2 = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
                                    if ((Vek1.X < 0 & Vek2.Y < 0) | (Vek1.Y < 0 & Vek2.X < 0)) Ugol = 3 * (Math.PI) / 2;
                                    if (Vek1.X > 0 & Vek2.Y > 0 | (Vek1.Y > 0 & Vek2.X > 0)) Ugol = (Math.PI) / 2;
                                    if (Vek1.X < 0 & Vek2.Y > 0 | (Vek1.Y > 0 & Vek2.X < 0)) Ugol = Math.PI;
                                    FPovorot(SpPoint.Last().TPoint, "", "#ПоворотГориз90ЛестУсил" + (Hsirin - 8).ToString(), Ugol, "0", Hsirin, Pom, ID_Les.ToString());
                                    double alf1 = Math.Atan(tan);
                                    double alf2 = Math.Atan(tan1);
                                    if (alf1 == 0) alf1 = Math.PI;
                                    if (alf1 == Math.PI) alf1 = 0;
                                    if (alf1 == Math.PI / 2) alf1 = -1 * Math.PI / 2;
                                    if (alf1 == 3 * Math.PI / 2) alf1 = Math.PI / 2;
                                    Point3d PoslT = polar1(SpPoint.Last().TPoint, Vek1, Otstup);
                                    TPosrL TPoslT = new TPosrL();
                                    TPoslT.NTPoint(PoslT);
                                    TPoslT.NVisot(SpPoint.Last().Visot);
                                    SpPoint[SpPoint.Count - 1] = TPoslT;
                                    Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, Pom, ID_Les.ToString(), Zagal, BildDetFast);
                                    ID = ID + 1;
                                    Point3d PervT = polar1(PostT.TPoint, Vek2, Otstup);
                                    TPosrL TPervT = new TPosrL();
                                    TPervT.NTPoint(PervT);
                                    TPervT.NVisot(PostT.Visot);
                                    SpPoint.Clear();
                                    SpPoint.Add(TPervT);
                                }
                                else
                                {
                                    Angel90gr(ref SpPoint, TPoint, Hsirin);
                                }
                            }
                        }
                        SpPoint.Add(TPoint);
                    }
                }
                if (SpPoint.Count > 1) Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, Pom, ID_Les.ToString(), Zagal, BildDetFast);
                double n = 0;
                foreach (TPosrL TT in SpPointVH)
                {
                    if (TT.Visot != "")
                    {
                        KrugIText(TT, SpPoint, n.ToString(), ID_Les.ToString());
                        n = n + 1;
                    }
                }
            }
            //this.Show();
        }//не интерактивный способ построения лестниц
        public void Lest_INtr_SP(string Tipor, double Hsirin, string Zagal, string Povorot, string Xvost, string Soed, string Adr, double Otstup, double HSag_Hvost, bool BildDetFast) 
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
                CreateLayer("НасыщениеСкрытое");
                CreateLayer("Насыщение");
                List<string> SpID = new List<string>();
                List<TPosrL> SpPoint = new List<TPosrL>();
                List<TPosrL> SpPointGl = new List<TPosrL>();
                this.Hide();
                double ID_Les = 0;
                SBORpl_ID(ref SpID);
                while (SpID.Exists(x=>x== ID_Les.ToString())) ID_Les = ID_Les +1;
                double ID = 0;
                string TVis = "0";
                while (TVis != "")
                {
                    TPosrL TPoint = FTpostr(TVis, SpPoint, ID.ToString());
                    SpPointGl.Add(TPoint);
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
                            if (Math.Abs(alf - Math.PI / 2)<0.1 | Math.Abs(alf - 3 * Math.PI / 2)<0.1) 
                             {
                                if (SpPoint.Last().TPoint.DistanceTo(SpPoint[SpPoint.Count - 2].TPoint) > Otstup & SpPoint.Last().TPoint.DistanceTo(TPoint.TPoint) > Otstup)
                                {
                                    Vector3d Vek1 = SpPoint.Last().TPoint.GetVectorTo(SpPoint[SpPoint.Count - 2].TPoint);
                                    Vector3d Vek2 = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
                                    if ((Vek1.X < 0 & Vek2.Y < 0) | (Vek1.Y < 0 & Vek2.X < 0)) Ugol = 3 * (Math.PI) / 2;
                                    if (Vek1.X > 0 & Vek2.Y > 0 | (Vek1.Y > 0 & Vek2.X > 0)) Ugol = (Math.PI) / 2;
                                    if (Vek1.X < 0 & Vek2.Y > 0 | (Vek1.Y > 0 & Vek2.X < 0)) Ugol = Math.PI;
                                    FPovorot(SpPoint.Last().TPoint, "", "#ПоворотГориз90ЛестУсил" + (Hsirin - 8).ToString(), Ugol, "0", Hsirin, this.textBox1.Text, ID_Les.ToString());
                                    double alf1 = Math.Atan(tan);
                                    double alf2 = Math.Atan(tan1);
                                    if (alf1 == 0) alf1 = Math.PI;
                                    if (alf1 == Math.PI) alf1 = 0;
                                    if (alf1 == Math.PI / 2) alf1 = -1 * Math.PI / 2;
                                    if (alf1 == 3 * Math.PI / 2) alf1 = Math.PI / 2;
                                    Point3d PoslT = polar1(SpPoint.Last().TPoint, Vek1, Otstup);
                                    TPosrL TPoslT = new TPosrL();
                                    TPoslT.NTPoint(PoslT);
                                    TPoslT.NVisot(SpPoint.Last().Visot);
                                    SpPoint[SpPoint.Count - 1] = TPoslT;
                                    Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, this.textBox1.Text, ID_Les.ToString(), Zagal, BildDetFast);
                                    ID = ID + 1;
                                    Point3d PervT = polar1(PostT.TPoint, Vek2, Otstup);
                                    TPosrL TPervT = new TPosrL();
                                    TPervT.NTPoint(PervT);
                                    TPervT.NVisot(PostT.Visot);
                                    SpPoint.Clear();
                                    SpPoint.Add(TPervT);
                                }
                                else 
                                {
                                    Angel90gr(ref SpPoint, TPoint, Hsirin);
                                }
                            }
                        }
                        SpPoint.Add(TPoint);
                    }
                }
                if (SpPoint.Count > 1) Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, this.textBox1.Text, ID_Les.ToString(), Zagal, BildDetFast);
                double n = 0;
                foreach (TPosrL TT in SpPointGl)
                {
                    if (TT.Visot != "")
                    {
                        KrugIText(TT, SpPoint, n.ToString(), ID_Les.ToString());
                        n = n + 1;
                    }
                }
            }
            this.Show();
        }//интерактивный способ построения лестниц
        public static void Angel90gr(ref List<TPosrL> SpPoint, TPosrL TPoint, double Hsirin) 
        {
            Vector3d Vek1 = SpPoint.Last().TPoint.GetVectorTo(SpPoint[SpPoint.Count - 2].TPoint);
            Vector3d Vek2 = SpPoint.Last().TPoint.GetVectorTo(TPoint.TPoint);
            Point3d PoslT = polar1(SpPoint.Last().TPoint, Vek1, Hsirin/2);
            Point3d PervT = polar1(SpPoint.Last().TPoint, Vek2, Hsirin / 2);
            double DistAftTPoint = SpPoint.Last().TPoint.DistanceTo(SpPoint[SpPoint.Count - 2].TPoint);
            double DistBeftTPoint = SpPoint.Last().TPoint.DistanceTo(TPoint.TPoint);
            SpPoint.Remove(SpPoint.Last());
               if (DistAftTPoint > Hsirin * 0.8) 
                {
                TPosrL TPoslT = new TPosrL();
                TPoslT.NTPoint(PoslT);
                TPoslT.NVisot(SpPoint.Last().Visot);
                SpPoint.Add(TPoslT);
                }
            if (DistBeftTPoint > Hsirin * 0.8)
            {
                TPosrL TPoslT = new TPosrL();
                TPoslT.NTPoint(PervT);
                TPoslT.NVisot(SpPoint.Last().Visot);
                SpPoint.Add(TPoslT);
            }
        }
        public void TR_INtr_SP(string Tipor, double Hsirin, string Zagal, string Povorot, string Xvost, string Soed, string Adr, double NdiamIZgib, double HSag_Hvost, string DobavSoed, string DobavKr,double HSag_Kr, bool BildDetFast)
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
                CreateLayer("НасыщениеСкрытое");
                CreateLayer("Насыщение");
                List<TPosrL> SpPoint = new List<TPosrL>();
                List<TPosrL> SpPointGl = new List<TPosrL>();
                List<string> SpID = new List<string>();
                this.Hide();
                double ID_Les = 0;
                SBORpl_ID(ref SpID);
                while (SpID.Exists(x => x == ID_Les.ToString())) ID_Les = ID_Les + 1;
                double ID = 0;
                string TVis = "0";
                while (TVis != "")
                {
                    TPosrL TPoint = FTpostr(TVis, SpPoint, ID.ToString());
                    SpPointGl.Add(TPoint);
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
                            double alf11 = Vek1.GetAngleTo(new Vector3d(1,0,0));
                            double alf21 = Vek2.GetAngleTo(new Vector3d(1, 0, 0));
                            //Application.ShowAlertDialog(alf11 + " " + alf21);
                            if (Math.Abs(alf11 - alf21) > 0.005 ) 
                            { 
                            FPovorot_TR(SpPoint[SpPoint.Count - 2].TPoint, SpPoint.Last().TPoint, TPoint.TPoint, ID_Les.ToString(), "ПоворотТр" + Vek3.GetAngleTo(Vek2).ToString("0.###") + "-" + Tipor, "0", Hsirin, this.textBox1.Text,6,1,ref Delt, ID_Les.ToString(), Tipor);
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
                            Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, this.textBox1.Text, ID_Les.ToString(), Zagal, BildDetFast);
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
                if (SpPoint.Count > 1) Postr(SpPoint, Hsirin / 2, ID.ToString(), HSag_Hvost, Adr, Xvost, Tipor, this.textBox1.Text, ID_Les.ToString(), Zagal, BildDetFast);
                double Dlin = 0;
                double n = 0;
                double Ang = 0;
                double Ang1 = 0;
                double DlinStor1 = 0;
                string Visot = "";
                Point3dCollection TkoorPL = new Point3dCollection();
                Point3d Point1 = SpPointGl[0].TPoint;
                foreach (TPosrL TT in SpPointGl)
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
                if (this.checkBox1.Checked == true)
                {
                    for (double RastTT = 40; RastTT < Dlin; RastTT = RastTT + HSag_Hvost)
                    {
                        Point_Ang(ref Point1, ref Ang, TkoorPL, RastTT, SpPoint, ref Visot, ref Ang1);
                        FXvost(Point1, "", ID_Les.ToString(), Ang, Visot, this.textBox1.Text, Xvost, DobavKr);
                    }
                    Point_Ang(ref Point1, ref Ang, TkoorPL, Dlin - 40, SpPoint, ref Visot, ref Ang1);
                    FXvost(Point1, "", ID_Les.ToString(), Ang, Visot, this.textBox1.Text, Xvost, DobavKr);
                }
                if (this.checkBox2.Checked == true)
                for (double RastTT = HSag_Kr; RastTT < Dlin; RastTT = RastTT + HSag_Kr)
                    {
                    Point_Ang(ref Point1, ref Ang, TkoorPL, RastTT, SpPoint, ref Visot, ref Ang1);
                    FSoedin(Point1, "", Soed, Ang, Visot, 0, this.textBox1.Text, ID_Les.ToString(), Zagal, DobavSoed);
                    }
            }
            this.Show();
        }//интерактивный способ построения труб
        public static void Postr(List<TPosrL> SpPoint, double Dles, string ID, double HSag_Hvost, string Adr, string Xvost, string COMPON, string POM, string InD, string Razdel,bool BildDetFast)
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
                IDd = IDd + 1;
                tan = (SpPoint[i + 1].TPoint.Y - SpPoint[i].TPoint.Y) / (SpPoint[i + 1].TPoint.X - SpPoint[i].TPoint.X);
                alf = Math.Atan(tan);
                alf1 = Math.Atan(tan) + Math.PI / 2;
                alf2 = Math.Atan(tan) - Math.PI / 2;
                T1 = polar(SpPoint[i].TPoint, alf1, Dles);
                T2 = polar(SpPoint[i].TPoint, alf2, Dles);
                T3 = polar(SpPoint[i + 1].TPoint, alf1, Dles);
                T4 = polar(SpPoint[i + 1].TPoint, alf2, Dles);
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
            Point3d Point1 = new Point3d();
            Point3d Point2 = new Point3d();
            if (Razdel == "Лес" & BildDetFast)
            {
                for (double RastTT = 40; RastTT < Dlin; RastTT = RastTT + HSag_Hvost)
                {
                    Point_Ang(ref Point1, ref Ang, TkoorPL11, RastTT, SpPoint, ref Visot, ref Ang1);
                    Point_Ang(ref Point2, ref Ang, TkoorPL22, RastTT, SpPoint, ref Visot, ref Ang1);
                    FXvost(Point1, "", InD + "-" + ID, Ang, Visot, POM, Xvost,"");
                    FXvost(Point2, "", InD + "-" + ID, Ang1, Visot, POM, Xvost,"");
                }
                Point_Ang(ref Point1, ref Ang, TkoorPL11, DlinStor1 - 40, SpPoint, ref Visot, ref Ang1);
                Point_Ang(ref Point2, ref Ang, TkoorPL22, DlinStor1 - 40, SpPoint, ref Visot, ref Ang1);
                FXvost(Point1, "", InD + "-" + ID, Ang, Visot, POM, Xvost,"");
                FXvost(Point2, "", InD + "-" + ID, Ang1, Visot, POM, Xvost,"");
            }
            Zagib(TkoorPL11, TkoorPL22, SpPoint, InD + "-" + ID);
            for (int i = TkoorPL22.Count - 1; i >= 0; i--) TkoorPL11.Add(TkoorPL22[i]);
            if (Razdel != "Тр")
            FPoly(TkoorPL11, SpPoint, COMPON, POM, InD, Razdel, InD + "-" + ID);
            else
            FPoly(TkoorPL11, SpPoint, COMPON, POM, InD, Razdel, InD);
        }//определение координат для построения лесницы полилинией
        public static void Zagib(Point3dCollection TkoorTL1, Point3dCollection TkoorTL2, List<TPosrL> SpPoint,string LINK) 
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
        }
        public static void Point_Ang(ref Point3d Poin, ref double Ang, Point3dCollection TkoorPL1,double Rast, List<TPosrL> SpPoint,ref string Visot, ref double Ang2)
        {
            double TDistLest = 0;
            double TDist = 0;
            double TDistOtr = 0;
            Vector3d VekT=new Vector3d();
            double TAngX = 0;
            double TAngY =0;
            for (int i = 0; i < TkoorPL1.Count - 1; i++) 
            {
                TDistOtr = TkoorPL1[i].DistanceTo(TkoorPL1[i+1]);
                VekT = TkoorPL1[i].GetVectorTo(TkoorPL1[i+1]);
                TAngX = VekT.GetAngleTo(new Vector3d(1, 0, 0));
                TAngY = VekT.GetAngleTo(new Vector3d(0, 1, 0));
                    if ( Rast > TDist & Rast < TDist + TDistOtr)
                    {
                    for (int i1 = 0; i1 < SpPoint.Count - 1; i1++)
                    {
                        if (Rast > TDistLest & Rast < TDistLest + SpPoint[i1].TPoint.DistanceTo(SpPoint[i1 + 1].TPoint))
                            Visot = SpPoint[i1].Visot;
                        TDistLest = TDistLest + SpPoint[i1].TPoint.DistanceTo(SpPoint[i1 + 1].TPoint);
                    }
                    Poin = polar1(TkoorPL1[i], VekT, Rast- TDist);
                    Ang = TAngX;
                    if (VekT.X >= 0 & VekT.Y >= 0) { Ang = TAngX; Ang2 = Ang + Math.PI / 2; };
                    if (VekT.X >= 0 & VekT.Y < 0) { Ang = VekT.GetAngleTo(new Vector3d(0, -1, 0)) + Math.PI; Ang2 = Ang - Math.PI / 2;};
                    if (VekT.X < 0 & VekT.Y <= 0) { Ang = VekT.GetAngleTo(new Vector3d(-1, 0, 0)); Ang2 = Ang + Math.PI / 2; };
                    if (VekT.X < 0 & VekT.Y > 0) { Ang = TAngX + Math.PI / 2; Ang2 = Ang - Math.PI / 2; }
                    return;
                    }
                TDist = TDist + TDistOtr;
            }
        }//угол в точке кривой 
        public static Point3d polar(Point3d XYZ1, double ugolX, double Rast)
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
                if (T11.X == T12.X) X = T11.X;
                if (T11.Y == T12.Y) Y = T11.Y;
                if (T21.X == T22.X) X = T21.X;
                if (T21.Y == T22.Y) Y = T21.Y;
                Tperes = new Point3d(X, Y, 0);
            }
            return Tperes;
        }//Определение точки пересечения
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
            Creat_BL(TPostr.TPoint, ID, TPostr.Visot, Name, "CWPoint", "CWAYPoint",0,"","");
        }
        public static void FXvost(Point3d Nkoor, string ID, string Name,double Ugol, string Visot, string Pom, string Xvost, string DopDet)
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
                    poly.AddVertexAt(0, new Point2d(Nkoor.X, Nkoor.Y+ Convert.ToDouble(ABC[0])), 0, 0, 0);
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
            Creat_BL(Nkoor, ID, Visot, Name, blkName, "Насыщение", Ugol,Pom, DopDet);
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
                        poly1.AddVertexAt(0, new Point2d(Nkoor.X - Convert.ToDouble(ABC[0]) / 2, Nkoor.Y ), 0, 0, 0);
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
                    Arc Ark1 = new Arc(new Point3d(Nkoor.X + 250 + Hscir / 2, Nkoor.Y - 250- Hscir / 2, 0),250, Math.PI / 2, Math.PI);
                    Ark1.Layer = "0";
                    acBlkTblRec.AppendEntity(Ark1);
                    Tx.AddNewlyCreatedDBObject(Ark1, true);

                    Arc Ark2 = new Arc(new Point3d(Nkoor.X + 250 + Hscir / 2, Nkoor.Y - 250-Hscir / 2, 0), 250+ Hscir, Math.PI / 2, Math.PI);
                    Ark2.Layer = "0";
                    acBlkTblRec.AppendEntity(Ark2);
                    Tx.AddNewlyCreatedDBObject(Ark2, true);

                    Polyline poly1 = new Polyline();
                    poly1.SetDatabaseDefaults();  
                    poly1.Layer = "0";
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X + Hscir / 2, Nkoor.Y - 250- Hscir / 2), 0, 0, 0);
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X + Hscir / 2, Nkoor.Y - 400- Hscir / 2), 0, 0, 0);
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X - Hscir / 2, Nkoor.Y - 400- Hscir / 2), 0, 0, 0);
                    poly1.AddVertexAt(0, new Point2d(Nkoor.X - Hscir / 2, Nkoor.Y - 250- Hscir / 2), 0, 0, 0);
                    acBlkTblRec.AppendEntity(poly1);
                    Tx.AddNewlyCreatedDBObject(poly1, true);

                    Polyline poly2 = new Polyline();
                    poly2.SetDatabaseDefaults();
                    poly2.Layer = "0";
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 250+ Hscir / 2, Nkoor.Y + Hscir / 2), 0, 0, 0);
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 400+ Hscir / 2, Nkoor.Y + Hscir / 2), 0, 0, 0);
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 400+ Hscir / 2, Nkoor.Y - Hscir / 2), 0, 0, 0);
                    poly2.AddVertexAt(0, new Point2d(Nkoor.X + 250+ Hscir / 2, Nkoor.Y - Hscir / 2), 0, 0, 0);
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
            Creat_BL_SpOID(Nkoor, ID, Visot, LINK, blkName, "Насыщение", Ugol, SpPrim,Pom,"Лес", blkName,"");
        }//отрисовка поворотов 90 граусов у лестниц
        public void FPovorot_TR(Point3d T1, Point3d T2, Point3d T3, string ID, string Name, string Visot, double Hscir, string Pom, double NdiamIZgib, double NdiamPrUHC,ref double Delt, string LINK, string Compon)
        {
            Vector3d Vek1 = T1.GetVectorTo(T2);
            Vector3d Vek2 = T2.GetVectorTo(T3);
            double Povorot = 0;
            double alfMV =Math.PI/2 - Vek1.GetAngleTo(Vek2);
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
            if (Math.Abs(T11.DistanceTo(T31) - (T11.DistanceTo(Tper1) + Tper1.DistanceTo(T31))) < 0.1) { Tper = Tper1;  }
            if (Math.Abs(T21.DistanceTo(T41) - (T21.DistanceTo(Tper2) + Tper2.DistanceTo(T41))) < 0.1) { Tper = Tper2;  }
            if (Vek1.GetAngleTo(Vek2) == Math.PI / 2) Tper= FTper90(T1,T2,T3, DeltRad);
            List<ObjectId> SpPrim = new List<ObjectId>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            EdI.Editor ed = doc.Editor;
            EdI.PromptPointOptions pPtOpts;
            double cosALF = DeltRad / T2.DistanceTo(Tper) ;
            Delt = T2.DistanceTo(Tper) * Math.Sqrt(1 - Math.Pow(cosALF,2));
            Povorot = OpredUgl(Vek1, Vek2);
            alfMV = 2* Math.Acos(cosALF);
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
            Creat_BL_SpOID(Tper, ID, (Delt*2).ToString("0.##"), LINK, blkName, "Насыщение", Povorot, SpPrim, Pom,"Тр", Compon,"");
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
            if (Vek1.X <= 0 & Vek1.Y >= 0 & Vek2.X >= 0 & Vek2.Y >= 0) Ugol = Math.PI/2 + V2X;//11-+++
            if (Vek1.X >= 0 & Vek1.Y >= 0 & Vek2.X <= 0 & Vek2.Y >= 0) Ugol = 0 + V1Y;//12++-+
            if (Vek1.X >= 0 & Vek1.Y <= 0 & Vek2.X <= 0 & Vek2.Y <= 0) Ugol = 0- (Math.PI - V2Y);//13+---
            //14-++-
            //15+--+
            //16+-++
            //17+--+
            if (Vek1.X <= 0 & Vek1.Y <= 0 & Vek2.X >= 0 & Vek2.Y <= 0) Ugol = V1Y;//18--+-
            return Ugol; 
        }
        public static void FPoly(Point3dCollection TkoorPL, List<TPosrL> SpPoint,string COMPON,string POM,string ID, string Razdel, string LINK)
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
                            //acAttDef.Layer = "НасыщениеСкрытое";
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
                            //acAttDef.Layer = "НасыщениеСкрытое";
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
                            //acAttDef.Layer = "НасыщениеСкрытое";
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
                        Pos = new Point3d(Pos.X, Pos.Y - 3, 0);
                        using (AttributeDefinition acAttDef = new AttributeDefinition())
                        {
                            acAttDef.Position = Pos;
                            acAttDef.Verifiable = true;
                            acAttDef.Invisible = true;
                            acAttDef.Prompt = "ДопДетали #: ";
                            acAttDef.Tag = "ДопДетали";
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
        static public void OBNNpoz(ref List<POZIZIA> SpPOZ, ref List<POZIZIA> SpPOZob)
        {
#region присвоение начальных значений переменных и создание фильтра для выборки
            string stKomp = "";
            string stTIP = "";
            string stLINK = "";
            string stDlina = "";
            string stVisot = "";
            string stVirez = "";
            string stKEI = "";
            string stHtoEto = "";
            string pom = "";
            string pomUDAL = "";
            string stKompUDAL = "";
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Editor ed = doc.Editor;
                TypedValue[] acTypValAr = new TypedValue[5];
                acTypValAr.SetValue(new TypedValue(0, "TEXT"), 0);
                acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
                acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 2);
                acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытые"), 3);
                acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
                // создаем фильтр
                SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
                PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
                if (acSSPrompt.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Нет деталей...");
                    return;
                }
                SelectionSet acSSet = acSSPrompt.Value;
                ObjectId[] ids = acSSet.GetObjectIds();
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    #endregion
                    #region чтение тасширеных данных текстового примитива
                    foreach (ObjectId ID in ids)
                    {
                        stKomp = "";
                        stTIP = "";
                        stLINK = "";
                        pomUDAL = "";
                        stKompUDAL = "";
                        //DBText bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as DBText;
                        DBText bref = (DBText)Tx.GetObject(ID, OpenMode.ForWrite);
                        bref.UpgradeOpen();
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 3) { stTIP = value.Value.ToString(); }
                                if (Schet == 4) { stLINK = value.Value.ToString(); pomUDAL = value.Value.ToString(); }
                                if (Schet == 5) { stKompUDAL = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                            #endregion
                            #region если в расширеных данных текстового примитива записан компонент и номер помещения
                            string[] stKompM = stKomp.Split('_');
                            if (stKompM.Length > 1)
                            {
                                string[] stKompM1 = stKompM[0].Split('#');
                                if (stKompM1.Length > 1)
                                    if (stKompM.Length > 1) stKomp = "#" + stKompM1[1].Replace('@', '.') + "_" + stKompM[1];
                                    else
                                        if (stKompM.Length > 1) stKomp = stKomp.Replace('@', '.') + "_" + stKompM[1];
                                //Application.ShowAlertDialog(stKomp);
                                //stKomp = stKompM1.Last().Replace('@', '.');                  
                                //if (SpPOZ.Exists(x => (x.Compon == stKomp || "#" + x.Compon == stKomp) & x.Pom == stKompM[1]) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                                //if (SpPOZ.Exists(x => x.Compon + "_" + x.Pom == stKomp | "#" + x.Compon + "_" + x.Pom == stKomp) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                                //if (SpPOZ.Exists(x => stKomp.Contains(x.Compon + "_" + x.Pom)) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                                if (SpPOZ.Exists(x => stKomp == ("#" + x.Compon + "_" + x.Pom)) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                                {
                                    //Application.ShowAlertDialog("#" + SpPOZ.Find(x => (x.Compon + "_" + x.Pom == stKomp | "#" + x.Compon + "_" + x.Pom == stKomp)).Compon + "_" + SpPOZ.Find(x => (x.Compon + "_" + x.Pom == stKomp | "#" + x.Compon + "_" + x.Pom == stKomp)).Pom);
                                    bref.TextString = SpPOZ.Find(x => stKomp == ("#" + x.Compon + "_" + x.Pom)).NOMpoz;
                                    //bref.TextString = SpPOZ.Find(x => (x.Compon == stKomp || "#" + x.Compon == stKomp) & x.Pom == stKompM[1]).NOMpoz;
                                }
                            }
                            #endregion
                            #region если в расширеных данных текстового примитива записана ссылка на блок и этот текстовый примитив является позицией
                            else
                            {

                                if (stTIP == "POZIZIA")
                                {
                                    string Handl = stKomp;
                                    BLOKPoHandl(Handl, ref stKomp, ref pom, ref stDlina, ref stVisot, ref stVirez, ref stKEI, ref stHtoEto, pomUDAL, stKompUDAL);
                                    if (SpPOZ.Exists(x => x.Compon == stKomp & x.Pom == pom) == true) { bref.TextString = SpPOZ.Find(x => x.Compon == stKomp & x.Pom == pom).NOMpoz; }
                                    if (SpPOZob.Exists(x => x.Ind == stKomp) == true) { bref.TextString = SpPOZob.Find(x => x.Ind == stKomp).NOMpoz; }
                                }
                                #endregion
                                #region если в расширеных данных текстового примитива записана ссылка на блок и этот текстовый примитив является текстом над полкой
                                if (stTIP == "T_NAD_POL" && stKomp != "")
                                {
                                    string Handl = stKomp;
                                    //if (stLINK != "") { Handl = stLINK; }
                                    //Application.ShowAlertDialog(stTIP + " " + Handl);
                                    BLOKPoHandl(Handl, ref stKomp, ref pom, ref stDlina, ref stVisot, ref stVirez, ref stKEI, ref stHtoEto, pomUDAL, stKompUDAL);
                                    if (SpPOZ.Exists(x => x.Compon == stKomp & x.Pom == pom) == true)
                                    {
                                        string strTextNP = "";
                                        string strTextPP = "";
                                        string NOMpoz = "без поз";
                                        double dDlin = 0;
                                        double dVisot = 0;
                                        if ((stDlina == "") == false) { dDlin = Convert.ToDouble(stDlina); }
                                        if ((stVisot == "") == false) { dVisot = Convert.ToDouble(stVisot); }
                                        TextPP_NP_POZ(ref SpPOZ, stKomp, pom, stKEI, stVirez, ref strTextNP, ref strTextPP, ref NOMpoz, dDlin, dVisot, stHtoEto, 1, ref SpPOZob);
                                        bref.TextString = strTextNP;
                                    }
                                }
                                #endregion
                                #region если в расширеных данных текстового примитива записана ссылка на блок и этот текстовый примитив является текстом под полкой
                                if (stTIP == "T_POD_POL" && stKomp != "")
                                {
                                    string Handl = stKomp;
                                    //if (stLINK != "") { Handl = stLINK; }
                                    //Application.ShowAlertDialog(stTIP + " " + Handl);
                                    BLOKPoHandl(Handl, ref stKomp, ref pom, ref stDlina, ref stVisot, ref stVirez, ref stKEI, ref stHtoEto, pomUDAL, stKompUDAL);
                                    if (SpPOZ.Exists(x => x.Compon == stKomp & x.Pom == pom) == true)
                                    {
                                        string strTextNP = "";
                                        string strTextPP = "";
                                        string NOMpoz = "без поз";
                                        double dDlin = 0;
                                        double dVisot = 0;
                                        if ((stDlina == "") == false) { dDlin = Convert.ToDouble(stDlina); }
                                        if ((stVisot == "") == false) { dVisot = Convert.ToDouble(stVisot); }
                                        TextPP_NP_POZ(ref SpPOZ, stKomp, pom, stKEI, stVirez, ref strTextNP, ref strTextPP, ref NOMpoz, dDlin, dVisot, stHtoEto, 1, ref SpPOZob);
                                        bref.TextString = strTextPP;
                                    }
                                }

                            }
                        }
                    }
                    Tx.Commit();
                }
                #endregion
            } 
        }//обновить номера позиций
        static void BLOKPoHandl(string Handl, ref string Compon, ref string pom, ref string stDlina, ref string stVisot, ref string stVirez, ref string stKEI, ref string stHtoEto, string pomUDAL, string stKompUDAL)
        {
            string stKomp = "";
            string stRazd = "";
            string stDobav = "";
            double dDlin = 0;
            double dVisot = 0;
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            int sist10 = Convert.ToInt32(Handl, 16);
            Handle H = new Handle(sist10);
            using (tr)
            {
                try
                {
                    ObjectId id = db.GetObjectId(false, H, -1);
                    //catch {return;}
                    //Application.ShowAlertDialog(id.ToString());
                    //if (id.IsErased == true) { Compon = "";  }
                    //if (id == false) { Compon = ""; return; }
                    Entity bref1 = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                    string[] TYPEm = bref1.GetType().ToString().Split('.');
                    string TYPE = TYPEm.Last();
                    if (TYPE == "BlockReference")//если позиция  блок
                    {
                        BlockReference bref = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;
                        if (bref.IsDynamicBlock)
                        {
                            DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); }
                                if (prop.PropertyName == "Расстояние1")
                                {
                                    dDlin = (Math.Ceiling(Convert.ToDouble(prop.Value.ToString()) / 100) / 10) * 1000; ;
                                    //dDlin = Convert.ToDouble(prop.Value.ToString()); 
                                    //dDlin = Math.Round(dDlin);
                                    stDlina = Convert.ToString(dDlin);
                                }
                                if (prop.PropertyName == "Расстояние2")
                                {
                                    dDlin = (Math.Ceiling(Convert.ToDouble(prop.Value.ToString()) / 100) / 10) * 1000; ;
                                    //dDlin = Convert.ToDouble(prop.Value.ToString()); 
                                    //dDlin = Math.Round(dDlin);
                                    stDlina = Convert.ToString(dDlin);
                                }
                            }
                            foreach (ObjectId idAtrRef in bref.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                        if (atrRef.Tag == "Помещение") { pom = atrRef.TextString; }
                                        if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                        if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); stVisot = Math.Round(dVisot, 0).ToString(); }
                                    }
                                }
                            }
                            ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 3) { stDobav = value.Value.ToString(); }
                                    if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                    Schet = Schet + 1;
                                }
                            }
                        }
                        else
                        {
                            ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { stKomp = value.Value.ToString(); }
                                    if (Schet == 2) { stVisot = value.Value.ToString(); }
                                    if (Schet == 3) { stDobav = value.Value.ToString(); }
                                    if (Schet == 4) { pom = value.Value.ToString(); }
                                    if (Schet == 5) { stRazd = value.Value.ToString(); }
                                    if (Schet == 6) { stKEI = value.Value.ToString(); }
                                    if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                    if (Schet == 9) { stDlina = value.Value.ToString(); }
                                    Schet = Schet + 1;
                                }
                            }
                            foreach (ObjectId idAtrRef in bref.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                        if (atrRef.Tag == "Помещение") { pom = atrRef.TextString; }
                                        if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                        if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); stVisot = Math.Round(dVisot, 0).ToString(); }
                                    }
                                }
                            }
                        }
                    }
                    else//если позиция не блок
                    {
                        Entity bref = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { pom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()); dDlin = Math.Round(dDlin / 10, 0); stDlina = Convert.ToString(dDlin * 10); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    string[] stKompM = stKomp.Split('#');
                    Compon = stKompM.Last().Replace('@', '.');
                    stVirez = stKompM.First();
                }
                catch { Compon = stKompUDAL; pom = pomUDAL; return; }
            }
        }//нод по метке
        static public void TextPP_NP_POZ(ref List<POZIZIA> SpPOZ, string stKomp, string stPom, string stKEI, string stVirez, ref string strTextNP, ref string strTextPP, ref string NOMpoz, double dDlin, double dVisot, string stHtoEto, double kol, ref List<POZIZIA> SpPOZob)
        {
            string[] VirezM = stVirez.Split('%');
            if (VirezM.Length > 1) { strTextNP = VirezM[1]; }
            if (dVisot == 0 & stKEI == "006") { strTextNP = "L=" + dDlin.ToString(); }
            if (dVisot > 0 & stKEI == "006") { strTextNP = "L=" + Convert.ToString(dDlin + 2 * dVisot); strTextPP = "2фл." + dVisot.ToString(); }
            if (dVisot > 0 & stKEI == "006" & stHtoEto == "хв") { strTextNP = "L=" + Convert.ToString(dVisot); if (kol > 1) { strTextPP = kol.ToString() + " шт."; } }
            if (dVisot > 0 & stKEI == "796") { strTextPP = "H=" + dVisot.ToString(); }
            if (SpPOZ.Exists(x => x.Compon == stKomp & x.Pom == stPom) == true) { NOMpoz = SpPOZ.Find(x => x.Compon == stKomp & x.Pom == stPom).NOMpoz; }
            if (SpPOZob.Exists(x => x.Ind == stKomp) == true) { NOMpoz = SpPOZob.Find(x => x.Ind == stKomp).NOMpoz; strTextNP = stKomp; strTextPP = SpPOZob.Find(x => x.Ind == stKomp).shabl; }
            if (strTextNP == "" & (strTextPP == "") == false) { strTextNP = strTextPP; strTextPP = ""; }
        }//текст позиций текст под полкой и над полкой
        public static void Creat_BL(Point3d BazeP, string ID, string Visot, string Name, string blkName, string Sloi,double Ugol, string Pom, string strDobav)
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
                    if(blkName.Contains("Уголок") == true)
                    SetDynamicBlkProperty(blkName, Visot, strDobav, Pom, "Доизол.д", "006", "хв", "", 0, "", Name, Ugol, Sloi);
                    else
                    SetDynamicBlkProperty(ID, Visot, "", "", "", "", "", "", 0, "", Name, Ugol, Sloi);
                    tr.Commit();
                }  
            }
            //this.Show();
        }//создание блока из последнено созданного примитива
        public static void Creat_BL_SpOID(Point3d BazeP, string ID, string Visot, string Name, string blkName, string Sloi, double Ugol,List<ObjectId> SpOID, string Pom, string Razd, string Compon, string Dobav)
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
                        else if(blkName.Contains("ПоворотГор"))
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
                    //Application.ShowAlertDialog(blkName);
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
        private void dataGridView7_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (this.dataGridView7.CurrentRow.Cells[0].Value == null) return;
            String IND = this.dataGridView7.CurrentRow.Cells[0].Value.ToString();
            List<Part> SpPartTCWAY = SpPart.FindAll(x => x.CWName == IND);
            List<Part> SpPartTCWAYSort = new List<Part>();
            this.dataGridView8.Rows.Clear();
            string StartPoint = "";
            string EndPoint = "";
            string StartPointTr = "";
            string EndPointTr = "";
            double CountStartPoint = 0;
            double CountEndPoint = 0;
            foreach (Part TPart in SpPartTCWAY)
            {
                StartPoint = TPart.StartPoint;
                EndPoint = TPart.EndPoint;
                List<Part> SpPartStartPoint = SpPartTCWAY.FindAll(x => x.StartPoint == StartPoint | x.EndPoint == StartPoint);
                List<Part> SpPartEndPoint = SpPartTCWAY.FindAll(x => x.StartPoint == EndPoint | x.EndPoint == EndPoint);
                if (SpPartStartPoint.Count == 1) { StartPointTr = StartPoint; CountStartPoint += 1; }
                if (SpPartEndPoint.Count == 1) { EndPointTr = EndPoint; CountEndPoint += 1; }
            }
            string TPointTr = StartPointTr;
            string Prodol = "Y";
            if(CountStartPoint==1 & CountEndPoint==1)
            while (TPointTr != EndPointTr | Prodol != "Y" )
            {
                if (SpPartTCWAY.Exists(x => x.StartPoint == TPointTr))
                {
                    Part TPart = SpPartTCWAY.Find(x => x.StartPoint == TPointTr);
                    TPointTr = TPart.EndPoint;
                    SpPartTCWAYSort.Add(TPart);
                }
                else
                    Prodol = "N";
            }
            foreach (Part TPart in SpPartTCWAYSort) this.dataGridView8.Rows.Add(TPart.CompName, TPart.StartPoint, TPart.EndPoint, TPart.Coordinates);
            this.label38.Text = "CountStartPoint=" + CountStartPoint + "|CountEndPoint=" + CountEndPoint;
        }//выбор трассы

        public void Tabl(ref List<POZIZIA> SpPOZ_PE)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("Точка начала построения");
                Database db = doc.Database;
                Editor ed = doc.Editor;
                BlockTableRecord acBlkTblRec;
                BlockTable acBlkTbl;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d XYZ = pPtRes.Value;
                Point3d T1 = new Point3d();
                Point3d T2 = new Point3d();
                Point3d T3 = new Point3d();
                Point3d T4 = new Point3d();
                Point3d T5 = new Point3d();
                Point3dCollection TkoorPL = new Point3dCollection();
                string label = "";
                string description = "";
                string kol = "";
                string note = "";

                string labelT = "";
                string descriptionT = "";
                string kolT = "";
                string noteT = "";
                string Compon = "";

                double Apos = 20;
                double Anaim = 110;
                double Akol = 10;
                double Aprim = 45;

                double B = 8;
                double aPOS = 2;
                double aNAIM = 2;
                double aKOL = 2;
                double aPRIM = 2;

                double b = 5;


                //List<POZIZIA> SpPOZ_PE = new List<POZIZIA>();
                //SBOR_OBOR_PE(ref SpPOZ_PE);

                Strok(ref T1, ref T2, ref T3, ref T4, ref T5, XYZ, Apos, Anaim, Akol, Aprim, 15, aPOS, aNAIM, aKOL, aPRIM, b, "Поз.", "Наименование", "Кол", "КЕИ","Обозначение");
                XYZ = new Point3d(XYZ.X, XYZ.Y - 15, XYZ.Z);

                foreach (POZIZIA Pos in SpPOZ_PE)
                {
                    label = Pos.NOMpoz;
                    description = Pos.polnNAIM.Replace("\n", "") ;
                    if (SpNamenkl.Exists(x => x.Compon == Pos.Compon)) description = SpNamenkl.Find(x => x.Compon == Pos.Compon).Hozain;
                    //description = Pos.description.Replace("&#xA", " ") + " " + Pos.nominal_description + " " + Pos.itt;
                    kol = Pos.Kol.ToString();
                    note = Pos.KEI;
                    Compon = Pos.Compon;
                    using (Transaction Tx = db.TransactionManager.StartTransaction())
                    {
                        acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        string label_razd = razbitStrok(4, label, ",");
                        string description_razd = razbitStrok(44, description, " ");
                        string note_razd = razbitStrok(15, note, " ");

                        string[] labelM = label_razd.Split('&');
                        string[] descriptionM = description_razd.Split('&');
                        string[] noteM = note_razd.Split('&');

                        double KolStr = Math.Max(labelM.Length, Math.Max(descriptionM.Length, noteM.Length));

                        for (int i = 0; i < KolStr; i++)
                        {
                            if (i <= labelM.Length - 1) labelT = labelM[i]; else labelT = "";
                            if (i <= descriptionM.Length - 1) descriptionT = descriptionM[i]; else descriptionT = "";
                            if (i <= noteM.Length - 1) noteT = noteM[i]; else noteT = "";
                            if (i != 0) { kol = ""; Compon = ""; }
                            Strok(ref T1, ref T2, ref T3, ref T4, ref T5, XYZ, Apos, Anaim, Akol, Aprim, B, aPOS, aNAIM, aKOL, aPRIM, b, labelT, descriptionT, kol, noteT, Compon);
                            XYZ = new Point3d(XYZ.X, XYZ.Y - B, XYZ.Z);
                        }
                        Tx.Commit();
                    }
                }
            }
        }
        public void Strok_tabl_L(string Baza, string NomLes, string POZ, string Xvost, string Plan, List<string> SpT, int max, Point3d TochT, double B)
        {
            //базавая точка
            double Abp = 430;
            double abp = 40;
            double bbp = 80;
            //номер лесницы
            double Anom = 430;
            double anom = 40;
            double bnom = 80;
            //позиция лестницы
            double Apos = 430;
            double apos = 40;
            double bpos = 80;
            //позиция хвостовика
            double Axvost = 430;
            double axvost = 40;
            double bxvost = 80;
            //позиция планки
            double Aplan = 430;
            double aplan = 40;
            double bplan = 80;
            //координаты
            double Akoor = 500;
            double akoor = 40;
            double bkoor = 80;
            string koor = "";
            LWPOIYiTEXT_hir(TochT, 0, NomLes, "", "", Anom, B, anom, bnom, 50);
            TochT = new Point3d(TochT.X + Anom, TochT.Y, 0);
            LWPOIYiTEXT_hir(TochT, 0, Baza, "", "", Abp, B, abp, bbp, 50);
            TochT = new Point3d(TochT.X + Abp, TochT.Y, 0);
            LWPOIYiTEXT_hir(TochT, 0, POZ, "", "", Apos, B, apos, bpos, 50);
            TochT = new Point3d(TochT.X + Apos, TochT.Y, 0);
            LWPOIYiTEXT_hir(TochT, 0, Xvost, "", "", Axvost, B, axvost, bxvost, 50);
            TochT = new Point3d(TochT.X + Axvost, TochT.Y, 0);
            LWPOIYiTEXT_hir(TochT, 0, Plan, "", "", Aplan, B, aplan, bplan, 50);
            TochT = new Point3d(TochT.X + Aplan, TochT.Y, 0);
            for (int i = 0; i < max; i++)
            {
                if (i < SpT.Count)
                    LWPOIYiTEXT_hir(TochT, 0, SpT[i].Split(' ')[0], SpT[i].Split(' ')[1], SpT[i].Split(' ')[2], Akoor, B, akoor, bkoor, 50);
                else
                    LWPOIYiTEXT_hir(TochT, 0, "", "", "", Akoor, B, akoor, bkoor, 50);
                TochT = new Point3d(TochT.X + Akoor, TochT.Y, 0);
            }
        }
        public void LWPOIYiTEXT_hir(Point3d Toch1, int Zvet, string IND1, string IND2, string IND3, double A, double B, double a, double b, double Htext)
        {
            Point3d Ttext1 = new Point3d(Toch1.X + a, Toch1.Y - b, 0);
            Point3d Ttext2 = new Point3d(Toch1.X + a, Toch1.Y - b * 2, 0);
            Point3d Ttext3 = new Point3d(Toch1.X + a, Toch1.Y - b * 3, 0);
            Point3d T1 = Toch1;
            Point3d T2 = new Point3d(Toch1.X + A, Toch1.Y, 0);
            Point3d T3 = new Point3d(Toch1.X + A, Toch1.Y - B, 0);
            Point3d T4 = new Point3d(Toch1.X, Toch1.Y - B, 0);
            Point3dCollection TkoorPL = new Point3dCollection();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = 0;
                poly.Closed = true;
                poly.Layer = "Выноски";
                int i = 0;
                foreach (Point3d pt in TkoorPL)
                {
                    poly.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    i = i + 1;
                }
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);

                if (IND1 != "" & IND1 != null)
                {
                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = Ttext1;
                    Poz.Height = Htext;
                    Poz.ColorIndex = Zvet;
                    Poz.TextString = IND1;
                    Poz.Layer = "Выноски";
                    btr.AppendEntity(Poz);
                    tr1.AddNewlyCreatedDBObject(Poz, true);
                }
                if (IND2 != "" & IND2 != null)
                {
                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = Ttext2;
                    Poz.Height = Htext;
                    Poz.ColorIndex = Zvet;
                    Poz.TextString = IND2;
                    Poz.Layer = "Выноски";
                    btr.AppendEntity(Poz);
                    tr1.AddNewlyCreatedDBObject(Poz, true);
                }
                if (IND3 != "" & IND3 != null)
                {
                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = Ttext3;
                    Poz.Height = Htext;
                    Poz.ColorIndex = Zvet;
                    Poz.TextString = IND3;
                    Poz.Layer = "Выноски";
                    btr.AppendEntity(Poz);
                    tr1.AddNewlyCreatedDBObject(Poz, true);
                }
                tr1.Commit();
            }
        }//Отрисовка прямоугольника и текста
        public void Strok(ref Point3d T1, ref Point3d T2, ref Point3d T3, ref Point3d T4, ref Point3d T5, Point3d XYZ, double Apos, double Anaim, double Akol, double Aprim, double B, double aPOS, double aNAIM, double aKOL, double aPRIM, double b, string pos, string Naim, string kol, string prim, string Compon)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("Точка начала построения");
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockTableRecord acBlkTblRec;
            BlockTable acBlkTbl;
            Point3dCollection TkoorPL = new Point3dCollection();
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, XYZ, Apos, B, aPOS, b);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                FPoly2d(TkoorPL);
                if (pos != "")
                {

                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = T5;
                    Poz.TextString = pos;
                    Poz.WidthFactor = 0.8;
                    Poz.Height = 3;
                    acBlkTblRec.AppendEntity(Poz);
                    Tx.AddNewlyCreatedDBObject(Poz, true);
                }

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, new Point3d(XYZ.X + Apos, XYZ.Y, XYZ.Z), Aprim, B, aPRIM, b);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                FPoly2d(TkoorPL);
                if (Compon != "")
                {

                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    Poz.Position = T5;
                    Poz.TextString = Compon;
                    Poz.WidthFactor = 0.8;
                    Poz.Height = 3;
                    acBlkTblRec.AppendEntity(Poz);
                    Tx.AddNewlyCreatedDBObject(Poz, true);
                }

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, new Point3d(XYZ.X  + Apos + Aprim, XYZ.Y, XYZ.Z), Anaim, B, aNAIM, b);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                FPoly2d(TkoorPL);
                if (Naim != "")
                {
                    DBText DOC_N = new DBText();
                    DOC_N.SetDatabaseDefaults();
                    DOC_N.Position = T5;
                    DOC_N.TextString = Naim;
                    DOC_N.WidthFactor = 0.8;
                    DOC_N.Height = 3;
                    acBlkTblRec.AppendEntity(DOC_N);
                    Tx.AddNewlyCreatedDBObject(DOC_N, true);
                }

                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, new Point3d(XYZ.X + Aprim + Apos + Anaim, XYZ.Y, XYZ.Z), Akol, B, aKOL, b);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                FPoly2d(TkoorPL);
                if (kol != "")
                {
                    DBText Poz_Korp = new DBText();
                    Poz_Korp.SetDatabaseDefaults();
                    Poz_Korp.Position = T5;
                    Poz_Korp.TextString = kol;
                    Poz_Korp.WidthFactor = 0.8;
                    Poz_Korp.Height = 3;
                    acBlkTblRec.AppendEntity(Poz_Korp);
                    Tx.AddNewlyCreatedDBObject(Poz_Korp, true);
                }


                RasKoor(ref T1, ref T2, ref T3, ref T4, ref T5, new Point3d(XYZ.X + Aprim + Apos + Anaim + Akol, XYZ.Y, XYZ.Z), Aprim, B, aPRIM, b);
                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                FPoly2d(TkoorPL);
                if (prim != "")
                {
                    DBText Prim = new DBText();
                    Prim.SetDatabaseDefaults();
                    Prim.Position = T5;
                    Prim.TextString = prim;
                    Prim.WidthFactor = 0.8;
                    Prim.Height = 3;
                    acBlkTblRec.AppendEntity(Prim);
                    Tx.AddNewlyCreatedDBObject(Prim, true);
                }

                Tx.Commit();
            }
        }
        public string razbitStrok(int Kol_sim, string Stroka, string Razdel)
        {
            string sobrStr;
            string razdStr;
            string[] StrokaM;
            if (Razdel == ",")
                StrokaM = Stroka.Split(',');
            else
                StrokaM = Stroka.Split(' ');
            sobrStr = "";
            razdStr = StrokaM[0];
            if (StrokaM.Length > 1)
            {
                for (int i = 1; i < StrokaM.Length; i++)
                {
                    if ((razdStr + " " + StrokaM[i]).Length <= Kol_sim)
                        razdStr = razdStr + Razdel + StrokaM[i];
                    else
                    {
                        sobrStr = sobrStr + "&" + razdStr;
                        razdStr = StrokaM[i];
                    }
                }
            }
            sobrStr = sobrStr + "&" + razdStr + "&";
            sobrStr = sobrStr.TrimStart('&');
            return sobrStr;
        }//разбиение строк
        public void ProstPOZ(Point3d T1, Point3d T2,ObjectId ObjID, List<POZIZIA> SpPOZ, List<POZIZIA> SpPOZob) 
        {
#region присвоение начальных значений переменных и создание переменных
                List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
                if (System.IO.File.Exists(@"C:\МАРШРУТ\SpIzNANOsPoz.txt")) HCteniPOZTXT(ref SpPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
                if (System.IO.File.Exists(@"C:\МАРШРУТ\SpIzNANOsPozObor.txt")) HCteniPOZTXT(ref SpPOZob, @"C:\МАРШРУТ\SpIzNANOsPozObor.txt");
                string stKomp = "";
                string stPom = "";
                string stRazd = "";
                string stKEI = "";
                string stDobav = "";
                string stHtoEto = "";
                string Handl_Det = "";
                string Hozain = "";
                string RAB = "";
                string LINKb = "";
                double dDlin = 0;
                double dVisot = 0;
                int Schet;
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;
#endregion
#region Чтение свойств динамического блока
                using (Transaction Tx = db.TransactionManager.StartTransaction())
                {
                    Entity bref1 = Tx.GetObject(ObjID, OpenMode.ForWrite) as Entity;
                    string[] TYPEm = bref1.GetType().ToString().Split('.');
                    string TYPE = TYPEm.Last();
                    if (TYPE == "BlockReference")//если позиция  блок
                    {
                        BlockReference bref = Tx.GetObject(ObjID, OpenMode.ForWrite) as BlockReference;
                        Handl_Det = bref.Handle.ToString();
                        if (bref.IsDynamicBlock)
                        {
                            foreach (ObjectId idAtrRef in bref.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                        if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                        if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                        if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") {if(double.TryParse(atrRef.TextString,out dVisot)) dVisot = Convert.ToDouble(atrRef.TextString); dVisot = Math.Round(dVisot, 0); }
                                    }
                                }
                            }
                            DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                            foreach (DynamicBlockReferenceProperty prop in props)
                            {
                                object[] values = prop.GetAllowedValues();
                                if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); }
                                if (prop.PropertyName == "Расстояние1")
                                {
                                    dDlin = Math.Round(Convert.ToDouble(prop.Value.ToString()));
                                    if (stRazd == "Кожухи") dDlin = (Math.Ceiling(Convert.ToDouble(prop.Value.ToString()) / 100) / 10) * 1000;
                                }
                            }
                            ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 3) { stDobav = value.Value.ToString(); }
                                    if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                    Schet = Schet + 1;
                                }
                            }
                        }
#endregion
#region Если блок статический
                        else
                        {
                            Handl_Det = bref.Handle.ToString();
                            ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { stKomp = value.Value.ToString(); }
                                    if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()); }
                                    if (Schet == 3) { stDobav = value.Value.ToString(); }
                                    if (Schet == 4) { stPom = value.Value.ToString(); }
                                    if (Schet == 5) { stRazd = value.Value.ToString(); }
                                    if (Schet == 6) { stKEI = value.Value.ToString(); }
                                    if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                    if (Schet == 8) { Hozain = value.Value.ToString(); }
                                    if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()); dDlin = Math.Round(dDlin, 0); }
                                    if (Schet == 10) { RAB = value.Value.ToString(); }
                                    if (Schet == 11) { LINKb = value.Value.ToString(); }
                                    Schet = Schet + 1;
                                }
                            }
                            foreach (ObjectId idAtrRef in bref.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Исполнение") { stKomp = atrRef.TextString; }
                                        if (atrRef.Tag == "Помещение") { stPom = atrRef.TextString; }
                                        if (atrRef.Tag == "Раздел_спецификации") { stRazd = atrRef.TextString; }
                                        if (atrRef.Tag == "КЕИ") { stKEI = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") { dVisot = Convert.ToDouble(atrRef.TextString); dVisot = Math.Round(dVisot, 0); }
                                        if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                                    }
                                }
                            }
                            //Application.ShowAlertDialog(RAB + "-" + LINKb);
                        }
                    }
#endregion
#region Если не блок
                    else//если позиция не блок
                    {
                        Entity bref = Tx.GetObject(ObjID, OpenMode.ForWrite) as Entity;
                        Handl_Det = bref.Handle.ToString();
                        ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out dVisot)) dVisot = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 8) { Hozain = value.Value.ToString(); }
                                if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Convert.ToDouble(value.Value.ToString()); dDlin = Math.Round(dDlin / 100, 0); dDlin = dDlin * 100; }
                                if (Schet == 10) { RAB = value.Value.ToString(); }
                                if (Schet == 11) { LINKb = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                        //Application.ShowAlertDialog(RAB + "-" + LINKb);
                    }
#endregion
#region Построение позиции блока
                    double dMas = Convert.ToDouble(Mas);
                    string[] stKompM = stKomp.Split('#');
                    stKomp = stKompM.Last().Replace('@', '.');
                    string[] VirezM = stKompM.First().Split('%');
                    string strTextNP = "";
                    string strTextPP = "";
                    string NOMpoz = "без поз";
                    if (RAB != "")
                    {
                    SBORBl_RAB(ref SpDOP_POZ, RAB, SpPOZsPOZ, LINKb);
                    SBORpl_RAB(ref SpDOP_POZ, RAB, SpPOZsPOZ, LINKb, stRazd, ref dDlin);
                    }
                TextPP_NP_POZ(ref SpNamenkl, stKomp, stPom, stKEI, VirezM, ref strTextNP, ref strTextPP, ref NOMpoz, dDlin, dVisot, stHtoEto, 1, ref SpPOZob);

                    PromptPointResult pPtRes;
                    PromptPointOptions pPtOpts = new PromptPointOptions("");
                    if (strTextNP == "" & (strTextPP == "") == false) { strTextNP = strTextPP; strTextPP = ""; }

                    Point3d ptStart = T1;
                    Point3d ZKr = T2;

                    BlockTable acBlkTbl;
                    BlockTableRecord acBlkTblRec;
                    // Open Model space for write
                    acBlkTbl = Tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    acBlkTblRec = Tx.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                    Point3d ptEnd = new Point3d();
                    Point3d txt = new Point3d();
                    Point3d txtNP = new Point3d();
                    Point3d txtPP = new Point3d();
                    Point3d TPol1 = new Point3d();
                    Point3d TPol2 = new Point3d();

                    RaschKoordin(ptStart, ZKr, ref ptEnd, ref txt, ref txtNP, ref txtPP, ref TPol1, ref TPol2, strTextNP);

                    // Define the new line
                    Line acLine = new Line(ptStart, ptEnd);
                    acLine.SetDatabaseDefaults();
                    acLine.ColorIndex = 1;
                    acLine.Layer = "Насыщение";
                    acLine.XData = new ResultBuffer(
                        new TypedValue(1001, "LAUNCH01"),
                        new TypedValue(1000, ""),
                        new TypedValue(1040, 0),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, LINKb),
                        new TypedValue(1000, "Позиция"));
                // Add the line to the drawing
                acBlkTblRec.AppendEntity(acLine);
                    Tx.AddNewlyCreatedDBObject(acLine, true);
                    Postpoen(ZKr, txt, txtNP, txtPP, TPol1, TPol2, NOMpoz, strTextNP, strTextPP, Handl_Det, "HANDL", "", stRazd, stPom, stKomp, LINKb);

                #endregion
                #region Построение дополнительных позиций
                    DOPpoz(ref SpDOP_POZ, ref SpNamenkl, stKomp, stPom, dDlin, dVisot, SpPOZsPOZ);
                    if (SpPOZob.Exists(x => x.Ind == stKomp) == true) { stDobav = SpPOZob.Find(x => x.Ind == stKomp).Dobav; stPom = SpPOZob.Find(x => x.Ind == stKomp).Pom; }
                    if (stDobav != "" & stDobav != "не важно") { RASKR_DOB(ref SpDOP_POZ, stDobav, stPom, SpPOZsPOZ); }
                    //if (RAB != "")
                    //{
                    //    SBORBl_RAB(ref SpDOP_POZ, RAB, SpPOZsPOZ, LINKb);
                    //    SBORpl_RAB(ref SpDOP_POZ, RAB, SpPOZsPOZ, LINKb);
                    //}
                    string LINK = Handl_Det;
                    foreach (POZIZIA POZ in SpDOP_POZ)
                    {
                        if (POZ.RazdelSp == "Доизол.д" || POZ.RazdelSp == "Доиз.Детали")
                        {
                            stKompM = POZ.Compon.Split('#');
                            stKomp = stKompM.Last().Replace('@', '.');
                            VirezM = stKompM.First().Split('%');
                            Handl_Det = POZ.Compon + "_" + POZ.Pom;
                            strTextNP = "";
                            strTextPP = "";
                            NOMpoz = "без поз";
                            if (ptStart.Y > ptEnd.Y)
                            { ZKr = new Point3d(ZKr.X, ZKr.Y - 200 * dMas, ZKr.Z); }
                            else
                            { ZKr = new Point3d(ZKr.X, ZKr.Y + 200 * dMas, ZKr.Z); }
                            //Application.ShowAlertDialog(POZ.KolF.ToString());
                            TextPP_NP_POZ(ref SpNamenkl, stKomp, POZ.Pom, POZ.KEI, VirezM, ref strTextNP, ref strTextPP, ref NOMpoz, POZ.Dlin, POZ.Visot, POZ.HtoEto, POZ.KolF, ref SpPOZob);
                            RaschKoordin(ptStart, ZKr, ref ptEnd, ref txt, ref txtNP, ref txtPP, ref TPol1, ref TPol2, strTextNP);
                            Postpoen(ZKr, txt, txtNP, txtPP, TPol1, TPol2, NOMpoz, strTextNP, strTextPP, Handl_Det, "COMPON", LINK, POZ.RazdelSp, stPom, stKomp, LINKb);
                        }
                    }
                    Tx.Commit();
#endregion
                }
            }//Проставление позиции
        public void Metka(Point3dCollection TkoorPL, string LINK, string Visot,double dMas,Point3d txt, string Kompon)
        {
            double DlinL = 0;
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
                //poly.Closed = true;
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
                    new TypedValue(1000, Kompon),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, "Метка"),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, ""),
                    new TypedValue(1000, LINK)
                    );
                //Application.ShowAlertDialog(DlinL.ToString());
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);


                DBText Poz = new DBText();
                Poz.SetDatabaseDefaults();
                Poz.Position = txt;
                Poz.Height = 50 * dMas;
                Poz.TextString = Visot;
                Poz.Layer = "Насыщение";
                Poz.XData = new ResultBuffer
                (
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, Kompon),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, "Метка"),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, ""),
                new TypedValue(1000, LINK)
                );
                btr.AppendEntity(Poz);
                tr1.AddNewlyCreatedDBObject(Poz, true);
                //btr.Dispose();
                tr1.Commit();
            }
        }//построение полилинии по списку координат
        public void AutoDim(Point3d Point1, Point3d Point2, double Corner)
        {
            double Dist = Point1.DistanceTo(Point2);
            Vector3d Vekt1 = Point1.GetVectorTo(Point2) / Dist;
            Point3d Point3 = Point1 + Vekt1 * (Dist / 2);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            int CollorInd = 0;
            // Append the point to the database
            using (tr1)
            {
 
                RotatedDimension Dim = new RotatedDimension();
                Dim.SetDatabaseDefaults();
                Dim.XLine1Point = Point1;
                Dim.XLine2Point = Point2;
                Dim.DimLinePoint = Point3;
                Dim.Rotation = Corner;
                Dim.ColorIndex = 1;
                Dim.Layer = "АвтоРазмеры";
                btr.AppendEntity(Dim);
                tr1.AddNewlyCreatedDBObject(Dim, true);
                tr1.Commit();
            }
        }//построение размера по списку координат


    }
    }
