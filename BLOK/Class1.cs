using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BLOK
{
    //using Platform = HostMgd;
    //using PlatformDb = Teigha;
    //using DbS = Teigha.DatabaseServices;
    //using EdI = HostMgd.EditorInput;
    //using Rt = Teigha.Runtime;
#if AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
#elif NANOCAD
    using Teigha.DatabaseServices;
    using Teigha.Runtime;
    using Teigha.Geometry;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
#endif
    public class Class1
    {
        public struct POZnm 
        {
            public Point3d KoorAC_NC, KoorMIR;
            public double Dlin;
            public void NKoorAC_NC(Point3d i) { KoorAC_NC = i; }
            public void NKoorMIR(Point3d i) { KoorMIR = i; }
            public void NDlin(double i) { Dlin = i; }
        }//позиция для нагрузки масс
        public struct POZIZIA
        {
            public string Compon, Dobav, Pom, RazdelSp, KEI, HtoEto, NOMpoz, Ind, shabl, rNAS, TiPObor, Prim, RAB, LINK, Hozain, CompHozain,Handl,ID;
            public double Dlin, Visot, Kol;
            public double KolF;
            public List<POZnm> spTPSK;
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
            public void NPrim(string i) { Prim = i; }
            public void NRAB(string i) { RAB = i; }
            public void NLINK(string i) { LINK = i; }
            public void NTiPObor(string i) {TiPObor = i;}
            public void NHozain(string i) { Hozain = i; }
            public void NCompHozain(string i) { CompHozain = i; }
            public void NHandl(string i) { Handl = i; }
            public void NID(string i) { ID = i; }
            public void addPOZnm(POZnm i) { spTPSK.Add(i); }
            public void NspTPSK(List<POZnm> i) { spTPSK = i; }

        }//позиция для спецификации
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
        }//плоскости

        [CommandMethod("SV_SP_DB")]
        static public void RemoveXdata()
        {
            Form3 form1 = new Form3();
            form1.Show();
        }//Чтение расширеных данных выбранного примитива 

        [CommandMethod("SborSp")]
        public void ZapLBizBD()
        {
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            List<POZIZIA> SpPOZsPOZ = new List<POZIZIA>();
            List<PLOSpr> SPPal = new List<PLOSpr>();
            FSPplos(ref SPPal);
            HCteniPOZTXT(ref SpPOZsPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
            SBORpl(ref SpPOZ, SPPal);
            SBORBl(ref SpPOZ, SPPal);
            SPvTXTfail(SpPOZ, SpPOZsPOZ);
        }//запись текстового файла со списком выставленых позиций
        [CommandMethod("SborDet")]
        public void ZapLBizBD_Det()
        {
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            List<POZIZIA> SpPOZsPOZ = new List<POZIZIA>();
            List<PLOSpr> SPPal = new List<PLOSpr>();
            FSPplos(ref SPPal);
            HCteniPOZTXT(ref SpPOZsPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
            SBORplDet(ref SpPOZ, SPPal);
            SBORBlDet(ref SpPOZ, SPPal);
            SPDetvTXTfail(SpPOZ, SpPOZsPOZ);
        }//запись текстового файла со списком выставленых деталей
        [CommandMethod("SborPoz")]
        public void ZapLBizBD_Poz()
        {
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            SBORTEXTDet(ref SpPOZ);
            SPTextTXTfail(SpPOZ);
        }//запись текстового файла со списком выставленых деталей
        [CommandMethod("796DB")]
        public void InsBlockRef()
        {
            Form1 form1 = new Form1();
            form1.Show();
        }//вызов формы для выставления ДБ
        [CommandMethod("plos")]
        public void PlosVID()
        {
            Form2 form1 = new Form2();
            form1.NSozd_Izm("Sozd");
            form1.Show();
        }//вызов формы для создания плоскостей
        [CommandMethod("SVplos")]
        public void SVPlosVID()
        {
            Form2 form1 = new Form2();
            form1.NSozd_Izm("Izm");
            form1.Show();
        }//вызов формы для изменения плоскостей
        [CommandMethod("ONP")]
        public void ObnNOMPoz()
        {
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            List<POZIZIA> SpPOZob = new List<POZIZIA>();
            HCteniPOZTXT(ref SpPOZ, @"C:\МАРШРУТ\SpIzNANOsPoz.txt");
            HCteniPOZTXT(ref SpPOZob, @"C:\МАРШРУТ\SpIzNANOsPozObor.txt");
            OBNNpoz(ref SpPOZ, ref SpPOZob);
        }//обновить номера позиций
        [CommandMethod("ODPr")]
        public void ObnRDDet()
        {
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            HCteniDETTXT(ref SpPOZ, @"C:\МАРШРУТ\SPizDetNANOExc.txt");
            foreach (POZIZIA poz in SpPOZ) BLOKPoHandl_Izm(poz);
        }//обновить расширеные данные деталей
        [CommandMethod("OTPr")]
        public void ObnRDPoz()
        {
            List<POZIZIA> SpPOZ = new List<POZIZIA>();
            HCteniTXTTXT(ref SpPOZ, @"C:\МАРШРУТ\SPizPozNANOExc.txt");
            foreach (POZIZIA poz in SpPOZ) TEXTPoHandl_Izm(poz);
        }//обновить расширеные данные деталей
        [CommandMethod("SDB")]
        public static void PLIN_R()
        {
        Document doc = Application.DocumentManager.MdiActiveDocument;
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
                Entity bref1 = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as Entity;
                string[] TYPEm = bref1.GetType().ToString().Split('.');
                string TYPE = TYPEm.Last();
                if (TYPE == "BlockReference")//если позиция  блок
                {
                    BlockReference bref = Tx.GetObject(prEntResult.ObjectId, OpenMode.ForWrite) as BlockReference;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            Application.ShowAlertDialog(prop.PropertyName + "-" + prop.Value.ToString());
                        }
                    }
                }
            }
        }//чтение свойств дин блока


        //[CommandMethod("AddDocEvent")]
        //public void AddDocEvent()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    doc.ImpliedSelectionChanged += new EventHandler(ImpliedSelectionChanged);
        //}
        //public void ImpliedSelectionChanged(object senderObj, EventArgs docColDocActEvtArgs)
        //{
        //    Application.ShowAlertDialog("Изменили выбор");
        //}

        static public void SBORBl(ref List<POZIZIA> SpPOZ, List<PLOSpr> SPPal)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPrim = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            string RAB = "";
            string LINK = "";
            string stCompHoz = "";
            double Out = 0;
            Point3d BAZt = new Point3d();
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
                    stPrim = "";
                    stHOZ = "";
                    RAB = "";
                    LINK = "";
                    stCompHoz = "";
                    dVisot = 0;
                    dDlin = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    BAZt = bref.Position;
                    stKomp = bref.Name;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") {stKomp = prop.Value.ToString(); }
                            if (prop.PropertyName == "Расстояние1") {dDlin=Math.Ceiling(Convert.ToDouble(prop.Value.ToString()) / 100) / 10; }
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
                                if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out Out)) dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                                //if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); if (stHtoEto != "" & stHtoEto != "Лес" & stHtoEto != "тр" & stHtoEto != "хв")stPrim = stHtoEto; }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out Out)) dDlin = Convert.ToDouble(value.Value.ToString()) / 1000; }
                                if (Schet == 10) { RAB = value.Value.ToString(); }
                                if (Schet == 11) { LINK = value.Value.ToString(); }
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
                                if (atrRef.Tag == "Высота_установки") { if (Double.TryParse(atrRef.TextString, out Out)) dVisot = Math.Ceiling(Convert.ToDouble(atrRef.TextString) / 100) / 10; }
                                if (atrRef.Tag == "Примечание" | atrRef.Tag == "ПРИМЕЧАНИЕ") { stPrim = atrRef.TextString; }
                                if (atrRef.Tag == "Что_это") { stHtoEto = atrRef.TextString; }
                            }
                        }
                    }
                    //Application.ShowAlertDialog(stKomp + ":" + stKEI + ":" + stHtoEto + ":" + dDlin.ToString());
                    POZnm Tnm = new POZnm();
                    Tnm.NKoorAC_NC(BAZt);
                    Tnm.NKoorMIR(MIRkoor(SPPal, BAZt));
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
                    tPOZ.NPrim(stPrim);
                    tPOZ.NRAB(RAB);
                    tPOZ.NHozain(stHOZ);
                    if (SpPOZ.Exists(x => x.RAB.Contains(stHOZ + ",") == true | x.RAB.Contains("," + stHOZ) == true | x.RAB.Contains(stHOZ) == true) & stHOZ != "") 
                    {
                        POZIZIA PozH = SpPOZ.Find(x => x.RAB.Contains(stHOZ + ",") == true | x.RAB.Contains("," + stHOZ) == true | x.RAB.Contains(stHOZ) == true);
                        stCompHoz = PozH.Compon;
                    }
                    tPOZ.NCompHozain(stCompHoz);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot);}
                    if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot);}
                    if (stKEI == "796") 
                        {
                         //Application.ShowAlertDialog(VirezM.Last().Contains("шт").ToString());
                        if (VirezM.Last().Contains("шт.") == true)
                        {
                            string[] KolM=VirezM.Last().Split('ш');
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
                            DOBOVLpoz(ref SpPOZ, tPOZ, Tnm);
                            DOPpoz(ref SpDOP_POZ, stKomp, stPom, dDlin, dVisot);
                            foreach (POZIZIA POZ in SpDOP_POZ) { DOBOVLpoz(ref SpPOZ, POZ, Tnm); }
                        }
                        SpDOP_POZ.Clear();
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками
        static public void SBORpl(ref List<POZIZIA> SpPOZ, List<PLOSpr> SPPal)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPrim = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            string RAB = "";
            string LINK = "";
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
                    stPrim = "";
                    RAB = "";
                    LINK = "";
                    stHOZ = "";
                    dVisot = 0;
                    dDlin = 0;
                    Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { if (Double.TryParse(value.Value.ToString(), out dVisot)) dVisot = Convert.ToDouble(value.Value.ToString()) / 1000; }
                                //if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { if (Double.TryParse(value.Value.ToString(), out dDlin)) dDlin = Math.Ceiling(Convert.ToDouble(value.Value.ToString()) / 100)/10; }
                                if (Schet == 10) { RAB = value.Value.ToString(); }
                                if (Schet == 11) { LINK = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    POZnm Tnm = new POZnm();
                    Tnm.NKoorAC_NC(bref.GetPointAtParameter(0.5));
                    Tnm.NKoorMIR(MIRkoor(SPPal, bref.GetPointAtParameter(0.5)));
                    Tnm.NDlin(dDlin);
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
                    tPOZ.NRAB(RAB);
                    tPOZ.NHozain(stHOZ);
                    tPOZ.NKolF(1);
                    tPOZ.NHtoEto(stHtoEto);
                    if (stKEI == "006" & stHtoEto == "") { tPOZ.NKol(dDlin + 2 * dVisot); }
                    if (stKEI == "006" & stHtoEto == "хв") { tPOZ.NKol(dVisot); }
                    if (stKEI == "006" & stHtoEto == "Лес") { tPOZ.NKol(dDlin); }
                    if (stKEI == "006" & stHtoEto == "тр") { tPOZ.NKol(dDlin); }
                    if (stKEI == "796") { tPOZ.NKol(1); }
                    if ((stKomp == "") == false) { DOBOVLpoz(ref SpPOZ, tPOZ, Tnm); }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных полилиниями
        static public void SBORBlDet(ref List<POZIZIA> SpPOZ, List<PLOSpr> SPPal)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPrim = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            string RAB = "";
            string LINK = "";
            string ID = "";
            string stCompHoz = "";
            Point3d BAZt = new Point3d();
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
                    stPrim = "";
                    stHOZ = "";
                    RAB = "";
                    LINK = "";
                    stCompHoz = "";
                    ID = "";
                    dVisot = 0;
                    dDlin = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    BAZt = bref.Position;
                    stKomp = bref.Name;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Исполнение") { stKomp = prop.Value.ToString(); }
                            if (prop.PropertyName == "Расстояние1") { dDlin = Convert.ToDouble(prop.Value.ToString()) ; }
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
                                if (Schet == 1) { stKomp = value.Value.ToString(); }
                                if (Schet == 2) { dVisot = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 3) { stDobav = value.Value.ToString(); }
                                if (Schet == 4) { stPom = value.Value.ToString(); }
                                if (Schet == 5) { stRazd = value.Value.ToString(); }
                                if (Schet == 6) { stKEI = value.Value.ToString(); }
                                if (Schet == 7) { stHtoEto = value.Value.ToString(); if (stHtoEto != "" & stHtoEto != "Лес" & stHtoEto != "тр" & stHtoEto != "хв")stPrim = stHtoEto; }
                                if (Schet == 8) { stHOZ = value.Value.ToString(); }
                                if (Schet == 9) { dDlin = Convert.ToDouble(value.Value.ToString()); }
                                if (Schet == 10) { RAB = value.Value.ToString(); }
                                if (Schet == 11) { LINK = value.Value.ToString(); }
                                if (Schet == 12) { ID = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                        }
                    }
                    POZIZIA tPOZ = new POZIZIA();
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NPrim(stPrim);
                    tPOZ.NHozain(stHOZ);
                    tPOZ.NRAB(RAB);
                    tPOZ.NLINK(LINK);
                    tPOZ.NPrim(stPrim);
                    tPOZ.NHtoEto(stHtoEto);
                    tPOZ.NHandl(bref.Handle.ToString());
                    tPOZ.NID(ID);
                    if ((stKomp == "") == false && (stKEI == "") == false) SpPOZ.Add(tPOZ);
                    SpDOP_POZ.Clear();
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками
        static public void SBORTEXTDet(ref List<POZIZIA> SpPOZ)
        {
            string stKomp = "";
            string stTIP = "";
            string stLINK = "";
            string pomUDAL = "";
            string stKompUDAL = "";
            Point3d BAZt = new Point3d();
            int Schet;
            List<POZIZIA> SpDOP_POZ = new List<POZIZIA>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
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
                ed.WriteMessage("Нет деталей изображенных блоками...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    stTIP = "";
                    stLINK = "";
                    pomUDAL = "";
                    stKompUDAL = "";
                    DBText bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as DBText;
                    BAZt = bref.Position;
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
                        }
                    //Application.ShowAlertDialog(stKomp + " " + stTIP + " " + stLINK + " " + stKompUDAL);
                    POZIZIA tPOZ = new POZIZIA();
                    tPOZ.NCompon(stKomp);;
                    tPOZ.NRAB(stTIP);
                    tPOZ.NHozain(stKompUDAL);
                    tPOZ.NLINK(stLINK);
                    tPOZ.NHandl(bref.Handle.ToString());
                    tPOZ.NCompHozain(bref.TextString);
                    if ((stKomp == "") == false ) SpPOZ.Add(tPOZ);
                }
                Tx.Commit();
            }
        }//Создание списков 
        static public void SBORplDet(ref List<POZIZIA> SpPOZ, List<PLOSpr> SPPal)
        {
            string stKomp = "";
            string stPom = "";
            string stRazd = "";
            string stKEI = "";
            string stDobav = "";
            string stHtoEto = "";
            string stPrim = "";
            string stHOZ = "";
            double dDlin = 0;
            double dVisot = 0;
            string RAB = "";
            string LINK = "";
            string ID = "";
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
                    stPrim = "";
                    RAB = "";
                    LINK = "";
                    stHOZ = "";
                    dVisot = 0;
                    dDlin = 0;
                    ID = "";
                    Polyline bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
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
                            if (Schet == 10) { RAB = value.Value.ToString(); }
                            if (Schet == 11) { LINK = value.Value.ToString(); }
                            if (Schet == 12) { ID = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    //POZnm Tnm = new POZnm();
                    //Tnm.NKoorAC_NC(bref.GetPointAtParameter(0.5));
                    //Tnm.NKoorMIR(MIRkoor(SPPal, bref.GetPointAtParameter(0.5)));
                    //Tnm.NDlin(dDlin);
                    POZIZIA tPOZ = new POZIZIA();
                    //string[] stKompM = stKomp.Split('#');
                    //stKomp = stKompM.Last().Replace('@', '.');
                    tPOZ.NCompon(stKomp);
                    tPOZ.NVisot(dVisot);
                    tPOZ.NDobav(stDobav);
                    tPOZ.NPom(stPom);
                    tPOZ.NRazdelSp(stRazd);
                    tPOZ.NKEI(stKEI);
                    tPOZ.NDlin(dDlin);
                    tPOZ.NRAB(RAB);
                    tPOZ.NHozain(stHOZ);
                    tPOZ.NLINK(LINK);
                    tPOZ.NKolF(1);
                    tPOZ.NPrim(stPrim);
                    tPOZ.NHtoEto(stHtoEto);
                    tPOZ.NHandl(bref.Handle.ToString());
                    tPOZ.NID(ID);
                    if ((stKomp == "") == false) { SpPOZ.Add(tPOZ); }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных полилиниями
        static public void DOPpoz(ref List<POZIZIA> SpDOP_POZ,  string compon, string pom, double Dlin, double Vis)
        {
            if (compon == "688-78.1611")
            {
                POZIZIA TPOZ = new POZIZIA();
                double Kol = (Dlin*1000) / 700;
                if (Kol >= 1)
                {
                    Kol = Math.Truncate((Dlin * 1000) / 700);
                    TPOZ.NNOMpoz("Без поз");
                    TPOZ.NPom(pom);
                    TPOZ.NRazdelSp("Доиз.Детали");
                    if (Vis < 0.250)
                    {
                        TPOZ.NCompon("Проставыш");
                        TPOZ.NKol(Kol * Vis);
                        TPOZ.NKolF(Convert.ToInt16(Kol));
                        TPOZ.NVisot(Vis);
                        TPOZ.NKEI("006");
                        TPOZ.NHtoEto("хв");
                    }
                    else
                    {
                        //if (Vis <= 0.210 & Vis >= 0.200) { TPOZ.NCompon("КЛИГ.746145.001"); }
                        //if (Vis <= 0.220 & Vis > 0.210) { TPOZ.NCompon("КЛИГ.746145.001-01"); }
                        //if (Vis <= 0.230 & Vis > 0.220) { TPOZ.NCompon("КЛИГ.746145.001-02"); }
                        //if (Vis <= 0.240 & Vis > 0.230) { TPOZ.NCompon("КЛИГ.746145.001-03"); }
                        //if (Vis <= 0.250 & Vis > 0.240) { TPOZ.NCompon("КЛИГ.746145.001-04"); }
                        if (Vis <= 0.260 & Vis > 0.250) { TPOZ.NCompon("КЛИГ.746145.001-05"); }
                        if (Vis <= 0.270 & Vis > 0.260) { TPOZ.NCompon("КЛИГ.746145.001-06"); }
                        if (Vis <= 0.280 & Vis > 0.270) { TPOZ.NCompon("КЛИГ.746145.001-07"); }
                        if (Vis <= 0.290 & Vis > 0.280) { TPOZ.NCompon("КЛИГ.746145.001-08"); }
                        if (Vis <= 0.300 & Vis > 0.290) { TPOZ.NCompon("КЛИГ.746145.001-09"); }
                        if (Vis <= 0.310 & Vis > 0.300) { TPOZ.NCompon("КЛИГ.746145.001-10"); }
                        if (Vis <= 0.320 & Vis > 0.310) { TPOZ.NCompon("КЛИГ.746145.001-11"); }
                        if (Vis <= 0.330 & Vis > 0.320) { TPOZ.NCompon("КЛИГ.746145.001-12"); }
                        if (Vis <= 0.340 & Vis > 0.330) { TPOZ.NCompon("КЛИГ.746145.001-13"); }
                        if (Vis <= 0.350 & Vis > 0.340) { TPOZ.NCompon("КЛИГ.746145.001-15"); }
                        if (Vis <= 0.380 & Vis > 0.350) { TPOZ.NCompon("КЛИГ.746145.001-16"); }
                        if (Vis > 380) { TPOZ.NCompon("Самый большой проставыш"); }
                        TPOZ.NKol(Kol);
                        TPOZ.NKolF(Convert.ToInt16(Kol));
                        TPOZ.NVisot(20);
                        TPOZ.NKEI("796");
                        TPOZ.NHtoEto("");
                    }
                    SpDOP_POZ.Add(TPOZ);
                }
                Kol = Dlin;
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("УзкаяЛентаНерж");
                TPOZ.NPom("");
                TPOZ.NKol(Kol);
                TPOZ.NKolF(1);
                TPOZ.NVisot(Vis);
                TPOZ.NDlin(Dlin);
                TPOZ.NHtoEto("");
                TPOZ.NRazdelSp("КрЛот/Лес");
                TPOZ.NKEI("006");
                SpDOP_POZ.Add(TPOZ);

                Kol = Math.Round((Dlin*1000) / 200);
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("УзкийЗамокНерж");
                TPOZ.NPom("");
                TPOZ.NHtoEto("");
                TPOZ.NKol(Kol);
                TPOZ.NKolF(Kol);
                TPOZ.NRazdelSp("КрЛот/Лес");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);
            }
            if (compon.Contains("305142.002") == true)
            {

                POZIZIA TPOZ = new POZIZIA();
                double Kol = Math.Ceiling((Dlin * 1000) / 2000);
                double Kolkr = Math.Ceiling(Dlin);

                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("КЛИГ.363613.006-008");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kolkr < 1) TPOZ.NKol(4); else TPOZ.NKol(Kolkr*4);
                if (Kolkr < 1) TPOZ.NKolF(4); else TPOZ.NKolF(Kolkr * 4);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);


                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("ТЛИШ.685614.002-13");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(1); else TPOZ.NKol(Kol);
                if (Kol < 1) TPOZ.NKolF(1); else TPOZ.NKolF(Kol);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);

                double KolPl = Kol * 2;
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("ТЛИШ.741124.011");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(2); else TPOZ.NKol(KolPl);
                if (Kol < 1) TPOZ.NKolF(2); else TPOZ.NKolF(KolPl);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);


                double KolBalt = Kol * 2;
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("01851363247");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(2); else TPOZ.NKol(KolBalt);
                if (Kol < 1) TPOZ.NKolF(2); else TPOZ.NKolF(KolBalt);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);


                double KolG = Kol * 4;
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("01880161010");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(4); else TPOZ.NKol(KolG);
                if (Kol < 1) TPOZ.NKolF(4); else TPOZ.NKolF(KolG);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);

                double KolH = Kol * 2;
                TPOZ.NNOMpoz("Без поз");
                TPOZ.NCompon("01990108060");
                TPOZ.NPom(pom);
                TPOZ.NHtoEto("");
                if (Kol < 1) TPOZ.NKol(2); else TPOZ.NKol(KolH);
                if (Kol < 1) TPOZ.NKolF(2); else TPOZ.NKolF(KolH);
                TPOZ.NRazdelSp("Кожухи");
                TPOZ.NKEI("796");
                SpDOP_POZ.Add(TPOZ);
            }
        }//Дополнительные позиции для перфоволосы
        static public void OBNNpoz(ref List<POZIZIA> SpPOZ, ref List<POZIZIA> SpPOZob)
        {
#region присвоение начальных значений переменных и создание фильтра для выборки
            string stKomp = "";
            string stTIP = "";
            string stLINK = "";
            string stDlina = "";
            string stVisot = "";
            string stVirez="";
            string stKEI = "";
            string stHtoEto = "";
            string pom = "";
            string pomUDAL = "";
            string stKompUDAL = "";
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
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
                            if (Schet == 1) { stKomp = value.Value.ToString();}
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
                            if (stKompM1.Length>1)
                                if (stKompM.Length>1) stKomp = "#" + stKompM1[1].Replace('@', '.') + "_" + stKompM[1];
                            else
                                if (stKompM.Length > 1) stKomp = stKomp.Replace('@', '.') + "_" + stKompM[1];
                            //Application.ShowAlertDialog(stKomp);
                            //stKomp = stKompM1.Last().Replace('@', '.');                  
                            //if (SpPOZ.Exists(x => (x.Compon == stKomp || "#" + x.Compon == stKomp) & x.Pom == stKompM[1]) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                            //if (SpPOZ.Exists(x => x.Compon + "_" + x.Pom == stKomp | "#" + x.Compon + "_" + x.Pom == stKomp) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                            //if (SpPOZ.Exists(x => stKomp.Contains(x.Compon + "_" + x.Pom)) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
                            if (SpPOZ.Exists(x => stKomp ==("#" + x.Compon + "_" + x.Pom)) == true & stTIP != "T_NAD_POL" & stTIP != "T_POD_POL")
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

                            if (stTIP == "POZIZIA" ) 
                        {
                            string Handl = stKomp;
                            BLOKPoHandl(Handl, ref stKomp, ref pom, ref stDlina, ref stVisot, ref stVirez, ref stKEI, ref stHtoEto, pomUDAL, stKompUDAL);
                            if (SpPOZ.Exists(x => x.Compon == stKomp & x.Pom == pom) == true) { bref.TextString = SpPOZ.Find(x => x.Compon == stKomp & x.Pom == pom).NOMpoz;}
                            if (SpPOZob.Exists(x => x.Ind == stKomp) == true) { bref.TextString = SpPOZob.Find(x => x.Ind == stKomp).NOMpoz;}
                        }
#endregion
#region если в расширеных данных текстового примитива записана ссылка на блок и этот текстовый примитив является текстом над полкой
                            if (stTIP == "T_NAD_POL" && stKomp!="")
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
                                    if ((stDlina == "") == false) { dDlin = Convert.ToDouble(stDlina);}
                                    if ((stVisot == "") == false) { dVisot = Convert.ToDouble(stVisot);}
                                    TextPP_NP_POZ(ref  SpPOZ, stKomp, pom, stKEI, stVirez, ref  strTextNP, ref  strTextPP, ref  NOMpoz, dDlin, dVisot, stHtoEto, 1, ref  SpPOZob);
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
                                    TextPP_NP_POZ(ref  SpPOZ, stKomp, pom, stKEI, stVirez, ref  strTextNP, ref  strTextPP, ref  NOMpoz, dDlin, dVisot, stHtoEto, 1, ref  SpPOZob);
                                    bref.TextString = strTextPP;
                                }
                            }

                        }
                    }
                }
                Tx.Commit();
            }
#endregion
        }//обновить номера позиций
        static void BLOKPoHandl(string Handl, ref string Compon, ref string pom, ref string stDlina, ref string stVisot, ref string stVirez, ref string stKEI, ref string stHtoEto,string pomUDAL,string stKompUDAL)
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
        static void BLOKPoHandl_Izm(POZIZIA IzmPoz)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            int sist10 = Convert.ToInt32(IzmPoz.Handl, 16);
            Handle H = new Handle(sist10);
            using (tr)
            {
                try
                {
                    ObjectId id = db.GetObjectId(false, H, -1);
                    Entity bref1 = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                    bref1.XData = new ResultBuffer(
                        new TypedValue(1001, "LAUNCH01"), 
                        new TypedValue(1000, IzmPoz.Compon),
                        new TypedValue(1040, IzmPoz.Visot),
                        new TypedValue(1000, IzmPoz.Dobav),
                        new TypedValue(1000, IzmPoz.Pom),
                        new TypedValue(1000, IzmPoz.RazdelSp),
                        new TypedValue(1000, IzmPoz.KEI),
                        new TypedValue(1000, IzmPoz.HtoEto),
                        new TypedValue(1000, IzmPoz.Hozain),
                        new TypedValue(1040, IzmPoz.Dlin),
                        new TypedValue(1000, IzmPoz.RAB),
                        new TypedValue(1000, IzmPoz.LINK),
                        new TypedValue(1000, IzmPoz.ID)
                        );
                 tr.Commit();
                }
                catch {  return; }
            }
        }//изменение расширеных данных деталей
        static void TEXTPoHandl_Izm(POZIZIA IzmPoz)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            int sist10 = Convert.ToInt32(IzmPoz.Handl, 16);
            Handle H = new Handle(sist10);
            using (tr)
            {
                try
                {
                    ObjectId id = db.GetObjectId(false, H, -1);
                    Entity bref1 = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                    bref1.XData = new ResultBuffer(
                        new TypedValue(1001, "LAUNCH01"),
                        new TypedValue(1000, IzmPoz.Compon),
                        new TypedValue(1000, IzmPoz.KEI),
                        new TypedValue(1000, IzmPoz.Pom),
                        new TypedValue(1000, IzmPoz.LINK)
                        );
                    tr.Commit();
                }
                catch { return; }
            }
        }//изменение расширеных данных позиций
        
        static void DOBOVLpoz(ref List<POZIZIA> SpPOZ, POZIZIA tPOZ, POZnm Tnm)
        {
            if (SpPOZ.Exists(x => x.Compon == tPOZ.Compon & x.Pom == tPOZ.Pom & x.RazdelSp == tPOZ.RazdelSp & x.CompHozain == tPOZ.CompHozain) == true)
            {
                POZIZIA izmPOZ = SpPOZ.Find(x => x.Compon == tPOZ.Compon & x.Pom == tPOZ.Pom & x.CompHozain == tPOZ.CompHozain);
                SpPOZ.Remove(izmPOZ);
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "") {izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin + 2 * tPOZ.Visot);}
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "" & izmPOZ.Compon == "688-78.1611" & tPOZ.Visot >= 0.25) { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin + 0.04);}
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "хв") {izmPOZ.NKol(izmPOZ.Kol + tPOZ.Visot);}
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "Лес"){izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin);}
                if (izmPOZ.KEI == "006" & izmPOZ.HtoEto == "тр") {izmPOZ.NKol(izmPOZ.Kol + tPOZ.Dlin);}
                if (izmPOZ.KEI == "796") { izmPOZ.NKol(izmPOZ.Kol + tPOZ.Kol); }
                Tnm.NDlin(tPOZ.Kol);
                izmPOZ.spTPSK.Add(Tnm);
                izmPOZ.NKolF(izmPOZ.KolF + tPOZ.KolF);
                izmPOZ.NDobav(izmPOZ.Dobav + "!" + tPOZ.Dobav);
                if (tPOZ.RAB != "") izmPOZ.NRAB(izmPOZ.RAB + "," + tPOZ.RAB);
                if (tPOZ.Hozain != "") izmPOZ.NHozain(izmPOZ.Hozain + "," + tPOZ.Hozain);
                //if (tPOZ.CompHozain != "") izmPOZ.NCompHozain(izmPOZ.CompHozain + "," + tPOZ.CompHozain);
                if (tPOZ.Prim != "" & tPOZ.Prim != null) izmPOZ.NPrim(izmPOZ.Prim + "," + tPOZ.Prim);
                SpPOZ.Add(izmPOZ);
            }
            else 
            {
             List<POZnm> TSppoz = new List<POZnm>();
             Tnm.NDlin(tPOZ.Kol);
             TSppoz.Add(Tnm);
             tPOZ.NspTPSK(TSppoz);
             SpPOZ.Add(tPOZ); 
            }
        }//добовление позиций в список

        public void SPvTXTfail(List<POZIZIA> strPoz, List<POZIZIA> strPozsPoz)
        {
            string NagrMass="";
            string Poz;
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\SPizNANO.txt", false, Encoding.GetEncoding("Windows-1251")))
            { 
                foreach (POZIZIA line in strPoz)
                {
                    NagrMass = "";
                    Poz = "без поз";
                    foreach (POZnm PozNM in line.spTPSK) { NagrMass = NagrMass + "#" + PozNM.Dlin.ToString() + "*" + PozNM.KoorMIR.X.ToString() + " " + PozNM.KoorMIR.Y.ToString() + " " + PozNM.KoorMIR.Z.ToString(); }
                    if (strPozsPoz.Exists(x => x.Compon == line.Compon && x.Pom == line.Pom)) { Poz = strPozsPoz.Find(x => x.Compon == line.Compon && x.Pom == line.Pom).NOMpoz; }
                    file.WriteLine(Poz + ":" + line.Compon + ":" + line.Kol.ToString("#.###") + ":" + line.Pom + ":" + line.KEI + ":" + line.RazdelSp + ":" + line.KolF.ToString("0.") + ":" + line.Dobav + ":" + NagrMass + ":" + line.Prim + ":" + line.Hozain + ":" + line.RAB + ":" + line.CompHozain); 
                }
                file.WriteLine(".");
            }
        }//запись списка блоков в текствоый файл
        public void SPDetvTXTfail(List<POZIZIA> strPoz, List<POZIZIA> strPozsPoz)
        {
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\SPizDetNANO.txt", false, Encoding.GetEncoding("Windows-1251")))
            {
                foreach (POZIZIA line in strPoz)
                {
                    file.WriteLine(line.Compon + ":" + line.Dlin.ToString("00.") + ":" + line.Visot.ToString() + ":" + line.Pom + ":" + line.KEI + ":" + line.RazdelSp + ":" + line.Dobav + ":" + line.Hozain + ":" + line.RAB + ":" + line.LINK + ":" + line.Handl + ":" + line.Prim + ":" + line.ID + ":" + line.HtoEto);
                }
            }
        }//запись списка блоков в текствоый файл
        public void SPTextTXTfail(List<POZIZIA> strPoz)
        {
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\SPizTextNANO.txt", false, Encoding.GetEncoding("Windows-1251")))
            {
                foreach (POZIZIA line in strPoz)
                {
                    file.WriteLine(line.Compon + ":" + line.Hozain + ":" + line.RAB + ":" + line.LINK + ":" + line.Handl + ":" + line.CompHozain);
                }
            }
        }//запись списка блоков в текствоый файл
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
                if (strRazd.Count() > 6)
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
        static public void HCteniDETTXT(ref List<POZIZIA> SpPOZ, string Adres)
        {
            string[] lines = System.IO.File.ReadAllLines(Adres, Encoding.GetEncoding("Windows-1251"));
            foreach (string Strok in lines)
            {
                POZIZIA TPOZ = new POZIZIA();
                string[] strRazd = Strok.Split(':');
                TPOZ.NCompon(strRazd[0]);
                TPOZ.NDlin(Convert.ToDouble(strRazd[1]));
                TPOZ.NVisot(Convert.ToDouble(strRazd[2]));
                TPOZ.NPom(strRazd[3]);
                TPOZ.NKEI(strRazd[4]);
                TPOZ.NRazdelSp(strRazd[5]);
                TPOZ.NDobav(strRazd[6]);
                TPOZ.NHozain(strRazd[7]);
                TPOZ.NRAB(strRazd[8]);
                TPOZ.NLINK(strRazd[9]);
                TPOZ.NHandl(strRazd[10]);
                TPOZ.NPrim(strRazd[11]);
                TPOZ.NID(strRazd[12]);
                TPOZ.NHtoEto(strRazd[13]);
                SpPOZ.Add(TPOZ);
            }
        }//чтение файла с расширенными данными для деталей
        static public void HCteniTXTTXT(ref List<POZIZIA> SpPOZ, string Adres)
        {
            string[] lines = System.IO.File.ReadAllLines(Adres, Encoding.GetEncoding("Windows-1251"));
            foreach (string Strok in lines)
            {
                POZIZIA TPOZ = new POZIZIA();
                string[] strRazd = Strok.Split(':');
                TPOZ.NCompon(strRazd[0]);
                TPOZ.NKEI(strRazd[1]);
                TPOZ.NPom(strRazd[2]);
                TPOZ.NLINK(strRazd[3]);                
                TPOZ.NHandl(strRazd[4]);
                TPOZ.NPrim(strRazd[5]);
                SpPOZ.Add(TPOZ);
            }
        }//чтение файла с расширенными данными для деталей
        static void FSPplos(ref List<PLOSpr> SPPal)
        {
#region переменные и фильтр
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            //List<PLOSpr> SPPal = new List<PLOSpr>();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "POLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Плоскости"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет плоскостей...");
                return;
            }
            SelectionSet acSSet = selRes.Value;
            using (DocumentLock docLock = doc.LockDocument())
            {
#endregion
#region список палуб
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
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
                }
#endregion
            }
        }//создание списка плоскостей
        static Point3d MIRkoor(List<PLOSpr> SPPal, Point3d Toch) 
        {
            Point3d MIRkoor = new Point3d(0,0,0);
            PLOSpr TPlos = new PLOSpr();
            TPlos.NomOSI("мимо");
            foreach (PLOSpr value in SPPal){if (Toch.X > value.Xmin && Toch.X < value.Xmax && Toch.Y > value.Ymin && Toch.Y < value.Ymax) { TPlos = value; }}
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
                MIRkoor = new Point3d(Xnow, Ynow, Znow);
            }
            return MIRkoor; 
        }//
    }
}
