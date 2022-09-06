using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Specialized;
using System.Net;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelNS = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;


namespace FeedGen
{
    public enum PlaySoundFlags : int
    {
        SND_SYNC = 0x0,     // play synchronously (default)
        SND_ASYNC = 0x1,    // play asynchronously
        SND_NODEFAULT = 0x2,    // silence (!default) if sound not found
        SND_MEMORY = 0x4,       // pszSound points to a memory file
        SND_LOOP = 0x8,     // loop the sound until next sndPlaySound
        SND_NOSTOP = 0x10,      // don't stop any currently playing sound
        SND_NOWAIT = 0x2000,    // don't wait if the driver is busy
        SND_ALIAS = 0x10000,    // name is a registry alias
        SND_ALIAS_ID = 0x110000,// alias is a predefined ID
        SND_FILENAME = 0x20000, // name is file name
        SND_RESOURCE = 0x40004, // name is resource name or atom
    }

    public partial class Form1 : Form
    {

        [DllImport("Kernel32.dll")]
        public static extern bool Beep(UInt32 frequency, UInt32 duration);

        Boolean ProcessComplete = false;
        string[] arURL = new string[8];
        int[] arYears = new int[8];
        string[] arMake = new string[8];
        string[] arMakeName = new string[8];
        int RecordsWritten = 0;
        bool RepeatedFAPItem;
        bool Processing = false;
        Dictionary<string, ArrayList> dicFAP = new Dictionary<string, ArrayList>();

        ExcelNS.Application oExcel ;
        ExcelNS.Workbook oWB;
        ExcelNS.Worksheet oSheet;

        bool SaveNow = false;

        //bool saveNow = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cboMakes.SelectedIndex = 0;

            arURL[0] = "http://www.rmeuropean.com/audi-parts.aspx";
            arURL[1] = "http://www.rmeuropean.com/bmw-parts.aspx";
            arURL[2] = "http://www.rmeuropean.com/mercedes-parts.aspx";
            arURL[3] = "http://www.rmeuropean.com/mini-parts.aspx";
            arURL[4] = "http://www.rmeuropean.com/porsche-parts.aspx";
            arURL[5] = "http://www.rmeuropean.com/saab-parts.aspx";
            arURL[6] = "http://www.rmeuropean.com/vw-parts.aspx";
            arURL[7] = "http://www.rmeuropean.com/volvo-parts.aspx";

            arYears[0] = 1978;
            arYears[1] = 1967;
            arYears[2] = 1954;
            arYears[3] = 2002;
            arYears[4] = 1956;
            arYears[5] = 1979;
            arYears[6] = 1978;
            arYears[7] = 1976;

            arYears[0] = 1990;
            arYears[1] = 1990;
            arYears[2] = 1990;
            arYears[3] = 2002;
            arYears[4] = 1990;
            arYears[5] = 1990;
            arYears[6] = 1990;
            arYears[7] = 1990;

            arMake[0] = "60";
            arMake[1] = "10";
            arMake[2] = "20";
            arMake[3] = "30";
            arMake[4] = "40";
            arMake[5] = "50";
            arMake[6] = "60";
            arMake[7] = "70";

            arMakeName[0] = "VW/Audi";
            arMakeName[1] = "BMW";
            arMakeName[2] = "Mercedes";
            arMakeName[3] = "Mini";
            arMakeName[4] = "Porsche";
            arMakeName[5] = "Saab";
            arMakeName[6] = "Volkswagen";
            arMakeName[7] = "Volvo";

            this.Show();
            this.Refresh();
            //StartButton_Click(null, null);
        }


        private void StartButton_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dicModels, dicGroups;
            ArrayList nvcolModels, nvcolGroups;
             

            ArrayList dicCat,colRMEData;
            string data;
            int index=0;
            string JSONData;
            int nPos1,nPos2;
            string SearchID;
            CarPart carPartItem;
            HashSet<string> uniqueCat;
            ArrayList FAPList;
            
            FAPData objFAP;

            ExcelNS.Range range;

            HashSet<string> uniqueItems;
            HashSet<string> nonFoundFAPItems;
            int LastRecWrote;

            bool FoundInFap = false;


            string LastRunFile; 
            string[] lastRunData = new string[8];
            bool resume1 = false, resume2 = false, resume3 = false, resume4 = false, resume5 = false, resume6 = false, resume7 = false;

            string LastCatWritten = "", LastRMENoWritten = "", LastFAPWritten = "";
            //  
            if (ProcessComplete || Processing)
            {
                try
                {
                    oWB.Saved = true;
                    try
                    {
                        oWB.Close(false, Type.Missing, Type.Missing);
                    }
                    catch { }
                    oExcel.Quit();
                }
                catch{}
                Application.Exit();
                return;
            }

            StartButton.Text = "Stop";
            Processing = true;

            if (chkResume.Checked)
                    resume1 = resume2 = resume3 = resume4 = resume5 = resume6 = resume7=true;

            oExcel = new ExcelNS.Application();
            oExcel.DisplayAlerts = false;

            nonFoundFAPItems = new HashSet<string>();

            if (File.Exists("NonFound.bin"))
                nonFoundFAPItems = DeSerializeNonFound();
            if (File.Exists("FAPData.bin"))
                dicFAP = DeSerializeFAPData();
            
            //if (dicFAP==null)
            //    dicFAP = new Dictionary<string, ArrayList>();

            string URL = "";

            string baseURL = arURL[cboMakes.SelectedIndex] ;
            index = cboMakes.SelectedIndex;

            if (!resume1)
            {
                setStatus("Copying template file");
                if (File.Exists(Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls"))
                {
                    if (MessageBox.Show("Output file already exists. Overwrite?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        File.Copy(Application.StartupPath + @"\template.xls", Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls", true);
                    else
                    {
                        ProcessComplete = true;
                        Processing = false;
                        StartButton.Text = "Close";
                        try
                        {
                            oExcel.Quit();
                        }
                        catch { }
                        return;
                    }
                }
                else
                    File.Copy(Application.StartupPath + @"\template.xls", Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls", true);
                
            }

            
            setStatus("Open output file");
            oWB = oExcel.Workbooks.Open(Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oSheet = (ExcelNS.Worksheet)oWB.Sheets[1];

            if (resume1)
                RecordsWritten = oSheet.UsedRange.Row - 2 + oSheet.UsedRange.Rows.Count;

            int TempVal = RecordsWritten;

            if (resume1)
            {

                lastRunData[0] = arURL[index];

                while (true)
                {
                    range = oSheet.get_Range("B" + (TempVal + 1), Type.Missing);
                    if (range.Cells.Value2 != null) break;
                    TempVal--;
                }
                lastRunData[1] = range.Cells.Value2.ToString();

                TempVal = RecordsWritten;
                while (true)
                {
                    range = oSheet.get_Range("C" + (TempVal + 1), Type.Missing);
                    if (range.Cells.Value2 != null) break;
                    TempVal--;
                }
                lastRunData[2] = range.Cells.Value2.ToString();

                TempVal = RecordsWritten;
                while (true)
                {
                    range = oSheet.get_Range("D" + (TempVal + 1), Type.Missing);
                    if (range.Cells.Value2 != null) break;
                    TempVal--;
                }
                lastRunData[3] = range.Cells.Value2.ToString();


                TempVal = RecordsWritten;
                while (true)
                {
                    range = oSheet.get_Range("E" + (TempVal + 1), Type.Missing);
                    if (range.Cells.Value2 != null) break;
                    TempVal--;
                }
                lastRunData[4] = range.Cells.Value2.ToString();

                TempVal = RecordsWritten;
                while (true)
                {
                    range = oSheet.get_Range("F" + (TempVal + 1), Type.Missing);
                    if (range.Cells.Value2 != null) break;
                    TempVal--;
                }
                lastRunData[5] = range.Cells.Value2.ToString();

                TempVal = RecordsWritten;
                while (true)
                {
                    range = oSheet.get_Range("N" + (TempVal + 1), Type.Missing);
                    if (range.Cells.Value2 != null) break;
                    TempVal--;
                }
                lastRunData[6] = range.Cells.Value2.ToString();
                //lastRunData[8] = RecordsWritten.ToString();
            }

            
            resume1 = false;

            Application.DoEvents();
            txtMake.Text = arMakeName[index];
            txtMake.Refresh();
            // For each year
            for (int year = int.Parse(txtStartYear.Text) ; year <= int.Parse(txtEndYear.Text); year++)
            {
                if (resume2)
                    if (year != int.Parse( lastRunData[1]))
                        continue;
                resume2 = false;
                Application.DoEvents();
                txtYear.Text = year.ToString() ;
                txtYear.Refresh();

                setStatus("Browsing for list of models...");

                //URL = baseURL + "?__EVENTTARGET=ctl00_MainContentHolder_submitBut1&__EVENTARGUMENT=MakeYear|{\"make\":\"60\",\"year\":\"" + year.ToString() + "\",\"model\":\"\",\"group\":\"\",\"keyword\":\"\",\"category\":\"\",\"search\":\"\",\"models\":[],\"groups\":[],\"categories\":[]}";
                //data = GetURLContents(URL);
                data = PostToURL(baseURL, "__EVENTTARGET=ctl00_MainContentHolder_submitBut1&__EVENTARGUMENT=MakeYear|{\"make\":\"" + arMake[index] + "\",\"year\":\"" + year.ToString() + "\",\"model\":\"\",\"group\":\"\",\"keyword\":\"\",\"category\":\"\",\"search\":\"\",\"models\":[],\"groups\":[],\"categories\":[]}");
                // Get the models available for that year
                setStatus("Collecting model data...");

                dicModels = GetModels(data, out nvcolModels);

                if (dicModels != null)
                {
                    //For each model
                    foreach (string strModel in dicModels.Keys)
                    {
                        if (resume3)
                            if (dicModels[strModel].ToUpper() != lastRunData[2].ToUpper())
                                continue;
                        resume3 = false;
                        uniqueItems = new HashSet<string>();

                        Application.DoEvents();
                        txtModel.Text = dicModels[strModel];
                        txtModel.Refresh();

                        JSONData = GetJSONModel(arMake[index], year.ToString(), strModel, nvcolModels);
                        //URL = baseURL + "?__EVENTTARGET=ctl00_MainContentHolder_submitBut2&__EVENTARGUMENT=Model|{\"make\":\"60\",\"year\":\"" + year.ToString() + "\",\"model\":\"" + retData[strModel] + "\",\"group\":\"\",\"keyword\":\"\",\"category\":\"\",\"search\":\"\",\"models\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"A4 1.8 Turbo\",\"value\":\"A4A18\"},{\"name\":\"A4 1.8 Turbo Quattro\",\"value\":\"A4Q18\"},{\"name\":\"A4 2.8\",\"value\":\"A4A28\"},{\"name\":\"A4 2.8 Quattro\",\"value\":\"A4Q28\"},{\"name\":\"A6 2.8\",\"value\":\"A6A28\"},{\"name\":\"A6 2.8 Quattro\",\"value\":\"A6Q28\"},{\"name\":\"A6 Quattro 2.7\",\"value\":\"A6Q27\"},{\"name\":\"A6 Quattro 4.2\",\"value\":\"A6Q42\"},{\"name\":\"A8 Quattro 4.2\",\"value\":\"A8Q\"},{\"name\":\"Beetle 1.8\",\"value\":\"BE18\"},{\"name\":\"Beetle 1.9\",\"value\":\"BE19\"},{\"name\":\"Beetle 2.0\",\"value\":\"BE20\"},{\"name\":\"Cabrio (VW)\",\"value\":\"GCAB\"},{\"name\":\"Eurovan 2.8\",\"value\":\"EURO9\"},{\"name\":\"Golf IV 1.8, German\",\"value\":\"GOLFM\"},{\"name\":\"Golf IV 1.9, Brazil\",\"value\":\"GOLFR\"},{\"name\":\"Golf IV 1.9, German\",\"value\":\"GOLFN\"},{\"name\":\"Golf IV 2.0, Brazil\",\"value\":\"GOLFT\"},{\"name\":\"Golf IV 2.0, German\",\"value\":\"GOLFO\"},{\"name\":\"Golf IV 2.8, German\",\"value\":\"GOLFP\"},{\"name\":\"Jetta IV 1.8\",\"value\":\"JETP\"},{\"name\":\"Jetta IV 1.9\",\"value\":\"JETR\"},{\"name\":\"Jetta IV 2.0\",\"value\":\"JETT\"},{\"name\":\"Jetta IV 2.8\",\"value\":\"JETV\"},{\"name\":\"Passat 1.8\",\"value\":\"PASL\"},{\"name\":\"Passat 2.8\",\"value\":\"PASM\"},{\"name\":\"Passat 2.8 4-Motion\",\"value\":\"PAST\"},{\"name\":\"S4 Quattro\",\"value\":\"S4B\"},{\"name\":\"TT\",\"value\":\"ATTA1\"},{\"name\":\"TT Quattro 1.8\",\"value\":\"ATTQ1\"}],\"groups\":[],\"categories\":[]}";
                        URL = baseURL + "?__EVENTTARGET=ctl00_MainContentHolder_submitBut2&__EVENTARGUMENT=Model|"  + JSONData;

                        setStatus("Browsing for list of groups...");

                        //data = GetURLContents(URL);
                        data = PostToURL(baseURL, "__EVENTTARGET=ctl00_MainContentHolder_submitBut2&__EVENTARGUMENT=Model|" + JSONData);

                        nPos1 = data.IndexOf("\"ctl00_MainContentHolder_hidSearchId\" value=\"");
                        if (nPos1 == -1) continue;
                        nPos2 = data.IndexOf("/>", nPos1 );
                        SearchID = data.Substring(nPos1 + 45, nPos2 - nPos1 - 47);

                        //Get the groups available for that model
                        setStatus("Collecting group data...");

                        dicGroups = GetGroups(data,out nvcolGroups);

                        //For each group
                        if (dicGroups == null) continue;
                        foreach (string strGroup in dicGroups.Keys)
                        {
                            if (resume4)
                                if (dicGroups [strGroup].ToUpper() != lastRunData[3].ToUpper())
                                    continue;
                            resume4 = false;
                            if (strGroup != "ALL")
                            {
                                Application.DoEvents();
                                txtGroup.Text = dicGroups[strGroup];
                                txtGroup.Refresh();

                                setStatus("Browsing for list of categories...");

                                JSONData = GetJSONGroup(arMake[index], year.ToString(), strModel, strGroup, SearchID, nvcolModels, nvcolGroups);
                                URL = baseURL + "?__EVENTTARGET=ctl00_MainContentHolder_submitBut3&__EVENTARGUMENT=Group|" + JSONData;

                                //data = GetURLContents(URL);
                                data = PostToURL(baseURL, "__EVENTTARGET=ctl00_MainContentHolder_submitBut3&__EVENTARGUMENT=Group|" + JSONData);

                                nPos1 = data.IndexOf("\"ctl00_MainContentHolder_hidSearchId\" value=\"");
                                if (nPos1 != -1)
                                {
                                    nPos2 = data.IndexOf("/>", nPos1);
                                    SearchID = data.Substring(nPos1 + 45, nPos2 - nPos1 - 47);
                                }

                                setStatus("Collecting category data...");
                                //Get the Categories available for that group
                                dicCat = GetCategories(data);

                                uniqueCat = new HashSet<string>();
                                //For each Catogory
                                if (dicCat!=null)
                                {                                 

                                    foreach (NameValuePair objCat in dicCat)
                                    {
                                        

                                        if (!uniqueCat.Add(objCat.Name))
                                        {
                                            setStatus("Category " + objCat.Value + " already processed. (Skipping)");
                                            continue;

                                        }

                                        if (resume5)
                                            if (objCat.Value.ToUpper() != lastRunData[4].ToUpper())
                                                continue;
                                        resume5 = false;
                                        Application.DoEvents();
                                        txtCat.Text = objCat.Value;
                                        txtCat.Refresh();



                                        JSONData = GetJSONCat(arMake[index], year.ToString(), strModel, strGroup, objCat.Name, SearchID, nvcolModels, nvcolGroups,dicCat);
                                        URL = baseURL + "?__EVENTTARGET=ctl00_MainContentHolder_submitBut4&__EVENTARGUMENT=Category|" + JSONData;

                                        setStatus("Collecting items for the category..." + objCat.Value);

                                        //Fetch item details
                                        //data = GetURLContents(URL);
                                        data = PostToURL(baseURL, "__EVENTTARGET=ctl00_MainContentHolder_submitBut4&__EVENTARGUMENT=Category|" + JSONData);

                                        nPos1 = data.IndexOf("\"ctl00_MainContentHolder_hidSearchId\" value=\"");
                                        if (nPos1 != -1)
                                        {
                                            nPos2 = data.IndexOf("/>", nPos1);
                                            SearchID = data.Substring(nPos1 + 45, nPos2 - nPos1 - 47);
                                        }

                                        colRMEData = GetRMEItems(data);
                                        if (colRMEData == null) 
                                            continue;
                                        foreach (RMEItem rmeItem in colRMEData)
                                        {
                                            

                                            if (resume6)
                                                if (rmeItem.PartNo != lastRunData[5])
                                                    continue;
                                            resume6 = false;
                                            if (nonFoundFAPItems.Contains(rmeItem.PartNo.Replace("-", "")))
                                            {
                                                setStatus(rmeItem.PartNo + " - Item found non existing in FAP");
                                                continue;
                                            }

                                            if (uniqueItems.Add(rmeItem.PartNo))
                                            {
                                                Application.DoEvents();

                                                if (dicFAP!=null && dicFAP.ContainsKey(rmeItem.PartNo))
                                                {
                                                    setStatus(rmeItem.PartNo + " - Part number found in cache");
                                                    
                                                    ArrayList FAPItemCol = dicFAP[rmeItem.PartNo];
                                                    RepeatedFAPItem = false;
                                                    foreach(FAPData objFAPItem in FAPItemCol)
                                                    {
                                                        carPartItem = new CarPart();
                                                        carPartItem.Make = arMakeName[index];
                                                        carPartItem.Year = year;
                                                        carPartItem.Model = dicModels[strModel];
                                                        carPartItem.Group = dicGroups[strGroup];
                                                        carPartItem.Category = objCat.Value;
                                                        carPartItem.PartNo = rmeItem.PartNo;
                                                        carPartItem.Description = rmeItem.Description;
                                                        carPartItem.Manufacturer = rmeItem.Manufacturer;
                                                        carPartItem.OriginalEquipment = rmeItem.OriginalEquipment;
                                                        carPartItem.Application = rmeItem.Application;
                                                        carPartItem.RMEListPrice = rmeItem.ListPrice;
                                                        carPartItem.RMEYourPrice = rmeItem.YourPrice;
                                                        carPartItem.RMEPartNo = rmeItem.PartNo.Replace("-", "");
                                                        carPartItem.FAP99YourPrice = objFAPItem.FAP99YourPrice;
                                                        carPartItem.FAPListPrice = objFAPItem.FAPListPrice;
                                                        carPartItem.FAPManufacturer = objFAPItem.FAPManufacturer;
                                                        carPartItem.CatalogDescription = objFAPItem.FAPCatalogDescription;
                                                        carPartItem.FAPPartNo =objFAPItem.FAPPartNo;

                                                        setStatus("Writing to output file from cache ...FAP Part No :" + objFAPItem.FAPPartNo);

                                                        WriteCarPart(carPartItem);

                                                        if (SaveNow)
                                                        {
                                                            setStatus("Saving...");
                                                            oWB.Save();
                                                            setStatus("Saving Done");
                                                            SaveNow = false;
                                                        }
                                                        
                                                        //setStatus("Saving File");
                                                        //oWB.Save();
                                                        //setStatus("Saving Done");

                                                        LastCatWritten = objCat.Value;
                                                        LastRMENoWritten = rmeItem.PartNo;
                                                        LastFAPWritten = carPartItem.FAPPartNo;
                                                        //File.WriteAllText(LastRunFile, arURL[index] + "|" + year.ToString() + "|" + strModel + "|" + strGroup + "|" + objCat.Name + "|" + rmeItem.PartNo + "|" + carPartItem.FAPPartNo + "|" + RecordsWritten.ToString());
                                                        RepeatedFAPItem = true;
                                                    }
                                                    
                                                }
                                                else
                                                {
                                                    setStatus("Browsing FAP information for item ..." + rmeItem.PartNo);
                                                    FoundInFap = false;
                                                    data = GetURLContents("http://www.fap99.com/searchitem.epc?lookfor=" + rmeItem.PartNo.Replace("-", ""));
                                                    //data = GetURLContents("http://www.fap99.com/ShopByVehicle.epc?q=2004-AUDI-A4--Cabriolet--/--L4_1.8l_turbo_amb-Body--Parts&yearid=2004%40%402004&makeid=AUDI%40%40AUDI%40%40X&engineid=1423314%40%40A4+CABRIOLET+%2F+L4_1.8L_Turbo_AMB%40%40A4+CABRIOLET&catid=Body+Parts&subcatid=Fresh%20Air%20Filter&mode=PA");

                                                    //MatchCollection m1 = Regex.Matches(data, @"(<u.*?>.*?</u>)", RegexOptions.Singleline);

                                                    //carPartItem.FAPPartNo= m1[1].Value.Replace("<u>", "").Replace("</u>","");

                                                    //Loop through each manufacturer

                                                    setStatus("Collecting FAP data ...");
                                                    nPos1 = 1;
                                                    RepeatedFAPItem = false;

                                                    FAPList = new ArrayList();

                                                    while (true)
                                                    {
                                                        Application.DoEvents();

                                                        if (nPos1 == -1 || data=="") break;
                                                        nPos1 = data.IndexOf("<span class=\"bigtext\">", nPos1);
                                                        if (nPos1 == -1) break;

                                                        nPos1 = data.IndexOf(">", nPos1);
                                                        nPos2 = data.IndexOf("<", nPos1);
                                                        string Manufacturer = data.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();

                                                        if (Manufacturer == "Other Parts") break;

                                                        int nPosNext1 = data.IndexOf("<span class=\"bigtext\">", nPos1);

                                                        //Loop through each item of the manufaturer
                                                        while (true)
                                                        {
                                                            objFAP = new FAPData();

                                                            Application.DoEvents();

                                                            nPos1 = data.IndexOf("id=\"txSku_", nPos1);

                                                            if (nPos1 == -1) break;

                                                            nPos1 = data.IndexOf("value=\"", nPos1);

                                                            if (nPos1 > nPosNext1 && nPosNext1 != -1)
                                                            {
                                                                //This item belongs to next manufacurer
                                                                nPos1 = nPosNext1;
                                                                break;
                                                            }
                                                            nPos1 = data.IndexOf("\"", nPos1);
                                                            nPos2 = data.IndexOf("\"", nPos1 + 1);
                                                            string PartNo = data.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();

                                                            nPos1 = data.IndexOf("darkredtext\">", nPos1);
                                                            nPos1 = data.IndexOf(">", nPos1);
                                                            nPos2 = data.IndexOf("<", nPos1 + 1);
                                                            string CatDesc = data.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();

                                                            nPos1 = data.IndexOf("partlistprice\">", nPos1);
                                                            nPos1 = data.IndexOf(">", nPos1);
                                                            nPos2 = data.IndexOf("<", nPos1 + 1);
                                                            string ListPrice = data.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim().Replace("$", "");


                                                            nPos1 = data.IndexOf("partsellprice\">", nPos1);
                                                            nPos1 = data.IndexOf(">", nPos1);
                                                            nPos2 = data.IndexOf("<", nPos1 + 1);
                                                            string YourPrice = data.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim().Replace("$", "");

                                                            carPartItem = new CarPart();
                                                            carPartItem.Make = arMakeName[index];
                                                            carPartItem.Year = year;
                                                            carPartItem.Model = dicModels[strModel];
                                                            carPartItem.Group = dicGroups[strGroup];
                                                            carPartItem.Category = objCat.Value;
                                                            carPartItem.PartNo = rmeItem.PartNo;
                                                            carPartItem.Description = rmeItem.Description;
                                                            carPartItem.Manufacturer = rmeItem.Manufacturer;
                                                            carPartItem.OriginalEquipment = rmeItem.OriginalEquipment;
                                                            carPartItem.Application = rmeItem.Application;
                                                            carPartItem.RMEListPrice = rmeItem.ListPrice;
                                                            carPartItem.RMEYourPrice = rmeItem.YourPrice;
                                                            carPartItem.RMEPartNo = rmeItem.PartNo.Replace("-", "");
                                                            if (YourPrice == "")
                                                                carPartItem.FAP99YourPrice = "Quote";
                                                            else
                                                                carPartItem.FAP99YourPrice = YourPrice;

                                                            carPartItem.FAPListPrice = double.Parse(ListPrice);
                                                            carPartItem.FAPManufacturer = Manufacturer;
                                                            carPartItem.CatalogDescription = CatDesc;
                                                            carPartItem.FAPPartNo = PartNo;

                                                            objFAP.FAP99YourPrice = carPartItem.FAP99YourPrice;
                                                            objFAP.FAPCatalogDescription= CatDesc;
                                                            objFAP.FAPListPrice = carPartItem.FAPListPrice;
                                                            objFAP.FAPManufacturer = Manufacturer;
                                                            objFAP.FAPPartNo = PartNo;

                                                            FAPList.Add(objFAP);

                                                            

                                                            setStatus("Writing to output file ...FAP Part No :" + PartNo);
                                                            FoundInFap = true;

                                                            
                                                            if (resume7)
                                                                if (carPartItem.FAPPartNo != lastRunData[6])
                                                                {
                                                                    RepeatedFAPItem = true;
                                                                    continue;
                                                                }
                                                            resume7 = false;

                                                            WriteCarPart(carPartItem);
                                                            if (SaveNow)
                                                            {
                                                                setStatus("Saving...");
                                                                oWB.Save();
                                                                setStatus("Done");
                                                                SaveNow = false;
                                                            }

                                                            //setStatus("Saving File");
                                                            //oWB.Save();
                                                            //setStatus("Saving Done");

                                                            LastCatWritten = objCat.Value;
                                                            LastFAPWritten = carPartItem.FAPPartNo;
                                                            LastRMENoWritten = rmeItem.PartNo;
                                                            RepeatedFAPItem = true;
                                                        }
                                                    }

                                                    if (!FoundInFap)
                                                    {
                                                        nonFoundFAPItems.Add(rmeItem.PartNo.Replace("-", ""));
                                                        SerializeNonFound(nonFoundFAPItems);
                                                    }
                                                    else
                                                    {
                                                        dicFAP.Add(rmeItem.PartNo, FAPList);
                                                        //SerializeFAPData(dicFAP);
                                                    }
                                                }

                                                if ( RecordsWritten >= 65000)
                                                {
                                                    //Start a new file
                                                    oWB.Save();
                                                    oWB.Saved = true;
                                                    oWB.Close(false, Type.Missing, Type.Missing);

                                                    File.Copy(Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls", Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + year.ToString() + ".xls", true);
                                                    File.Copy(Application.StartupPath + @"\template.xls", Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls", true);

                                                    oWB = oExcel.Workbooks.Open(Application.StartupPath + @"\" + arMakeName[index].Replace("/", "") + ".xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                                                    oSheet = (ExcelNS.Worksheet)oWB.Sheets[1];
                                                    RecordsWritten =0;
                                                    SaveNow = true;
                                                }


                                            }
                                            else
                                            {
                                                setStatus("Part number already added for year and model (skipping)..." + rmeItem.PartNo);
                                            }
                                        }
                                    }
                                    setStatus("Saving File");
                                    oWB.Save();
                                    setStatus("Saving Done");
                                }
                            }
                        }
                    }
                }
            }
            oWB.Close(false, Type.Missing, Type.Missing);
            

            try
            {
                oWB.Close(false, Type.Missing, Type.Missing);
            }
            catch { }
            oExcel.Quit();
            setStatus("Process Complete.");
            
            Beep(1000, 600);

            ProcessComplete = true;
            Processing = false;
            StartButton.Text = "Close";
        }


        private void SerializeNonFound(HashSet<string> item)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream("NonFound.bin",
                         FileMode.Create,
                         FileAccess.Write, FileShare.None);
            formatter.Serialize(stream, item);
            stream.Close();
        }

        private HashSet<string> DeSerializeNonFound()
        {
            HashSet<string> item;
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream("NonFound.bin",
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.Read);
            item = (HashSet<string>)formatter.Deserialize(stream);
            stream.Close();
            return item;
        }

        private void SerializeFAPData(Dictionary<string, ArrayList> item)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream("FAPData.bin",
                         FileMode.Create,
                         FileAccess.Write, FileShare.None);
            formatter.Serialize(stream, item);
            stream.Close();
        }

        private Dictionary<string, ArrayList> DeSerializeFAPData()
        {
            Dictionary<string, ArrayList> item;
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream("FAPData.bin",
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.Read);
            item = (Dictionary<string, ArrayList>)formatter.Deserialize(stream);
            stream.Close();
            return item;
        }

        

        #region Stable Methods

        public void setStatus(string Text)
        {
            //TODO: Remove
            //return;
            if (Status.Lines.Length > 500)
                Status.Clear();
            Status.AppendText(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + ": " + Text + Environment.NewLine);
        }
         
        //private string GetJSONModel(string Make, string Year, string model, Dictionary<string, string> ModelDic)
        private string GetJSONModel(string Make, string Year, string model, ArrayList ModelDic)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);

            JsonWriter jsonWriter = new JsonTextWriter(sw);
            jsonWriter.Formatting = Newtonsoft.Json.Formatting.None;

            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("make");
            jsonWriter.WriteValue(Make);
            jsonWriter.WritePropertyName("year");
            jsonWriter.WriteValue(Year);
            jsonWriter.WritePropertyName("model");
            jsonWriter.WriteValue(model);
            jsonWriter.WritePropertyName("group");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("keyword");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("category");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("search");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("models");
            jsonWriter.WriteStartArray();

            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("name");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("value");
            jsonWriter.WriteValue("");
            jsonWriter.WriteEndObject();

            foreach (NameValuePair sKey in ModelDic)
            {

                jsonWriter.WriteStartObject();
                jsonWriter.WritePropertyName("name");
                jsonWriter.WriteValue(sKey.Name);
                jsonWriter.WritePropertyName("value");
                jsonWriter.WriteValue(sKey.Value);
                jsonWriter.WriteEndObject();
            }


            jsonWriter.WriteEndArray();
            jsonWriter.WritePropertyName("groups");
            jsonWriter.WriteStartArray();
            jsonWriter.WriteEndArray();
            jsonWriter.WritePropertyName("categories");
            jsonWriter.WriteStartArray();
            jsonWriter.WriteEndArray();
            jsonWriter.WriteEndObject();
            return sb.ToString();
        }


        //private string GetJSONGroup(string Make, string Year, string model, string group, string SearchID, Dictionary<string, string> ModelDic, Dictionary<string, string> GroupDic)
        private string GetJSONGroup(string Make, string Year, string model, string group, string SearchID, ArrayList ModelDic, ArrayList GroupDic)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);

            JsonWriter jsonWriter = new JsonTextWriter(sw);
            jsonWriter.Formatting = Newtonsoft.Json.Formatting.None;

            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("make");
            jsonWriter.WriteValue(Make);
            jsonWriter.WritePropertyName("year");
            jsonWriter.WriteValue(Year);
            jsonWriter.WritePropertyName("model");
            jsonWriter.WriteValue(model);
            jsonWriter.WritePropertyName("group");
            jsonWriter.WriteValue(group);
            jsonWriter.WritePropertyName("keyword");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("category");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("search");
            jsonWriter.WriteValue(SearchID);
            jsonWriter.WritePropertyName("models");

            jsonWriter.WriteStartArray();
            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("name");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("value");
            jsonWriter.WriteValue(" ");
            jsonWriter.WriteEndObject();
            foreach (NameValuePair sKey in ModelDic)
            {

                jsonWriter.WriteStartObject();
                jsonWriter.WritePropertyName("name");
                jsonWriter.WriteValue(sKey.Name);
                jsonWriter.WritePropertyName("value");
                jsonWriter.WriteValue(sKey.Value);
                jsonWriter.WriteEndObject();
            }
            jsonWriter.WriteEndArray();

            
            jsonWriter.WritePropertyName("groups");
            jsonWriter.WriteStartArray();
            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("name");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("value");
            jsonWriter.WriteValue(" ");
            jsonWriter.WriteEndObject();
            foreach (NameValuePair sKey in GroupDic)
            {

                jsonWriter.WriteStartObject();
                jsonWriter.WritePropertyName("name");
                jsonWriter.WriteValue(sKey.Name);
                jsonWriter.WritePropertyName("value");
                jsonWriter.WriteValue(sKey.Value);
                jsonWriter.WriteEndObject();
            }


            jsonWriter.WriteEndArray();



            jsonWriter.WritePropertyName("categories");
            jsonWriter.WriteStartArray();
            jsonWriter.WriteEndArray();
            jsonWriter.WriteEndObject();
            return sb.ToString();
        }

        private string GetJSONCat(string Make, string Year, string model, string group, string category, string SearchID, ArrayList ModelDic, ArrayList GroupDic, ArrayList alCat)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);

            JsonWriter jsonWriter = new JsonTextWriter(sw);
            jsonWriter.Formatting = Newtonsoft.Json.Formatting.None;

            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("make");
            jsonWriter.WriteValue(Make);
            jsonWriter.WritePropertyName("year");
            jsonWriter.WriteValue(Year);
            jsonWriter.WritePropertyName("model");
            jsonWriter.WriteValue(model);
            jsonWriter.WritePropertyName("group");
            jsonWriter.WriteValue(group);
            jsonWriter.WritePropertyName("keyword");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("category");
            jsonWriter.WriteValue(category);
            jsonWriter.WritePropertyName("search");
            jsonWriter.WriteValue(SearchID);
            jsonWriter.WritePropertyName("models");

            jsonWriter.WriteStartArray();
            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("name");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("value");
            jsonWriter.WriteValue(" ");
            jsonWriter.WriteEndObject();
            foreach (NameValuePair sKey in ModelDic)
            {

                jsonWriter.WriteStartObject();
                jsonWriter.WritePropertyName("name");
                jsonWriter.WriteValue(sKey.Name);
                jsonWriter.WritePropertyName("value");
                jsonWriter.WriteValue(sKey.Value);
                jsonWriter.WriteEndObject();
            }
            jsonWriter.WriteEndArray();


            jsonWriter.WritePropertyName("groups");
            jsonWriter.WriteStartArray();
            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("name");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("value");
            jsonWriter.WriteValue(" ");
            jsonWriter.WriteEndObject();
            foreach (NameValuePair sKey in GroupDic)
            {

                jsonWriter.WriteStartObject();
                jsonWriter.WritePropertyName("name");
                jsonWriter.WriteValue(sKey.Name);
                jsonWriter.WritePropertyName("value");
                jsonWriter.WriteValue(sKey.Value);
                jsonWriter.WriteEndObject();
            }
            jsonWriter.WriteEndArray();



            jsonWriter.WritePropertyName("categories");
            jsonWriter.WriteStartArray();
            jsonWriter.WriteStartObject();
            jsonWriter.WritePropertyName("name");
            jsonWriter.WriteValue("");
            jsonWriter.WritePropertyName("value");
            jsonWriter.WriteValue(" ");
            jsonWriter.WriteEndObject();
            foreach (NameValuePair sKey in alCat)
            {

                jsonWriter.WriteStartObject();
                jsonWriter.WritePropertyName("name");
                jsonWriter.WriteValue(sKey.Value);
                jsonWriter.WritePropertyName("value");
                jsonWriter.WriteValue(sKey.Name);
                jsonWriter.WriteEndObject();
            }
            jsonWriter.WriteEndArray();

            jsonWriter.WriteEndObject();
            return sb.ToString();
        }

        private string PostToURL(string url,string data)
        {
            //data = "ctl00$ScriptManager1=ctl00$MainContentHolder$UpdatePanel4|ctl00_MainContentHolder_submitBut4&__EVENTTARGET=ctl00_MainContentHolder_submitBut4&__EVENTARGUMENT=Category|{\"make\":\"10\",\"year\":\"2009\",\"model\":\"E90335XD\",\"group\":\"B4\",\"keyword\":\"\",\"category\":\"6184\",\"search\":\"529357\",\"models\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"128i (E82 chassis)\",\"value\":\"E82128C\"},{\"name\":\"128i Conv. (E88 chassis)\",\"value\":\"E88128CV\"},{\"name\":\"135i (E82 chassis)\",\"value\":\"E82135C\"},{\"name\":\"135i Conv. (E88 chassis)\",\"value\":\"E88135CV\"},{\"name\":\"328i (E90 chassis)\",\"value\":\"E90328\"},{\"name\":\"328i Conv. (E93 chassis)\",\"value\":\"E93328\"},{\"name\":\"328i Coupe (E92 chassis)\",\"value\":\"E92328C\"},{\"name\":\"328i Wagon (E91 chassis)\",\"value\":\"E91328\"},{\"name\":\"328i xDrive (E90 chassis)\",\"value\":\"E90328XD\"},{\"name\":\"328i xDrv Coupe (E92 Chassis)\",\"value\":\"E92328CXD\"},{\"name\":\"328i xDrv Wagon (E91 chassis)\",\"value\":\"E91328XD\"},{\"name\":\"335d (E90 chassis)\",\"value\":\"E90335D\"},{\"name\":\"335i (E90 chassis)\",\"value\":\"E90335\"},{\"name\":\"335i Conv. (E93 chassis)\",\"value\":\"E93335\"},{\"name\":\"335i Coupe (E92 chassis)\",\"value\":\"E92335C\"},{\"name\":\"335i xDrive (E90 chassis)\",\"value\":\"E90335XD\"},{\"name\":\"335i xDrv Coupe (E92 chassis)\",\"value\":\"E92335CXD\"},{\"name\":\"528i (E60 chassis)\",\"value\":\"E60528\"},{\"name\":\"528i xDrive (E60 chassis)\",\"value\":\"E60528XD\"},{\"name\":\"535i (E60 chassis)\",\"value\":\"E60535\"},{\"name\":\"535i xDrive (E60 chassis)\",\"value\":\"E60535XD\"},{\"name\":\"535i xDrive Wagon (E61 chassis)\",\"value\":\"E61535WXD\"},{\"name\":\"550i (E60 chassis)\",\"value\":\"E60550\"},{\"name\":\"650i (E63 chassis)\",\"value\":\"E63650C\"},{\"name\":\"650i Conv. (E64 chassis)\",\"value\":\"E64650CV\"},{\"name\":\"750i (F01 chassis)\",\"value\":\"F01750\"},{\"name\":\"750Li (F02 chasis)\",\"value\":\"F02750L\"},{\"name\":\"M3 Conv. (E93 chassis)\",\"value\":\"E93M3CV\"},{\"name\":\"M3 Coupe (E92 chassis)\",\"value\":\"E92M3C\"},{\"name\":\"M3 Sedan (E90 chassis)\",\"value\":\"E90M3\"},{\"name\":\"M5 (E60 chassis)\",\"value\":\"E60M5\"},{\"name\":\"M6 (E63 chassis)\",\"value\":\"E63M6\"},{\"name\":\"M6 Conv. (E64 chassis)\",\"value\":\"E64M6CV\"},{\"name\":\"X3 xDrive30i (E83 chassis)\",\"value\":\"E83X330XD\"},{\"name\":\"X5 xDrive30i (E70 chassis)\",\"value\":\"E70X530XD\"},{\"name\":\"X5 xDrive35d (E70 chassis)\",\"value\":\"E70X535D\"},{\"name\":\"X5 xDrive48i (E70 chassis)\",\"value\":\"E70X548XD\"},{\"name\":\"X6 xDrive35i (E71 chassis)\",\"value\":\"E71X635XD\"},{\"name\":\"X6 xDrive50i (E71 chassis)\",\"value\":\"E71X650XD\"},{\"name\":\"Z4 sDrive30i (E89 chassis)\",\"value\":\"E89Z430\"},{\"name\":\"Z4 sDrive35i (E89 chassis)\",\"value\":\"E89Z435\"}],\"groups\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"All Groups\",\"value\":\"ALL\"},{\"name\":\"Belts\",\"value\":\"B2\"},{\"name\":\"Body\",\"value\":\"B4\"},{\"name\":\"Brakes\",\"value\":\"B6\"},{\"name\":\"Cooling System\",\"value\":\"C2\"},{\"name\":\"Drive Shafts, Axles, Differentials\",\"value\":\"D2\"},{\"name\":\"Engine\",\"value\":\"E2\"},{\"name\":\"Exhaust\",\"value\":\"E4\"},{\"name\":\"Fuel/Air Intake System\",\"value\":\"F2\"},{\"name\":\"Heating, A/C\",\"value\":\"H2\"},{\"name\":\"Ignition, Alternator, Starter, Battery\",\"value\":\"I2\"},{\"name\":\"Lighting\",\"value\":\"L2\"},{\"name\":\"Pedals, Levers\",\"value\":\"P2\"},{\"name\":\"Relays, Motors, Switches, Wiper\",\"value\":\"R2\"},{\"name\":\"Supplies and Miscellaneous\",\"value\":\"Z2\"},{\"name\":\"Suspension, Steering System\",\"value\":\"S2\"},{\"name\":\"Transmission, Clutch\",\"value\":\"T2\"}],\"categories\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"Actuator - Hatch Lock\",\"value\":\"5919\"},{\"name\":\"Actuator - Tailgate Lock\",\"value\":\"5919\"},{\"name\":\"Actuator - Trunk Lock\",\"value\":\"5919\"},{\"name\":\"Air Chanel - Radiator\",\"value\":\"6632\"},{\"name\":\"Air Collector\",\"value\":\"6632\"},{\"name\":\"Air Duct - Radiator\",\"value\":\"6632\"},{\"name\":\"Air Duct Collector\",\"value\":\"6632\"},{\"name\":\"Base - License Plate\",\"value\":\"6155\"},{\"name\":\"Bracket - Bumper Cover\",\"value\":\"6133\"},{\"name\":\"Bumper Carrier - Front\",\"value\":\"6120\"},{\"name\":\"Bumper Carrier - Rear\",\"value\":\"6122\"},{\"name\":\"Bumper Cover - Front\",\"value\":\"6128\"},{\"name\":\"Bumper Cover - Rear\",\"value\":\"6130\"},{\"name\":\"Bumper Cover Clamp\",\"value\":\"6133\"},{\"name\":\"Bumper Cover End Support\",\"value\":\"6133\"},{\"name\":\"Bumper Cover Guide\",\"value\":\"6133\"},{\"name\":\"Bumper Cover Mount\",\"value\":\"6133\"},{\"name\":\"Bumper Cover Support\",\"value\":\"6133\"},{\"name\":\"Bumper Tow Hook Flap\",\"value\":\"6184\"},{\"name\":\"Catch - Hood\",\"value\":\"6524\"},{\"name\":\"CD Holder\",\"value\":\"7308\"},{\"name\":\"CD Magazine\",\"value\":\"7308\"},{\"name\":\"Clamp - Seat Belt\",\"value\":\"7366\"},{\"name\":\"Clip - Door Panel\",\"value\":\"6602\"},{\"name\":\"Clip - Interior Moulding\",\"value\":\"6605\"},{\"name\":\"Door - Front\",\"value\":\"6016\"},{\"name\":\"Door - Rear\",\"value\":\"6017\"},{\"name\":\"Door Emblem\",\"value\":\"6225\"},{\"name\":\"Door Lock Mechanism\",\"value\":\"5935\"},{\"name\":\"Door Panel Clip\",\"value\":\"6602\"},{\"name\":\"Ejector - Fuel Door\",\"value\":\"6045\"},{\"name\":\"Emblem - Door\",\"value\":\"6225\"},{\"name\":\"Emblem - Fender\",\"value\":\"6225\"},{\"name\":\"Emblem - Hatch\",\"value\":\"6225\"},{\"name\":\"Emblem - Hood\",\"value\":\"6225\"},{\"name\":\"Emblem - Roundel\",\"value\":\"6225\"},{\"name\":\"Emblem - Trunk\",\"value\":\"6225\"},{\"name\":\"Emblem Grommet\",\"value\":\"6224\"},{\"name\":\"Engine Hood\",\"value\":\"6012\"},{\"name\":\"Fender\",\"value\":\"6014\"},{\"name\":\"Fender Emblem\",\"value\":\"6225\"},{\"name\":\"Fender Liner\",\"value\":\"6620\"},{\"name\":\"Frame - License Plate\",\"value\":\"6141\"},{\"name\":\"Front Bumper Cover\",\"value\":\"6128\"},{\"name\":\"Front Panel - Radiator Support\",\"value\":\"6008\"},{\"name\":\"Fuel Door Ejector\",\"value\":\"6045\"},{\"name\":\"Fuel Door Latch\",\"value\":\"6045\"},{\"name\":\"Gas Door Ejector\",\"value\":\"6045\"},{\"name\":\"Gas Door Latch\",\"value\":\"6045\"},{\"name\":\"Glove Box Catch\",\"value\":\"7323\"},{\"name\":\"Glove Box Latch\",\"value\":\"7323\"},{\"name\":\"Grille - Kidney\",\"value\":\"6209\"},{\"name\":\"Grille - Radiator\",\"value\":\"6209\"},{\"name\":\"Grommet - Emblem\",\"value\":\"6224\"},{\"name\":\"Hatch Emblem\",\"value\":\"6225\"},{\"name\":\"Hatch Lock Actuator\",\"value\":\"5919\"},{\"name\":\"Holder - CD\",\"value\":\"7308\"},{\"name\":\"Hood\",\"value\":\"6012\"},{\"name\":\"Hood Bracket\",\"value\":\"6528\"},{\"name\":\"Hood Catch\",\"value\":\"6524\"},{\"name\":\"Hood Emblem\",\"value\":\"6225\"},{\"name\":\"Hood Hinge\",\"value\":\"6528\"},{\"name\":\"Hood Lock\",\"value\":\"6538\"},{\"name\":\"Hood Safety Catch\",\"value\":\"6524\"},{\"name\":\"Hood Support\",\"value\":\"6528\"},{\"name\":\"Impact Strip/License Plate Holder\",\"value\":\"6155\"},{\"name\":\"Interior Moulding Clip\",\"value\":\"6605\"},{\"name\":\"Interior Panel Clip\",\"value\":\"6605\"},{\"name\":\"Jack Pad\",\"value\":\"6641\"},{\"name\":\"Kidney Grille\",\"value\":\"6209\"},{\"name\":\"Latch - Fuel Door\",\"value\":\"6045\"},{\"name\":\"License Plate Base\",\"value\":\"6155\"},{\"name\":\"License Plate Frame\",\"value\":\"6141\"},{\"name\":\"License Plate Holder\",\"value\":\"6155\"},{\"name\":\"Lock - Door Mechanism\",\"value\":\"5935\"},{\"name\":\"Lock - Hood\",\"value\":\"6538\"},{\"name\":\"Lock Actuator - Hatch\",\"value\":\"5919\"},{\"name\":\"Lock Actuator - Tailgate\",\"value\":\"5919\"},{\"name\":\"Lock Actuator - Trunk\",\"value\":\"5919\"},{\"name\":\"Lock Panel - Glove Box\",\"value\":\"7323\"},{\"name\":\"Magazine - CD\",\"value\":\"7308\"},{\"name\":\"Radiator Grille\",\"value\":\"6209\"},{\"name\":\"Radiator Support\",\"value\":\"6008\"},{\"name\":\"Radiator Support Baffle\",\"value\":\"6008\"},{\"name\":\"Rear Body Panel\",\"value\":\"6022\"},{\"name\":\"Rear Bumper Cover\",\"value\":\"6130\"},{\"name\":\"Rocker Panel Clip\",\"value\":\"6670\"},{\"name\":\"Roundel - Hood & Trunk\",\"value\":\"6225\"},{\"name\":\"Safety Catch - Hood\",\"value\":\"6524\"},{\"name\":\"Seat Belt Buckle Button Stop\",\"value\":\"7367\"},{\"name\":\"Seat Belt Clamp\",\"value\":\"7366\"},{\"name\":\"Tail Panel\",\"value\":\"6022\"},{\"name\":\"Tail Panel Trim\",\"value\":\"6022\"},{\"name\":\"Tailgate Lock Actuator\",\"value\":\"5919\"},{\"name\":\"Tow Hook Cover\",\"value\":\"6184\"},{\"name\":\"Trunk Emblem\",\"value\":\"6225\"},{\"name\":\"Trunk Lid\",\"value\":\"6026\"},{\"name\":\"Trunk Lock Actuator\",\"value\":\"5919\"},{\"name\":\"Window Regulator\",\"value\":\"6712\"}]}&ctl00$MainContentHolder$ContentColumnControl1$ctl00$KeywordField=&ctl00$MainContentHolder$hidSelection=Category&ctl00$MainContentHolder$hidSSFState={\"make\":\"10\",\"year\":\"2009\",\"model\":\"E90335XD\",\"group\":\"B4\",\"keyword\":\"\",\"category\":\"6184\",\"search\":\"529357\",\"models\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"128i (E82 chassis)\",\"value\":\"E82128C\"},{\"name\":\"128i Conv. (E88 chassis)\",\"value\":\"E88128CV\"},{\"name\":\"135i (E82 chassis)\",\"value\":\"E82135C\"},{\"name\":\"135i Conv. (E88 chassis)\",\"value\":\"E88135CV\"},{\"name\":\"328i (E90 chassis)\",\"value\":\"E90328\"},{\"name\":\"328i Conv. (E93 chassis)\",\"value\":\"E93328\"},{\"name\":\"328i Coupe (E92 chassis)\",\"value\":\"E92328C\"},{\"name\":\"328i Wagon (E91 chassis)\",\"value\":\"E91328\"},{\"name\":\"328i xDrive (E90 chassis)\",\"value\":\"E90328XD\"},{\"name\":\"328i xDrv Coupe (E92 Chassis)\",\"value\":\"E92328CXD\"},{\"name\":\"328i xDrv Wagon (E91 chassis)\",\"value\":\"E91328XD\"},{\"name\":\"335d (E90 chassis)\",\"value\":\"E90335D\"},{\"name\":\"335i (E90 chassis)\",\"value\":\"E90335\"},{\"name\":\"335i Conv. (E93 chassis)\",\"value\":\"E93335\"},{\"name\":\"335i Coupe (E92 chassis)\",\"value\":\"E92335C\"},{\"name\":\"335i xDrive (E90 chassis)\",\"value\":\"E90335XD\"},{\"name\":\"335i xDrv Coupe (E92 chassis)\",\"value\":\"E92335CXD\"},{\"name\":\"528i (E60 chassis)\",\"value\":\"E60528\"},{\"name\":\"528i xDrive (E60 chassis)\",\"value\":\"E60528XD\"},{\"name\":\"535i (E60 chassis)\",\"value\":\"E60535\"},{\"name\":\"535i xDrive (E60 chassis)\",\"value\":\"E60535XD\"},{\"name\":\"535i xDrive Wagon (E61 chassis)\",\"value\":\"E61535WXD\"},{\"name\":\"550i (E60 chassis)\",\"value\":\"E60550\"},{\"name\":\"650i (E63 chassis)\",\"value\":\"E63650C\"},{\"name\":\"650i Conv. (E64 chassis)\",\"value\":\"E64650CV\"},{\"name\":\"750i (F01 chassis)\",\"value\":\"F01750\"},{\"name\":\"750Li (F02 chasis)\",\"value\":\"F02750L\"},{\"name\":\"M3 Conv. (E93 chassis)\",\"value\":\"E93M3CV\"},{\"name\":\"M3 Coupe (E92 chassis)\",\"value\":\"E92M3C\"},{\"name\":\"M3 Sedan (E90 chassis)\",\"value\":\"E90M3\"},{\"name\":\"M5 (E60 chassis)\",\"value\":\"E60M5\"},{\"name\":\"M6 (E63 chassis)\",\"value\":\"E63M6\"},{\"name\":\"M6 Conv. (E64 chassis)\",\"value\":\"E64M6CV\"},{\"name\":\"X3 xDrive30i (E83 chassis)\",\"value\":\"E83X330XD\"},{\"name\":\"X5 xDrive30i (E70 chassis)\",\"value\":\"E70X530XD\"},{\"name\":\"X5 xDrive35d (E70 chassis)\",\"value\":\"E70X535D\"},{\"name\":\"X5 xDrive48i (E70 chassis)\",\"value\":\"E70X548XD\"},{\"name\":\"X6 xDrive35i (E71 chassis)\",\"value\":\"E71X635XD\"},{\"name\":\"X6 xDrive50i (E71 chassis)\",\"value\":\"E71X650XD\"},{\"name\":\"Z4 sDrive30i (E89 chassis)\",\"value\":\"E89Z430\"},{\"name\":\"Z4 sDrive35i (E89 chassis)\",\"value\":\"E89Z435\"}],\"groups\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"All Groups\",\"value\":\"ALL\"},{\"name\":\"Belts\",\"value\":\"B2\"},{\"name\":\"Body\",\"value\":\"B4\"},{\"name\":\"Brakes\",\"value\":\"B6\"},{\"name\":\"Cooling System\",\"value\":\"C2\"},{\"name\":\"Drive Shafts, Axles, Differentials\",\"value\":\"D2\"},{\"name\":\"Engine\",\"value\":\"E2\"},{\"name\":\"Exhaust\",\"value\":\"E4\"},{\"name\":\"Fuel/Air Intake System\",\"value\":\"F2\"},{\"name\":\"Heating, A/C\",\"value\":\"H2\"},{\"name\":\"Ignition, Alternator, Starter, Battery\",\"value\":\"I2\"},{\"name\":\"Lighting\",\"value\":\"L2\"},{\"name\":\"Pedals, Levers\",\"value\":\"P2\"},{\"name\":\"Relays, Motors, Switches, Wiper\",\"value\":\"R2\"},{\"name\":\"Supplies and Miscellaneous\",\"value\":\"Z2\"},{\"name\":\"Suspension, Steering System\",\"value\":\"S2\"},{\"name\":\"Transmission, Clutch\",\"value\":\"T2\"}],\"categories\":[{\"name\":\"\",\"value\":\" \"},{\"name\":\"Actuator - Hatch Lock\",\"value\":\"5919\"},{\"name\":\"Actuator - Tailgate Lock\",\"value\":\"5919\"},{\"name\":\"Actuator - Trunk Lock\",\"value\":\"5919\"},{\"name\":\"Air Chanel - Radiator\",\"value\":\"6632\"},{\"name\":\"Air Collector\",\"value\":\"6632\"},{\"name\":\"Air Duct - Radiator\",\"value\":\"6632\"},{\"name\":\"Air Duct Collector\",\"value\":\"6632\"},{\"name\":\"Base - License Plate\",\"value\":\"6155\"},{\"name\":\"Bracket - Bumper Cover\",\"value\":\"6133\"},{\"name\":\"Bumper Carrier - Front\",\"value\":\"6120\"},{\"name\":\"Bumper Carrier - Rear\",\"value\":\"6122\"},{\"name\":\"Bumper Cover - Front\",\"value\":\"6128\"},{\"name\":\"Bumper Cover - Rear\",\"value\":\"6130\"},{\"name\":\"Bumper Cover Clamp\",\"value\":\"6133\"},{\"name\":\"Bumper Cover End Support\",\"value\":\"6133\"},{\"name\":\"Bumper Cover Guide\",\"value\":\"6133\"},{\"name\":\"Bumper Cover Mount\",\"value\":\"6133\"},{\"name\":\"Bumper Cover Support\",\"value\":\"6133\"},{\"name\":\"Bumper Tow Hook Flap\",\"value\":\"6184\"},{\"name\":\"Catch - Hood\",\"value\":\"6524\"},{\"name\":\"CD Holder\",\"value\":\"7308\"},{\"name\":\"CD Magazine\",\"value\":\"7308\"},{\"name\":\"Clamp - Seat Belt\",\"value\":\"7366\"},{\"name\":\"Clip - Door Panel\",\"value\":\"6602\"},{\"name\":\"Clip - Interior Moulding\",\"value\":\"6605\"},{\"name\":\"Door - Front\",\"value\":\"6016\"},{\"name\":\"Door - Rear\",\"value\":\"6017\"},{\"name\":\"Door Emblem\",\"value\":\"6225\"},{\"name\":\"Door Lock Mechanism\",\"value\":\"5935\"},{\"name\":\"Door Panel Clip\",\"value\":\"6602\"},{\"name\":\"Ejector - Fuel Door\",\"value\":\"6045\"},{\"name\":\"Emblem - Door\",\"value\":\"6225\"},{\"name\":\"Emblem - Fender\",\"value\":\"6225\"},{\"name\":\"Emblem - Hatch\",\"value\":\"6225\"},{\"name\":\"Emblem - Hood\",\"value\":\"6225\"},{\"name\":\"Emblem - Roundel\",\"value\":\"6225\"},{\"name\":\"Emblem - Trunk\",\"value\":\"6225\"},{\"name\":\"Emblem Grommet\",\"value\":\"6224\"},{\"name\":\"Engine Hood\",\"value\":\"6012\"},{\"name\":\"Fender\",\"value\":\"6014\"},{\"name\":\"Fender Emblem\",\"value\":\"6225\"},{\"name\":\"Fender Liner\",\"value\":\"6620\"},{\"name\":\"Frame - License Plate\",\"value\":\"6141\"},{\"name\":\"Front Bumper Cover\",\"value\":\"6128\"},{\"name\":\"Front Panel - Radiator Support\",\"value\":\"6008\"},{\"name\":\"Fuel Door Ejector\",\"value\":\"6045\"},{\"name\":\"Fuel Door Latch\",\"value\":\"6045\"},{\"name\":\"Gas Door Ejector\",\"value\":\"6045\"},{\"name\":\"Gas Door Latch\",\"value\":\"6045\"},{\"name\":\"Glove Box Catch\",\"value\":\"7323\"},{\"name\":\"Glove Box Latch\",\"value\":\"7323\"},{\"name\":\"Grille - Kidney\",\"value\":\"6209\"},{\"name\":\"Grille - Radiator\",\"value\":\"6209\"},{\"name\":\"Grommet - Emblem\",\"value\":\"6224\"},{\"name\":\"Hatch Emblem\",\"value\":\"6225\"},{\"name\":\"Hatch Lock Actuator\",\"value\":\"5919\"},{\"name\":\"Holder - CD\",\"value\":\"7308\"},{\"name\":\"Hood\",\"value\":\"6012\"},{\"name\":\"Hood Bracket\",\"value\":\"6528\"},{\"name\":\"Hood Catch\",\"value\":\"6524\"},{\"name\":\"Hood Emblem\",\"value\":\"6225\"},{\"name\":\"Hood Hinge\",\"value\":\"6528\"},{\"name\":\"Hood Lock\",\"value\":\"6538\"},{\"name\":\"Hood Safety Catch\",\"value\":\"6524\"},{\"name\":\"Hood Support\",\"value\":\"6528\"},{\"name\":\"Impact Strip/License Plate Holder\",\"value\":\"6155\"},{\"name\":\"Interior Moulding Clip\",\"value\":\"6605\"},{\"name\":\"Interior Panel Clip\",\"value\":\"6605\"},{\"name\":\"Jack Pad\",\"value\":\"6641\"},{\"name\":\"Kidney Grille\",\"value\":\"6209\"},{\"name\":\"Latch - Fuel Door\",\"value\":\"6045\"},{\"name\":\"License Plate Base\",\"value\":\"6155\"},{\"name\":\"License Plate Frame\",\"value\":\"6141\"},{\"name\":\"License Plate Holder\",\"value\":\"6155\"},{\"name\":\"Lock - Door Mechanism\",\"value\":\"5935\"},{\"name\":\"Lock - Hood\",\"value\":\"6538\"},{\"name\":\"Lock Actuator - Hatch\",\"value\":\"5919\"},{\"name\":\"Lock Actuator - Tailgate\",\"value\":\"5919\"},{\"name\":\"Lock Actuator - Trunk\",\"value\":\"5919\"},{\"name\":\"Lock Panel - Glove Box\",\"value\":\"7323\"},{\"name\":\"Magazine - CD\",\"value\":\"7308\"},{\"name\":\"Radiator Grille\",\"value\":\"6209\"},{\"name\":\"Radiator Support\",\"value\":\"6008\"},{\"name\":\"Radiator Support Baffle\",\"value\":\"6008\"},{\"name\":\"Rear Body Panel\",\"value\":\"6022\"},{\"name\":\"Rear Bumper Cover\",\"value\":\"6130\"},{\"name\":\"Rocker Panel Clip\",\"value\":\"6670\"},{\"name\":\"Roundel - Hood & Trunk\",\"value\":\"6225\"},{\"name\":\"Safety Catch - Hood\",\"value\":\"6524\"},{\"name\":\"Seat Belt Buckle Button Stop\",\"value\":\"7367\"},{\"name\":\"Seat Belt Clamp\",\"value\":\"7366\"},{\"name\":\"Tail Panel\",\"value\":\"6022\"},{\"name\":\"Tail Panel Trim\",\"value\":\"6022\"},{\"name\":\"Tailgate Lock Actuator\",\"value\":\"5919\"},{\"name\":\"Tow Hook Cover\",\"value\":\"6184\"},{\"name\":\"Trunk Emblem\",\"value\":\"6225\"},{\"name\":\"Trunk Lid\",\"value\":\"6026\"},{\"name\":\"Trunk Lock Actuator\",\"value\":\"5919\"},{\"name\":\"Window Regulator\",\"value\":\"6712\"}]}&ctl00$MainContentHolder$DropDownMake=10&ctl00$MainContentHolder$DropDownYear=2009&ctl00$MainContentHolder$SelectModel=E90335XD&ctl00$MainContentHolder$hidSearchId=529357&ctl00$MainContentHolder$SelectGroup=B4&ctl00$MainContentHolder$keywordBox=&ctl00$MainContentHolder$SelectCategory=6184&ctl00$MainContentHolder$GridView1$ctl02$QtyField=1&ctl00$MainContentHolder$GridView1$ctl03$QtyField=1&ctl00$MainContentHolder$GridView1$ctl04$QtyField=1&ctl00$MainContentHolder$GridView1$ctl05$QtyField=1";
            int i=0;
            while (i < 3)
            {
                try
                {
                    WebRequest Request = WebRequest.Create(url);
                    Request.Method = "POST";
                    Request.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                    Stream requestStream = Request.GetRequestStream();
                    ASCIIEncoding ASCIIEncoding = new ASCIIEncoding();
                    byte[] PostData = ASCIIEncoding.GetBytes(data);
                    requestStream.Write(PostData, 0, PostData.Length);
                    requestStream.Close();
                    StreamReader reader = new StreamReader(Request.GetResponse().GetResponseStream());
                    string retData = reader.ReadToEnd();
                    return retData;
                }
                catch (Exception ex)
                {
                    return "";
                    i++;
                }
            }
            return "";
            
        }

        private string GetURLContents(string url)
        {

            //NetworkCredential myCred = new NetworkCredential("rkrishnamoorthy", "hotpot@2010", "GDCCHENNAI");

            //WebProxy prox = new WebProxy("http://172.16.6.61:8080", true, null, myCred);

            System.Net.WebClient Client = new System.Net.WebClient();
            //Client.Proxy = prox;


            try
            {
                Stream strm = Client.OpenRead(url);
                StreamReader sr = new StreamReader(strm);
                string line;
                System.Text.StringBuilder data = new System.Text.StringBuilder();


                do
                {
                    line = sr.ReadLine();
                    data.AppendLine(line);
                }
                while (line != null);

                strm.Close();
                return data.ToString();
            }
            catch { return ""; }
        }

        private Dictionary<string, string> GetModels(string HTMLData,out ArrayList nvcolIn)
        {            
            int Pos1 = HTMLData.IndexOf("ctl00$MainContentHolder$SelectModel", 0);
            string HTMLSnippet;
            int Pos2;
            Dictionary<string, string> retDict = new Dictionary<string, string>();
            XmlDocument xmlDoc = new XmlDocument();

            nvcolIn = new ArrayList();

            if (Pos1 > 0)
            {
                Pos1 = HTMLData.LastIndexOf("<select", Pos1);
                Pos2 = HTMLData.IndexOf("</select>", Pos1);
                HTMLSnippet = HTMLData.Substring(Pos1, Pos2 - Pos1);
                HTMLSnippet = HTMLSnippet + "</select>";
                xmlDoc.LoadXml(HTMLSnippet);
                XmlNodeList nodes = xmlDoc.SelectNodes("/select/option");

                if (nodes.Count == 1 && nodes[0].Attributes["value"].Value.Trim() == "")
                {
                    //No data found
                    return null;
                }

                else
                {
                    foreach (XmlNode objNode in nodes)
                    {
                        if (objNode.Attributes["value"].Value.Trim() != "")
                        {
                            try
                            {
                                retDict.Add(objNode.Attributes["value"].Value, objNode.InnerText);                                
                            }
                            catch
                            {
                            }
                            nvcolIn.Add( new NameValuePair( objNode.Attributes["value"].Value, objNode.InnerText));
                        }
                    }
                    return retDict;
                }
            }
            else
                return null;
        }

        private Dictionary<string, string> GetGroups(string HTMLData, out ArrayList nvcolIn)
        {
            int Pos1 = HTMLData.IndexOf("ctl00$MainContentHolder$SelectGroup", 0);
            string HTMLSnippet;
            int Pos2;
            Dictionary<string, string> retDict = new Dictionary<string, string>();
            XmlDocument xmlDoc = new XmlDocument();

            nvcolIn = new ArrayList();

            if (Pos1 > 0)
            {
                Pos1 = HTMLData.LastIndexOf("<select", Pos1);
                Pos2 = HTMLData.IndexOf("</select>", Pos1);
                HTMLSnippet = HTMLData.Substring(Pos1, Pos2 - Pos1);
                HTMLSnippet = HTMLSnippet + "</select>";
                xmlDoc.LoadXml(HTMLSnippet);
                XmlNodeList nodes = xmlDoc.SelectNodes("/select/option");

                if (nodes.Count == 1 && nodes[0].Attributes["value"].Value.Trim() == "")
                {
                    //No data found
                    return null;
                }

                else
                {
                    foreach (XmlNode objNode in nodes)
                    {
                        if (objNode.Attributes["value"].Value.Trim() != "")
                        {
                            retDict.Add(objNode.Attributes["value"].Value, objNode.InnerText);
                            nvcolIn.Add( new NameValuePair( objNode.Attributes["value"].Value, objNode.InnerText));
                        }
                    }
                    return retDict;
                }
            }
            else
                return null;
        }

        private ArrayList GetCategories(string HTMLData)
        {
            int Pos1 = HTMLData.IndexOf("ctl00$MainContentHolder$SelectCategory", 0);
            string HTMLSnippet;
            int Pos2;
            ArrayList retDict = new ArrayList();
            XmlDocument xmlDoc = new XmlDocument();
            NameValuePair objNamVal;
            
            if (Pos1 > 0)
            {
                Pos1 = HTMLData.LastIndexOf("<select", Pos1);
                Pos2 = HTMLData.IndexOf("</select>", Pos1);
                HTMLSnippet = HTMLData.Substring(Pos1, Pos2 - Pos1);
                HTMLSnippet = HTMLSnippet + "</select>";
                xmlDoc.LoadXml(HTMLSnippet);
                XmlNodeList nodes = xmlDoc.SelectNodes("/select/option");

                if (nodes.Count == 1 && nodes[0].Attributes["value"].Value.Trim() == "")
                {
                    //No data found
                    return null;
                }

                else
                {
                    foreach (XmlNode objNode in nodes)
                    {
                        objNamVal = new NameValuePair();
                        objNamVal.Name=objNode.Attributes["value"].Value;
                        objNamVal.Value=objNode.InnerText;

                        if (objNode.Attributes["value"].Value.Trim() != "")
                            retDict.Add(objNamVal);
                    }
                    return retDict;
                }
            }
            else
                return null;
        }
        #endregion

        public ArrayList GetRMEItems(string HTMLContent)
        {
            ArrayList retList = new ArrayList();
            RMEItem rmeItem;
            CarPart carPartItem;
            string strSL;
            string TagData;
            int nPos1, nPos2, slNo;
            double price;

            nPos1 = 1;
            slNo = 2;
            if (HTMLContent == "") return null;

            while (true)
            {
                strSL = slNo.ToString("00");

                rmeItem = new RMEItem();
                
                nPos1 = HTMLContent.IndexOf("Part Number:", nPos1);
                if (nPos1 == -1) break;
                nPos1 = HTMLContent.IndexOf("<tr><td>", nPos1);
                nPos1 = HTMLContent.IndexOf("<td>", nPos1);
                nPos1 = HTMLContent.IndexOf(">", nPos1);
                nPos2 = HTMLContent.IndexOf("<", nPos1);
                rmeItem.PartNo = HTMLContent.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();

                nPos1 = HTMLContent.IndexOf("ctl00_MainContentHolder_GridView1_ctl" + strSL + "_NameAnchor", nPos1);
                if (nPos1 > 0)
                {                   
                    nPos1 = HTMLContent.IndexOf(">", nPos1);
                    nPos2 = HTMLContent.IndexOf("<", nPos1);
                    rmeItem.Description = HTMLContent.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();
                    nPos1 = HTMLContent.IndexOf("Manufacturer:</b>", nPos1);
                    nPos1 = HTMLContent.IndexOf(">", nPos1);
                    nPos2 = HTMLContent.IndexOf("<", nPos1);
                    rmeItem.Manufacturer = HTMLContent.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();
                    if (HTMLContent.Substring(nPos2, 100).Contains("Original Equipment"))
                        rmeItem.OriginalEquipment = "Yes";
                    else
                        rmeItem.OriginalEquipment = "No";
                    nPos1 = HTMLContent.IndexOf("Application:</b>", nPos1);
                    nPos1 = HTMLContent.IndexOf(">", nPos1);
                    nPos2 = HTMLContent.IndexOf("<", nPos1 );
                    rmeItem.Application = HTMLContent.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim();
                    //Find List Price
                    while (true)
                    {
                        nPos1 = HTMLContent.IndexOf("</td><td>", nPos1);
                        nPos2 = HTMLContent.IndexOf("</td><td>", nPos1+1);
                        TagData = HTMLContent.Substring(nPos1 + 9, nPos2 - nPos1 - 9);
                        if (double.TryParse(TagData, out price))
                        {
                            rmeItem.ListPrice = price;
                            break;
                        }
                        nPos1 = nPos2;
                    }
                    nPos1 = HTMLContent.IndexOf(":bold\">$", nPos1);
                    nPos1 = HTMLContent.IndexOf("$", nPos1);
                    nPos2 = HTMLContent.IndexOf("<", nPos1 );
                    rmeItem.YourPrice = double.Parse( HTMLContent.Substring(nPos1 + 1, nPos2 - nPos1 - 1).Trim());
                    retList.Add(rmeItem);
                }
                else
                    //End of page
                    break;
                slNo++;
            }
            return retList;
        }

        private void WriteCarPart(CarPart itemToWrite)
        {
            ExcelNS.Range range;
            string CellIndex;

            CellIndex = (RecordsWritten + 2).ToString();

            range = oSheet.get_Range("A" + CellIndex , Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem?  itemToWrite.Make:"";

            range = oSheet.get_Range("B" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Year.ToString() : "";

            range = oSheet.get_Range("C" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Model : "";

            range = oSheet.get_Range("D" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Group : "";

            range = oSheet.get_Range("E" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Category : "";

            range = oSheet.get_Range("F" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.PartNo : "";

            range = oSheet.get_Range("G" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Description : "";

            range = oSheet.get_Range("H" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Manufacturer : "";

            range = oSheet.get_Range("I" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.OriginalEquipment : "";

            range = oSheet.get_Range("J" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.Application : "";

            range = oSheet.get_Range("K" + CellIndex, Type.Missing);
            range.Cells.Value2 = itemToWrite.RMEListPrice.ToString();

            range = oSheet.get_Range("L" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.RMEYourPrice.ToString() : "";

            range = oSheet.get_Range("M" + CellIndex, Type.Missing);
            range.Cells.Value2 = !RepeatedFAPItem ? itemToWrite.RMEPartNo : "";

            range = oSheet.get_Range("N" + CellIndex, Type.Missing);
            range.Cells.Value2 = itemToWrite.FAPPartNo;

            range = oSheet.get_Range("O" + CellIndex, Type.Missing);
            range.Cells.Value2 = Left(itemToWrite.FAPManufacturer,3);

            range = oSheet.get_Range("P" + CellIndex, Type.Missing);
            range.Cells.Value2 = itemToWrite.FAPManufacturer;

            range = oSheet.get_Range("Q" + CellIndex, Type.Missing);
            range.Cells.Value2 = itemToWrite.CatalogDescription;

            range = oSheet.get_Range("R" + CellIndex, Type.Missing);
            range.Cells.Value2 = itemToWrite.FAPListPrice;

            range = oSheet.get_Range("S" + CellIndex, Type.Missing);
            range.Cells.Value2 = itemToWrite.FAP99YourPrice;
                        
            //oWB.Save();
            RecordsWritten++;
        }

        public static string Left(string text, int length)
        {
            if (length < 0)
                throw new ArgumentOutOfRangeException("length", length, "length must be > 0");
            else if (length == 0 || text.Length == 0)
                return "";
            else if (text.Length <= length)
                return text;
            else
                return text.Substring(0, length);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            setStatus("Saving Cache...");
            try
            {
                SerializeFAPData(dicFAP);
                setStatus("Cache saved");
            }
            catch (Exception ex)
            {
                setStatus("Error is saving : " + ex.Message);
            }
        } 
    }

    #region Item Classes
    class NameValuePair
    {
        public string Name;
        public string Value;

        public NameValuePair()
        {
        }

        public NameValuePair(string sName,string sValue)
        {
            Name = sName;
            Value = sValue; 

        }
    }


    class CarPart
    {
        public string Make;
        public int Year;
        public string Model;
        public string Group;
        public string Category;
        public string PartNo;
        public string Description;
        public string Manufacturer;
        public string OriginalEquipment;
        public string Application;
        public double RMEListPrice;
        public double RMEYourPrice;
        public string RMEPartNo;
        public string FAPPartNo;
        public string FAPManufacturer;
        public string CatalogDescription;
        public double FAPListPrice;
        public string FAP99YourPrice;

    }

    [Serializable]
    class FAPData
    {
        public string FAPPartNo;
        public string FAPManufacturer;
        public string FAPCatalogDescription;
        public double FAPListPrice;
        public string FAP99YourPrice;
    }

    class RMEItem
    {
        public string Description;
        public string Manufacturer;
        public string OriginalEquipment;
        public string Application;
        public double ListPrice;
        public double YourPrice;
        public string PartNo;
    }

    #endregion
}
