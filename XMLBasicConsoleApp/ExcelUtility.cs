using System;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;

using Microsoft.Office.Interop.Excel;

namespace XMLBasicConsoleApp
{

    public class StringList : ArrayList
    {
        public StringList()
        {
            //
            // TODO: Fügen Sie hier die Konstruktorlogik hinzu
            //
        }

        public StringList(int newSize) : base(newSize)
        {
            //
            // TODO: Fügen Sie hier die Konstruktorlogik hinzu
            //
        }

        public new String this[int index]
        {
            get
            {
                return (String)(base[index]);
            }
            set
            {
                base[index] = value;
            }
        }

        public static bool isAllChrInPool(string str, string pool)
        {
            for (int i = 0; i < str.Length; i++)
            {
                if (pool.IndexOf(str.Substring(i, 1)) < 0)
                    return false;
            }
            return true;
        }

        public static int countOccurence(string outStr, string inChar)
        {
            int oc = 0;
            for (int i = 0; i < outStr.Length; i++)
            {
                if (outStr[i] == inChar[0])
                    oc++;
            }
            return oc;
        }
    }

    public class String2Dim : ArrayList
    {



        public String2Dim()
        {
            //
            // TODO: Fügen Sie hier die Konstruktorlogik hinzu
            //
        }

        public String2Dim(int newSize) : base(newSize)
        {
        }

        public String2Dim(int newRows, int newCols) : base(newRows)
        {
            for (int r = 0; r < newRows; r++)
            {
                StringList sl = new StringList(newCols);
                for (int c = 0; c < newCols; c++)
                {
                    sl.Add("");
                }
                this.Add(sl);
            }
        }

        public new StringList this[int index]
        {
            get
            {
                return (StringList)(base[index]);
            }
            set
            {
                base[index] = value;
            }
        }
    }

    public class String3Dim : ArrayList
    {
        public String3Dim()
        {
            //
            // TODO: Fügen Sie hier die Konstruktorlogik hinzu
            //
        }

        public String3Dim(int newX, int newY, int newZ) : base(newX)
        {
            for (int x = 0; x < newX; x++)
            {
                this.Add(new String2Dim(newY, newZ));
            }
        }

        public new String2Dim this[int index]
        {
            get
            {
                return (String2Dim)(base[index]);
            }
            set
            {
                base[index] = value;
            }
        }
    }


    public class ExcelApplicationExtended
	{
        public Microsoft.Office.Interop.Excel.Application excelApp = null;
		public string stringGuid = System.Guid.NewGuid().ToString().ToUpper();

		public ExcelApplicationExtended()
		{
            excelApp = new Microsoft.Office.Interop.Excel.Application();
			excelApp.DisplayAlerts = false;
            //excelApp.Caption = stringGuid;
            //Microsoft.Office.Core.LanguageSettings test = excelApp.LanguageSettings;
            //LanguageSettings langSettings = 
            //(LanguageSettings) thisApplication.LanguageSettings;
            //int lcid = 
            //langSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
		}

		public void close()
		{
			this.excelApp.Quit();
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
				Debug.WriteLine("Sleeping...");
				System.Threading.Thread.Sleep(1000);
				Debug.WriteLine("End Excel");

			}
			catch { }
			finally
			{
				excelApp = null;
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();  // Scheint zu funktionieren - Excel wird entladen...

			// brutal: Window-Process beenden....
			//foreach (System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
			//{
			//    if (process.MainWindowTitle == this.stringGuid)
			//    {
			//        process.Kill();
			//        break;
			//    }
			//}
		}
	}


    public class ExcelUtility
    {
        public static void Run(string fileName)
        {
            ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();

            String3Dim data = ExcelUtility.loadFromExcel(excelApplicationExtended.excelApp, fileName);
            for (int w = 0; w < data.Count; w++)
            {
                string csvFileName = fileName + w.ToString("000") + ".csv";
                //saveCSV File
                ExcelUtility.saveToCSV(csvFileName, data[w]);

            }

            excelApplicationExtended.close();
        }



        public static String3Dim loadFromExcel(string pathFileExt)
        {
			ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();

			String3Dim data = loadFromExcel(excelApplicationExtended.excelApp, pathFileExt);

			excelApplicationExtended.close();

            return data;
        }

        public static String3Dim loadFromExcel(Microsoft.Office.Interop.Excel._Application excelApp, string pathFileExt)
        {
			string path = Path.GetPathRoot(pathFileExt);
			
			String3Dim data = new String3Dim();

            // OK wird geladen
            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(
                pathFileExt,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                true,
                System.Reflection.Missing.Value
                );
            


            // Excel 2003 macht Probleme: wenn "xls" im tmp-File enthalten, dann werden die Umlaute nicht korrekt in tmp_File geschrieben (Fehler im Excel???)
            string tmpFile = pathFileExt.Replace(".xls", ".txt");
            // die copy Datei wg. Flush?)
            string tmpFileCopy = pathFileExt.Replace(".xls", "copy.txt");
            
            for (int w = 0; w < wb.Sheets.Count; w++)
            {
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[w + 1];

                #region Methode SaveAS macht Probleme mit Formaten ( local -Parameter geht nicht )

                // Probleme probleme Komma /Semikolon
                ws.SaveAs(
                    tmpFile,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    false,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value,
                    true  // ( wird ignoriert: Excel verwendet US-Formate, immerhin gehen die Umlaute...
                    );

                File.Copy(tmpFile, tmpFileCopy, true);  // wg. Flush?

                // String2Dim data1 = ExcelUtility.loadFromCSV_USA(tmpFileCopy);
                String2Dim data1 = ExcelUtility.loadFromCSV(tmpFileCopy);   

                data.Add(data1);

                #endregion SaveAS


                #region Cell by Cell / Methode funktioniert ist aber 10 X langsamer
                /*
                // values helfen nur der Optimierung......
                object[,] values = (object[,])ws.UsedRange.Value2;  // Datumsformate gehen verloren 
                if (values != null)
                {
                    int intRows = ws.UsedRange.Rows.Count;
                    int intCols = ws.UsedRange.Columns.Count;
                    String2Dim data1 = new String2Dim(intRows, intCols);
                    for (int r = 0; r < intRows; r++)
                    {
                        for (int c = 0; c < intCols; c++)
                        {
                            if (values[r + 1, c + 1] != null)
                            {
                                string text = (string)((Microsoft.Office.Interop.Excel.Range)ws.Cells[r + 1, c + 1]).Text;
                                if (text.Contains(";"))
                                {
                                    Debug.Fail("Warnung:  Excel - Zelle enthält Semikolons (;) " + text);
                                }
                                data1[r][c] = text.Trim();
                            }
                        }
                        // Nur Test: Console.WriteLine(r);
                    }
                    data.Add(data1);
                }
                */
                
                #endregion Cell by Cell

            }

            wb.Close(
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value
                );

			//File.Delete( tmpFileCopy);
			//File.Delete( tmpFile );

            return data;
        }



        //public string[] GetRow(string RangeStart, string RangeEnd)
        //{
        //    try
        //    {
        //        if (iRowsRead == Worksheet.UsedRange.Count) 
        //            return null;

        //        iRowsRead++;
        //        Range range = Worksheet.get_Range(RangeStart + iRowsRead.ToString(), RangeEnd + iRowsRead.ToString());
        //        Array myvalues = (Array)range.Cells.Value2;
        //        return saryCurrentRow = ConvertToStringArray(myvalues);
        //    }
        //    catch (COMException ex)
        //    {
        //        throw new COMException("Die Excel-Tabelle ist geschlossen worden!", ex);
        //    }
        //}

        //private string[] ConvertToStringArray(Array values)
        //{
        //    string[] sArray = new string[values.Length];
        //    for (int i = 1; i <= values.Length; i++)
        //    {
        //        if (values.GetValue(1, i) == null)
        //            sArray[i - 1] = "";
        //        else
        //            sArray[i - 1] = values.GetValue(1, i).ToString().Trim();
        //    }
        //    return sArray;
        //}

        public static void saveToExcel(string pathFileExt, String2Dim[] data)
        {
			ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();

			saveToExcel(excelApplicationExtended.excelApp, pathFileExt, data);

			excelApplicationExtended.close();
        }

        public static void saveToExcel(Microsoft.Office.Interop.Excel._Application excelApp, string pathFileExt, String2Dim[] data)
        {
            excelApp.Workbooks.Close();
            Workbook wb = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            for (int s = 0; s < data.Length; s++)
            {
                String2Dim dataSheet = data[s];
                Worksheet ws = null;
                if (wb.Sheets.Count < s)
                {
                    ws = (Worksheet)(wb.Worksheets.Add(
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value
                        ));
                }
                else
                {
                    ws = ((Worksheet)(wb.Worksheets[s + 1]));
                }

                for (int r = 0; r < dataSheet.Count; r++)
                {
                    StringList dataRow = dataSheet[r];
                    for (int c = 0; c < dataRow.Count; c++)
                    {
                        ws.Cells[r + 1, c + 1] = dataRow[c];
                        if (c == 3)
                            ((Microsoft.Office.Interop.Excel.Range)(ws.Cells[r + 1, c + 1])).NumberFormat = "#0,000";
                    }
                }
            }
            int counter = 0;

            FileInfo fi = new FileInfo(pathFileExt);
            string ext = fi.Extension;

            string newPathFile = pathFileExt;
            while (File.Exists(newPathFile))
            {
                newPathFile = pathFileExt.Substring(0, pathFileExt.Length - (ext.Length)) + "_" + (counter++).ToString() + "." + ext;
            }

            wb.SaveCopyAs(newPathFile);
        }


		public static String2Dim loadFromCSV(string pathFile)
		{
			String2Dim data = new String2Dim();
			FileInfo fileCSV = new FileInfo(pathFile);
			StreamReader sr = fileCSV.OpenText();
			while (sr.Peek() >= 0)
			{
				string line = sr.ReadLine();
				string[] tokens = line.Split(';');
				StringList sl = new StringList(tokens.Length);
				for (int t = 0; t < tokens.Length; t++)
				{
					sl.Add(tokens[t]);
				}
				data.Add(sl);
			}
			sr.Close();
			return data;
		}

        public static String2Dim loadFromCSV_USA(string pathFile)
        {
            String2Dim data = new String2Dim();
            //StreamReader sr = new StreamReader(pathFile, System.Text.Encoding.UTF7);  // verschluckt "+"
            StreamReader sr = new StreamReader(pathFile, System.Text.Encoding.Default);  // liest +, umlaute ( möglicherweise ging unter XP, NET 1.1 nicht?)

            
            //StreamReader sr=new StreamReader(@"C:\Test\text1.txt",System.Text.Encoding.GetEncoding(1252));
            //StreamReader sr=new StreamReader(@"C:\Test\text1.txt",System.Text.Encoding.GetEncoding(437));
            //StreamReader sr=new StreamReader(@"C:\Test\text1.txt",System.Text.Encoding.GetEncoding(850));


            while (sr.Peek() >= 0)
            {
                string line = sr.ReadLine();
                string[] tokens = line.Split(',');  // Problem mit Dezimalpunkt z.B. 1,200.30
                StringList sl = new StringList(tokens.Length);
                for (int t = 0; t < tokens.Length; t++)
                {
                    tokens[t] = convertFrom_USA( tokens[t].Trim() );
                    sl.Add(tokens[t]);
                }
                data.Add(sl);
            }
            sr.Close();
            return data;
        }

        public static string convertFrom_USA(string token)
        {
            if (isFloatNumber_USA(token))
                return token.Replace(".", ",");
            if (isDate_USA(token))
            {
                string[] tokens = token.Split('/');
                DateTime dt = new DateTime(Int32.Parse(tokens[2]), Int32.Parse(tokens[0]), Int32.Parse(tokens[1] ) );
                return dt.ToShortDateString();
            }
            return token;
        }

        public static bool isFloatNumber_USA(string token)
        {
            if (StringList.isAllChrInPool(token, "-0123456789."))
                if (StringList.countOccurence(token, ".") == 1)
                    return true;
            return false;
        }

        public static bool isDate_USA(string token)
        {
            if (StringList.isAllChrInPool(token, @"0123456789/"))
                if (StringList.countOccurence(token, @"/") == 2)
                    return true;
            return false;
        }

        
        public static void saveToCSV(string pathFile, String2Dim data)
        {
            FileInfo fileCSV = new FileInfo(pathFile);
            StreamWriter sw = fileCSV.CreateText();
            for (int r = 0; r < data.Count; r++)
            {
                string line = "";
                for (int c = 0; c < data[r].Count; c++)
                {
                    line += data[r][c] + (c < data[r].Count - 1 ? ";" : "");
                }
                sw.WriteLine(line);
            }
            sw.Close();
            return;
        }

        public static void extractCSVFromPath(string pathName)
        {
            if (Directory.Exists(pathName) )
            {
				ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();
				
                String[] filesXLS = Directory.GetFiles(pathName, "*.xls");
                for (int f = 0; f < filesXLS.Length; f++)
                {
					String3Dim data = loadFromExcel(excelApplicationExtended.excelApp, filesXLS[f]);
                    for (int w = 0; w < data.Count; w++)
                    {
                        string csvFileName = filesXLS[f] + w.ToString("000") + ".csv";
                        //saveCSV File
                        saveToCSV( csvFileName, data[w] );
                    }
                }
			
				excelApplicationExtended.close();

			}
        }

        //		public static String3Dim loadFromExcelViaCSV( Excel.Application excelApp, string pathFile )
        //		{
        //			String3Dim data = new String3Dim();;
        //
        //			StringList fileList = new StringList();
        //
        //			Excel.Workbook wb = excelApp.Workbooks.Open( pathFile,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value
        //				);
        //			
        //			for ( int w = 0; w < wb.Sheets.Count; w++ )
        //			{
        //				Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[w+1];
        //				string pathFileCSV = pathFile+"_"+w.ToString("000")+".csv";
        //
        //				object[,] values = (object[,])ws.UsedRange.Value2;
        //				if ( values != null )
        //				{
        //					ws.SaveAs( 
        //						pathFileCSV,
        //						Excel.XlFileFormat.xlCSVWindows,
        //						System.Reflection.Missing.Value,
        //						System.Reflection.Missing.Value,
        //						System.Reflection.Missing.Value,
        //						System.Reflection.Missing.Value,
        //						System.Reflection.Missing.Value,
        //						System.Reflection.Missing.Value,
        //						System.Reflection.Missing.Value
        //						);
        //					fileList.Add( pathFileCSV );
        //				}
        //			}
        //
        //			wb.Close( 
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value,
        //				System.Reflection.Missing.Value
        //				);
        //
        //			for ( int w = 0; w < fileList.Count; w++ )
        //			{
        //				data.Add( loadFromCSV( fileList[w] ) );
        //			}
        //			return data;
        //		}



        public static void doMannheimer()
        {
            String2Dim dataTarget = new String2Dim();

            StringList newRecord = new StringList();
            newRecord.Add(" ");
            newRecord.Add("N:");
            newRecord.Add("BAP:");
            newRecord.Add("G:");
            newRecord.Add("ALTER:");  // Alter
            newRecord.Add("BTR:");  // Beitrag
            newRecord.Add("KEA:");  // KEA
            newRecord.Add("KMA:");  // KMA
            newRecord.Add("KMV:");  // KMV

            dataTarget.Add(newRecord);

            string pathName = @"C:\Projekte\Dokumente\Kv\Mannheimer\Eigene\Test";
            if (Directory.Exists(pathName))
            {
                ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();
                String[] filesXLS = Directory.GetFiles(pathName, "*.xls");
                for (int f = 0; f < filesXLS.Length; f++)
                {
                    try
                    {
                        // String3Dim fileData = loadFromExcel(filesXLS[f]);
                        String3Dim fileData = loadFromExcelSlow(filesXLS[f]);
                        String2Dim ws0Data = fileData[0];

                        string tarifName = ws0Data[1][1];
                        string stand = ws0Data[3][1].ToString().Substring(7);
                        for (int r = 7; r < 200; r++)
                        {
                            StringList oldRecord = ws0Data[r];

                            if (oldRecord[0].Trim() == "")
                            {
                                break;
                            }

                            newRecord = new StringList();

                            newRecord.Add(" ");
                            newRecord.Add(tarifName);
                            newRecord.Add(stand);
                            newRecord.Add("M");
                            newRecord.Add(oldRecord[0]);  // Alter
                            newRecord.Add(oldRecord[1]);  // Beitrag
                            newRecord.Add(oldRecord[3]);  // KEA
                            newRecord.Add(oldRecord[4]);  // KMA
                            newRecord.Add(oldRecord[5]);  // KMV

                            dataTarget.Add(newRecord);

                            newRecord = new StringList();

                            newRecord.Add(" ");
                            newRecord.Add(tarifName);
                            newRecord.Add(stand);
                            newRecord.Add("W");
                            newRecord.Add(oldRecord[0]);  // Alter
                            newRecord.Add(oldRecord[7]);  // Beitrag
                            newRecord.Add(oldRecord[9]);  // KEA
                            newRecord.Add(oldRecord[10]);  // KMA
                            newRecord.Add(oldRecord[11]);  // KMV

                            dataTarget.Add(newRecord);
                        }
                    }
                    catch
                    {
                        Debug.Write("fehler " + filesXLS[f]);
                    }
                    Debug.Write(f.ToString());
                }
                excelApplicationExtended.close();

                ExcelUtility.saveToCSV(pathName+"\\output.txt", dataTarget);
                            
            }
        }


        public static String3Dim loadFromExcelSlow(string pathFileExt)
        {
            ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();

            String3Dim data = loadFromExcelSlow(excelApplicationExtended.excelApp, pathFileExt);

            excelApplicationExtended.close();

            return data;
        }

        public static String3Dim loadFromExcelSlow(Microsoft.Office.Interop.Excel._Application excelApp, string pathFileExt)
        {
            string path = Path.GetPathRoot(pathFileExt);

            String3Dim data = new String3Dim();

            // OK wird geladen
            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(
                pathFileExt,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                true,
                System.Reflection.Missing.Value
                );



            // Excel 2003 macht Probleme: wenn "xls" im tmp-File enthalten, dann werden die Umlaute nicht korrekt in tmp_File geschrieben (Fehler im Excel???)
            string tmpFile = pathFileExt.Replace(".xls", ".txt");
            // die copy Datei wg. Flush?)
            string tmpFileCopy = pathFileExt.Replace(".xls", "copy.txt");

            for (int w = 0; w < wb.Sheets.Count; w++)
            {
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[w + 1];

                //#region Methode SaveAS macht Probleme mit Formaten ( local -Parameter geht nicht )

                //// Probleme probleme Komma /Semikolon
                //ws.SaveAs(
                //    tmpFile,
                //    Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows,
                //    System.Reflection.Missing.Value,
                //    System.Reflection.Missing.Value,
                //    System.Reflection.Missing.Value,
                //    false,
                //    System.Reflection.Missing.Value,
                //    System.Reflection.Missing.Value,
                //    System.Reflection.Missing.Value,
                //    true  // ( wird ignoriert: Excel verwendet US-Formate, immerhin gehen die Umlaute...
                //    );

                //File.Copy(tmpFile, tmpFileCopy, true);  // wg. Flush?

                //String2Dim data1 = ExcelUtility.loadFromCSV_USA(tmpFileCopy);

                //data.Add(data1);

                //#endregion SaveAS


                #region Cell by Cell / Methode funktioniert ist aber 10 X langsamer
                
                // values helfen nur der Optimierung......
                object[,] values = (object[,])ws.UsedRange.Value2;  // Datumsformate gehen verloren 
                if (values != null)
                {
                    int intRows = ws.UsedRange.Rows.Count;
                    int intCols = ws.UsedRange.Columns.Count;
                    String2Dim data1 = new String2Dim(intRows, intCols);
                    for (int r = 0; r < intRows; r++)
                    {
                        for (int c = 0; c < intCols; c++)
                        {
                            if (values[r + 1, c + 1] != null)
                            {
                                string text = (string)((Microsoft.Office.Interop.Excel.Range)ws.Cells[r + 1, c + 1]).Text;
                                if (text.Contains(";"))
                                {
                                    Debug.Fail("Warnung:  Excel - Zelle enthält Semikolons (;) " + text);
                                }
                                data1[r][c] = text.Trim();
                            }
                        }
                        // Nur Test: Console.WriteLine(r);
                    }
                    data.Add(data1);
                }
                #endregion Cell by Cell
            }

            wb.Close(
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value,
                System.Reflection.Missing.Value
                );

            //File.Delete( tmpFileCopy);
            //File.Delete( tmpFile );

            return data;
        }

        public static void doHUK()
        {
            String2Dim dataTarget = new String2Dim();

            StringList newRecord = new StringList();
            newRecord.Add(" ");
            newRecord.Add("N:");
            newRecord.Add("BAP:");
            newRecord.Add("G:");
            newRecord.Add("ALTER:");  // Alter
            newRecord.Add("BTR:");  // Beitrag
            newRecord.Add("KEA:");  // KEA
            newRecord.Add("KMA:");  // KMA
            newRecord.Add("KMV:");  // KMV

            dataTarget.Add(newRecord);

            string pathName = @"C:\Projekte\Dokumente\Kv\HUK\Eigene\Test";
            if (Directory.Exists(pathName))
            {
                ExcelApplicationExtended excelApplicationExtended = new ExcelApplicationExtended();
                String[] filesXLS = Directory.GetFiles(pathName, "*.xls");
                for (int f = 0; f < filesXLS.Length; f++)
                {
                    try
                    {
                        // String3Dim fileData = loadFromExcel(filesXLS[f]);
                        String3Dim fileData = loadFromExcelSlow(filesXLS[f]);

                        for (int w = 0; w < fileData.Count; w++)
                        {

                            String2Dim ws0Data = fileData[w];

                            string tarifName = ws0Data[1][1];
                            string stand = ""; // ws0Data[3][1].ToString().Substring(7);
                            for (int r = 12; r < 100; r++)
                            {
                                if (r < ws0Data.Count)
                                {

                                    StringList orgRecord = ws0Data[r];
                                    StringList oldRecord = new StringList();
                                    for (int s = 0; s < 10; s++)
                                    {
                                        oldRecord.Add("");
                                        if (s < orgRecord.Count)
                                            oldRecord[s] = orgRecord[s];
                                        else
                                            oldRecord[s] = "";
                                    }

                                    if (oldRecord[1].Trim() != "")   // Alter
                                    {

                                        newRecord = new StringList();

                                        newRecord.Add(" ");
                                        newRecord.Add(tarifName);
                                        newRecord.Add(stand);
                                        newRecord.Add("M");
                                        newRecord.Add(oldRecord[1]);  // Alter
                                        newRecord.Add(oldRecord[2]);  // Beitrag
                                        newRecord.Add(oldRecord[4]);  // KEA
                                        newRecord.Add(oldRecord[6]);  // KMA
                                        newRecord.Add(oldRecord[8]);  // KMV

                                        dataTarget.Add(newRecord);

                                        newRecord = new StringList();

                                        newRecord.Add(" ");
                                        newRecord.Add(tarifName);
                                        newRecord.Add(stand);
                                        newRecord.Add("W");
                                        newRecord.Add(oldRecord[1]);  // Alter
                                        newRecord.Add(oldRecord[3]);  // Beitrag
                                        newRecord.Add(oldRecord[5]);  // KEA
                                        newRecord.Add(oldRecord[7]);  // KMA
                                        newRecord.Add(oldRecord[9]);  // KMV

                                        dataTarget.Add(newRecord);
                                    }
                                }
                            }
                        }
                    }
                    catch(Exception se)
                    {
                        Debug.Write("fehler " + filesXLS[f]+"  "+se.Message);
                    }
                    Debug.Write(f.ToString());
                }
                excelApplicationExtended.close();

                ExcelUtility.saveToCSV(pathName + "\\output.txt", dataTarget);

            }
        }



    }
}
/*
Sub alle()
Dim zelle As Range
On Error Resume Next
    For Each Worksheet In ThisWorkbook.Sheets
        For Each zelle In Worksheet.UsedRange
            If zelle <> "" Then zelle = zelle * 1
            If Left(zelle, 1) = "'" Then
               zelle = Right(zelle, Len(zelle) - 1)
            End If
        Next zelle
    Next Worksheet
    
End Sub
 * 
 * 
Sub alle()
Dim zelle As Range
On Error Resume Next
    For Each Worksheet In ThisWorkbook.Sheets
        For Each zelle In Worksheet.UsedRange
            'If zelle <> "" Then zelle = zelle * 1
            If zelle.PrefixCharacter = "'" Then
               zelle = Right(zelle, Len(zelle) - 1)
            End If
        Next zelle
    Next Worksheet
    
End Sub

*/