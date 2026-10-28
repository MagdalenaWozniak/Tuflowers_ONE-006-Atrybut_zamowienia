using Hydra;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using ZZnaCSV;
using System.Collections.Generic;
using static ZZnaCSV.ZS;
using System.Text;
using System;
using System.Data.SqlClient;
using System.Data;
using cdn_api;

[assembly: CallbackAssemblyDescription(Name: "ONE-006 Atrybut zamówienia",
                                           Description: "Dodanie atrybutu EksportZSMobile na dokumencie ZS",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2021.0.2",
                                           Created: "23-12-2021")]


namespace ZZnaCSV
{
    [SubscribeProcedure(Procedures.ZaNLista, "Atrybut na ZS")]
    public class ZSnaFS_Class : Callback
    {
        private int XLWersjaAPI = 20202;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "192.168.1.202";
        private string XLSQLDatabase = "Tuflowers-Mob";//"ERPXL_PoloMarket"
        private string XLUser = "ADMIN"; //
        private string XLPass = "NE!w0rk12!";//"Dupa2003!";
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        private List<ListaZS> ZSy = new List<ListaZS>();
        int listID;
        public override void Init()
        {

            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }



        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {

            ClaWindow Parent = GetWindow();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?TabOdb"];
            RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            //RapButton.TextRaw = "ph";
            RapButton.ToolTipRaw = "Dodanie atrybutu EksportZSMobile";
            RapButton.IconRaw = "C:\\OneSupport\\atrybut.ico";

            filtrBtn = GetWindow().AllChildren["?BRW1::TotalWorking"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 23, filtrBtn.Bounds.Y, 20, 15);
            listID = GetWindow().AllChildren["?ListaZam"].Id;      //tu sprawdzic

            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);

            //pobrane z pola w zakładce Raporty
            return true;
        }

        public bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?BRW1::TotalWorking"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 23, filtrBtn.Bounds.Y, 20, 15);

            return true;
        }

        public bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {

            try
            {

                ADODB._Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcId, listID, GetCallbackThread());
                // if(recordset!=null)
                // MessageBox.Show(Convert.ToString(listID), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ZbieranieZS(recordset);
                //proceduraZS();
                if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 22) == "ERPXL_Tuflowers_Mobile")
                    XLSQLDatabase = "Tuflowers-Mob";
                else if(Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 15) == "ERPXL_Tuflowers")
                    XLSQLDatabase = "Tuflowers";

                XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
                nowyUzytkownik.Wersja = XLWersjaAPI;
                nowyUzytkownik.Winieta = XLWinieta;
                nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
                nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
                nowyUzytkownik.SerwerKlucza = XLSQLServer;
                nowyUzytkownik.Baza = XLSQLDatabase;
                nowyUzytkownik.OpeIdent = XLUser;
                nowyUzytkownik.OpeHaslo = XLPass;
                nowyUzytkownik.PortHTTPSerweraKlucza = 1;
                nowyUzytkownik.RodzajSerweraKlucza = 0;
                nowyUzytkownik.SerwerKlucza = "TF-SQL::5000185869";
                nowyUzytkownik.ProgramID = "ONE-006 Atrybut zamówienia";

                int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

                if (blad != 0)
                {
                    MessageBox.Show("Błąd logowania: "+blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

               

                foreach (var item in ZSy)
                 {
                    XLAtrybutInfo_20202 atrybut5 = new XLAtrybutInfo_20202();
                    atrybut5.Wersja = XLWersjaAPI;
                    atrybut5.GIDFirma = item.GIDFirma;
                    atrybut5.GIDLp = item.GIDLp;
                    atrybut5.GIDNumer = item.GIDNumer;
                    atrybut5.GIDSubLp = 0;
                    atrybut5.GIDTyp = item.GIDTyp;
                    atrybut5.Wartosc = "Wyslij do ph";
                    atrybut5.Klasa = "EksportZSMobile";
                    int blad_atr=cdn_api.cdn_api.XLDodajAtrybut(sesja, atrybut5);

                    if(blad_atr!=0)
                    {
                        MessageBox.Show("Błąd atrybutu: " + blad_atr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    //MessageBox.Show("GIDTyp "+Convert.ToString(item.GIDTyp), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //MessageBox.Show("GIDFirma " + Convert.ToString(item.GIDFirma), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //MessageBox.Show("GIDNumer " + Convert.ToString(item.GIDNumer), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //MessageBox.Show("GIDLp " + Convert.ToString(item.GIDLp), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    
                }

                cdn_api.cdn_api.XLLogout(sesja);

            }
            catch { }
            ZSy.Clear();
            MessageBox.Show("Zakończono dodawanie atrybutów", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }

        bool ZbieranieZS(ADODB._Recordset recordset)
        {

            try
            {
                StringBuilder dokumenty = new StringBuilder();
                if (recordset != null)
                {
                    int count = recordset.RecordCount;

                    if (count != 0)
                    {
                        recordset.MoveFirst();

                        while (recordset.EOF == false)
                        {
                            ADODB.Fields fields = recordset.Fields;

                            ZSy.Add(new ListaZS()
                            {
                                GIDTyp = Convert.ToInt16(fields[0].Value),
                                GIDFirma = Convert.ToInt32(fields[1].Value),
                                GIDNumer = Convert.ToInt32(fields[2].Value),
                                GIDLp = Convert.ToInt16(fields[3].Value),
                            });
                            recordset.MoveNext();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nie zaznaczono dokumentów. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("Nie zaznaczono dokumentów. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd pobierania zaznaczonych ZS --> " + ex.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
        }

        void proceduraZS()
        {
            string connectionString = "";
            if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 22) == "ERPXL_Tuflowers_Mobile")
                connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_Mobile; User ID = onesupport; Password=Ancr3d0$$db";
                //connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_Mobile;persist security info=True; Integrated Security=SSPI;";
            if (connectionString != "")
            {
                foreach (var item in ZSy)
                {
                    string komunikat = "";
                    procedura(connectionString, item.GIDTyp, item.GIDNumer, ref komunikat);
                    MessageBox.Show(komunikat, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        void procedura(string connectionString, int GIDTyp, int GIDNumer, ref string komunikat)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("cdn.ONE_PrzeksztalcIndeksyNaZS_Z", conn);


                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add("@komunikat", SqlDbType.VarChar, 8000);
                    cmd.Parameters["@komunikat"].Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(new SqlParameter("@gidtyp", GIDTyp));
                    cmd.Parameters.Add(new SqlParameter("@gidnumer", GIDNumer));
                   // cmd.Parameters.Add(new SqlParameter("@komunikat", komunikat));
                    

                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results, printing each to console
                        while (rdr.Read())
                        {

                            //Console.WriteLine("Product: {0,-35} Total: {1,2}", rdr["ProductName"], rdr["Total"]);
                            //MessageBox.Show("W procedurze: "+Convert.ToString(rdr["komunikat"]), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    komunikat = Convert.ToString(cmd.Parameters["@komunikat"].Value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Złe wywołanie procedury: " + ex.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public override void Cleanup()
        {

        }

    }
}
