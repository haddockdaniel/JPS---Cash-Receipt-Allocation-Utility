using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties
        //152557.82
        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
                tbDepDate.Text = DateTime.Now.Date.ToShortDateString();
                tbChkDate.Text = DateTime.Now.Date.ToShortDateString();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            Cursor.Current = Cursors.WaitCursor;         
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch...", 1, 10); 
            Application.DoEvents();

            string BatRecAmount = tbChkAmt.Text;
            string SQLC = "select max(case when spname='CurAcctPrdYear' then cast(spnbrvalue as varchar(4)) else '' end) as PrdYear, max(Case when spname = 'CurAcctPrdNbr' then case " +
                " when spnbrvalue<9 then '0' + cast(spnbrvalue as varchar(1)) else cast(spnbrvalue as varchar(2)) end  else '' end) as PrdNbr from sysparam";
            DataSet myRSSysParm = _jurisUtility.RecordsetFromSQL(SQLC);

            DataTable dtSP = myRSSysParm.Tables[0];

            if (dtSP.Rows.Count == 0)
            { MessageBox.Show("Incorrect SysParams"); }
            else
            {
                foreach (DataRow dr in dtSP.Rows)
                {
                    PYear = dr["PrdYear"].ToString();
                    PNbr = dr["PrdNbr"].ToString();

                }
            }

                    //DataTable d1 = (DataTable)dataGridView1.DataSource;

                    string SQL = "select crbbatchnbr as Batch, crbcomment, crbstatus, crbdateentered, crbreccount, crbbatchtotal from cashreceiptsbatch where crbstatus NOT IN ('P','D') and crbcomment like 'JPS-Cash Alloc Tool%' and convert(varchar(10),crbdateentered,101) =convert(varchar(10),convert(date,getdate()),101)  ";
            DataSet myRSBatch = _jurisUtility.RecordsetFromSQL(SQL);

            DataTable dtBatch = myRSBatch.Tables[0];
            if (dtBatch.Rows.Count == 0)
            { CreateBatch(dtBatch); }
            else
            {
                DialogResult dg = MessageBox.Show("An unposted cash receipts batch created by this utility for today already exists. Would you like to add a new record to the existing batch?  If you select no, you will need to review the existing batch and delete or post", "Batch Already Exists", MessageBoxButtons.YesNo);
                if (dg == DialogResult.Yes)
                { 
                    foreach (DataRow dr in dtBatch.Rows)
                    {
                        singleBatch = dr["Batch"].ToString();
                        string BatchTotal = dr["crbbatchtotal"].ToString();
                        string RecTotal = dr["crbreccount"].ToString();
                        string s1 = "update cashreceiptsbatch set crbreccount=(cast'" + RecTotal + "' as int)  + 1, crbbatchtotal=(cast'" + BatchTotal +  "' as money)  + cast('" + BatRecAmount + "' as money) where crbbatchnbr=" + singleBatch;
                        _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                    }
            }
            }

            CreateBatchRecord();
            CreateBatchAR();
            CreateBatchPPD();        
            CreateBatchTrust();         
            CreateBatchOther();    
            CreateBatchFees();       
            CreateBatchExps();
            
            Cursor.Current = Cursors.Default;
            toolStripStatusLabel.Text = "Utility Completed.";
            statusStrip.Refresh();
            UpdateStatus("Utility Completed.", 1, 1);
            Application.DoEvents();

            string sqlTbl = "drop Table #ARAlloc";
            _jurisUtility.ExecuteNonQueryCommand(0, sqlTbl);

            string cmt = Application.ProductName.ToString();
            WriteLog(cmt);

            MessageBox.Show("Cash Receipt Allocation Completed.");
         
        }


        private void CreateBatch(DataTable dBat)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch...", 1, 10);
            Application.DoEvents();

            string BatRecAmount = tbChkAmt.Text;
            string BatDepDate = tbDepDate.Text;
            string MYFolder = PYear + "-" + PNbr;

            if (dBat.Rows.Count == 0)
            {
                string SQL = "Insert into CashReceiptsBatch(crbbatchnbr, crbcomment, crbstatus, crbreccount, crbenteredby,crbdateentered, crbpostedby, crbdateposted, crbbatchtotal)" +
                    " Values( (select spnbrvalue from sysparam where spname='LastBatchCash') + 1 ,'JPS-Cash Alloc Tool-' + '" + BatDepDate + "' ," +
                    "'U' , 1 , 1 ,convert(varchar(10),getdate(),101) , 1 , convert(varchar(10),getdate(),101) , cast('" + BatRecAmount + "' as money) )";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                SQL = "Update sysparam set spnbrvalue=spnbrvalue + 1 where spname='LastBatchCash'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                SQL = "select max(case when spname='CurAcctPrdYear' then cast(spnbrvalue as varchar(4)) else '' end) as PrdYear, " +
                   "max(Case when spname='CurAcctPrdNbr' then case when spnbrvalue<9 then '0' + cast(spnbrvalue as varchar(1)) else cast(spnbrvalue as varchar(2)) end  else '' end) as PrdNbr, " +
                   "max(case when spname='LastSysNbrDocTree' then spnbrvalue else 0 end) as DTree from sysparam";
                DataSet myRSSysParm = _jurisUtility.RecordsetFromSQL(SQL);

                DataTable dtSP = myRSSysParm.Tables[0];

                if (dtSP.Rows.Count == 0)
                { MessageBox.Show("Incorrect SysParams"); }
                else
                {
                    foreach (DataRow dr in dtSP.Rows)
                    {
                        string LastSys = dr["DTree"].ToString();

                        string SPSql = "Select dtdocid from documenttree where dtparentid=35 and dtdocclass='5300' and dttitle='" + MYFolder + "'";
                        DataSet spMY = _jurisUtility.RecordsetFromSQL(SPSql);
                        DataTable dtMY = spMY.Tables[0];
                        if (dtMY.Rows.Count == 0)
                        {
                            string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                  "select (select max(dtdocid)  + 1, 'Y', 5300,'F', 35,'" + MYFolder + "' from documenttree ";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'F', dtdocid,'SMGR'" +
                                " from documenttree where dtparentid=35 and dttitle='" + MYFolder + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " + 
                                " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR')," +
                                "'JPS-Cash Alloc Tool-' + '" + BatDepDate + "', " +
                                "(select  crbbatchnbr from cashreceiptsbatch where crbcomment like 'JPS-Cash Alloc Tool-' + '" + BatDepDate + "' and crbstatus not in ('D','P') " +
                                " and convert(datetime,crbdateentered,101)=convert(datetime,convert(date,getdate()),101)) from documenttree  where dtparentid=(select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR')";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);


                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);
                        }
                        else
                        {
                            string SMGRSql = "Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR'";
                            DataSet spSMGR = _jurisUtility.RecordsetFromSQL(SMGRSql);
                            DataTable dtSMGR = spSMGR.Tables[0];
                            if (dtSMGR.Rows.Count == 0)
                            {
                                string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                               "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'F', dtdocid,'SMGR'" +
                               " from documenttree where dtparentid=35 and dttitle='" + MYFolder + "'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                    "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                    " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR')," +
                                    "'JPS-Cash Alloc Tool-' + '" + BatDepDate + "', " +
                                    "(select crbbatchnbr from cashreceiptsbatch where crbcomment like'JPS-Cash Alloc Tool-' + '" + BatDepDate + "' and crbstatus not in ('D','P') " +
                                    " and convert(datetime,crbdateentered,101)=convert(datetime,convert(date,getdate()),101)) from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            }
                            else
                            {
                                string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                    "select (select max(dtdocid) from documenttree) + 1, 'Y', 5300,'R', " +
                                    " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR')," +
                                    "'JPS-Cash Alloc Tool-' + '" + BatDepDate + "', " +
                                    "(select crbbatchnbr from cashreceiptsbatch where crbcomment like 'JPS-Cash Alloc Tool-' + '" + BatDepDate + "' and crbstatus not in ('D','P') " +
                                    " and convert(datetime,crbdateentered,101)=convert(datetime,convert(date,getdate()),101)) from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=35 and dttitle='" + MYFolder + "') and dttitle='SMGR'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            }
                        }
                    }
                }
            }
            string sqlB = "select crbbatchnbr from cashreceiptsbatch where crbcomment='JPS-Cash Alloc Tool-' + '" + BatDepDate + "' and crbstatus='U' and crbreccount=1 " +
                " and convert(varchar(10),crbdateentered,101) =convert(varchar(10),getdate(),101)  and crbbatchtotal= cast('" + BatRecAmount + "' as money) ";
            DataSet spBatch = _jurisUtility.RecordsetFromSQL(sqlB);
            DataTable dtB = spBatch.Tables[0];
            if (dtB.Rows.Count == 0)
            { MessageBox.Show("Error Creating Cash Receipt Batch"); }
            else
            {
                foreach (DataRow dr in dtB.Rows)
                {
                    singleBatch = dr["crbbatchnbr"].ToString();
                 }

            }
        }


        private void CreateBatchRecord()
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Record...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch Record...", 2, 10);
            Application.DoEvents();

            string SQL = "Insert into cashreceipt(crbatch, crrecnbr,crposted, crdate, crprdyear, crprdnbr, crchecknbr, crcheckdate, crcheckamt, crpayor, crarcsh, crppdcsh, crtrustcsh, crnonclicsh)" +
                "select crbbatchnbr, crbreccount,'N',convert(datetime,'" + depDate + "',101) ," + PYear + "," + PNbr + ",'" + CkNbr + "','" + CkDate + "',cast('" + singleCheck + "' as money),'" + Payor +
              "',cast('" + singleAR + "' as money),cast('" + singlePPD + "' as money),cast('" + singleTrust + "' as money),cast('" + singleOther + "' as money) " +
              "from cashreceiptsbatch where crbbatchnbr=" + singleBatch + " and crbbatchnbr not in (select crbatch from cashreceipt)";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


        }

        private void CreateBatchAR()
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch AR...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch AR...", 3, 10);
            Application.DoEvents();

            string SQL = "Insert into CRARAlloc(crabatch, crarecnbr, cramatter, crabillnbr, cradate, crachecknbr, cracheckdate, crapayor, crafeeamt, cracshexpamt, crancshexpamt, craprepostfee, craprepostcshexp, craprepostncshexp, crabank, crasurchgamt , cratax1amt, cratax2amt, cratax3amt, crainterestamt, craprepostsurchg , crapreposttax1, crapreposttax2, crapreposttax3, craprepostinterest) " +
                " Select  crbbatchnbr, crbreccount, matter, billnbr,convert(datetime,'" + depDate + "',101) ,'" + CkNbr + "','" + CkDate + "','" + Payor + "',fees, cshexp, ncshexp, feeAR, cshAR, ncshAR, ofcbankcode,0,0,0,0,0,0,0,0,0,0 " +
                " from (select matter, billnbr, sum(case when itype='Fee' then allocamt else 0 end) as fees, sum(case when ITYpe='Cost' and exptype='C' then allocamt else 0 end) as cshexp, sum(case when ITYpe='Cost' and exptype='N' then allocamt else 0 end) as NCshExp " + 
                " from #ARAlloc group by matter, billnbr) AR " +
                "inner join matter on matsysnbr=matter inner join officecode on matofficecode=ofcofficecode " + 
                " Inner join (select armmatter, armbillnbr, sum(armfeebld - armfeercvd + armfeeadj) as FeeAR, sum(armcshexpbld - armcshexprcvd + armcshexpadj) as CshAR,sum(armncshexpbld - armncshexprcvd + armncshexpadj) as ncshAR " +
                " from armatalloc group by armmatter, armbillnbr) ARM on matter=armmatter and billnbr=armbillnbr, cashreceiptsbatch where crbbatchnbr=" + singleBatch;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


            SQL = "Update ARMatalloc set armpendfee= crafeeamt, armpendcshexp=cracshexpamt, armpendncshexp=crancshexpamt " +
                " from (select distinct cramatter, crabillnbr, crafeeamt, cracshexpamt, crancshexpamt from craralloc " +
                " inner join #ARAlloc on matter=cramatter and billnbr=crabillnbr)CR where cramatter=armmatter and crabillnbr=armbillnbr";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


        }

        private void CreateBatchPPD()
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch PPD...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch PPD...", 4, 10);
            Application.DoEvents();

            if (radioButton2.Checked == true)
            { singleClient = this.cbClient.GetItemText(this.cbClient.SelectedItem).Split(' ')[0];
              singleMatter = this.cbMatter.GetItemText(this.cbMatter.SelectedItem).Split(' ')[0];

                string SQL = "Insert into CRPPDAlloc(crpbatch, crprecnbr, crpmatter, crpamount) " +
                     "select crbbatchnbr, crbreccount, matsysnbr," + singlePPD + " from matter inner join client on matclinbr=clisysnbr, cashreceiptsbatch " +
                     "where dbo.jfn_formatclientcode in ('" + singleClient + "') and dbo.jfn_formatmattercode(matcode) in ('" + singleMatter + "') and crbbatchnbr=" + singleBatch;

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
             }

        private void CreateBatchTrust()
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Trust...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch Trust...", 5, 10);
            Application.DoEvents();

            if (radioButton3.Checked == true)
            {
                singleClient = this.cbClient.GetItemText(this.cbClient.SelectedItem).Split(' ')[0];
                singleMatter = this.cbMatter.GetItemText(this.cbMatter.SelectedItem).Split(' ')[0];
                singleBank = this.cbBank.GetItemText(this.cbBank.SelectedItem).Split(' ')[0];
                
                string SQL = "Insert into CRTrustAlloc(crtbatch, crtrecnbr, crtseqnbr, crtmatter, crtbank, crtamount) " +
                     "select crbbatchnbr, crbreccount,1, matsysnbr,'" + singleBank + "'," + singleTrust + " from matter inner join client on matclinbr=clisysnbr, cashreceiptsbatch " +
                     "where dbo.jfn_formatclientcode in ('" + singleClient + "') and dbo.jfn_formatmattercode(matcode) in ('" + singleMatter + "') and crbbatchnbr=" + singleBatch;

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
        }

        private void CreateBatchOther()
        {   Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Non-Client...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt  Batch Non-Client...", 5, 10);
        Application.DoEvents();
            if (radioButton4.Checked == true)
            {
                singleBank = this.cbBank.GetItemText(this.cbBank.SelectedItem).Split(' ')[0];
                singleGL = this.cbGL.GetItemText(this.cbGL.SelectedItem).Split(' ')[0];

             //   string SQL = "Insert into CRNonCliAlloc(crnbatch, crnrecnbr, crnseqnbr,crnbank, crncreditaccount, crnreference, crnamount) " +
         //   "select crbbatchnbr, crbreccount,1, '" + singleBank + "',(select chtsysnbr from chartofaccounts where dbo.jnf_formatchartofaccount('" + singleGL + "')," + singleOther + "  +" +
         //   "from cashreceiptsbatch " +
          //  "where  crbbatchnbr=" + singleBatch;

                string SQL = "Insert into CRNonCliAlloc(crnbatch, crnrecnbr, crnseqnbr,CRNBankCode, crncreditaccount, crnreference, crnamount) " +
                            " select crbbatchnbr, crbreccount,1, '" + singleBank + "',(select chtsysnbr from chartofaccounts where dbo.jfn_formatchartofaccount(Chtsysnbr) ='" + singleGL + "'),'" + tbReference.Text + "', " + singleOther + 
                            " from cashreceiptsbatch " +
                            " where  crbbatchnbr=" + singleBatch; // add reference text

                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }
        }

        private void CreateBatchFees()
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Fees...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch Fees...", 6, 10);
            Application.DoEvents();


            string SQL = "Insert into CRFeeAlloc(crfbatch, crfrecnbr, crfmatter, crfbillnbr, crftkpr, crftaskcd, crfactivitycd, crfprepost, crfamount) " +
                " Select  crbbatchnbr, crbreccount, matter, billnbr,tkpr, case when taskcd='' then null else taskcd end, case when actcd='' then null else actcd end, prepost, amt " +
                " from (select matter, billnbr, tkpr, isnull(taskcd,'') as taskcd, isnull(actcd,'') as actcd, sum(allocAmt) as amt from #ARAlloc where IType='Fee' group by matter, billnbr, tkpr, isnull(taskcd,''), isnull(actcd,'')) AR " +
                " Inner join (select arftmatter, arftbillnbr, arfttkpr, isnull(arfttaskcd,'') as ARFTTask, isnull(arftactivitycd,'') as ActivityCd, sum(arftactualamtbld - arftrcvd + arftadj) as Prepost " +
                " from arftaskalloc group by arftmatter, arftbillnbr, arfttkpr, isnull(arfttaskcd,''), isnull(arftactivitycd,'')) ARM on matter=arftmatter and billnbr=arftbillnbr and arfttkpr=tkpr and taskcd=arfttask and actcd=activitycd, cashreceiptsbatch " +
                " where crbbatchnbr=" + singleBatch;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


            SQL = "Update arftaskalloc set arftpend=crfamount " +
                " from crfeealloc where crfbatch=" + singleBatch + " and arftmatter=crfmatter and arftbillnbr=crfbillnbr and crftkpr=arfttkpr and isnull(crftaskcd,'')=isnull(arfttaskcd,'') " +
                "  and isnull(crfactivitycd,'')=isnull(arftactivitycd,'')";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

        }

        private void CreateBatchExps()
        { Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Cash Receipt Batch Exps...";
            statusStrip.Refresh();
            UpdateStatus("Creating Cash Receipt Batch Exps...", 8, 10);
            Application.DoEvents();


            string SQL = "Insert into CRExpAlloc(crebatch, crerecnbr, crematter, crebillnbr, creexpcd, creexptype, creprepost, creamount) " +
                " Select  crbbatchnbr, crbreccount, matter, billnbr,expcd, exptype, prepost, amt " +
                " from (select matter, billnbr,expcd, exptype, sum(allocAmt) as amt from #ARAlloc where IType='Cost' group by matter, billnbr,expcd, exptype) AR " +
                " Inner join (select arematter, arebillnbr, areexpcd, areexptype, sum(arebldamount - arercvd + areadj) as Prepost " +
                " from arexpalloc group by arematter, arebillnbr, areexpcd, areexptype) ARM on matter=arematter and billnbr=arebillnbr and areexpcd=expcd and areexptype=exptype, cashreceiptsbatch " + 
                " where crbbatchnbr=" + singleBatch;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);


            SQL = "Update arexpalloc set arepend=creamount " +
                " from crexpalloc where crebatch=" + singleBatch + " and crematter=arematter and arebillnbr=crebillnbr and areexpcd=creexpcd and areexptype=creexptype ";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

        }




        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion



        public string singleMatter = "";
        public string singleClient = "";
        public string singleBank = "";
        public string singleGL = "";
        public string singleBatch = "";
        public string singleCheck = "";
        public string singleAR = "";
        public string singleRemain= "0";
        public string singlePPD = "0";
        public string singleTrust = "0";
        public string singleOther = "0";
        public string PYear = "";
        public string PNbr = "";
        public string depDate = "";
        public string CkDate = "";
        public string Payor = "";
        public string CkNbr = "";

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName;
                fileName = dlg.FileName.ToString();
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;';");

                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);

                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                dataGridView1.DataSource = DtSet.Tables[0];
                MyConnection.Close();

                DataTable dt = DtSet.Tables[0];

                string depDate2 = dt.Rows[1]["Deposit_Date"].ToString();
                tbDepDate.Text = depDate2;
                string chkNbr = dt.Rows[1]["Check_Number"].ToString();
                tbChkNbr.Text = chkNbr;
                string chkDate = dt.Rows[1]["Check_Date"].ToString();
                tbChkDate.Text = chkDate;
                string Payor2 = dt.Rows[1]["Payor"].ToString();
                tbPayor.Text = Payor2;
                object ChkAmt = dt.Rows[1]["CheckAmount"];
                tbChkAmt.Text = String.Format("{0, 0:C2}", ChkAmt);
                object AllocAmt;
                AllocAmt = dt.Compute("Sum(Amt_to_Allocate)", "");
                tbAllocAmt.Text = String.Format("{0, 0:C2}", AllocAmt);
                Decimal ckA = Convert.ToDecimal(ChkAmt);
                Decimal AlA = Convert.ToDecimal(AllocAmt);
                Decimal RB = ckA - AlA;
                tbBalance.Text = String.Format("{0, 0:C2}", RB);
                if (RB < 0)
                {
                    tbBalance.ForeColor = System.Drawing.Color.Red;
                    MessageBox.Show("Allocations Exceed total Check Amount. Review spreadsheet and adjust allocations.");
                    button1.Enabled = false;
                }

                if (RB > 0)
                {
                    tbBalance.ForeColor = System.Drawing.Color.Green;
                    MessageBox.Show("Check Amount exceeds Allocations.  Select options for remaining balance.");
                    groupBox1.Enabled = true;
                    tbReference.Text = "Cash Allocation Tool";

                    string CliIndex;
                    cbClient.ClearItems();
                    string SQLCli = "select Client from ( select dbo.jfn_formatclientcode(clicode) + '   ' +  clireportingname as Client from Client) Emp order by Client";
                    DataSet myRSCli = _jurisUtility.RecordsetFromSQL(SQLCli);

                    if (myRSCli.Tables[0].Rows.Count == 0)
                        cbClient.SelectedIndex = 0;
                    else
                    {
                        foreach (DataTable table in myRSCli.Tables)
                        {

                            foreach (DataRow dr in table.Rows)
                            {
                                CliIndex = dr["Client"].ToString();
                                cbClient.Items.Add(CliIndex);
                            }
                        }

                    }


                }


                singleCheck = tbChkAmt.Text;
                singleAR = tbAllocAmt.Text;
                singleRemain = tbBalance.Text;
                depDate = tbDepDate.Text;
                CkDate = tbChkDate.Text;
                Payor = tbPayor.Text;
                CkNbr = tbChkNbr.Text;

                string sqlTbl = "create Table #ARAlloc(billnbr int, matter int, Itype varchar(10), tkpr int, taskcd varchar(4), actcd varchar(4), expcd varchar(4), exptype varchar(1), allocamt money)";
                _jurisUtility.ExecuteNonQueryCommand(0, sqlTbl);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string BillNbr = dt.Rows[i]["Invoice_Number"].ToString();
                    string Matter = dt.Rows[i]["Matter_Sys"].ToString();
                    string Tkpr = dt.Rows[i]["Tkpr_Sys"].ToString();
                    string TC = dt.Rows[i]["Task_Code"].ToString();
                    string IType = dt.Rows[i]["Item_Type"].ToString();
                    string AC = dt.Rows[i]["Activity_Code"].ToString();
                    string EC = dt.Rows[i]["Expense_Code"].ToString();
                    string ET = dt.Rows[i]["ExpenseType"].ToString();
                    string AA = dt.Rows[i]["Amt_to_Allocate"].ToString();

                    string sqlIns = "Insert into #ARAlloc Values(cast('" + BillNbr + "' as int),cast('" + Matter + "' as int),'" + IType + "',case when '" + Tkpr + "'='' then 0 else cast('" + Tkpr + "' as int) end,'" + TC + "','" + AC + "','" + EC + "','" + ET + "',cast('" + AA + "' as money))";
                    
                    _jurisUtility.ExecuteNonQueryCommand(0, sqlIns);
                }

            }

 }



        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                singlePPD = singleRemain;
                cbClient.Enabled = true;
                cbMatter.Enabled = true;
                cbBank.Enabled = false;
                cbGL.Enabled = false;

                lblClient.Enabled = true;
                lblMatter.Enabled = true;
                lblBank.Enabled = false;
                lblGL.Enabled = false;

                string CliIndex;
                cbClient.ClearItems();
                string SQLCli = "select Client from ( select dbo.jfn_formatclientcode(clicode) + '   ' +  clireportingname as Client from Client) Emp order by Client";
                DataSet myRSCli = _jurisUtility.RecordsetFromSQL(SQLCli);

                if (myRSCli.Tables[0].Rows.Count == 0)
                    cbClient.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSCli.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            CliIndex = dr["Client"].ToString();
                            cbClient.Items.Add(CliIndex);
                        }
                    }

                }
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton3.Checked == true)
            {
                singleTrust = singleRemain;
                cbClient.Enabled = true;
                cbMatter.Enabled = true;
                cbBank.Enabled = true;
                cbGL.Enabled = false;

                lblClient.Enabled = true;
                lblMatter.Enabled = true;
                lblBank.Enabled = true;
                lblGL.Enabled = false;


            }
        }

        private void cbClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbMatter.Enabled = true;
            singleClient = this.cbClient.GetItemText(this.cbClient.SelectedItem).Split(' ')[0];
            cbMatter.SelectedIndex = -1;
            string MatIndex;
            cbMatter.ClearItems();
            string SQLMat = " select dbo.jfn_formatmattercode(matcode) + '   ' +  matreportingname as Matter from Matter inner join client on matclinbr=clisysnbr where dbo.jfn_Formatclientcode(Clicode)='" + singleClient.ToString() + "' order by  dbo.jfn_formatmattercode(matcode)";
            DataSet myRSMat = _jurisUtility.RecordsetFromSQL(SQLMat);

            if (myRSMat.Tables[0].Rows.Count == 0)
                cbMatter.SelectedIndex = 0;
            else
            {
                foreach (DataTable table in myRSMat.Tables)
                {

                    foreach (DataRow dr in table.Rows)
                    {
                        MatIndex = dr["Matter"].ToString();
                        cbMatter.Items.Add(MatIndex);
                    }
                }

            }
        }

        private void cbMatter_SelectedIndexChanged(object sender, EventArgs e)
        {
            singleMatter = this.cbMatter.GetItemText(this.cbMatter.SelectedItem).Split(' ')[0];

            if (radioButton3.Checked == true)
            {
                cbBank.Enabled = true;
                cbBank.SelectedIndex = -1;
                string BnkIndex;
                cbBank.ClearItems();
                string SQLBank = " select bnkcode + '   ' +  bnkdesc as Bank from BankAccount where bnkaccttype='T'";
                DataSet myRSBank = _jurisUtility.RecordsetFromSQL(SQLBank);

                if (myRSBank.Tables[0].Rows.Count == 0)
                    cbBank.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSBank.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            BnkIndex = dr["Bank"].ToString();
                            cbMatter.Items.Add(BnkIndex);
                        }
                    }

                }
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
          
            if (radioButton4.Checked == true)
            {
                singleOther = singleRemain;
                cbClient.Enabled = false;
                cbMatter.Enabled = false;
                cbBank.Enabled = true;
                cbGL.Enabled = true;
                tbReference.Enabled = true;
                lblRef.Enabled = true;
                lblClient.Enabled = false;
                lblMatter.Enabled = false;
                lblBank.Enabled = true;
                lblGL.Enabled = true;

              
                cbBank.SelectedIndex = -1;
                string BnkIndex;
                cbBank.ClearItems();
                string SQLBank = " select bnkcode + '   ' +  bnkdesc as Bank from BankAccount where bnkaccttype='T'";
                DataSet myRSBank = _jurisUtility.RecordsetFromSQL(SQLBank);

                if (myRSBank.Tables[0].Rows.Count == 0)
                    cbBank.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSBank.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            BnkIndex = dr["Bank"].ToString();
                            cbBank.Items.Add(BnkIndex);
                        }
                    }

                }
            }
        }

        private void cbBank_SelectedIndexChanged(object sender, EventArgs e)
        {
            singleBank = this.cbBank.GetItemText(this.cbBank.SelectedItem).Split(' ')[0];

            if (radioButton4.Checked == true)
            {
                
                lblGL.Enabled = false;
                cbGL.Enabled = true;
                cbGL.SelectedIndex = -1;
                string GLIndex;
                cbGL.ClearItems();
                string SQLGL = " select dbo.jfn_formatchartofaccount(chtsysnbr)  + '   ' +  chtdesc as ChartAcct from Chartofaccounts where chtactive='Y'";
                DataSet myRSGL = _jurisUtility.RecordsetFromSQL(SQLGL);

                if (myRSGL.Tables[0].Rows.Count == 0)
                    cbGL.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in myRSGL.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            GLIndex = dr["ChartAcct"].ToString();
                            cbGL.Items.Add(GLIndex);
                        }
                    }

                }
            }

        }

        private void cbGL_SelectedIndexChanged(object sender, EventArgs e)
        {
            singleGL = this.cbGL.GetItemText(this.cbGL.SelectedItem).Split(' ')[0];
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            cbClient.Enabled = false;
            cbMatter.Enabled = false;
            cbBank.Enabled = false;
            cbGL.Enabled = false;
            lblClient.Enabled = false;
            lblMatter.Enabled = false;
            lblBank.Enabled = false;
            lblGL.Enabled = false;
        }
    }
}
