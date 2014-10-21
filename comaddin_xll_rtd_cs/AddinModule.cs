using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using AddinExpress.MSO;

namespace comaddin_xll_rtd_cs
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("1BCE9B7F-93B7-40F4-9787-66D960E99CD2"), ProgId("comaddin_xll_rtd_cs.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public AddinModule()
        {
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler
        }

        private AddinExpress.MSO.ADXExcelAppEvents adxExcelEvents;
 
        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;
 
        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.adxExcelEvents = new AddinExpress.MSO.ADXExcelAppEvents(this.components);
            // 
            // adxExcelEvents
            // 
            this.adxExcelEvents.WorkbookOpen += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxExcelEvents_WorkbookOpen);
            // 
            // AddinModule
            // 
            this.AddinName = "comaddin_xll_rtd_cs";
            this.HandleShortcuts = true;
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
            this.AddinStartupComplete += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinStartupComplete);

        }
        #endregion
 
        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

        private bool isInitialized = false;
        private MyCommonData commonData;

        private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
        {
            isInitialized = true;
            MessageBox.Show("Add In Initilaized");
            Log.WriteLine("COM add-in. AppDomain=" + AppDomain.CurrentDomain.FriendlyName);
            commonData = MyCommonData.CurrentInstance;
            if (!commonData.IsInitialized)
            {
                Log.WriteLine("COM add-in. Initializing common data...");
                commonData.Initialize();
            }
            Log.WriteLine("COM add-in. Common data are initialized.");
            Log.WriteLine("COM add-in. Checking other objects...");
            commonData.CheckObjects();
        }

        public bool IsInitialized
        {
            get { return isInitialized; }
        }

        private void adxExcelEvents_WorkbookOpen(object sender, object hostObj)
        {
            MessageBox.Show("Derpy derp");
        }

    }
}

