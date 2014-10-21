using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace comaddin_xll_rtd_cs
{
    /// <summary>
    ///   Add-in Express RTD Server Module
    /// </summary>
    [GuidAttribute("BC3AF8A4-D3F0-4ED7-A1F6-64E837E7B77B"), ProgId("comaddin_xll_rtd_cs.RTDServerModule1")]
    public class RTDServerModule1 : AddinExpress.RTD.ADXRTDServerModule
    {
        public RTDServerModule1()
        {
            InitializeComponent();
        }

        private AddinExpress.RTD.ADXRTDTopic adxrtdTopic1;
 
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
            this.adxrtdTopic1 = new AddinExpress.RTD.ADXRTDTopic(this.components);
            // 
            // adxrtdTopic1
            // 
            this.adxrtdTopic1.String01 = "test";
            this.adxrtdTopic1.Tag = "";
            this.adxrtdTopic1.RefreshData += new AddinExpress.RTD.ADXRefreshData_EventHandler(this.adxrtdTopic1_RefreshData);
            // 
            // RTDServerModule1
            // 
            this.RTDInitialize += new System.EventHandler(this.RTDServerModule1_RTDInitialize);

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
        public static void RTDServerRegister(Type t)
        {
            AddinExpress.RTD.ADXRTDServerModule.ADXRTDServerRegister(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void RTDServerUnregister(Type t)
        {
            AddinExpress.RTD.ADXRTDServerModule.ADXRTDServerUnregister(t);
        }
 
        #endregion

        private object adxrtdTopic1_RefreshData(object sender)
        {
            Random rnd = new Random();
            return rnd.Next(1000);
        }

        bool isInitialized = false;
        private MyCommonData commonData;
        private void RTDServerModule1_RTDInitialize(object sender, EventArgs e)
        {
            isInitialized = true;
            Log.WriteLine("RTD. AppDomain=" + AppDomain.CurrentDomain.FriendlyName);
            commonData = MyCommonData.CurrentInstance;
            if (!commonData.IsInitialized)
            {
                Log.WriteLine("RTD. Initializing common data...");
                commonData.Initialize();
            }
            Log.WriteLine("RTD. Common data are initialized.");
            Log.WriteLine("RTD. Checking other objects...");
            commonData.CheckObjects();
        }

        public bool IsInitialized
        {
            get { return isInitialized; }
        }

    }
}

