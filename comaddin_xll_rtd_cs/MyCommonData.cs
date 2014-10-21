using System;
using System.Collections.Generic;
using System.Text;

namespace comaddin_xll_rtd_cs
{
    class MyCommonData
    {
        private static MyCommonData currentInstance;
        private MyCommonData()
        {
            currentInstance = this;
        }

        public static MyCommonData CurrentInstance
        {
            get 
            { 
                if (currentInstance == null) currentInstance = new MyCommonData();
                return currentInstance;
            } 
        }

        private bool isInitialized;
        public bool IsInitialized
        {
            get { return isInitialized; }
        }

        public void Initialize()
        {
            if (!isInitialized)
            {
                isInitialized = true;
            }
        }

        public void CheckObjects()
        {
            Log.WriteLine("  MyCommonData. AppDomain=" + AppDomain.CurrentDomain.FriendlyName);

            comaddin_xll_rtd_cs.AddinModule addinModule = AddinExpress.MSO.ADXAddinModule.CurrentInstance as comaddin_xll_rtd_cs.AddinModule;
            if (addinModule != null)
                Log.WriteLine("  MyCommonData. addinModule.IsInitialized=" + addinModule.IsInitialized.ToString());
            else
                Log.WriteLine("  MyCommonData. addinModule not created");

            comaddin_xll_rtd_cs.XLLModule1 xllModule = AddinExpress.MSO.ADXXLLModule.CurrentInstance as comaddin_xll_rtd_cs.XLLModule1;
            if (xllModule != null)
                Log.WriteLine("  MyCommonData. xllModule.IsInitialized=" + xllModule.IsInitialized.ToString());
            else
                Log.WriteLine("  MyCommonData. xllModule not created");

            comaddin_xll_rtd_cs.RTDServerModule1 rtdModule = AddinExpress.RTD.ADXRTDServerModule.CurrentInstance as comaddin_xll_rtd_cs.RTDServerModule1;
            if (rtdModule != null)
                Log.WriteLine("  MyCommonData. rtdModule.IsInitialized=" + rtdModule.IsInitialized.ToString());
            else
                Log.WriteLine("  MyCommonData. rtdModule not created");
        }
    }
}
