using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace comaddin_xll_rtd_cs
{
    /// <summary>
    ///   Add-in Express XLL Add-in Module
    /// </summary>
    [ComVisible(true)]
    public class XLLModule1 : AddinExpress.MSO.ADXXLLModule
    {
        public XLLModule1()
        {
            InitializeComponent();
        }
 
        #region Component Designer generated code
 
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
 
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            // 
            // XLLModule1
            // 
            this.AddinName = "comaddin_xll_rtd_cs";
            this.OnInitialize += new System.EventHandler(this.XLLModule1_OnInitialize);

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
        public static void RegisterXLL(Type t)
        {
            AddinExpress.MSO.ADXXLLModule.RegisterXLLInternal(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void UnregisterXLL(Type t)
        {
            AddinExpress.MSO.ADXXLLModule.UnregisterXLLInternal(t);
        }
 
        #endregion
 
        #region Define your UDFs in this section
 
        /// <summary>
        /// The container for user-defined functions (UDFs). Every UDF is a public static (Public Shared in VB.NET) method that returns a value of any base type: string, double, integer.
        /// </summary>
        internal static class XLLContainer
        {
            /// <summary>
            /// Required by Add-in Express. Please do not modify this method.
            /// </summary>
            internal static XLLModule1 Module
            {
                get
                {
                    return AddinExpress.MSO.ADXXLLModule.
                        CurrentInstance as comaddin_xll_rtd_cs.XLLModule1;
                }
            }
 
            #region Sample function
 
            // Demonstrates how to handle all parameter types available for UDFs.
            // Uncomment the code, click Register Add-in Express Project in the Build menu, and run Excel.

            //public static string AllSupportedExcelTypes(object arg)
            //{
            //    if (arg is double)
            //        return "Double: " + (double)arg;
            //    else if (arg is string)
            //        return "String: " + (string)arg;
            //    else if (arg is bool)
            //        return "Boolean: " + (bool)arg;
            //    else if (arg is AddinExpress.MSO.ADXExcelError)
            //        return "ExcelError: " + arg.ToString();
            //    else if (arg is object[,])
            //        return string.Format("Array[{0},{1}]", ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
            //    else if (arg is System.Reflection.Missing)
            //        return "Missing";
            //    else if (arg == null)
            //        return "Empty";
            //    else if (arg is AddinExpress.MSO.ADXExcelRef)
            //    {
            //        AddinExpress.MSO.ADXExcelRef reference = arg as AddinExpress.MSO.ADXExcelRef;
            //        return string.Format("Reference [{0},{1},{2},{3}]", reference.ColumnFirst, reference.RowFirst, reference.ColumnLast, reference.RowLast);
            //    }
            //    else if (arg is short)
            //        return "Short: " + (short)arg;
            //    else
            //        return "Unknown Type";
            //}
 
            #endregion

            public static string CallRTD()
            {
                if (Module.IsInFunctionWizard)
                {
                    return "This UDF calls an RTD server.";
                }
                return Module.CallWorksheetFunction(AddinExpress.MSO.ADXExcelWorksheetFunction.Rtd, new object[3] { "comaddin_xll_rtd_cs.RTDServerModule1", "", "test" }).ToString();
            }
        }
 
        #endregion

        bool isInitialized = false;
        private MyCommonData commonData;
        private void XLLModule1_OnInitialize(object sender, EventArgs e)
        {
            isInitialized = true;
            Log.WriteLine("XLL. AppDomain=" + AppDomain.CurrentDomain.FriendlyName);
            commonData = MyCommonData.CurrentInstance;
            if (!commonData.IsInitialized)
            {
                Log.WriteLine("XLL. Initializing common data...");
                commonData.Initialize();
            }
            Log.WriteLine("XLL. Common data are initialized.");
            Log.WriteLine("XLL. Checking other objects...");
            commonData.CheckObjects();
        }

        public bool IsInitialized
        {
            get { return isInitialized; }
        }

    }
}

