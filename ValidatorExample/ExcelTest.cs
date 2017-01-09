// <copyright file="ExcelTest.cs" company="More4Apps">
// Copyright (c) More4Apps. All rights reserved.
// </copyright>
namespace UnitTest
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Vbe.Interop;
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Unit Testing Class
    /// Procedures for interfacing unit tests with Excel VBA modules/classes
    /// </summary>
    public class ExcelTest : IDisposable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelTest"/> class
        /// </summary>
        /// <param name="sClass">Validator Class Name</param>
        public ExcelTest(string sClass = null)
        {
            // Grabs the active Excel application here
            // This could be enhanced in the future to specify an Excel file
            this.ExcelApp = new Excel.Application();
            this.ExcelApp = (Excel.Application)Marshal.GetActiveObject("excel.application");
            this.ExcelApp.Visible = true;
            this.ExcelApp.DisplayAlerts = true;

            // Bevan - Initialise our class if a variable is passed in
            if (!string.IsNullOrEmpty(sClass)) {
                this.Class = GetVBAClass(this.ExcelApp, sClass);
            }
        }

        /// <summary>
        /// Gets or sets the ExcelApp Object
        /// Call the _run2 function to run a function
        /// </summary>
        public Excel.Application ExcelApp { get; set; } // Excel application link

        /// <summary>
        /// Gets or sets the Class (if used)
        /// Need to test this - try to call validator functions
        /// </summary>
        public object Class { get; set; } // Excel application link

        /// <inheritdoc/>
        public void Dispose()
        {
            // Clean Excel Up
            // Perform any object clean up here.
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(this.Class);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(this.ExcelApp);

            // If you are inheriting from another class that
            // also implements IDisposable, don't forget to
            // call base.Dispose() as well.
            this.Dispose();
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Bevan Test Add Component method
        /// We can't use reflection to call the classes by name
        /// </summary>
        /// <param name="xlApp">Excel application object</param>
        /// <param name="sClassName">Validator Class Name</param>
        /// <returns>The Validator Class Reference</returns>
        private static object GetVBAClass(Excel.Application xlApp, string sClassName)
        {
            // Grab the Class component being passed in from name
            VBProject xlProj = xlApp.ActiveWorkbook.VBProject;
            VBComponent compVal = xlProj.VBComponents.Item(sClassName);

            // Function name to run
            string sFunctionName = "UNIT_TEST" + sClassName;

            // Add a new module/function
            VBComponent compModule = xlProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            compModule.CodeModule.InsertLines(
                compModule.CodeModule.CountOfLines + 1,
                "Public Function " + sFunctionName + "() As " + sClassName + "\r\n Set " + sFunctionName + " = New " + sClassName + "\r\n End Function");

            // Run the function
            object val = xlApp._Run2(sFunctionName);

            // Remove the function
            xlProj.VBComponents.Remove(compModule);

            // Clean the COM references
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(compModule);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(compVal);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlProj);

            // Return the object
            return val;
        }
    }
}
