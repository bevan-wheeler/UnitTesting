﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ValidatorExample
{
    /// <summary>
    /// Summary description for SampleExample
    /// </summary>
    [TestClass]
    public class SetupExample
    {
        //Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext) { }

        //Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void MyClassCleanup() { }

        //Use TestInitialize to run code before running each test
        [TestInitialize()]
        public void MyTestInitialize() { }

        //Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void MyTestCleanup() { }

        //Macro Test 
        [TestMethod]
        public void TestMethod1()
        {
            // TODO: Add test logic here
         
            //Arrange

            //Act

            //Assert
        }
    }
}
