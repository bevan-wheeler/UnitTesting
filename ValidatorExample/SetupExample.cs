namespace UnitTest
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// SetupExample class
    /// </summary>
    [TestClass]
    public class SetupExample
    {
        private static ExcelTest xlTest; // Testing object
        private bool bResult;        // Result boolean

        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class
        /// </summary>
        /// <param name="testContext">Current Testing Context</param>
        [ClassInitialize]
        public static void MyClassInitialize(TestContext testContext)
        {
            // Generate our new testing object
            xlTest = new ExcelTest();

            // Create a new instance of our class
            // This is only needed for class objects
            xlTest.SetClass("ValidatorPriceListDates");
        }

        /// <summary>
        /// Use ClassCleanup to run code after all tests in a class have run
        /// </summary>
        [ClassCleanup]
        public static void MyClassCleanup()
        {
            xlTest.Dispose();
        }

        /// <summary>
        /// Use TestInitialize to run code before running each test
        /// </summary>
        [TestInitialize]
        public void MyTestInitialize()
        {
            // Add Initialise Test code here
            this.bResult = false;
        }

        /// <summary>
        /// Use TestCleanup to run code after each test has run
        /// </summary>
        [TestCleanup]
        public void MyTestCleanup()
        {
            // Add Termination Test code here
            this.bResult = false;
        }

        /// <summary>
        /// TestModule1 sample
        /// Runs a procedure (not a validator[class] procedure)
        /// </summary>
        [TestMethod]
        public void TestModule1()
        {
            // TODO: Add test logic here

            // Arrange
            // sParam1 = "Test";

            // Act
            // bResult = xlTest.ExcelApp._Run2("Procedure", sParam1);

            // Assert
            // Assert.IsFalse(bActual);
        }
    }
}
