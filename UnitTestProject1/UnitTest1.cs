using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace UnitTestProject1
{
    [TestClass]
    public class CalculateCostTests
    {
        [TestMethod]
        public void Test1_BigNumbers()
        {
            // Проверка парсинга большого числа
            string bigNumber = "10000.0";
            double result = double.Parse(bigNumber.Replace('.', ','));

            // Проверяем, что число действительно большое
            Assert.IsTrue(result > 9999, $"Число {result} должно быть больше 9999");
        }

        [TestMethod]
        public void Test2_NegativeNumbers()
        {
            // Проверка парсинга отрицательного числа
            string negative = "-10.5";
            double result = double.Parse(negative.Replace('.', ','));

            // Проверяем, что число отрицательное
            Assert.IsTrue(result < 0, $"Число {result} должно быть отрицательным");
        }

        [TestMethod]
        public void Test3_EmptyFields()
        {
            // Проверка пустой строки
            string empty = "";

            // Проверяем, что строка пустая
            Assert.IsTrue(string.IsNullOrWhiteSpace(empty), "Строка должна быть пустой");

            // Проверяем, что пустую строку нельзя распарсить
            bool canParse = double.TryParse(empty, out double result);
            Assert.IsFalse(canParse, "Пустую строку нельзя парсить");
        }
    }
}
