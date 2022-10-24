using System;
using System.Text;
using System.Security.Cryptography;

namespace Init_Data
{


    /// <summary>
    /// 
    /// </summary>
    public class DataGrid_Init_05
    {
        public string Name01 { get; set; }
        public string Name02 { get; set; }
        public string Name03 { get; set; }
        public string Name04 { get; set; }
        public string Name05 { get; set; }
        public int NameId { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public class DataGrid_Init_10
    {
        public string Name01 { get; set; }
        public string Name02 { get; set; }
        public string Name03 { get; set; }
        public string Name04 { get; set; }
        public string Name05 { get; set; }
        public string Name06 { get; set; }
        public string Name07 { get; set; }
        public string Name08 { get; set; }
        public string Name09 { get; set; }
        public string Name10 { get; set; }
        public int NameId { get; set; }
    }
    /// <summary>
    /// 
    /// </summary>
    public class Combo_InitList
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public class CryptoEngine
    {
        public static string Encrypt(string input, string key)
        {
            byte[] inputArray = UTF8Encoding.UTF8.GetBytes(input);
            TripleDESCryptoServiceProvider tripleDES = new TripleDESCryptoServiceProvider();
            tripleDES.Key = UTF8Encoding.UTF8.GetBytes(key);
            tripleDES.Mode = CipherMode.ECB;
            tripleDES.Padding = PaddingMode.PKCS7;
            ICryptoTransform cTransform = tripleDES.CreateEncryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length);
            tripleDES.Clear();
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
        public static string Decrypt(string input, string key)
        {
            byte[] inputArray = Convert.FromBase64String(input);
            TripleDESCryptoServiceProvider tripleDES = new TripleDESCryptoServiceProvider();
            tripleDES.Key = UTF8Encoding.UTF8.GetBytes(key);
            tripleDES.Mode = CipherMode.ECB;
            tripleDES.Padding = PaddingMode.PKCS7;
            ICryptoTransform cTransform = tripleDES.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length);
            tripleDES.Clear();
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
    }

}
