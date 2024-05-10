using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace BOD_Utility
{
    static class clsEncryptionDecryption
    {
        public static string EncryptString(string strData, string strKey)
        {

            byte[] key = { }; //Encryption Key   
            byte[] IV = { 10, 20, 30, 40, 50, 60, 70, 80 };
            byte[] inputByteArray;

            try
            {
                key = Encoding.UTF8.GetBytes(strKey.ToString());
                // DESCryptoServiceProvider is a cryptography class defind in c#.  
                DESCryptoServiceProvider ObjDES = new DESCryptoServiceProvider();
                inputByteArray = Encoding.UTF8.GetBytes(strData);
                MemoryStream Objmst = new MemoryStream();
                CryptoStream Objcs = new CryptoStream(Objmst, ObjDES.CreateEncryptor(key, IV), CryptoStreamMode.Write);
                Objcs.Write(inputByteArray, 0, inputByteArray.Length);
                Objcs.FlushFinalBlock();

                return Convert.ToBase64String(Objmst.ToArray());//encrypted string  
            }
            catch (Exception)
            {
                //if (!exceptionsList.Contains(ee.Message))
                //{
                //    InsertError(ee.Message + ":" + ee.ToString());
                //}
                return "";
            }

        }

        public static string DecryptData(string strData, string strKey)
        {

            string rtMesssage = "";
            try
            {
                byte[] key = { };// Key   
                byte[] IV = { 10, 20, 30, 40, 50, 60, 70, 80 };
                byte[] inputByteArray = new byte[strData.Length];


                key = Encoding.UTF8.GetBytes(strKey);
                DESCryptoServiceProvider ObjDES = new DESCryptoServiceProvider();
                inputByteArray = Convert.FromBase64String(strData);

                MemoryStream Objmst = new MemoryStream();
                CryptoStream Objcs = new CryptoStream(Objmst, ObjDES.CreateDecryptor(key, IV), CryptoStreamMode.Write);
                Objcs.Write(inputByteArray, 0, inputByteArray.Length);
                Objcs.FlushFinalBlock();

                Encoding encoding = Encoding.UTF8;
                rtMesssage = encoding.GetString(Objmst.ToArray());
            }
            catch (Exception)
            { }
            return rtMesssage;
        }

        public static string Encrypt(string clearText)                                          //added on 25-2-19 by Amey
        {
            string EncryptionKey = "nerve123";
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return clearText;
        }

        public static string Decrypt(string cipherText)                                                     //added on 25-2-19 by Amey
        {
            string EncryptionKey = "nerve123";
            cipherText = cipherText.Replace(" ", "+");
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }

        #region Decryption

        public static DataTable DecryptData(DataTable dt)
        {
            try
            {
                int a = 0;
                #region added by navin on 29-06-2018 speed improvement
                IEnumerator enum_Clientifo = dt.Rows.GetEnumerator();
                DataRow dr_Clientifo;
                while (enum_Clientifo.MoveNext())
                {
                    dr_Clientifo = (DataRow)enum_Clientifo.Current;
                    for (int i = a; i < dr_Clientifo.ItemArray.Length; i++)
                    {
                        dr_Clientifo[i] = DecryptData(dr_Clientifo[i].ToString(), "Nerve123");
                    }
                }
                #endregion
                return dt;
            }
            catch (Exception)
            {
                return dt;
            }
        }

        public static string DecryptString(string strData, string strKey)
        {
            try
            {
                byte[] key = { };// Key   
                byte[] IV = { 10, 20, 30, 40, 50, 60, 70, 80 };
                byte[] inputByteArray = new byte[strData.Length];

                key = Encoding.UTF8.GetBytes(strKey);
                DESCryptoServiceProvider ObjDES = new DESCryptoServiceProvider();
                inputByteArray = Convert.FromBase64String(strData);

                MemoryStream Objmst = new MemoryStream();
                CryptoStream Objcs = new CryptoStream(Objmst, ObjDES.CreateDecryptor(key, IV), CryptoStreamMode.Write);
                Objcs.Write(inputByteArray, 0, inputByteArray.Length);
                Objcs.FlushFinalBlock();

                Encoding encoding = Encoding.UTF8;
                return encoding.GetString(Objmst.ToArray());
            }
            catch (Exception)
            {
                return "";
            }
        }

        #endregion
    }
}
