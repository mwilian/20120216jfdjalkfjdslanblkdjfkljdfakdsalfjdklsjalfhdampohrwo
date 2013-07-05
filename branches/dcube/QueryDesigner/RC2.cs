using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace dCube
{
    public class RC2
    {
        public static void ChoseOperationMode(RC2CryptoServiceProvider MyRC2, string operationMode)
        {
            if (MyRC2 != null)
            {
                switch (operationMode)
                {
                    case "CBC":
                        MyRC2.Mode = CipherMode.CBC;
                        break;
                    case "CFB":
                        MyRC2.Mode = CipherMode.CFB;
                        break;
                    case "CTS":
                        MyRC2.Mode = CipherMode.CTS;
                        break;
                    case "OFB":
                        MyRC2.Mode = CipherMode.OFB;
                        break;
                    default:
                        MyRC2.Mode = CipherMode.ECB;
                        break;
                }
            }
        }
        public static void ChosePaddingMode(RC2CryptoServiceProvider MyRC2, string paddingMode)
        {
            if (MyRC2 != null)
            {
                switch (paddingMode)
                {
                    case "PKCS7":
                        MyRC2.Padding = PaddingMode.PKCS7;
                        break;
                    case "X923":
                        MyRC2.Padding = PaddingMode.ANSIX923;
                        break;
                    case "ISO10126":
                        MyRC2.Padding = PaddingMode.ISO10126;
                        break;
                    default:
                        MyRC2.Padding = PaddingMode.None;
                        break;
                }
            }
        }
        public static void MaHoaFile(string inputFile, string outputFile, string szSecureKey, string strIV, string paddingMode, string operationMode)
        {
            FileStream inStream, outStream;
            inStream = new FileStream(inputFile, FileMode.Open, FileAccess.Read);
            outStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);

            RC2CryptoServiceProvider MyRC2 = new RC2CryptoServiceProvider();
            MyRC2.Key = ASCIIEncoding.ASCII.GetBytes(szSecureKey);
            MyRC2.IV = ASCIIEncoding.ASCII.GetBytes(strIV);
            ICryptoTransform MyRC2_Ecryptor = MyRC2.CreateEncryptor();
            ChoseOperationMode(MyRC2, operationMode);
            ChosePaddingMode(MyRC2, paddingMode);
            CryptoStream myEncryptStream;
            myEncryptStream = new CryptoStream(outStream, MyRC2_Ecryptor, CryptoStreamMode.Write);
            byte[] byteBuffer = new byte[100];
            long nTotalByteInput = inStream.Length, nTotalByteWritten = 0;
            int nCurReadLen = 0;

            while (nTotalByteWritten < nTotalByteInput)
            {
                nCurReadLen = inStream.Read(byteBuffer, 0, byteBuffer.Length);
                myEncryptStream.Write(byteBuffer, 0, nCurReadLen);
                nTotalByteWritten += nCurReadLen;
            }
            myEncryptStream.Close();
            inStream.Close();
            outStream.Close();

        }
        public static void GiaiMaFile(string inputFile, string outputFile, string szSecureKey, string strIV, string paddingMode, string operationMode)
        {
            FileStream inStream, outStream;
            inStream = new FileStream(inputFile, FileMode.Open, FileAccess.Read);
            outStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);

            RC2CryptoServiceProvider MyRC2 = new RC2CryptoServiceProvider();
            MyRC2.Key = ASCIIEncoding.ASCII.GetBytes(szSecureKey);
            MyRC2.IV = ASCIIEncoding.ASCII.GetBytes(strIV);
            ChoseOperationMode(MyRC2, operationMode);
            ChosePaddingMode(MyRC2, paddingMode);

            ICryptoTransform MyRC2_Decryptor = MyRC2.CreateDecryptor();

            CryptoStream myDecryptStream;
            myDecryptStream = new CryptoStream(outStream, MyRC2_Decryptor, CryptoStreamMode.Write);

            byte[] byteBuffer = new byte[100];
            long nTotalByteInput = inStream.Length, nTotalByteWritten = 0;
            int nCurReadLen = 0;

            while (nTotalByteWritten < nTotalByteInput)
            {
                nCurReadLen = inStream.Read(byteBuffer, 0, byteBuffer.Length);
                myDecryptStream.Write(byteBuffer, 0, nCurReadLen);
                nTotalByteWritten += nCurReadLen;
            }

            inStream.Close();
            outStream.Close();
            myDecryptStream.Close();

        }
        public static string EncryptString(string MainString, string szSecureKey, string strIV, string paddingMode, string operationMode)
        {
            MemoryStream memory = new MemoryStream();
            RC2CryptoServiceProvider MyRC2 = new RC2CryptoServiceProvider();
            MyRC2.Key = ASCIIEncoding.ASCII.GetBytes(szSecureKey);
            MyRC2.IV = ASCIIEncoding.ASCII.GetBytes(strIV);
            ICryptoTransform MyRC2_Ecryptor = MyRC2.CreateEncryptor();
            ChoseOperationMode(MyRC2, operationMode);
            ChosePaddingMode(MyRC2, paddingMode);

            CryptoStream myEncryptStream = new CryptoStream(memory, MyRC2_Ecryptor, CryptoStreamMode.Write);
            StreamWriter streamwriter = new StreamWriter(myEncryptStream);
            streamwriter.WriteLine(MainString);
            streamwriter.Close();
            myEncryptStream.Close();
            byte[] buffer = memory.ToArray();
            memory.Close();
            return Convert.ToBase64String(buffer);
        }

        public static string DecryptString(string MainString, string szSecureKey, string strIV, string paddingMode, string operationMode)
        {
            byte[] buffer = Convert.FromBase64String(MainString.Trim());
            MemoryStream memory = new MemoryStream(buffer);
            RC2CryptoServiceProvider MyRC2 = new RC2CryptoServiceProvider();
            MyRC2.Key = ASCIIEncoding.ASCII.GetBytes(szSecureKey);
            MyRC2.IV = ASCIIEncoding.ASCII.GetBytes(strIV);
            ChoseOperationMode(MyRC2, operationMode);
            ChosePaddingMode(MyRC2, paddingMode);

            ICryptoTransform MyRC2_Decryptor = MyRC2.CreateDecryptor();

            CryptoStream myDecryptStream = new CryptoStream(memory, MyRC2_Decryptor, CryptoStreamMode.Read);
            StreamReader sr = new StreamReader(myDecryptStream);

            // Read the stream as a string.
            string val = sr.ReadToEnd();

            // Close the streams.
            sr.Close();
            myDecryptStream.Close();
            memory.Close();
            return val;

        }
    }
}
