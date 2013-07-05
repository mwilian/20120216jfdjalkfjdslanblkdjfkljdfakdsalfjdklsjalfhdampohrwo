using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;

public class MyHash
{
    private static HashAlgorithm m_Hash;
    public MyHash(string val)
    {
        switch (val)
        {
            case "SHA1":
                m_Hash = new SHA1CryptoServiceProvider();
                break;
            case "MD5":
                m_Hash = new MD5CryptoServiceProvider();
                break;
            case "SHA256":
                m_Hash = new SHA256Managed();
                break;
            case "SHA384":
                m_Hash = new SHA384Managed();
                break;
            case "SHA512":
                m_Hash = new SHA512Managed();
                break;
        }
    }
    public static void Init(string val)
    {
        switch (val)
        {
            case "SHA1":
                m_Hash = new SHA1CryptoServiceProvider();
                break;
            case "MD5":
                m_Hash = new MD5CryptoServiceProvider();
                break;
            case "SHA256":
                m_Hash = new SHA256Managed();
                break;
            case "SHA384":
                m_Hash = new SHA384Managed();
                break;
            case "SHA512":
                m_Hash = new SHA512Managed();
                break;
        }
    }
    //public byte[] Hash(string input)
    //{
    //    byte[] inputData = Encoding.UTF8.GetBytes(input);
    //    byte[] result;
    //    result = m_Hash.ComputeHash(inputData);
    //    return result;
    //}
    public static string Hash(string input, string mode)
    {
        Init(mode);
        byte[] inputData = Encoding.UTF8.GetBytes(input);
        byte[] result;
        result = m_Hash.ComputeHash(inputData);
        string outString = "";
        for (int i = 0; i < result.Length; i++)
        {
            outString = outString + result[i].ToString("X2");
        }
        return outString;
    }
    public static string Hash(Stream input, string mode)
    {
        Init(mode);
        //byte[] inputData = Encoding.UTF8.GetBytes(input);
        byte[] result;
        result = m_Hash.ComputeHash(input);
        string outString = "";
        for (int i = 0; i < result.Length; i++)
        {
            outString = outString + result[i].ToString("X2");
        }
        return outString;
    }
}

