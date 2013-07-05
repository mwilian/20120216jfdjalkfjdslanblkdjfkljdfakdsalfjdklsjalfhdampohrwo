
using System;
using System.Collections.Generic;
using System.Text;

namespace dCube
{
    class ReadEN
    {
        private string[] strSo ={ "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight","Nine" };
        private string[] strDonViNho ={ "linh", "lăm", "mười", "mươi", "mốt", "trăm" };
        private string[] strDonViLon = { "", "Thousand", "Million", "Billion" };
        private string[] strMainGroup;
        private string[] strSubGroup;
        private string Len1(string strA)
        {
            return strSo[int.Parse(strA)];
        }
        private string Len2(string strA)
        {
            if (strA.Substring(0, 1) == "0")
            {
                return strDonViNho[0] + " " + Len1(strA.Substring(1, 1));
            }
            else if (strA.Substring(0, 1) == "1")
            {
                if (strA.Substring(1, 1) == "5")
                {
                    return strDonViNho[2] + " " + strDonViNho[1];
                }
                else if (strA.Substring(1, 1) == "0")
                {
                    return strDonViNho[2];
                }
                else
                {
                    return strDonViNho[2] + " " + Len1(strA.Substring(1, 1));
                }
            }
            else
            {
                if (strA.Substring(1, 1) == "5")
                {
                    return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3] + " " + strDonViNho[1];
                }
                else if (strA.Substring(1, 1) == "0")
                {
                    return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3];
                }
                else if (strA.Substring(1, 1) == "1")
                {
                    return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3] + " " + strDonViNho[4];
                }
                else
                {
                    return Len1(strA.Substring(0, 1)) + " " + strDonViNho[3] + " " + Len1(strA.Substring(1, 1));
                }
            }
        }
        private string Len3(string strA)
        {
            if ((strA.Substring(0, 3) == "000"))
            {
                return null;
            }
            else if ((strA.Substring(1, 2) == "00"))
            {
                return Len1(strA.Substring(0, 1)) + " " + strDonViNho[5];
            }
            else
            {
                return Len1(strA.Substring(0, 1)) + " " + strDonViNho[5] + " " + Len2(strA.Substring(1, strA.Length - 1));
            }
        }
        /////////////////////
        private string FullLen(string strSend)
        {
            bool boKTNull = false;
            string strKQ ="";
            string strA = strSend.Trim();
            int iIndex = strA.Length - 9;
            int iPreIndex = 0;

            if (strSend.Trim() == "")
            {
                return Len1("0");
            }
            //tra ve khong neu la khong
            for (int i = 0; i < strA.Length; i++)
            {
                if (strA.Substring(i, 1) != "0")
                {
                    break;
                }
                else if (i == strA.Length - 1)
                {
                    return strSo[0];
                }
            }
            int k = 0;
            while (strSend.Trim().Substring(k++, 1) == "0")
            {
                strA = strA.Remove(0, 1);
            }
            //
            if (strA.Length < 9)
            {
                iPreIndex = strA.Length;
            }
            //
            if ((strA.Length % 9) != 0)
            {
                strMainGroup = new string[strA.Length / 9 + 1];
            }
            else
            {
                strMainGroup = new string[strA.Length / 9];
            }
            //nguoc
            for (int i = strMainGroup.Length - 1; i >= 0; i--)
            {
                if (iIndex >= 0)
                {
                    iPreIndex = iIndex;
                    strMainGroup[i] = strA.Substring(iIndex, 9);
                    iIndex -= 9;
                }
                else
                {
                    strMainGroup[i] = strA.Substring(0, iPreIndex);
                }

            }
            /////////////////////////////////
            //tach moi maingroup thanh nhieu subgroup
            //xuoi
            for (int j = 0; j < strMainGroup.Length; j++)
            {
                //gan lai gia tri
                iIndex = strMainGroup[j].Length - 3;
                if (strMainGroup[j].Length < 3)
                {
                    iPreIndex = strMainGroup[j].Length;
                }
                ///
                if ((strMainGroup[j].Length % 3) != 0)
                {
                    strSubGroup = new string[strMainGroup[j].Length / 3 + 1];
                }
                else
                {
                    strSubGroup = new string[strMainGroup[j].Length / 3];
                }
                for (int i = strSubGroup.Length - 1; i >= 0; i--)
                {
                    if (iIndex >= 0)
                    {
                        iPreIndex = iIndex;
                        strSubGroup[i] = strMainGroup[j].Substring(iIndex, 3);
                        iIndex -= 3;
                    }
                    else
                    {
                        strSubGroup[i] = strMainGroup[j].Substring(0, iPreIndex);
                    }
                }
                //duyet subgroup de lay string
                for (int i = 0; i < strSubGroup.Length; i++)
                {
                    boKTNull = false;//phai de o day
                    if ((j == strMainGroup.Length - 1) && (i == strSubGroup.Length - 1))
                    {
                        if (strSubGroup[i].Length < 3)
                        {
                            if (strSubGroup[i].Length == 1)
                            {
                                strKQ +=Len1(strSubGroup[i]);
                            }
                            else
                            {
                                strKQ += Len2(strSubGroup[i]);
                            }
                        }
                        else
                        {
                            strKQ += Len3(strSubGroup[i]);
                        }
                    }
                    else
                    {
                        if (strSubGroup[i].Length < 3)
                        {
                            if (strSubGroup[i].Length == 1)
                            {
                                strKQ += Len1(strSubGroup[i]) + " ";
                            }
                            else
                            {
                                strKQ += Len2(strSubGroup[i]) + " ";
                            }
                        }
                        else
                        {
                            if (Len3(strSubGroup[i]) == null)
                            {
                                boKTNull = true;
                            }
                            else
                            {
                                strKQ += Len3(strSubGroup[i]) + " ";
                            }
                        }
                    }
                    //dung
                    if (!boKTNull)
                    {
                        if (strSubGroup.Length - 1 - i != 0)
                        {
                            strKQ += strDonViLon[strSubGroup.Length - 1 - i] + " ";
                        }
                        else
                        {
                            strKQ += strDonViLon[strSubGroup.Length - 1 - i] + " ";
                        }

                    }
                }
                //dung
                if (j != strMainGroup.Length - 1)
                {
                    if (!boKTNull)
                    {
                        strKQ = strKQ.Substring(0, strKQ.Length - 1) + strDonViLon[3] + " ";
                    }
                    else
                    {
                        strKQ = strKQ.Substring(0, strKQ.Length - 1) + " " + strDonViLon[3] + " ";
                    }
                }
            }
            //xoa ky tu trang
            strKQ = strKQ.Trim();
            //xoa dau , neu co
            if (strKQ.Substring(strKQ.Length - 1, 1) == ".")
            {
                strKQ = strKQ.Remove(strKQ.Length - 1, 1);
            }
            return strKQ;

            ////////////////////////////////////


        }
        public string Convert(string strSend, char charInSeparator, string strOutSeparator)
        {
            if (strOutSeparator == "")
            {
                return "Lỗi dấu phân cách đầu ra rỗng";
            }
            if (strSend == "")
            {
                return Len1("0");
            }

            string[] strTmp = new string[2];
            try
            {
                
                strTmp = strSend.Split(charInSeparator);
                string strTmpRight = strTmp[1];
                for (int i = strTmpRight.Length - 1; i >= 0; i--)
                { 
                    if(strTmpRight.Substring(i,1)=="0")
                    {
                        strTmpRight = strTmpRight.Remove(i,1);
                    }
                    else
                    {
                        break;
                    }
                }
                if (strTmpRight != "")
                {
                    string strRight = "";
                    for (int i = 0; i < strTmpRight.Length; i++)
                    {
                        strRight += Len1(strTmpRight.Substring(i, 1)) + " ";
                    }


                    return FullLen(strTmp[0]) + " " + strOutSeparator + " " + strRight.TrimEnd();
                }
                else
                {
                    return FullLen(strTmp[0]);
                }
            }
            catch
            {
                return FullLen(strTmp[0]);
            }

        }

    }
}
