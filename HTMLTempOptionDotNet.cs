using System;
using System.Collections.Generic;
using System.Text;
using System.Web;
using System.Data;
using System.Web.Caching;

namespace HTMLTempOptionDotNet
{
    public class HTMLDNetClassClass
    {
        #region 私有变量
        StringBuilder sbTemp = new StringBuilder();

        StringBuilder sbSub;

        StringBuilder sbSubTotal;

        string fullpath;

        string subText;

        string beginBlock;
        string endBlock;

        int beginPos;
        int endPos;

        static Dictionary<string, string> temps = new Dictionary<string, string>();
        bool allowCache;
        #endregion


        public HTMLDNetClassClass()
        {
        }

        public HTMLDNetClassClass(bool allowCache)
        {
            this.allowCache = allowCache;
        }

        public void OpenHtmlFile(string filename,bool UTF8=false)
        {
            sbTemp.Clear();
            try
            {
                //包含\即认为是物理路径，直接使用，不包含则认为是虚拟路径，需要转换
                if (filename.IndexOf("\\") > -1)
                {
                    fullpath = filename;
                }
                else
                {
                    fullpath = HttpContext.Current.Server.MapPath(filename);
                }
                fullpath = fullpath.ToLower();

                string tempContent;
                if (allowCache && temps.ContainsKey(fullpath))
                {
                    tempContent = temps[fullpath];
                }
                else
                {
                    //tempContent = System.IO.File.ReadAllText(fullpath, Encoding.GetEncoding("gb2312"));
                    if (!UTF8)
                    {
                        tempContent = ReadTempFile(fullpath, Encoding.GetEncoding("gb2312"));
                    }
                    else
                    {
                        tempContent = ReadTempFile(fullpath, Encoding.GetEncoding("utf-8"));
                    }
                    if (!temps.ContainsKey(fullpath))
                        temps.Add(fullpath, tempContent);
                }
                sbTemp.Append(tempContent);
            }
            catch
            { 
                
            }
        }

        public string ReadTempFile(string fullpath, Encoding encoding)
        {
            string strFileContent = HttpRuntime.Cache[fullpath] as string;
            //string strFileContent = null;
            if (strFileContent == null)
            {
                strFileContent = System.IO.File.ReadAllText(fullpath,encoding);
                CacheDependency dep = new CacheDependency(fullpath);
                HttpRuntime.Cache.Insert(fullpath, strFileContent, dep);
                //strFileContent = "File:" + strFileContent;
            }
            else
            {
                if (encoding.EncodingName.ToLower() == "utf-8")
                { 
                    
                }
                //strFileContent = "Cache:" + strFileContent; ;
            }
            return strFileContent;
        }

        /// <summary>
        /// 首选编码的代码页名称
        /// </summary>
        /// <param name="srcName">原编码格式</param>
        /// <param name="convToName">要转换成的编码格式</param>
        /// <param name="value">需要转换的字符串</param>
        /// <returns>返回转换后的字符串</returns>
        public static string EnCodeCovert(string srcName, string convToName, string value)
        {
            System.Text.Encoding srcEncode = System.Text.Encoding.GetEncoding(srcName);
            System.Text.Encoding convToEncode = System.Text.Encoding.GetEncoding(convToName);
            byte[] bytes = srcEncode.GetBytes(value);
            System.Text.Encoding.Convert(srcEncode, convToEncode, bytes, 0, bytes.Length);
            return convToEncode.GetString(bytes);
        }

        public void ReplaceVar(string tagName, string newValue)
        {
            sbTemp.Replace("%%" + tagName + "%%", newValue);
        }

        /// <summary>
        /// 仅为了兼容，其实什么都不做
        /// </summary>
        public void CloseFile()
        {
        }

        public void ListSub(string subName)
        {
            if (string.IsNullOrEmpty(subName))
                throw new Exception("LIST方法中区块名不能为空");

            if (sbTemp.Length == 0)
                throw new Exception("模板文件不存在");

            beginBlock = "<!--start " + subName + " -->";
            endBlock = "<!--end " + subName + " -->";


            beginPos = sbTemp.ToString().IndexOf(beginBlock);
            endPos = sbTemp.ToString().IndexOf(endBlock);

             

            if (beginPos == -1 || endPos == -1 || endPos < beginPos)
            {
                //string message = "区块" + subName + "不存在或不完全";
                //throw new Exception(message);
                //System.Web.HttpContext.Current.Response.Write(message);
                throw new Exception("LIST方法中区块名未找到");
            }

            int sublength = endPos - beginPos + endBlock.Length;
            char[] chars = new char[sublength];
            sbTemp.CopyTo(beginPos, chars, 0, sublength);
            subText = new string(chars);
            sbSub = new StringBuilder();
            sbSub.Append(subText);
            sbSubTotal = new StringBuilder();
            //sbSub = new StringBuilder();
            //sbSub.Append(chars);
        }

        public void ReplaceSub(string tagName, string newValue)
        {
            if (sbSub == null)
            {
                if (subText == null)
                {
                    //throw new Exception("请先用list方法打开区块");
                    System.Web.HttpContext.Current.Response.Write("请先用list方法打开区块");
                    //subText = "";
                }
                sbSub = new StringBuilder();
                sbSub.Append(subText);
            }

            sbSub.Replace("%%" + tagName + "%%", newValue);
        }

        public void ReplaceSub(DataTable dt)
        {
            foreach (DataRow dataRow in dt.Rows)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    string tagName = column.ColumnName;
                    string newValue = dataRow[tagName].ToString();

                    ReplaceSub(tagName, newValue);
                }
                ReplaceSubAll();
            }
        }

        public void ReplaceSubAll()
        {
            sbSubTotal.Append(sbSub.ToString());
            sbSub = null;
        }

        public void ReplaceSubHtml()
        {
            try
            {
	            sbTemp.Replace(subText, sbSubTotal.ToString());
	            sbTemp.Replace(beginBlock, "");
	            sbTemp.Replace(endBlock, "");
	
	            beginPos = -1;
	            endPos = -1;
	            sbSubTotal = null;
	            sbSub = null;
	     }
	     catch
	     {
	     }    
        }



        public void ReplaceSubNull()
        {
            sbTemp.Replace(subText, subText.ToString());
        }

        public string GetHtml()
        {
            return sbTemp.ToString();
        }

        public void SetHtml(string sV) 
        {
            //sbTemp.Clear();
            sbTemp.Remove(0, sbTemp.Length);//���StringBuilder�ķ���
            sbTemp.Append(sV);

        }

        public string Ver()
        {
            return "1.03 For .Net 2010-4-18";
        }          
    }
}