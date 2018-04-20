using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CommonTool
{
    public static class FtpTransferTaskAllot
    {
//2018/04/10 modify by daijun #792【SPC改善】SPCの検査結果比較で結果比較する時に比較失敗したのメッセージを出しました(#8328) start
        public static List<List<T>> TaskAllot<T>(List<T> listFtpParameter, int threadCount)
        {
            List<List<T>> ret = new List<List<T>>();

            int nThreadNum = threadCount;
            int nFtpThreadNum = nThreadNum / 2 + 1;

            int nBegin = 0;
            int nLength = 0;
            if (listFtpParameter.Count < nFtpThreadNum)
            {
                nFtpThreadNum = listFtpParameter.Count;
                nLength = 1;
            }
            else
            {
                nLength = listFtpParameter.Count / nFtpThreadNum;
            }

            for (int i = 0; i < nFtpThreadNum; i++)
            {
                List<T> listFtpParameterTemp = new List<T>();

                if (i == nFtpThreadNum - 1)
                {
                    listFtpParameterTemp.AddRange(listFtpParameter.GetRange(nBegin, listFtpParameter.Count - nBegin));
                }
                else
                {
                    listFtpParameterTemp.AddRange(listFtpParameter.GetRange(nBegin, nLength));
                    nBegin += nLength;
                }

                ret.Add(listFtpParameterTemp);
            }

            return ret;
        }
//2018/04/10 modify by daijun #792【SPC改善】SPCの検査結果比較で結果比較する時に比較失敗したのメッセージを出しました(#8328) end
    }
}
