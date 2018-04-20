using System;
using TopssLogger;

namespace CommonMath
{
  public static class SpcCalculate
  {
      /// <summary>
      /// Calculatec Average
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <param name="digits">The number of digits after the decimal point. The default value is 6</param>
      /// <returns>average value of data source </returns>
      static double CalculateAverage(double[] datasource, Byte digits = 6)
      {
          //平均計算
          double result = 0;
          foreach(double d in datasource)
          {
            result += d;
          }
          return Math.Round((result / datasource.Length), digits, MidpointRounding.AwayFromZero);
      }

      /// <summary>
      /// Calculate Standard Deviation
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <param name="averageValue">average value</param>
      /// <returns></returns>
      static double CalculateStandardDeviation(double[] datasource, double averageValue)
      {
          //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　START
		  if (datasource == null)
          {
              return double.NaN;
          }
          if (datasource.Length == 0)
          {
              return double.NaN;
          }
		  //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　END
		  
          //標準偏差計算
          double average = averageValue;//calcAverage(datasource);
          double temp = 0;
          double tempSum = 0;
          foreach(double d in datasource)
          {
            temp = Math.Pow((d - average), 2);//平均との差の2乗
            tempSum += temp;//総和
          }
          //総和を個数で割ったものの平方根を計算
          double result = Math.Sqrt((tempSum / datasource.Length));//Math.Round((tempSum / datasource.Length), 3, MidpointRounding.AwayFromZero));
          return result;
      }

      /// <summary>
      /// CP値計算
      /// </summary>
      /// <param name="datasource">データソース</param>
      /// <param name="usl">上限規格値</param>
      /// <param name="lsl">下限規格値</param>
      /// <param name="cp">Cp値</param>
      /// <param name="cpu">CpUpper値</param>
      /// <param name="cpl">CpLower値</param>
      /// <param name="cpk">Cpk値</param>
      /// <param name="digits">The number of digits after the decimal point. The default value is 6</param>
      public static bool CalculateCpValue(double[] datasource, double usl, double lsl, out double cp, out double cpu, out double cpl, out double cpk, Byte digits = 6)
      {
          //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　START
          //cp = 0;
          //cpu = 0;
          //cpl = 0;
          //cpk = 0;
		  cp = double.NaN;
          cpu = double.NaN;
          cpl = double.NaN;
          cpk = double.NaN;

          if (datasource == null)
              return false;
          
          if (datasource.Length < 1)
            return false;
          //平均値計算
          double average = CalculateAverage(datasource);
          //標準偏差計算
          double sigma = CalculateStandardDeviation(datasource, average);
          if (double.IsNaN(sigma))
          {
              return false;
          }
          double epsilon = Math.Pow(10, 0 - digits);
          //if (Math.Abs(sigma - epsilon) < 0)
          if(Math.Abs(sigma) <= epsilon)
          {
            return false;//sigmaが0の場合は計算できない
          }
		  //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　END
          LogManager.GetLogger().Write(LogLevel.Error, "datasource.length = " + datasource.Length + " usl = " + usl + " lsl = " + lsl + " average = " + average + " sigma = " + sigma + " digits = " + digits);
          //Cp値計算
          cp = Math.Round((usl - lsl) / (6 * sigma), digits, MidpointRounding.AwayFromZero);
          cpu = Math.Round((usl - average) / (3 * sigma), digits, MidpointRounding.AwayFromZero);
          cpl = Math.Round((average - lsl) / (3 * sigma), digits, MidpointRounding.AwayFromZero);
          cpk = Math.Min(cpu, cpl);//CpuまたはCplのどちらか小さい方
          LogManager.GetLogger().Write(LogLevel.Error, " cp = " + cp + " cpu = " + cpu + " cpl = " + cpl + " cpk = " + cpk);
          return true;
      }

      /// <summary>
      /// Calculate Standard Deviation
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <returns></returns>
      public static double CalculateStandardDeviation(double[] datasource)
      {
     	  //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　START
          if (datasource == null)
          {
              return Double.NaN;
          }
          if (datasource.Length == 0)
          {
              return Double.NaN;
          }
		  //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　end
          //標準偏差計算
          //平均計算
          double result = 0;
          foreach (double d in datasource)
          {
              result += d;
          }
          double average = Math.Round((result / datasource.Length), 3, MidpointRounding.AwayFromZero);//calcAverage(datasource);
          double temp = 0;
          double tempSum = 0;
          foreach (double d in datasource)
          {
              temp = Math.Pow((d - average), 2);//平均との差の2乗
              tempSum += temp;//総和
          }
          //総和を個数で割ったものの平方根を計算
          return Math.Sqrt((tempSum / datasource.Length));//Math.Round((tempSum / datasource.Length), 3, MidpointRounding.AwayFromZero));
      }

      /// <summary>
      /// Get Center Line(CL)
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <returns></returns>
      public static double GetCenterLine(double[] datasource)
      {
          return CalculateAverage(datasource);
      }

      /// <summary>
      /// CL値計算
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <param name="lcl1">Lower CL 値1</param>
      /// <param name="ucl1">Upper CL 値1</param>
      /// <param name="lcl2">Lower CL 値2</param>
      /// <param name="ucl2">Upper CL 値2</param>
      /// <param name="digits">The number of digits after the decimal point. The default value is 6</param>
      /// <returns></returns>
      public static bool CalculateClValue(double[] datasource, out double lcl1, out double ucl1, out double lcl2, out double ucl2, Byte digits = 6)
      {
          //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　START
          //lcl1 = 0;
          //ucl1 = 0;
          //lcl2 = 0;
          //ucl2 = 0;
		  
		  lcl1 = Double.NaN;
          ucl1 = Double.NaN;
          lcl2 = Double.NaN;
          ucl2 = Double.NaN;

          if (datasource.Length < 1)
              return false;

          //平均値計算
          double average = CalculateAverage(datasource);
          //標準偏差計算
          double sigma = CalculateStandardDeviation(datasource, average);
          if (double.IsNaN(sigma))
          {
              return false;
          }
          double epsilon = Math.Pow(10, 0 - digits);
          if (Math.Abs(sigma) <= epsilon)
          {
              return false;//sigmaが0の場合は計算できない
          }
		  //CHN 2017/09/25 modify by oujianbo 分析画面で生産能力指数の計算が正しくない　END
          //CL値計算
          lcl1 = Math.Round(average - 3 * sigma, digits, MidpointRounding.AwayFromZero);
          ucl1 = Math.Round(average + 3 * sigma, digits, MidpointRounding.AwayFromZero);
          lcl2 = Math.Round(average - 6 * sigma, digits, MidpointRounding.AwayFromZero);
          ucl2 = Math.Round(average + 6 * sigma, digits, MidpointRounding.AwayFromZero);
          return true;
      }
  }
}
