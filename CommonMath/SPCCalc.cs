using System;

namespace CommonMath
{
  public static class SpcCalculate
  {
      /// <summary>
      /// Calculatec Average
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <returns>average value of data source </returns>
      static double CalculateAverage(double[] datasource)
      {
          //平均計算
          double result = 0;
          foreach(double d in datasource)
          {
            result += d;
          }
          return Math.Round((result / datasource.Length), 3, MidpointRounding.AwayFromZero);
      }

      /// <summary>
      /// Calculate Standard Deviation
      /// </summary>
      /// <param name="datasource">data source</param>
      /// <param name="averageValue">average value</param>
      /// <returns></returns>
      static double CalculateStandardDeviation(double[] datasource, double averageValue)
      {
          if (datasource.Length == 0)
          {
              return double.Epsilon;
          }
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
      public static bool CalculateCpValue(double[] datasource, double usl, double lsl, out double cp, out double cpu, out double cpl, out double cpk)
      {
          cp = 0;
          cpu = 0;
          cpl = 0;
          cpk = 0;

          if (datasource.Length < 1)
            return false;

          //平均値計算
          double average = CalculateAverage(datasource);
          //標準偏差計算
          double sigma = CalculateStandardDeviation(datasource, average);
          if (Math.Abs(sigma) < double.Epsilon)
          {
            return false;//sigmaが0の場合は計算できない
          }

          //Cp値計算
          cp = Math.Round((usl - lsl) / (6 * sigma), 3, MidpointRounding.AwayFromZero);
          cpu = Math.Round((usl - average) / (3 * sigma), 3, MidpointRounding.AwayFromZero);
          cpl = Math.Round((average - lsl) / (3 * sigma), 3, MidpointRounding.AwayFromZero);
          cpk = Math.Min(cpu, cpl);//CpuまたはCplのどちらか小さい方
          return true;
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
      /// <returns></returns>
      public static bool CalculateClValue(double[] datasource, out double lcl1, out double ucl1, out double lcl2, out double ucl2)
      {
          lcl1 = 0;
          ucl1 = 0;
          lcl2 = 0;
          ucl2 = 0;

          if (datasource.Length < 1)
              return false;

          //平均値計算
          double average = CalculateAverage(datasource);
          //標準偏差計算
          double sigma = CalculateStandardDeviation(datasource, average);
          if (Math.Abs(sigma) < double.Epsilon)
          {
              return false;//sigmaが0の場合は計算できない
          }

          //CL値計算
          lcl1 = Math.Round(average - 3 * sigma, 3, MidpointRounding.AwayFromZero);
          ucl1 = Math.Round(average + 3 * sigma, 3, MidpointRounding.AwayFromZero);
          lcl2 = Math.Round(average - 6 * sigma, 3, MidpointRounding.AwayFromZero);
          ucl2 = Math.Round(average + 6 * sigma, 3, MidpointRounding.AwayFromZero);
          return true;
      }
  }
}
