using System.Runtime.InteropServices;
using LSEXT;
using System;

namespace ResultFormattingExtension
{
    [ComVisible(true)]
    [ProgId("ResultFormattingExtension.ResultFormattingExtension")]
    public class ResultFormattingExtension : IResultFormat
    {
        public ResultFormattingExtension(){}
      
        public ResultEntryFormat Format(ref LSExtensionParameters Parameters, ResultEntryPhase Phase)
        {
            if (Phase == ResultEntryPhase.reConvertUnits)
            {
                try
                {
                    double rawNumericResult = Parameters.Parameter("raw_numeric_result").Value;
                    double dilutionFactor = Parameters.Parameter("dilution_factor").Value;
                    dilutionFactor = dilutionFactor == 0 ? 1 : dilutionFactor;
                    if (FractionalPart(dilutionFactor) == 5)
                    {
                        Parameters.Parameter("calculated_numeric_result").Value = rawNumericResult * 50;
                    }
                    else
                    {
                        Parameters.Parameter("calculated_numeric_result").Value = rawNumericResult*
                                                                                  Math.Pow(10, dilutionFactor);
                    }
                    return LSEXT.ResultEntryFormat.rfSkipDefault;
                }
                catch(COMException exception)
                {
                //    Logger.WriteLogFile(exception);
                    
                }
            }
            return LSEXT.ResultEntryFormat.rfDoDefault;
        }

        public ResultFieldChange FieldChange(ref LSExtensionParameters Parameters)
        {
            return LSEXT.ResultFieldChange.rcAllow;
        }
        private  Int32 FractionalPart(double n)
        {
            string s = n.ToString("#.#########", System.Globalization.CultureInfo.InvariantCulture);
            return Int32.Parse(s.Substring(s.IndexOf(".") + 1));
        }
    }
}

