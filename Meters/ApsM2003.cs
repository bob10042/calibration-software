using System;
using System.Threading;

namespace CalTestGUI
{
    public class ApsM2003 : BaseMeter
    {
        public override string GetIdentification()
        {
            try
            {
                Write("*IDN?");
                return Read();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to get identification from APS M2003: {ex.Message}");
            }
        }

        public override string Measure(string parameter)
        {
            try
            {
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE:DC":
                        Write("MEAS:VOLT:DC?");
                        break;
                    case "VOLTAGE:AC":
                        Write("MEAS:VOLT:AC?");
                        break;
                    case "CURRENT:DC":
                        Write("MEAS:CURR:DC?");
                        break;
                    case "CURRENT:AC":
                        Write("MEAS:CURR:AC?");
                        break;
                    case "RESISTANCE":
                        Write("MEAS:RES?");
                        break;
                    case "FREQUENCY":
                        Write("MEAS:FREQ?");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported measurement parameter: {parameter}");
                }
                return Read();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to measure {parameter} with APS M2003: {ex.Message}");
            }
        }

        public void SetRange(string parameter, double range)
        {
            try
            {
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE:DC":
                        Write($"VOLT:DC:RANG {range}");
                        break;
                    case "VOLTAGE:AC":
                        Write($"VOLT:AC:RANG {range}");
                        break;
                    case "CURRENT:DC":
                        Write($"CURR:DC:RANG {range}");
                        break;
                    case "CURRENT:AC":
                        Write($"CURR:AC:RANG {range}");
                        break;
                    case "RESISTANCE":
                        Write($"RES:RANG {range}");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported range parameter: {parameter}");
                }
                Thread.Sleep(200); // Allow time for range change to take effect
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set {parameter} range on APS M2003: {ex.Message}");
            }
        }

        public void Reset()
        {
            try
            {
                Write("*RST");
                Thread.Sleep(1000); // Allow time for reset to complete
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to reset APS M2003: {ex.Message}");
            }
        }

        public void ConfigureFilter(bool enable, int count = 10)
        {
            try
            {
                if (enable)
                {
                    Write($"SENS:AVER:COUN {count}");
                    Write("SENS:AVER ON");
                }
                else
                {
                    Write("SENS:AVER OFF");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to configure filter on APS M2003: {ex.Message}");
            }
        }

        public void SetAutoRange(string parameter, bool enable)
        {
            try
            {
                string command = enable ? "ON" : "OFF";
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE:DC":
                        Write($"VOLT:DC:RANG:AUTO {command}");
                        break;
                    case "VOLTAGE:AC":
                        Write($"VOLT:AC:RANG:AUTO {command}");
                        break;
                    case "CURRENT:DC":
                        Write($"CURR:DC:RANG:AUTO {command}");
                        break;
                    case "CURRENT:AC":
                        Write($"CURR:AC:RANG:AUTO {command}");
                        break;
                    case "RESISTANCE":
                        Write($"RES:RANG:AUTO {command}");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported parameter for autorange: {parameter}");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set autorange for {parameter} on APS M2003: {ex.Message}");
            }
        }
    }
}
