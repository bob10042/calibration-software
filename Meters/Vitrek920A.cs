using System;
using System.Threading;

namespace CalTestGUI
{
    public class Vitrek920A : BaseMeter
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
                throw new Exception($"Failed to get identification from Vitrek 920A: {ex.Message}");
            }
        }

        public override string Measure(string parameter)
        {
            try
            {
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE":
                        Write("MEAS:VOLT?");
                        break;
                    case "CURRENT":
                        Write("MEAS:CURR?");
                        break;
                    case "POWER":
                        Write("MEAS:POW?");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported measurement parameter: {parameter}");
                }
                return Read();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to measure {parameter} with Vitrek 920A: {ex.Message}");
            }
        }

        public void ConfigureForPowerMeasurement()
        {
            try
            {
                Write("CONF:POW");
                Thread.Sleep(500); // Allow time for configuration to take effect
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to configure Vitrek 920A for power measurement: {ex.Message}");
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
                throw new Exception($"Failed to reset Vitrek 920A: {ex.Message}");
            }
        }

        public void SetRange(string parameter, double range)
        {
            try
            {
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE":
                        Write($"VOLT:RANG {range}");
                        break;
                    case "CURRENT":
                        Write($"CURR:RANG {range}");
                        break;
                    case "POWER":
                        Write($"POW:RANG {range}");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported range parameter: {parameter}");
                }
                Thread.Sleep(200); // Allow time for range change to take effect
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set {parameter} range on Vitrek 920A: {ex.Message}");
            }
        }
    }
}
