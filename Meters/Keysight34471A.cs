using System;
using System.Threading;

namespace CalTestGUI
{
    public class Keysight34471A : BaseMeter
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
                throw new Exception($"Failed to get identification from Keysight 34471A: {ex.Message}");
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
                throw new Exception($"Failed to measure {parameter} with Keysight 34471A: {ex.Message}");
            }
        }

        public void ConfigureResolution(string parameter, double resolution)
        {
            try
            {
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE:DC":
                        Write($"VOLT:DC:RES {resolution}");
                        break;
                    case "VOLTAGE:AC":
                        Write($"VOLT:AC:RES {resolution}");
                        break;
                    case "CURRENT:DC":
                        Write($"CURR:DC:RES {resolution}");
                        break;
                    case "CURRENT:AC":
                        Write($"CURR:AC:RES {resolution}");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported resolution parameter: {parameter}");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to configure resolution for {parameter} on Keysight 34471A: {ex.Message}");
            }
        }

        public void SetIntegrationTime(double nplc)
        {
            try
            {
                Write($"SENS:VOLT:DC:NPLC {nplc}");
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set integration time on Keysight 34471A: {ex.Message}");
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
                throw new Exception($"Failed to reset Keysight 34471A: {ex.Message}");
            }
        }

        public void ConfigureAutoZero(bool enable)
        {
            try
            {
                string state = enable ? "ON" : "OFF";
                Write($"SENS:ZERO:AUTO {state}");
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to configure auto zero on Keysight 34471A: {ex.Message}");
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
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set {parameter} range on Keysight 34471A: {ex.Message}");
            }
        }

        public void SetAutoRange(string parameter, bool enable)
        {
            try
            {
                string state = enable ? "ON" : "OFF";
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE:DC":
                        Write($"VOLT:DC:RANG:AUTO {state}");
                        break;
                    case "VOLTAGE:AC":
                        Write($"VOLT:AC:RANG:AUTO {state}");
                        break;
                    case "CURRENT:DC":
                        Write($"CURR:DC:RANG:AUTO {state}");
                        break;
                    case "CURRENT:AC":
                        Write($"CURR:AC:RANG:AUTO {state}");
                        break;
                    case "RESISTANCE":
                        Write($"RES:RANG:AUTO {state}");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported parameter for autorange: {parameter}");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set autorange for {parameter} on Keysight 34471A: {ex.Message}");
            }
        }
    }
}
