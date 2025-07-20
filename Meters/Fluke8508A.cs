using System;
using System.Threading;

namespace CalTestGUI
{
    public class Fluke8508A : BaseMeter
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
                throw new Exception($"Failed to get identification from Fluke 8508A: {ex.Message}");
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
                throw new Exception($"Failed to measure {parameter} with Fluke 8508A: {ex.Message}");
            }
        }

        public void SetResolution(int digits)
        {
            try
            {
                if (digits < 4 || digits > 8)
                    throw new ArgumentException("Resolution must be between 4 and 8 digits");
                
                Write($"RESOLUTION {digits}");
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set resolution on Fluke 8508A: {ex.Message}");
            }
        }

        public void SelectInput(string input)
        {
            try
            {
                switch (input.ToUpper())
                {
                    case "FRONT":
                        Write("INPUT FRONT");
                        break;
                    case "REAR":
                        Write("INPUT REAR");
                        break;
                    case "SCAN":
                        Write("INPUT SCAN");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported input selection: {input}");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to select input on Fluke 8508A: {ex.Message}");
            }
        }

        public void ConfigureFilter(bool enable, int count = 10)
        {
            try
            {
                if (enable)
                {
                    Write($"FILTER ON {count}");
                }
                else
                {
                    Write("FILTER OFF");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to configure filter on Fluke 8508A: {ex.Message}");
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
                throw new Exception($"Failed to reset Fluke 8508A: {ex.Message}");
            }
        }

        public void SetRange(string parameter, double range)
        {
            try
            {
                switch (parameter.ToUpper())
                {
                    case "VOLTAGE:DC":
                        Write($"RANGE VDC, {range}");
                        break;
                    case "VOLTAGE:AC":
                        Write($"RANGE VAC, {range}");
                        break;
                    case "CURRENT:DC":
                        Write($"RANGE IDC, {range}");
                        break;
                    case "CURRENT:AC":
                        Write($"RANGE IAC, {range}");
                        break;
                    case "RESISTANCE":
                        Write($"RANGE OHMS, {range}");
                        break;
                    default:
                        throw new ArgumentException($"Unsupported range parameter: {parameter}");
                }
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to set {parameter} range on Fluke 8508A: {ex.Message}");
            }
        }

        public void ConfigureAutoZero(bool enable)
        {
            try
            {
                Write(enable ? "AUTO ON" : "AUTO OFF");
                Thread.Sleep(200);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to configure auto zero on Fluke 8508A: {ex.Message}");
            }
        }

        public void Calibrate()
        {
            try
            {
                Write("CAL");
                Thread.Sleep(5000); // Allow time for calibration to complete
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to calibrate Fluke 8508A: {ex.Message}");
            }
        }
    }
}
