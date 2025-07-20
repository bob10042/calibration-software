using System;
using System.Threading;
using System.Windows.Forms;
using CalTestGUI.Services;

namespace CalTestGUI.Models.TestRunners
{
    public abstract class BaseTestRunner
    {
        protected readonly DeviceCommunicationService communicationService;
        protected readonly ExcelService excelService;
        protected double meas1, meas2, meas3, meas4, measLL, pps1;

        protected BaseTestRunner(DeviceCommunicationService communicationService, ExcelService excelService)
        {
            this.communicationService = communicationService;
            this.excelService = excelService;
        }

        protected void CustomRest()
        {
            Thread.Sleep(2000);
        }

        protected void Newtons4thControls()
        {
            communicationService.WriteN4l("*RST");
            communicationService.WriteN4l("TRG");
            communicationService.WriteN4l("KEYBOARD,DISABLE");
            communicationService.WriteN4l("SYST;POWER");
            communicationService.WriteN4l("COUPLI,PHASE1,AC+DC");
            communicationService.WriteN4l("COUPLI,PHASE2,AC+DC");
            communicationService.WriteN4l("COUPLI,PHASE3,AC+DC");
            communicationService.WriteN4l("APPLIC,NORMAL,SPEED3");
            communicationService.WriteN4l("BANDWI,WIDE");
            communicationService.WriteN4l("DATALO,DISABLE");
            communicationService.WriteN4l("RESOLU,HIGH");
            communicationService.WriteN4l("EFFICI,0");
            communicationService.WriteN4l("FAST,ON");
            communicationService.WriteN4l("FREQFI,OFF");
            communicationService.WriteN4l("SPEED,HIGH");
            communicationService.WriteN4l("DISPLAY,VOLTAGE");
            communicationService.WriteN4l("ZOOM,1,3,5,6,7");
            Thread.Sleep(5000);
            communicationService.WriteN4l("MULTIL,0");
            communicationService.WriteN4l("MULTIL,1,1,50");
            communicationService.WriteN4l("MULTIL,2,2,50");
            communicationService.WriteN4l("MULTIL,3,3,50");
            communicationService.WriteN4l("MULTIL,4,1,1");
            communicationService.WriteN4l("MULTIL,5,1,79");
        }

        protected void MeasurePhase(int phase, Action<double> setMeasurement)
        {
            double measurement = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                communicationService.WriteN4l("MULTIL?");
                var result = communicationService.ReadN4l().Split(',')[phase];
                if (double.TryParse(result, out double value))
                {
                    measurement += value;
                }
            }
            measurement /= 10;
            setMeasurement(measurement);
        }

        protected void MeasurePPSVoltage(string command, Action<double> setMeasurement)
        {
            double measurement = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                communicationService.WritePPS(command);
                var result = communicationService.ReadPPS();
                if (double.TryParse(result, out double value))
                {
                    measurement += value;
                }
            }
            measurement /= 10;
            setMeasurement(measurement);
        }

        protected void SetVoltageAndMeasure(int row, double voltage, Action measureAction)
        {
            communicationService.WritePPS($":VOLT,{voltage}");
            Thread.Sleep(15000);
            measureAction();
        }

        protected void SetFrequencyAndMeasure(int row, double frequency, Action measureAction)
        {
            communicationService.WritePPS($":FREQ,{frequency}");
            Thread.Sleep(15000);
            measureAction();
        }

        protected abstract void ConfigureDevice(string userFreq);
        public abstract void RunTest(TestConfiguration config);
    }
}
