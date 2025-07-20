using System;
using System.Threading;
using CalTestGUI.Services;

namespace CalTestGUI.Models.TestRunners
{
    public class ThreePhaseTestRunner : BaseTestRunner
    {
        public ThreePhaseTestRunner(DeviceCommunicationService communicationService, ExcelService excelService) 
            : base(communicationService, excelService)
        {
        }

        protected override void ConfigureDevice(string userFreq)
        {
            communicationService.WritePPS("*LLO");
            communicationService.WritePPS(":OUTP,OFF;:PROG:EXEC,0;:FREQ,60;:FREQ:SPAN,600");
            Thread.Sleep(8000);
            communicationService.WritePPS($":VOLT,0;:FORM,3;:FREQ,{userFreq?.Trim() ?? "60"}" +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void RunVoltageTest(int startRow)
        {
            // Initial measurement with longer wait time
            var voltage = excelService.GetCellValue(startRow, 2);
            if (voltage.HasValue)
            {
                SetVoltageAndMeasure(startRow, voltage.Value, () =>
                {
                    Thread.Sleep(30000); // Longer initial wait
                    MeasureAllPhases(startRow);
                });
            }

            // Subsequent measurements
            for (int row = startRow + 3; row < startRow + 24; row += 3)
            {
                voltage = excelService.GetCellValue(row, 2);
                if (voltage.HasValue)
                {
                    SetVoltageAndMeasure(row, voltage.Value, () =>
                    {
                        MeasureAllPhases(row);
                    });
                }
            }
        }

        private void MeasureAllPhases(int startRow)
        {
            // Measure Phase 1
            MeasurePPSVoltage(":MEAS:VOLT1?", value =>
            {
                pps1 = value;
                excelService.SetMeasurement(startRow, 4, value);
            });
            MeasurePhase(0, value =>
            {
                meas1 = value;
                excelService.SetMeasurement(startRow, 8, value);
            });

            // Measure Phase 2
            MeasurePPSVoltage(":MEAS:VOLT2?", value =>
            {
                pps1 = value;
                excelService.SetMeasurement(startRow + 1, 4, value);
            });
            MeasurePhase(1, value =>
            {
                meas2 = value;
                excelService.SetMeasurement(startRow + 1, 8, value);
            });

            // Measure Phase 3
            MeasurePPSVoltage(":MEAS:VOLT3?", value =>
            {
                pps1 = value;
                excelService.SetMeasurement(startRow + 2, 4, value);
            });
            MeasurePhase(2, value =>
            {
                meas3 = value;
                excelService.SetMeasurement(startRow + 2, 8, value);
            });
        }

        private void ConfigureFrequencyResponse()
        {
            communicationService.WritePPS(":OUTP,OFF;:PROG:EXEC,0;:FORM,3;:VOLT,0;:FREQ,50;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHASe2,120;" +
                ":PHASe3,240;:RANG,0;:RAMP,0;:VOLT,120;:OUTP,ON");
            CustomRest();
        }

        private void RunFrequencyTest(int startRow)
        {
            ConfigureFrequencyResponse();

            // Initial measurement with longer wait time
            var frequency = excelService.GetCellValue(startRow, 2);
            if (frequency.HasValue)
            {
                SetFrequencyAndMeasure(startRow, frequency.Value, () =>
                {
                    Thread.Sleep(30000); // Longer initial wait
                    MeasureFrequencyResponse(startRow);
                });
            }

            // Subsequent measurements
            for (int row = startRow + 3; row < startRow + 15; row += 3)
            {
                frequency = excelService.GetCellValue(row, 2);
                if (frequency.HasValue)
                {
                    SetFrequencyAndMeasure(row, frequency.Value, () =>
                    {
                        MeasureFrequencyResponse(row);
                    });
                }
            }
        }

        private void MeasureFrequencyResponse(int startRow)
        {
            // Measure Frequency
            MeasurePhase(3, value =>
            {
                meas4 = value;
                excelService.SetMeasurement(startRow, 4, value);
            });

            // Measure Phase 1
            MeasurePhase(0, value =>
            {
                meas1 = value;
                excelService.SetMeasurement(startRow, 8, value);
            });

            // Measure Phase 2
            MeasurePhase(1, value =>
            {
                meas2 = value;
                excelService.SetMeasurement(startRow + 1, 8, value);
            });

            // Measure Phase 3
            MeasurePhase(2, value =>
            {
                meas3 = value;
                excelService.SetMeasurement(startRow + 2, 8, value);
            });
        }

        public override void RunTest(TestConfiguration config)
        {
            Newtons4thControls();
            ConfigureDevice(config.UserFreq);

            // Run voltage tests starting at row 77
            RunVoltageTest(77);

            // Run frequency response tests starting at row 226
            RunFrequencyTest(226);

            // Configure high frequency span
            communicationService.WritePPS(":OUTP,OFF");
            CustomRest();
            communicationService.WritePPS(":FREQ:SPAN,1200");
            Thread.Sleep(8000);
            communicationService.WritePPS(":FREQ,1000");
            communicationService.WritePPS(":OUTP,ON");
            CustomRest();

            // Run high frequency tests starting at row 235
            RunFrequencyTest(235);

            // Reset frequency settings
            communicationService.WritePPS(":OUTP,OFF");
            CustomRest();
            communicationService.WritePPS(":FREQ,60");
            CustomRest();
            communicationService.WritePPS(":FREQ:SPAN,600");
            Thread.Sleep(8000);
        }
    }
}
