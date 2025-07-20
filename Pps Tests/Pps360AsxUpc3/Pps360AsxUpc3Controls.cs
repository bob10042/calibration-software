using System.Threading;

namespace CalTestGUI.Pps360AsxUpc3
{
    public partial class Pps360AsxUpc3Core
    {
        protected void Newtons4thControls()
        {
            WriteN4l("*RST");
            WriteN4l("TRG");                      // initiates a new measurement, resets the range and smoothing
            WriteN4l("KEYBOARD,DISABLE");         // disable keyboard
            WriteN4l("SYST;POWER");               // selects power meter
            WriteN4l("COUPLI,PHASE1,AC+DC");      // coupling ac + dc
            WriteN4l("COUPLI,PHASE2,AC+DC");      // coupling ac + dc
            WriteN4l("COUPLI,PHASE3,AC+DC");      // coupling ac + dc
            WriteN4l("APPLIC,NORMAL,SPEED3");     // select application mode
            WriteN4l("BANDWI,WIDE");              // bandwidth @ 3MHz
            WriteN4l("DATALO,DISABLE");           // datalog disabled
            WriteN4l("RESOLU,HIGH");              // sets the data resolution                             
            WriteN4l("EFFICI,0");                 // efficiency calculation disabled
            WriteN4l("FAST,ON");                  // this will disable the screen for faster resolutions 
            WriteN4l("FREQFI,OFF");               // turns the frequency filter off                     
            WriteN4l("SPEED,HIGH");               // slows the speed down 
            WriteN4l("DISPLAY,VOLTAGE");          // display voltage mode
            WriteN4l("ZOOM,1,3,5,6,7");           // sets the zoom level for the display
            Thread.Sleep(5000);
            WriteN4l("MULTIL,0");                 // reset multilogs 
            WriteN4l("MULTIL,1,1,50");            // index 1 == phase1, rms voltage
            WriteN4l("MULTIL,2,2,50");            // index 2 == phase2, rms voltage
            WriteN4l("MULTIL,3,3,50");            // index 3 == phase3, rms voltage       
            WriteN4l("MULTIL,4,1,1");             // index 4 == phase1, frequency
            WriteN4l("MULTIL,5,1,79");            // index 5 == phase-phase, rms voltage 
        }

        protected void ThreePhaseControls()
        {
            Write_PPS("*LLO");
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FREQ,60;:FREQ:SPAN,600");
            Thread.Sleep(8000);
            Write_PPS(":VOLT,0;:FORM,3;:FREQ," + (UserFreq?.Trim() ?? "60") +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        protected void SpitPhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,2;:FREQ,50;" +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        protected void SinglePhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,1;:FREQ," + (UserFreq?.Trim() ?? "60") +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        protected void ThreePhaseFreqRespControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FORM,3;:VOLT,0;:FREQ,50;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHASe2,120;" +
                ":PHASe3,240;:RANG,0;:RAMP,0;:VOLT,120;:OUTP,ON");
            CustomRest();
        }
    }
}
