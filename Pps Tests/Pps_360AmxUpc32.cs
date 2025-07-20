﻿using System;
using CalTestGUI.Pps360AmxUpc32;

namespace CalTestGUI
{
    public class Pps_360AmxUpc32
    {
        private readonly Pps360AmxUpc32Communication _communication;
        private readonly Pps360AmxUpc32Core _core;
        private readonly Pps360AmxUpc32Controls _controls;
        private readonly Pps360AmxUpc32Measurements _measurements;
        private readonly Pps360AmxUpc32TestExecution _testExecution;

        public Pps_360AmxUpc32()
        {
            _communication = new Pps360AmxUpc32Communication();
            _core = new Pps360AmxUpc32Core(_communication);
            _controls = new Pps360AmxUpc32Controls(_communication, _core);
            _measurements = new Pps360AmxUpc32Measurements(_communication);
            _testExecution = new Pps360AmxUpc32TestExecution(_communication, _core, _controls, _measurements);
        }

        public string VisaName
        {
            get => _core.VisaName;
            set => _core.VisaName = value;
        }

        public string TestFreq
        {
            get => _core.TestFreq;
            set => _core.TestFreq = value;
        }

        public string Company
        {
            get => _core.Company;
            set => _core.Company = value;
        }

        public string CertNum
        {
            get => _core.CertNum;
            set => _core.CertNum = value;
        }

        public string CaseNum
        {
            get => _core.CaseNum;
            set => _core.CaseNum = value;
        }

        public string SerNum
        {
            get => _core.SerNum;
            set => _core.SerNum = value;
        }

        public string UserFreq
        {
            get => _core.UserFreq;
            set => _core.UserFreq = value;
        }

        public string PlantNum
        {
            get => _core.PlantNum;
            set => _core.PlantNum = value;
        }

        public string PortName_Newtowns_4th
        {
            get => _communication.PortName_Newtowns_4th;
            set => _communication.PortName_Newtowns_4th = value;
        }

        public bool IsOpen_Newtowns_4th => _communication.IsOpen_Newtowns_4th;

        public void OpenSession()
        {
            _core.Initialize();
        }

        public void Open_Newtowns_4th()
        {
            _communication.Open_Newtowns_4th();
        }

        public void ClosePort()
        {
            _communication.ClosePort();
        }

        public void Close()
        {
            _communication.Close();
        }

        public void RunTest()
        {
            _testExecution.RunTest();
        }
    }
}
