using System;
using System.IO;
using System.IO.Ports;
using System.Windows.Forms;
using NationalInstruments.Visa;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Threading;
using System.Data;
using System.Drawing;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace CalTestGUI    
{
    public partial class frmMain : Form
    {
        GWI myGWI = new GWI();
        Pps_360AmxUpc_32 my_360AMX_UPC32 = new Pps_360AmxUpc_32();
        Pps_360AsxUpc_3 my_360ASX_UPC3 = new Pps_360AsxUpc_3();
        Pps_3150Afx my_3150AFX = new Pps_3150Afx();
        Pps_118AcxUpc1 my_118ACX_UPC1 = new Pps_118AcxUpc1();
        Pps_115AcxUpc_1 my_115ACX_UPC1 = new Pps_115AcxUpc_1();

        private string[] ports = SerialPort.GetPortNames();

        public frmMain()
        {
            InitializeComponent();
            customizeDesign();
            HidePanels();
            panelBckGrd.Dock = DockStyle.Fill;
            panelBckGrd.Show();
        }

        void Form1_Load(object sender, EventArgs e)
        {                            
            CloseAllPorts();
        }
        
        private void CloseAllPorts()
        {
            myGWI.ClosePort();
            my_360ASX_UPC3.ClosePort();
            my_360AMX_UPC32.ClosePort();
            my_3150AFX.ClosePort();
            my_118ACX_UPC1.ClosePort();
            my_115ACX_UPC1.ClosePort();
            RichTextPpsAte.Clear();
            RichTextPpsAte.ScrollBars = RichTextBoxScrollBars.None;
            progBarPpsGpib.Value = 0;
            progBarPatTest.Value = 0;
        }

        private void customizeDesign()
        {
            panelCalibrationMenu.Visible = false;
            panelProductMenu.Visible = false;
            panelAboutUsMenu.Visible = false;
            panelMetersMenu.Visible = false;
        }

        private void hideSubMenu()
        {
            if (panelCalibrationMenu.Visible == true)
                panelCalibrationMenu.Visible = false;
            if (panelProductMenu.Visible == true)
                panelProductMenu.Visible = false;
            if (panelAboutUsMenu.Visible == true)
                panelAboutUsMenu.Visible = false;
            if (panelMetersMenu.Visible == true)
                panelMetersMenu.Visible = false;
        }
        private void showSubMenu(Panel subMenu)
        {
            if (subMenu.Visible == false)
            {
                hideSubMenu();
                subMenu.Visible = true;
            }
            else subMenu.Visible = false;
        }        
       
        private void btnPpsPanel_Click(object sender, EventArgs e)
        {
            CmbBoxN4LPrt.Items.AddRange(ports);
            //Adds Visa Adreess to pacific source combo box items 
            using (var rmSession = new ResourceManager())
            {
                try
                {
                    var resources = rmSession.Find("(ASRL|GPIB|TCPIP|USB)?*");
                    foreach (string s in resources)
                    {
                        CmbBoxPacificSelect.Items.Add(s);
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            HidePanels();
            GrBoxPpsCalSet.Dock = DockStyle.Fill;
            RichTextPpsAte.Clear();
            RichTextPpsAte.ScrollBars = RichTextBoxScrollBars.None;
            RichTextPpsAte.ScrollToCaret();
            GrBoxPpsCalSet.Show();
            RichTextPpsAte.Multiline = true;
            RichTextPpsAte.WordWrap = true;
            RichTextPpsAte.Visible = true;
            RadioBtnPps400Hz.Checked = true;
            radioBtnNoSoak.Checked = true;
        }
             
        private void btnAcPowerSuppliesPanel_Click(object sender, EventArgs e)
        {
            hideSubMenu();
        }
        
        private void btnMCodesPanel_Click(object sender, EventArgs e)
        {
            hideSubMenu();
        }
        
        private void btnAssemblyDrawingsPanel_Click(object sender, EventArgs e)
        {
            hideSubMenu();
        } 
                
        private void btnPowerRatingsPanel_Click(object sender, EventArgs e)
        {
            hideSubMenu();
        }
        
        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void panelMainDisplay_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
       
        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void btnHomeScreenPanel_Click_3(object sender, EventArgs e)
        {
            hideSubMenu();
            HidePanels();
            panelBckGrd.Show();
            panelBckGrd.Dock = DockStyle.Fill;
            CloseAllPorts();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void panelBckGrd_Paint(object sender, PaintEventArgs e)
        {

        }
        private void HidePanels()
        {
            panelBckGrd.Hide();
            panelPatTest.Hide();
            GrBoxPpsCalSet.Hide();
            panelMetersLanding.Hide();
        }
               
        private void panelPpsTst_Paint(object sender, PaintEventArgs e)
        {

        }
       
        private void panelPpsTst_Paint_1(object sender, PaintEventArgs e)
        {

        }
                                   
        private void panelPpsTstGpib_Paint(object sender, PaintEventArgs e)
        {
            
        }
      
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
　　 　 　 　
        private void panelCalibrationMenu_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnPatTestPanel_Click_2(object sender, EventArgs e)
        {
            showSubMenu(panelCalibrationMenu);
        }

        private void btnPatTest_Click(object sender, EventArgs e)
        {
            cmbBoxGwInstekUsb.Items.AddRange(ports);
            HidePanels();
            panelPatTest.Dock = DockStyle.Fill;
            panelPatTest.Show();            
        }
       
        private void btnCircuitDiagramsPanel_Click(object sender, EventArgs e)
        {
            showSubMenu(panelProductMenu);
        }

        private void btnProductSpecificationPanel_Click(object sender, EventArgs e)
        {
            showSubMenu(panelAboutUsMenu);
        }
        
        private void cmbBxSltUut_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void cmbBoxGwInstekUsb_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if ((cmbBoxGwInstekUsb.Text != "") && (cmbSupplyVoltage.Text != "") && (cmbCurrRating.Text != "")
                && (cmbCableLength.Text != ""))
            {
                btnPatConnect.Enabled = true;
            }
            else { btnPatConnect.Enabled = false; }
        }

        private void btnPatConnect_Click(object sender, EventArgs e)
        {
            myGWI.ClosePort();
            try
            {               
                myGWI.PortName = cmbBoxGwInstekUsb.Text;
                myGWI.Open();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            btnPatConnect.Enabled = false;
            btnPatClose.Enabled = true;
            btnClassOne.Enabled = true;
            btnClassTwo.Enabled = true;
            btnIecLeadTest.Enabled = true;
            cmbBoxGwInstekUsb.Enabled = false;
            cmbSupplyVoltage.Enabled = false;
            cmbCurrRating.Enabled = false;
            cmbCableLength.Enabled = false;
            progBarPatTest.Value = 100;
        }        

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnPatClose_Click(object sender, EventArgs e)
        {
            myGWI.ClosePort();
            btnPatClose.Enabled = false;
            btnClassOne.Enabled = false;
            btnClassTwo.Enabled = false;
            btnIecLeadTest.Enabled = false;
            btnPatConnect.Enabled = true;
            cmbBoxGwInstekUsb.Enabled = true;
            cmbSupplyVoltage.Enabled = true;
            cmbCurrRating.Enabled = true;
            cmbCableLength.Enabled = true;
        }

        private void panelPatTest_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void groupBox7_Enter(object sender, EventArgs e)
        {

        }
         
        private void pictureBox6_Click(object sender, EventArgs e)
        {

        }

        private void btnApsPanel_Click(object sender, EventArgs e)
        {

        }

        private void btnSpare1_Click(object sender, EventArgs e)
        {

        }

        private void btnSpare4_Click(object sender, EventArgs e)
        {

        }

        private void btnPowerSuppliesPanel_Click(object sender, EventArgs e)
        {

        }

        private void btnElectronicLoadsPanel_Click(object sender, EventArgs e)
        {

        }
                          
        public void cmbCableLength_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            myGWI.ResLimit = cmbCableLength.Text;
            if ((cmbBoxGwInstekUsb.Text != "") && (cmbSupplyVoltage.Text != "") && (cmbCurrRating.Text != "")
                && (cmbCableLength.Text != ""))
            {
                btnPatConnect.Enabled = true;
            }
            else { btnPatConnect.Enabled = false; }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void cmbCurrRating_SelectedIndexChanged(object sender, EventArgs e)
        {
            myGWI.CurrRating = cmbCurrRating.Text;
            if ((cmbBoxGwInstekUsb.Text != "") && (cmbSupplyVoltage.Text != "") && (cmbCurrRating.Text != "")
                && (cmbCableLength.Text != ""))
            {
                btnPatConnect.Enabled = true;
            }
            else { btnPatConnect.Enabled = false; }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
               
        private void cmbSupplyVoltage_SelectedIndexChanged(object sender, EventArgs e)
        {
            myGWI.VoltRating = cmbSupplyVoltage.Text;
            if ((cmbBoxGwInstekUsb.Text != "") && (cmbSupplyVoltage.Text != "") && (cmbCurrRating.Text != "")
                && (cmbCableLength.Text != ""))
            {
                btnPatConnect.Enabled = true;
            }
            else { btnPatConnect.Enabled = false; }
        }

        private void groupBox10_Enter(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnClassOne_Click(object sender, EventArgs e)
        {
            myGWI.Class_1();
            CloseAllPorts();
            btnPatClose.Enabled = false;
            btnClassOne.Enabled = false;
            btnClassTwo.Enabled = false;
            btnIecLeadTest.Enabled = false;
            btnPatConnect.Enabled = true;
            cmbBoxGwInstekUsb.Enabled = true;
            cmbSupplyVoltage.Enabled = true;
            cmbCurrRating.Enabled = true;
            cmbCableLength.Enabled = true;
        }
        private void btnClassTwo_Click(object sender, EventArgs e)
        {
            myGWI.Class_2();
            CloseAllPorts();
            btnPatClose.Enabled = false;
            btnClassOne.Enabled = false;
            btnClassTwo.Enabled = false;
            btnIecLeadTest.Enabled = false;
            btnPatConnect.Enabled = true;
            cmbBoxGwInstekUsb.Enabled = true;
            cmbSupplyVoltage.Enabled = true;
            cmbCurrRating.Enabled = true;
            cmbCableLength.Enabled = true;
        }

        private void btnIecLeadTest_Click(object sender, EventArgs e)
        {
            myGWI.IecLeadTest();
            CloseAllPorts();
            btnPatClose.Enabled = false;
            btnClassOne.Enabled = false;
            btnClassTwo.Enabled = false;
            btnIecLeadTest.Enabled = false;
            btnPatConnect.Enabled = true;
            cmbBoxGwInstekUsb.Enabled = true;
            cmbSupplyVoltage.Enabled = true;
            cmbCurrRating.Enabled = true;
            cmbCableLength.Enabled = true;
        }
        private void groupBox9_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox10_Enter_1(object sender, EventArgs e)
        {

        }
                               
        private void panelPpsTst_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbBoxN4LPrtComPort_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void BtnPpsClose_Click(object sender, EventArgs e)
        {
            BtnPpsClose.Enabled = false; 
            BtnPpsConnect.Enabled = true;
            btnPpsRunTest.Enabled = false;
            RadioBtnPps400Hz.Enabled = true;
            RadioBtnPps50Hz.Enabled = true;
            RadioBtnPps400Hz.Enabled = true;
            RadioBtnPps50Hz.Enabled = true;
            CmbBoxPacificSelect.Enabled = true;
            CmbBoxN4LPrt.Enabled = true;
            CmbBoxPacificUnit.Enabled = true;
            CmbBoxTrans.Enabled = true;
            TextBoxPpsEnterPacificSerNum.Enabled = true;
            TextBoxPpsEnterCertNum.Enabled = true;
            textBoxPpsEnterCaseNum.Enabled = true;
            textBoxPpsEnterName.Enabled = true;
            radioBtnNoSoak.Enabled = true;
            radioBtnPreSoak.Enabled = true;
            radioBtnNoSoak.Enabled = true;
            radioBtnPreSoak.Enabled = true;
            textBoxPpsEnterPlantNum.Enabled = true;
            CloseAllPorts();
        }

        private void progBarPpsGpib_Click(object sender, EventArgs e)
        {

        }

        private void BtnPpsConnect_Click(object sender, EventArgs e)
        {
            CloseAllPorts();
            RichTextPpsAte.Clear();
            RichTextPpsAte.ScrollBars = RichTextBoxScrollBars.None;            
            try
            {
                if (CmbBoxPacificUnit.Text == "360_AMX_UPC32")
                {
                    my_360AMX_UPC32.PortName_Newtowns_4th = CmbBoxN4LPrt.Text;
                    my_360AMX_UPC32.OpenSession();
                    my_360AMX_UPC32.Open_Newtowns_4th();
                }

                else if (CmbBoxPacificUnit.Text == "360_ASX_UPC3")
                {
                    my_360ASX_UPC3.PortName_Newtowns_4th = CmbBoxN4LPrt.Text;
                    my_360ASX_UPC3.OpenSession();
                    my_360ASX_UPC3.Open_Newtowns_4th();
                }
                else if (CmbBoxPacificUnit.Text == "3150_AFX")
                { 
                    my_3150AFX.PortName_Newtowns_4th = CmbBoxN4LPrt.Text;
                    my_3150AFX.OpenSession();
                    my_3150AFX.Open_Newtowns_4th();
                }
                else if (CmbBoxPacificUnit.Text == "118_ACX_UCP1")
                {
                    my_118ACX_UPC1.PortName_Newtowns_4th = CmbBoxN4LPrt.Text;
                    my_118ACX_UPC1.OpenSession();
                    my_118ACX_UPC1.Open_Newtowns_4th();
                }
                else if (CmbBoxPacificUnit.Text == "115_ACX_UCP1")
                {
                    my_115ACX_UPC1.PortName_Newtowns_4th = CmbBoxN4LPrt.Text;
                    my_115ACX_UPC1.OpenSession();
                    my_115ACX_UPC1.Open_Newtowns_4th();
                }
                else
                { }　　 　 　 　
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            RichTextBoxEntries();          
            BtnPpsClose.Enabled = true;
            BtnPpsConnect.Enabled = false;
            btnPpsRunTest.Enabled = true;
            RadioBtnPps400Hz.Enabled = false;
            RadioBtnPps50Hz.Enabled = false;
            CmbBoxPacificSelect.Enabled = false;
            CmbBoxN4LPrt.Enabled = false;
            CmbBoxPacificUnit.Enabled = false;
            CmbBoxTrans.Enabled = false;
            TextBoxPpsEnterPacificSerNum.Enabled = false;
            TextBoxPpsEnterCertNum.Enabled = false;
            textBoxPpsEnterCaseNum.Enabled = false;
            textBoxPpsEnterName.Enabled = false;
            radioBtnNoSoak.Enabled = false;
            radioBtnPreSoak.Enabled = false;
            radioBtnNoSoak.Enabled = false;
            radioBtnPreSoak.Enabled = false;
            textBoxPpsEnterPlantNum.Enabled = false;
            btnPpsRunTest.Enabled = true;
            progBarPpsGpib.Value = 100;
        }

        private void RichTextBoxEntries()
        {
            my_115ACX_UPC1.Company = textBoxPpsEnterName.Text;
            my_115ACX_UPC1.CertNum = TextBoxPpsEnterCertNum.Text;
            my_115ACX_UPC1.CaseNum = textBoxPpsEnterCaseNum.Text;
            my_115ACX_UPC1.serNum = TextBoxPpsEnterPacificSerNum.Text;
            my_115ACX_UPC1.PlantNum = textBoxPpsEnterPlantNum.Text;

            my_118ACX_UPC1.Company = textBoxPpsEnterName.Text;
            my_118ACX_UPC1.CertNum = TextBoxPpsEnterCertNum.Text;
            my_118ACX_UPC1.CaseNum = textBoxPpsEnterCaseNum.Text;
            my_118ACX_UPC1.serNum = TextBoxPpsEnterPacificSerNum.Text;
            my_118ACX_UPC1.PlantNum = textBoxPpsEnterPlantNum.Text;

            my_360AMX_UPC32.Company = textBoxPpsEnterName.Text;
            my_360AMX_UPC32.CertNum = TextBoxPpsEnterCertNum.Text;
            my_360AMX_UPC32.CaseNum = textBoxPpsEnterCaseNum.Text;
            my_360AMX_UPC32.serNum = TextBoxPpsEnterPacificSerNum.Text;
            my_360AMX_UPC32.PlantNum = textBoxPpsEnterPlantNum.Text;

            my_360ASX_UPC3.Company = textBoxPpsEnterName.Text;
            my_360ASX_UPC3.CertNum = TextBoxPpsEnterCertNum.Text;
            my_360ASX_UPC3.CaseNum = textBoxPpsEnterCaseNum.Text;
            my_360ASX_UPC3.serNum = TextBoxPpsEnterPacificSerNum.Text;
            my_360ASX_UPC3.PlantNum = textBoxPpsEnterPlantNum.Text;

            my_3150AFX.Company = textBoxPpsEnterName.Text;
            my_3150AFX.CertNum = TextBoxPpsEnterCertNum.Text;
            my_3150AFX.CaseNum = textBoxPpsEnterCaseNum.Text;
            my_3150AFX.serNum = TextBoxPpsEnterPacificSerNum.Text;
            my_3150AFX.PlantNum = textBoxPpsEnterPlantNum.Text;
        }

        private void btnPpsRunTest_Click(object sender, EventArgs e)
        {
            btnPpsRunTest.Enabled = false;
            if (CmbBoxPacificUnit.Text == "360_AMX_UPC32")
            {
                my_360AMX_UPC32.RunTest();
            }

            else if (CmbBoxPacificUnit.Text == "360_ASX_UPC3")
            {
                my_360ASX_UPC3.RunTest();
            }
            else if (CmbBoxPacificUnit.Text == "3150_AFX")
            {
                my_3150AFX.RunTest();
            }
            else if (CmbBoxPacificUnit.Text == "118_ACX_UCP1")
            {
                my_118ACX_UPC1.RunTest();
            }
            else if (CmbBoxPacificUnit.Text == "115_ACX_UCP1")
            {
                my_115ACX_UPC1.RunTest();
            }
            else
            { }
            CloseAllPorts();
        }

        private void CmbBoxN4LPrt_Click(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }              

        private void cmbBoxSelectComm_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CmbBoxPacificSelect_Click(object sender, EventArgs e)
        {
            if (CmbBoxPacificUnit.Text == "360_AMX_UPC32")
            {
                my_360AMX_UPC32.visaName = CmbBoxPacificSelect.Text;
            }

            else if (CmbBoxPacificUnit.Text == "360_ASX_UPC3")
            {
                my_360ASX_UPC3.visaName = CmbBoxPacificSelect.Text;
            }
            else if (CmbBoxPacificUnit.Text == "3150_AFX")
            { 
                my_3150AFX.visaName = CmbBoxPacificSelect.Text;
            }
            else if (CmbBoxPacificUnit.Text == "118_ACX_UCP1")
            {
                my_118ACX_UPC1.visaName = CmbBoxPacificSelect.Text;
            }
            else if (CmbBoxPacificUnit.Text == "115_ACX_UCP1")
            {
                my_115ACX_UPC1.visaName = CmbBoxPacificSelect.Text;
            }
            else
            { }

            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }

            CloseAllPorts();
        }

        private void CmbBoxPacificUnit_Click(object sender, EventArgs e)
        {
            if (CmbBoxPacificUnit.Text == "360_AMX_UPC32")
            {
                my_360AMX_UPC32.visaName = CmbBoxPacificSelect.Text;
            }
          
            else if (CmbBoxPacificUnit.Text == "360_ASX_UPC3")
            {
                my_360ASX_UPC3.visaName = CmbBoxPacificSelect.Text;
            }
            else if (CmbBoxPacificUnit.Text == "3150_AFX")
            {
                my_3150AFX.visaName = CmbBoxPacificSelect.Text;
            }
            else if (CmbBoxPacificUnit.Text == "118_ACX_UCP1")
            {
                my_118ACX_UPC1.visaName = CmbBoxPacificSelect.Text;
            }
            else if (CmbBoxPacificUnit.Text == "115_ACX_UCP1")
            {
                my_115ACX_UPC1.visaName = CmbBoxPacificSelect.Text;
            }
            else
            { }

            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        private void panelPpsAte_Paint(object sender, PaintEventArgs e)
        {

        }

        private void TextBoxAddress_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBoxPpsN4l_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBoxPpsSelectUut_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBoxPpsTrans_TextChanged(object sender, EventArgs e)
        {

        }

        private void CmbBoxTrans_Click(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        public void RichTextPpsATE_TextChanged(object sender, EventArgs e)
        {                                   
                       
        }

        private void GrBoxPpsAteData_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        public void RadioBtnPps400Hz_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioBtnPps50Hz.Checked == true)
            {
                my_360AMX_UPC32.TestFreq = " "; my_360AMX_UPC32.UserFreq = "60";
                my_360ASX_UPC3.TestFreq = " "; my_360ASX_UPC3.UserFreq = "60";
                my_3150AFX.TestFreq = " "; my_3150AFX.UserFreq = "60";
                my_118ACX_UPC1.TestFreq = " "; my_118ACX_UPC1.UserFreq = "60";
                my_115ACX_UPC1.TestFreq = " "; my_115ACX_UPC1.UserFreq = "60";
            }
            else if (RadioBtnPps400Hz.Checked == true)
            {
                my_360AMX_UPC32.TestFreq = "A"; my_360AMX_UPC32.UserFreq = "400";
                my_360ASX_UPC3.TestFreq = "A"; my_360ASX_UPC3.UserFreq = "400";
                my_3150AFX.TestFreq = "A"; my_3150AFX.UserFreq = "400";
                my_118ACX_UPC1.TestFreq = "A"; my_118ACX_UPC1.UserFreq = "400";
                my_115ACX_UPC1.TestFreq = "A"; my_115ACX_UPC1.UserFreq = "400";
            }
            else { }
        }       

        public void RadioBtnPps50Hz_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioBtnPps50Hz.Checked == true)
            {
                my_360AMX_UPC32.TestFreq = " "; my_360AMX_UPC32.UserFreq = "60";
                my_360ASX_UPC3.TestFreq = " " ; my_360ASX_UPC3.UserFreq = "60";
                my_3150AFX.TestFreq = " " ; my_3150AFX.UserFreq = "60";
                my_118ACX_UPC1.TestFreq = " "; my_118ACX_UPC1.UserFreq = "60";
                my_115ACX_UPC1.TestFreq = " "; my_115ACX_UPC1.UserFreq = "60";
            }
            else if (RadioBtnPps400Hz.Checked == true)
            {
                my_360AMX_UPC32.TestFreq = "A"; my_360AMX_UPC32.UserFreq = "400";
                my_360ASX_UPC3.TestFreq = "A"; my_360ASX_UPC3.UserFreq = "400";
                my_3150AFX.TestFreq = "A"; my_3150AFX.UserFreq = "400";
                my_118ACX_UPC1.TestFreq = "A"; my_118ACX_UPC1.UserFreq = "400";
                my_115ACX_UPC1.TestFreq = "A"; my_115ACX_UPC1.UserFreq = "400";
            }
            else { }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxPpsCaseNum_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxPpsEnterName_Click(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        private void TextBoxPpsEnterPacificSerNum_Click(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        private void TextBoxPpsEnterCertNum_Click(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        private void textBoxPpsEnterCaseNum_Click(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        private void textBoxPpsEnterPlantNum_TextChanged(object sender, EventArgs e)
        {
            if ((CmbBoxPacificSelect.Text != "") && (CmbBoxN4LPrt.Text != "") && (CmbBoxPacificUnit.Text != "")
                && (CmbBoxTrans.Text != "") && (TextBoxPpsEnterPacificSerNum.Text != "")
                && (TextBoxPpsEnterCertNum.Text != "") && (textBoxPpsEnterCaseNum.Text != "")
                && (textBoxPpsEnterName.Text != "") && (textBoxPpsEnterPlantNum.Text != ""))
            {
                BtnPpsConnect.Enabled = true;
            }
            else { BtnPpsConnect.Enabled = false; }
        }

        private void btnMeters_Click(object sender, EventArgs e)
        {
            showSubMenu(panelMetersMenu);
            HidePanels();
            panelMetersLanding.Show();
        }

        #region Meters Menu Methods
        private void btnApsM2003_Click(object sender, EventArgs e)
        {
            OpenMeterTemplate("APS M2003", "apsm2003");
            hideSubMenu();
        }

        private void btnFluke8508A_Click(object sender, EventArgs e)
        {
            OpenMeterTemplate("Fluke 8508A", "fluke8508a");
            hideSubMenu();
        }

        private void btnKeysight34471A_Click(object sender, EventArgs e)
        {
            OpenMeterTemplate("Keysight 34471A", "keysight34471a");
            hideSubMenu();
        }

        private void btnVitrek920A_Click(object sender, EventArgs e)
        {
            OpenMeterTemplate("Vitrek 920A", "vitrek920a");
            hideSubMenu();
        }

        private void btnSelectMeter_Click(object sender, EventArgs e)
        {
            // Show the meters submenu when the "View Available Meters" button is clicked
            showSubMenu(panelMetersMenu);
        }

        private void OpenMeterTemplate(string meterName, string meterType)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xls)|*.xls|All files (*.*)|*.*";
            openDialog.Title = $"Select {meterName} Calibration Template";
            openDialog.InitialDirectory = Path.Combine(System.Windows.Forms.Application.StartupPath, "Templates");

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Initialize the selected meter with communication setup
                    InitializeMeterCommunication(meterType, openDialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error opening template for {meterName}: {ex.Message}", 
                                  "Template Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void InitializeMeterCommunication(string meterType, string templatePath)
        {
            // Show VISA resource selection dialog
            using (var resourceDialog = new Form())
            {
                resourceDialog.Text = $"Select VISA Resource for {meterType.ToUpper()}";
                resourceDialog.Size = new System.Drawing.Size(400, 200);
                resourceDialog.StartPosition = FormStartPosition.CenterParent;

                var label = new System.Windows.Forms.Label() { Text = "Select VISA Resource:", Location = new System.Drawing.Point(10, 20), Size = new System.Drawing.Size(200, 20) };
                var comboBox = new System.Windows.Forms.ComboBox() { Location = new System.Drawing.Point(10, 50), Size = new System.Drawing.Size(350, 20), DropDownStyle = ComboBoxStyle.DropDownList };
                var btnOK = new System.Windows.Forms.Button() { Text = "Connect", Location = new System.Drawing.Point(200, 100), Size = new System.Drawing.Size(75, 30), DialogResult = DialogResult.OK };
                var btnCancel = new System.Windows.Forms.Button() { Text = "Cancel", Location = new System.Drawing.Point(285, 100), Size = new System.Drawing.Size(75, 30), DialogResult = DialogResult.Cancel };

                // Populate VISA resources
                PopulateVisaResources(comboBox);

                resourceDialog.Controls.AddRange(new Control[] { label, comboBox, btnOK, btnCancel });
                resourceDialog.AcceptButton = btnOK;
                resourceDialog.CancelButton = btnCancel;

                if (resourceDialog.ShowDialog() == DialogResult.OK && comboBox.SelectedItem != null)
                {
                    string visaAddress = comboBox.SelectedItem.ToString();
                    EstablishMeterConnection(meterType, visaAddress, templatePath);
                }
            }
        }

        private void PopulateVisaResources(ComboBox comboBox)
        {
            try
            {
                using (var rmSession = new ResourceManager())
                {
                    var resources = rmSession.Find("(ASRL|GPIB|TCPIP|USB)?*");
                    foreach (string resource in resources)
                    {
                        comboBox.Items.Add(resource);
                    }
                    if (comboBox.Items.Count > 0)
                    {
                        comboBox.SelectedIndex = 0;
                    }
                    else
                    {
                        // Add helpful message when no resources found
                        comboBox.Items.Add("No VISA resources detected");
                        comboBox.Items.Add("Check instrument connections");
                        comboBox.Items.Add("Install NI-VISA Runtime if needed");
                        comboBox.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"VISA Driver Error: {ex.Message}\n\n" +
                              "This may indicate:\n" +
                              "• VISA drivers not properly installed\n" +
                              "• No instruments connected\n" +
                              "• Instruments not powered on\n\n" +
                              "Recommendation: Install NI-VISA Runtime\n" +
                              "Download from: ni.com/downloads", 
                              "VISA Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
                // Add error status to comboBox
                comboBox.Items.Add("VISA Error - Check installation");
                comboBox.SelectedIndex = 0;
            }
        }

        private void EstablishMeterConnection(string meterType, string visaAddress, string templatePath)
        {
            try
            {
                BaseMeter meter = CreateMeterInstance(meterType);
                if (meter != null)
                {
                    meter.VisaAddress = visaAddress;
                    meter.OpenSession();
                    
                    // Test communication
                    string identification = meter.GetIdentification();
                    
                    MessageBox.Show($"Connected to {meterType.ToUpper()}!\n\n" +
                                  $"Identification: {identification}\n" +
                                  $"VISA Address: {visaAddress}\n" +
                                  $"Template: {templatePath}\n\n" +
                                  "Ready for calibration data streaming.",
                                  "Connection Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // TODO: Open Excel template and start data streaming
                    OpenExcelTemplate(templatePath, meter);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to connect to {meterType.ToUpper()}: {ex.Message}", 
                              "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private BaseMeter CreateMeterInstance(string meterType)
        {
            switch (meterType.ToLower())
            {
                case "apsm2003":
                    return new ApsM2003();
                case "fluke8508a":
                    return new Fluke8508A();
                case "keysight34471a":
                    return new Keysight34471A();
                case "vitrek920a":
                    return new Vitrek920A();
                default:
                    throw new ArgumentException($"Unknown meter type: {meterType}");
            }
        }

        private void OpenExcelTemplate(string templatePath, BaseMeter meter)
        {
            try
            {
                // Initialize Excel application
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook workbook = excelApp.Workbooks.Open(templatePath);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                MessageBox.Show($"Excel template opened: {templatePath}\n\n" +
                              "You can now stream data from the meter to the template.\n" +
                              "Use the meter's measurement methods to populate the template.",
                              "Template Ready", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // TODO: Implement continuous data streaming to Excel template
                // This would involve reading measurements from the meter and writing to specific cells
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening Excel template: {ex.Message}", 
                              "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }    
}
