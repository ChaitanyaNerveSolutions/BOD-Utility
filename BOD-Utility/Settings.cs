using NerveLog;
using NerveUtility;
using System;
using System.Data;
using System.Windows.Forms;

namespace BOD_Utility
{
    public partial class Settings : DevExpress.XtraEditors.XtraForm
    {
        string ApplicationPath = Application.StartupPath + "\\";
        DataSet ds_Config = new DataSet();
        NerveLogger _logger = new NerveLogger(true, true, ApplicationName: "BOD-Utility");
        public Settings()
        {
            InitializeComponent();
            _logger.Initialize(ApplicationPath);
        }

        private void Settings_Shown(object sender, EventArgs e)
        {
            try
            {
                ds_Config = NerveUtils.XMLC(ApplicationPath + "config.xml");
                var dRow = ds_Config.Tables["AUTOMATICSETTINGS"].Rows[0];
                spEdit_Attempts.Text = dRow["ATTEMPTS"].STR();
                spEdit_Interval.Text = dRow["INTERVAL"].STR();
                var startApiText = dRow["STARTNOTIS"].STR();
                if (startApiText == "Yes") { radioBtn_Yes.Checked = true; }
                else { radioBtn_No.Checked = true; }
                txt_FromAddress.Text = dRow["FROMEMAIL"].STR();
                txt_ToAddress.Text = dRow["TOEMAIL"].STR();
                txt_Password.Text = dRow["PASSWORD"].STR();
                txt_SMTP.Text = dRow["SMTP"].STR();
            }
            catch (Exception error) { _logger.Error(error); }
        }

        private void Settings_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                var dRow = ds_Config.Tables["AUTOMATICSETTINGS"].Rows[0];
                dRow["ATTEMPTS"] = spEdit_Attempts.Text.Trim('.');
                dRow["INTERVAL"] = spEdit_Interval.Text;
                if (radioBtn_Yes.Checked) { dRow["STARTNOTIS"] = "Yes"; }
                else { dRow["STARTNOTIS"] = "No"; }
                dRow["FROMEMAIL"] = txt_FromAddress.Text;
                dRow["TOEMAIL"] = txt_ToAddress.Text;
                dRow["PASSWORD"] = txt_Password.Text;
                dRow["SMTP"] = txt_SMTP.Text;
                NerveUtils.XMLW(ds_Config, "auto", "config.xml");
            }
            catch (Exception error) { _logger.Error(error); }

        }
    }
}