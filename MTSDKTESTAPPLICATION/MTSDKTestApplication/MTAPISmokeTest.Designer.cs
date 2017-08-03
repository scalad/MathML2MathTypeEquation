// MTAPISmokeTestDesigner.cs --- Copyright (c) 2008-2010 by Design Science, Inc.
// Purpose:
// $Header: /MathType/Windows/SDK/DotNET/MTSDKTESTAPPLICATION/MTSDKTestApplication/MTAPISmokeTest.Designer.cs 3     1/19/10 2:02p Jimm $

namespace MTSDKTestApplication
{
    partial class MTAPISmokeTest
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (_isMTAPIConnected)
            {
                _MTSDK.MTAPIDisconnectMgn();
            }

            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.StartAPI = new System.Windows.Forms.Button();
            this.EndPI = new System.Windows.Forms.Button();
            this.APIVersion = new System.Windows.Forms.Button();
            this.EqnOnClipboard = new System.Windows.Forms.Button();
            this.ClearCB = new System.Windows.Forms.Button();
            this.GetLastDim = new System.Windows.Forms.Button();
            this.MTOpenFile = new System.Windows.Forms.Button();
            this.GetCBPrefs = new System.Windows.Forms.Button();
            this.GetPrefsFromFile = new System.Windows.Forms.Button();
            this.PerfsToUIForm = new System.Windows.Forms.Button();
            this.GetDefaultPrefs = new System.Windows.Forms.Button();
            this.SetMTPrefs = new System.Windows.Forms.Button();
            this.GetTranslatorsInfo = new System.Windows.Forms.Button();
            this.EnumTranslators = new System.Windows.Forms.Button();
            this.FormReset = new System.Windows.Forms.Button();
            this.FormAddVarSub = new System.Windows.Forms.Button();
            this.FormSetTrans = new System.Windows.Forms.Button();
            this.FormSetPrefs = new System.Windows.Forms.Button();
            this.FormEqn = new System.Windows.Forms.Button();
            this.MTXFormGetStatus = new System.Windows.Forms.Button();
            this.MTPreviewDialog = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.MTShowAboutBox = new System.Windows.Forms.Button();
            this.MTGetURL = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            //
            // StartAPI
            //
            this.StartAPI.Location = new System.Drawing.Point(5, 19);
            this.StartAPI.Name = "StartAPI";
            this.StartAPI.Size = new System.Drawing.Size(145, 23);
            this.StartAPI.TabIndex = 0;
            this.StartAPI.Text = "MTAPIConnect";
            this.StartAPI.UseVisualStyleBackColor = true;
            this.StartAPI.Click += new System.EventHandler(this.StartAPI_Click);
            //
            // EndPI
            //
            this.EndPI.Location = new System.Drawing.Point(6, 48);
            this.EndPI.Name = "EndPI";
            this.EndPI.Size = new System.Drawing.Size(144, 23);
            this.EndPI.TabIndex = 1;
            this.EndPI.Text = "MTAPIDisconnect";
            this.EndPI.UseVisualStyleBackColor = true;
            this.EndPI.Click += new System.EventHandler(this.EndPI_Click);
            //
            // APIVersion
            //
            this.APIVersion.Location = new System.Drawing.Point(5, 77);
            this.APIVersion.Name = "APIVersion";
            this.APIVersion.Size = new System.Drawing.Size(145, 23);
            this.APIVersion.TabIndex = 2;
            this.APIVersion.Text = "MTAPIVersion";
            this.APIVersion.UseVisualStyleBackColor = true;
            this.APIVersion.Click += new System.EventHandler(this.APIVersion_Click);
            //
            // EqnOnClipboard
            //
            this.EqnOnClipboard.Location = new System.Drawing.Point(6, 19);
            this.EqnOnClipboard.Name = "EqnOnClipboard";
            this.EqnOnClipboard.Size = new System.Drawing.Size(144, 23);
            this.EqnOnClipboard.TabIndex = 0;
            this.EqnOnClipboard.Text = "MTEquationOnClipboard";
            this.EqnOnClipboard.UseVisualStyleBackColor = true;
            this.EqnOnClipboard.Click += new System.EventHandler(this.EqnOnClipboard_Click);
            //
            // ClearCB
            //
            this.ClearCB.Location = new System.Drawing.Point(6, 77);
            this.ClearCB.Name = "ClearCB";
            this.ClearCB.Size = new System.Drawing.Size(144, 23);
            this.ClearCB.TabIndex = 2;
            this.ClearCB.Text = "MTClearClipboard";
            this.ClearCB.UseVisualStyleBackColor = true;
            this.ClearCB.Click += new System.EventHandler(this.ClearCB_Click);
            //
            // GetLastDim
            //
            this.GetLastDim.Location = new System.Drawing.Point(6, 48);
            this.GetLastDim.Name = "GetLastDim";
            this.GetLastDim.Size = new System.Drawing.Size(136, 23);
            this.GetLastDim.TabIndex = 1;
            this.GetLastDim.Text = "MTGetLastDimension";
            this.GetLastDim.UseVisualStyleBackColor = true;
            this.GetLastDim.Click += new System.EventHandler(this.GetLastDim_Click);
            //
            // MTOpenFile
            //
            this.MTOpenFile.Location = new System.Drawing.Point(6, 19);
            this.MTOpenFile.Name = "MTOpenFile";
            this.MTOpenFile.Size = new System.Drawing.Size(136, 23);
            this.MTOpenFile.TabIndex = 0;
            this.MTOpenFile.Text = "MTOpenFileDialog";
            this.MTOpenFile.UseVisualStyleBackColor = true;
            this.MTOpenFile.Click += new System.EventHandler(this.MTOpenFile_Click);
            //
            // GetCBPrefs
            //
            this.GetCBPrefs.Location = new System.Drawing.Point(6, 48);
            this.GetCBPrefs.Name = "GetCBPrefs";
            this.GetCBPrefs.Size = new System.Drawing.Size(145, 23);
            this.GetCBPrefs.TabIndex = 1;
            this.GetCBPrefs.Text = "MTGetPrefsFromClipboard";
            this.GetCBPrefs.UseVisualStyleBackColor = true;
            this.GetCBPrefs.Click += new System.EventHandler(this.GetCBPrefs_Click);
            //
            // GetPrefsFromFile
            //
            this.GetPrefsFromFile.Location = new System.Drawing.Point(12, 106);
            this.GetPrefsFromFile.Name = "GetPrefsFromFile";
            this.GetPrefsFromFile.Size = new System.Drawing.Size(130, 23);
            this.GetPrefsFromFile.TabIndex = 3;
            this.GetPrefsFromFile.Text = "MTGetPrefsFromFile";
            this.GetPrefsFromFile.UseVisualStyleBackColor = true;
            this.GetPrefsFromFile.Click += new System.EventHandler(this.GetPrefsFromFile_Click);
            //
            // PerfsToUIForm
            //
            this.PerfsToUIForm.Location = new System.Drawing.Point(12, 48);
            this.PerfsToUIForm.Name = "PerfsToUIForm";
            this.PerfsToUIForm.Size = new System.Drawing.Size(130, 23);
            this.PerfsToUIForm.TabIndex = 1;
            this.PerfsToUIForm.Text = "MTConvertPrefsToUIForm";
            this.PerfsToUIForm.UseVisualStyleBackColor = true;
            this.PerfsToUIForm.Click += new System.EventHandler(this.PerfsToUIForm_Click);
            //
            // GetDefaultPrefs
            //
            this.GetDefaultPrefs.Location = new System.Drawing.Point(12, 19);
            this.GetDefaultPrefs.Name = "GetDefaultPrefs";
            this.GetDefaultPrefs.Size = new System.Drawing.Size(130, 23);
            this.GetDefaultPrefs.TabIndex = 0;
            this.GetDefaultPrefs.Text = "MTGetPrefsMTDefault";
            this.GetDefaultPrefs.UseVisualStyleBackColor = true;
            this.GetDefaultPrefs.Click += new System.EventHandler(this.GetDefaultPrefs_Click);
            //
            // SetMTPrefs
            //
            this.SetMTPrefs.Location = new System.Drawing.Point(12, 77);
            this.SetMTPrefs.Name = "SetMTPrefs";
            this.SetMTPrefs.Size = new System.Drawing.Size(130, 23);
            this.SetMTPrefs.TabIndex = 2;
            this.SetMTPrefs.Text = "MTSetMTPrefs";
            this.SetMTPrefs.UseVisualStyleBackColor = true;
            this.SetMTPrefs.Click += new System.EventHandler(this.SetMTPrefs_Click);
            //
            // GetTranslatorsInfo
            //
            this.GetTranslatorsInfo.Location = new System.Drawing.Point(4, 48);
            this.GetTranslatorsInfo.Name = "GetTranslatorsInfo";
            this.GetTranslatorsInfo.Size = new System.Drawing.Size(145, 23);
            this.GetTranslatorsInfo.TabIndex = 1;
            this.GetTranslatorsInfo.Text = "MTGetTranslatorsInfo";
            this.GetTranslatorsInfo.UseVisualStyleBackColor = true;
            this.GetTranslatorsInfo.Click += new System.EventHandler(this.GetTranslatorsInfo_Click);
            //
            // EnumTranslators
            //
            this.EnumTranslators.Location = new System.Drawing.Point(5, 19);
            this.EnumTranslators.Name = "EnumTranslators";
            this.EnumTranslators.Size = new System.Drawing.Size(144, 23);
            this.EnumTranslators.TabIndex = 0;
            this.EnumTranslators.Text = "MTEnumTranslators";
            this.EnumTranslators.UseVisualStyleBackColor = true;
            this.EnumTranslators.Click += new System.EventHandler(this.EnumTranslators_Click);
            //
            // FormReset
            //
            this.FormReset.Location = new System.Drawing.Point(16, 136);
            this.FormReset.Name = "FormReset";
            this.FormReset.Size = new System.Drawing.Size(130, 23);
            this.FormReset.TabIndex = 4;
            this.FormReset.Text = "MTXFormReset";
            this.FormReset.UseVisualStyleBackColor = true;
            this.FormReset.Click += new System.EventHandler(this.FormReset_Click);
            //
            // FormAddVarSub
            //
            this.FormAddVarSub.Location = new System.Drawing.Point(16, 20);
            this.FormAddVarSub.Name = "FormAddVarSub";
            this.FormAddVarSub.Size = new System.Drawing.Size(130, 23);
            this.FormAddVarSub.TabIndex = 0;
            this.FormAddVarSub.Text = "MTXFormAddVarSub";
            this.FormAddVarSub.UseVisualStyleBackColor = true;
            this.FormAddVarSub.Click += new System.EventHandler(this.FormAddVarSub_Click);
            //
            // FormSetTrans
            //
            this.FormSetTrans.Location = new System.Drawing.Point(16, 49);
            this.FormSetTrans.Name = "FormSetTrans";
            this.FormSetTrans.Size = new System.Drawing.Size(130, 23);
            this.FormSetTrans.TabIndex = 1;
            this.FormSetTrans.Text = "MTXFormSetTranslator";
            this.FormSetTrans.UseVisualStyleBackColor = true;
            this.FormSetTrans.Click += new System.EventHandler(this.FormSetTrans_Click);
            //
            // FormSetPrefs
            //
            this.FormSetPrefs.Location = new System.Drawing.Point(16, 78);
            this.FormSetPrefs.Name = "FormSetPrefs";
            this.FormSetPrefs.Size = new System.Drawing.Size(130, 23);
            this.FormSetPrefs.TabIndex = 2;
            this.FormSetPrefs.Text = "MTXFormSetPrefs";
            this.FormSetPrefs.UseVisualStyleBackColor = true;
            this.FormSetPrefs.Click += new System.EventHandler(this.FormSetPrefs_Click);
            //
            // FormEqn
            //
            this.FormEqn.Location = new System.Drawing.Point(16, 164);
            this.FormEqn.Name = "FormEqn";
            this.FormEqn.Size = new System.Drawing.Size(130, 23);
            this.FormEqn.TabIndex = 5;
            this.FormEqn.Text = "MTXFormEqn";
            this.FormEqn.UseVisualStyleBackColor = true;
            this.FormEqn.Click += new System.EventHandler(this.FormEqn_Click);
            //
            // MTXFormGetStatus
            //
            this.MTXFormGetStatus.Location = new System.Drawing.Point(16, 107);
            this.MTXFormGetStatus.Name = "MTXFormGetStatus";
            this.MTXFormGetStatus.Size = new System.Drawing.Size(130, 23);
            this.MTXFormGetStatus.TabIndex = 3;
            this.MTXFormGetStatus.Text = "MTXFormGetStatus";
            this.MTXFormGetStatus.UseVisualStyleBackColor = true;
            this.MTXFormGetStatus.Click += new System.EventHandler(this.MTXFormGetStatus_Click);
            //
            // MTPreviewDialog
            //
            this.MTPreviewDialog.Location = new System.Drawing.Point(6, 77);
            this.MTPreviewDialog.Name = "MTPreviewDialog";
            this.MTPreviewDialog.Size = new System.Drawing.Size(136, 23);
            this.MTPreviewDialog.TabIndex = 2;
            this.MTPreviewDialog.Text = "MTPreviewDialog";
            this.MTPreviewDialog.UseVisualStyleBackColor = true;
            this.MTPreviewDialog.Click += new System.EventHandler(this.MTPreviewDialog_Click);
            //
            // groupBox1
            //
            this.groupBox1.Controls.Add(this.APIVersion);
            this.groupBox1.Controls.Add(this.StartAPI);
            this.groupBox1.Controls.Add(this.EndPI);
            this.groupBox1.Location = new System.Drawing.Point(7, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(160, 108);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Initial API calls";
            //
            // groupBox2
            //
            this.groupBox2.Controls.Add(this.ClearCB);
            this.groupBox2.Controls.Add(this.EqnOnClipboard);
            this.groupBox2.Controls.Add(this.GetCBPrefs);
            this.groupBox2.Location = new System.Drawing.Point(7, 128);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(161, 110);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Clipboard API Calls";
            //
            // groupBox3
            //
            this.groupBox3.Controls.Add(this.EnumTranslators);
            this.groupBox3.Controls.Add(this.GetTranslatorsInfo);
            this.groupBox3.Location = new System.Drawing.Point(8, 254);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(159, 84);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Translator API calls";
            //
            // groupBox4
            //
            this.groupBox4.Controls.Add(this.FormAddVarSub);
            this.groupBox4.Controls.Add(this.FormSetTrans);
            this.groupBox4.Controls.Add(this.FormSetPrefs);
            this.groupBox4.Controls.Add(this.MTXFormGetStatus);
            this.groupBox4.Controls.Add(this.FormReset);
            this.groupBox4.Controls.Add(this.FormEqn);
            this.groupBox4.Location = new System.Drawing.Point(192, 4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(160, 198);
            this.groupBox4.TabIndex = 3;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Equation Xfrm API calls";
            //
            // groupBox5
            //
            this.groupBox5.Controls.Add(this.GetDefaultPrefs);
            this.groupBox5.Controls.Add(this.PerfsToUIForm);
            this.groupBox5.Controls.Add(this.SetMTPrefs);
            this.groupBox5.Controls.Add(this.GetPrefsFromFile);
            this.groupBox5.Location = new System.Drawing.Point(196, 219);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(155, 138);
            this.groupBox5.TabIndex = 4;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "MathType Preferences";
            //
            // groupBox6
            //
            this.groupBox6.Controls.Add(this.MTGetURL);
            this.groupBox6.Controls.Add(this.MTShowAboutBox);
            this.groupBox6.Controls.Add(this.MTOpenFile);
            this.groupBox6.Controls.Add(this.GetLastDim);
            this.groupBox6.Controls.Add(this.MTPreviewDialog);
            this.groupBox6.Location = new System.Drawing.Point(376, 11);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(157, 168);
            this.groupBox6.TabIndex = 5;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Other MathType API calls";
            //
            // MTShowAboutBox
            //
            this.MTShowAboutBox.Location = new System.Drawing.Point(6, 106);
            this.MTShowAboutBox.Name = "MTShowAboutBox";
            this.MTShowAboutBox.Size = new System.Drawing.Size(136, 23);
            this.MTShowAboutBox.TabIndex = 4;
            this.MTShowAboutBox.Text = "MTShowAboutBox";
            this.MTShowAboutBox.UseVisualStyleBackColor = true;
            this.MTShowAboutBox.Click += new System.EventHandler(this.MTShowAboutBox_Click);
            //
            // MTGetURL
            //
            this.MTGetURL.Location = new System.Drawing.Point(6, 135);
            this.MTGetURL.Name = "MTGetURL";
            this.MTGetURL.Size = new System.Drawing.Size(136, 23);
            this.MTGetURL.TabIndex = 6;
            this.MTGetURL.Text = "MTGetURL";
            this.MTGetURL.UseVisualStyleBackColor = true;
            this.MTGetURL.Click += new System.EventHandler(this.MTGetURL_Click);
            //
            // Form1
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(565, 412);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "MathType SDK Smoke test application";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button StartAPI;
        private System.Windows.Forms.Button EndPI;
        private System.Windows.Forms.Button APIVersion;
        private System.Windows.Forms.Button EqnOnClipboard;
        private System.Windows.Forms.Button ClearCB;
        private System.Windows.Forms.Button GetLastDim;
        private System.Windows.Forms.Button MTOpenFile;
        private System.Windows.Forms.Button GetCBPrefs;
        private System.Windows.Forms.Button GetPrefsFromFile;
        private System.Windows.Forms.Button PerfsToUIForm;
        private System.Windows.Forms.Button GetDefaultPrefs;
        private System.Windows.Forms.Button SetMTPrefs;
        private System.Windows.Forms.Button GetTranslatorsInfo;
        private System.Windows.Forms.Button EnumTranslators;
        private System.Windows.Forms.Button FormReset;
        private System.Windows.Forms.Button FormAddVarSub;
        private System.Windows.Forms.Button FormSetTrans;
        private System.Windows.Forms.Button FormSetPrefs;
        private System.Windows.Forms.Button FormEqn;
        private System.Windows.Forms.Button MTXFormGetStatus;
        private System.Windows.Forms.Button MTPreviewDialog;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Button MTShowAboutBox;
        private System.Windows.Forms.Button MTGetURL;
    }
}

