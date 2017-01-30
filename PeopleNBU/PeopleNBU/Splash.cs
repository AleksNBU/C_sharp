using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PeopleNBU
{
    public partial class Splash : Form
    {
        #region FIELDS

        Timer timer = new Timer();
        bool fadeIn = true;
        bool fadeOut = true;

        #endregion

        #region METHODS

        public Splash()
        {
            InitializeComponent();

            ExtraFormSettings();
            SetAndStartTimer();
        }

        private void SetAndStartTimer()
        {
            timer.Interval = 100;
            timer.Tick += new EventHandler(t_Tick);
            timer.Start();
        }

        private void ExtraFormSettings()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.Opacity = 0.01;
            this.BackgroundImage = Properties.Resources.lady;
        }

        #endregion

        #region EVENTS

        void t_Tick(object sender, EventArgs e)
        {
            if (fadeIn)
            {
                if (this.Opacity < 1.0)
                {
                    this.Opacity += 0.01;
                }
                else
                {
                    fadeIn = false;
                    fadeOut = true;
                }
            }
            else if (fadeOut)
            {
                if (this.Opacity > 0)
                {
                    this.Opacity -= 0.01;
                }
                else
                {
                    fadeOut = false;
                }
            }

            if (!(fadeIn || fadeOut))
            {
                timer.Stop();
                this.Close();
            }
        }

        #endregion

    }
}

