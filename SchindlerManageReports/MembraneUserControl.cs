using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SchindlerManageReports
{
    public partial class MembraneUserControl : Form
    {

        #region Constructors + Fields

        public MembraneUserControl()
        {
            this.SetStyle(ControlStyles.DoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.UserPaint, true);
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        }

        private Color color1 = Color.White;
        private Color color2 = Color.FromArgb(206, 227, 247);


        #endregion

        #region Properties

        public Color Color1
        {
            get { return color1; }
            set { color1 = value; }
        }

        public Color Color2
        {
            get { return color2; }
            set { color2 = value; }
        }

        #endregion

        #region Methods

        [DebuggerStepThrough]
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);
            LinearGradientBrush BackgroundGradientBrush =
                new LinearGradientBrush(
                new Rectangle(0, 0, this.Width, this.Height),
                color1,
                color2,
                LinearGradientMode.Vertical);
            e.Graphics.FillRectangle(BackgroundGradientBrush, e.ClipRectangle);
            if (BackgroundGradientBrush != null) BackgroundGradientBrush.Dispose();
        }

        protected override void OnCreateControl()
        {
            base.OnCreateControl();


        }

        #endregion
    }
}