using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using i00SpellCheck;

namespace CSharpTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //enable the spell check
            //this will enable the spell check on ALL POSSIBLE CONTROLS ON THIS form AND ALL POSSIBLE CONTROLS ON ALL OWNED FORMS AS THEY OPEN automatically :)
            this.EnableSpellCheck(null);

            ////To enable spell check on single line textboxes you will need to call:
            //TextBox1.EnableSpellCheck(null);

            ////if you wanted to pass in options you can do so by going:
            //i00SpellCheck.SpellCheckSettings SpellCheckSettings = new i00SpellCheck.SpellCheckSettings();
            //SpellCheckSettings.DoSubforms = true; //Specifies if owned forms should be automatically spell checked
            //SpellCheckSettings.AllowAdditions = true; //Specifies if you want to allow the user to add words to the dictionary
            //SpellCheckSettings.AllowIgnore = true; //Specifies if you want to allow the user ignore words
            //SpellCheckSettings.AllowRemovals = true; //Specifies if you want to allow users to delete words from the dictionary
            //SpellCheckSettings.AllowInMenuDefs = true; //Specifies if the in menu definitions should be shown for correctly spelled words
            //SpellCheckSettings.AllowChangeTo = true; //Specifies if "Change to..." (to change to a synonym) should be shown in the menu for correctly spelled words
            //this.EnableSpellCheck(SpellCheckSettings);

            ////You can also enable spell checking on an individual Control (if supported):
            //TextBox1.EnableSpellCheck(null);

            ////To disable the spell check on a Control:
            //TextBox1.DisableSpellCheck();

            ////To see if the spell check is enabled on a Control:
            //bool SpellCheckEnabled = TextBox1.IsSpellCheckEnabled();

            ////To change options on an individual Control:
            //TextBox1.SpellCheck(true, null).Settings.AllowAdditions = true;
            //TextBox1.SpellCheck(true, null).Settings.AllowIgnore = true;
            //TextBox1.SpellCheck(true, null).Settings.AllowRemovals = true;
            //TextBox1.SpellCheck(true, null).Settings.ShowMistakes = true;
            ////etc

            ////To show a spellcheck dialog for an individual Control:
            ////...Sorry I don't know how to convert this ... let me know if you do!

            ////To load a custom dictionary from a saved file:
            i00SpellCheck.FlatFileDictionary Dictionary = new i00SpellCheck.FlatFileDictionary("c:\\Custom.dic", false);

            ////To create a new blank dictionary and save it as a file
            //i00SpellCheck.FlatFileDictionary Dictionary = new i00SpellCheck.FlatFileDictionary("c:\\Custom.dic", true);
            //Dictionary.Add("CustomWord1");
            //Dictionary.Add("CustomWord2");
            //Dictionary.Add("CustomWord3");
            //Dictionary.Save(Dictionary.Filename, true);

            ////To Load a custom dictionary for an individual Control:
            //TextBox1.SpellCheck(true, null).CurrentDictionary = Dictionary;

            ////To Open the dictionary editor for a dictionary associated with a Control:
            ////NOTE: this should only be done after the dictionary has loaded (Control.SpellCheck.CurrentDictionary.Loading = False)
            //TextBox1.SpellCheck(true, null).CurrentDictionary.ShowUIEditor();

            ////Repaint all of the controls that use the same dictionary...
            //TextBox1.SpellCheck(true, null).InvalidateAllControlsWithSameDict(true);


            //set the object for the property grid
            PropertyGrid1.SelectedObject = TextBox1.SpellCheck(true, null);

            //everything below here is for cosmetics...

            UpdateEnabledCheck();
 
            var ToolBoxIcon = new ToolboxBitmapAttribute(typeof(PropertyGrid));
            tsbProperties.Image = ToolBoxIcon.GetImage(typeof(PropertyGrid), false);

            TextBox1.SelectionStart = 0;
            TextBox1.SelectionLength = 0;

            var ico = Icon.ExtractAssociatedIcon("i00SpellCheck.exe");
            using (ico) {
                var b=new Bitmap(16,16);
                var g = Graphics.FromImage(b);
                using (g)
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.High;
                    g.DrawIcon(ico,new Rectangle(0, 0, b.Width, b.Height));
                }
                tsbSpellCheck.Image = b;
            }

            this.Icon = Icon.ExtractAssociatedIcon("i00SpellCheck.exe");
        }

        //show and hide the property grid
        private void tsbProperties_Click(object sender, EventArgs e)
        {
            tsbProperties.Checked = ! tsbProperties.Checked;
            PropertyGrid1.Visible = tsbProperties.Checked;
        }

        #region "Enable / Disable Spell Check"

        private void UpdateEnabledCheck()
        {
	        var ts = (ToolStrip)tsiEnabled.Owner;
            System.Windows.Forms.VisualStyles.CheckBoxState state = TextBox1.IsSpellCheckEnabled() ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal : System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal;
	        var Size = System.Windows.Forms.CheckBoxRenderer.GetGlyphSize(this.CreateGraphics(), state);

	        int bWidth = 0;
	        int bHeight = 0;

	        bWidth = ts.ImageScalingSize.Width;
	        bHeight = ts.ImageScalingSize.Height;

	        Point Offset = new Point(0, 0);

	        if (Size.Width < ts.ImageScalingSize.Width) {
		        Offset.X = Convert.ToInt32(((ts.ImageScalingSize.Width - Size.Width) / 2));
	        } else {
		        bWidth = Size.Width;
	        }
	        if (Size.Height < ts.ImageScalingSize.Height) {
		        Offset.Y = Convert.ToInt32(((ts.ImageScalingSize.Height - Size.Height) / 2));
	        } else {
		        bHeight = Size.Height;
	        }


	        Bitmap b = new Bitmap(bWidth, bHeight);
            Graphics g = Graphics.FromImage(b); 
            using (g) {
		        g.TranslateTransform(Offset.X, Offset.Y);
		        System.Windows.Forms.CheckBoxRenderer.DrawCheckBox(g, new Point(0, 0), state);
	        }
	        tsiEnabled.Image = b;
	        tsiEnabled.Visible = true;
        }

        private void tsiEnabled_Click(object sender, EventArgs e)
        {
            if (TextBox1.IsSpellCheckEnabled())
            {
                TextBox1.DisableSpellCheck();
            }
            else
            {
                TextBox1.EnableSpellCheck(null);
            }
            UpdateEnabledCheck();
        }

        #endregion

        //show the spellcheck dialog
        private void tsbSpellCheck_Click(object sender, EventArgs e)
        {
            var iSpellCheckDialog = TextBox1.SpellCheck(false,null) as i00SpellCheck.SpellCheckControlBase.iSpellCheckDialog;
            if (iSpellCheckDialog != null)
            {
                iSpellCheckDialog.ShowDialog();
            }
        }

    }
}
