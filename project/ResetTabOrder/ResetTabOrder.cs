using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;

namespace ResetTabOrder
{
    /// <summary>
    /// 重設 Tab Index 順序
    /// </summary>
    public static class ResetTabOrder
    {
        private static Dictionary<Control, int> allControls = new Dictionary<Control, int>();

        /// <summary>
        /// 執行重設
        /// </summary>
        /// <param name="activeDoc">目前執行畫面</param>
        public static void DoWork(Document activeDoc)
        {
            allControls.Clear();

            IDesignerHost host = (IDesignerHost)activeDoc.ActiveWindow.Object;
            IComponent cf = host.RootComponent;

            bool isForm = cf is Form;
            if (!isForm) MessageBox.Show(".....");

            Form f = cf as Form;
            ResetSubControls(f);

            WriteToDesign(activeDoc.FullName);

        }

        private static void ResetSubControls(Control parentCtrl)
        {
            List<Control> subCtrl = new List<Control>();

            foreach (Control sbCtrl in parentCtrl.Controls)
            {
                subCtrl.Add(sbCtrl);
                allControls.Add(sbCtrl, sbCtrl.TabIndex);
            }

            Control[] reorderSubCtrl
                = subCtrl.AsEnumerable()
                         .OrderBy(c => c.Location.Y)
                         .ThenBy(c => c.Location.X)
                         .ToArray();

            for (int i = 0; i < reorderSubCtrl.Length; i++)
            {
                Control c = reorderSubCtrl[i];
                c.TabIndex = i;
                ResetSubControls(c);
            }
        }

        private static void WriteToDesign(string docPath)
        {
            string designerPath =
                Path.ChangeExtension(docPath, ".Designer.cs");

            string designContent = File.ReadAllText(designerPath);

            foreach (var c in allControls)
            {
                string oldSetting = string.Format("{0}.TabIndex = {1};", c.Key.Name, c.Value);
                string newSetting = string.Format("{0}.TabIndex = {1};", c.Key.Name, c.Key.TabIndex);
                designContent = designContent.Replace(oldSetting, newSetting);
            }

            File.WriteAllText(designerPath, designContent);
        }
    }
}