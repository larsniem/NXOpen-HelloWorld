using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NXOpen;
using NXOpen.UF;

namespace HelloWorld
{
    public class Programm
    {
        public static Session theSession;
        public static ListingWindow lw;
        public static Part workPart;
        public static UFSession UfSession;

        // Entry point in case of execution of the .dll by NX
        public static void Main(string[] args)
        {
            theSession = Session.GetSession();
            UfSession = UFSession.GetUFSession();
            lw = theSession.ListingWindow;
            workPart = theSession.Parts.Work;

            //HelloWorld(lw);

            //var pts = GetAllPoints(workPart);
            //var pts = AskPoints();
            var pts = AskPointsBlockStyler();

            if (pts.Length > 0)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV-File (*.csv) | *.csv";
                saveFileDialog.Title = "Export CSV";
                var response = saveFileDialog.ShowDialog();
                if (response == DialogResult.OK && !String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                {
                    WriteCSV(pts, saveFileDialog.FileName);
                }
            }
        }

        public static void HelloWorld(ListingWindow lw)
        {
            lw.Open();
            lw.WriteFullline("Hello World!");
        }

        public static Point3d[] GetAllPoints(Part workPart)
        {
            var points = workPart.Points.ToArray();
            
            Point3d[] pts = new Point3d[points.Length];
            for (int i = 0; i < points.Length; i++)
            {
                pts[i] = points[i].Coordinates;
            }

            pts = points.Select(p => p.Coordinates).ToArray();

            return pts;
        }

        public static Point3d[] AskPoints()
        {
            var selectionoptions = new UFUi.SelectionOption();
            
            // Type Filter (right next to the 'Menu' dropdown menu in the NX GUI)
            var selectionFilter = new UFUi.Mask();
            selectionFilter.object_type = UFConstants.UF_point_type;
            selectionFilter.object_subtype = UFConstants.UF_point_subtype;
            selectionFilter.solid_type = 0;
            selectionoptions.mask_triples = new[] { selectionFilter };
            selectionoptions.num_mask_triples = 1;

            // Only objects in workpart, not in an assembly or other components of an assembly
            // aka. Selection Scope (right next the 'Type Filter in the NX GUI)
            selectionoptions.scope = UFConstants.UF_UI_SEL_SCOPE_WORK_PART;

            int response;
            int pointCount;
            Tag[] pointTags;
            UfSession.Ui.SelectByClass("Select Points to export as CSV.", ref selectionoptions, out response, out pointCount, out pointTags);

            Point3d[] pts = new Point3d[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                // The dialog only returns the Tags (ID as unsigend integer) of the points. The object concrete object will be obtained via the global object manager  
                Point point = (Point) UfSession.GetObjectManager().GetTaggedObject(pointTags[i]);
                pts[i] = point.Coordinates;
            }

            return pts;
        }

        public static Point3d[] AskPointsBlockStyler()
        {
            var dialog = new Dialog();
            var response = dialog.Show();
            return dialog.Pts;
        }


        public static void WriteCSV(Point3d[] pts, string filepath)
        {
            string[] csv = new string[pts.Length];

            // The escape character '\t' represents a tab
            for (int i = 0; i < pts.Length; i++)
            {
                csv[i] = String.Format("{0}\t{1}\t{2}", pts[i].X, pts[i].Y, pts[i].Z);
            }

            File.WriteAllLines(filepath, csv);
        }

        // Option to 'free' the library at the end of an execution/call of the library.
        public static int GetUnloadOption(string dummy) { return (int)NXOpen.Session.LibraryUnloadOption.Immediately; }
    }
}
