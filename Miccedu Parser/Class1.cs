using System.Windows;
using System.Windows.Controls;

namespace miccedux
{
    class Class1
    {
        public static ClassExcelTable exceltable { get; private set; } = new ClassExcelTable();

        public static Frame Window_Frame { get; private set; } = null;

        public static Window Window_App { get; private set; } = null;

        public static bool ExcelAlive { get; set; } = false;

        public static void SetWindow(Window New_Window)
        {
            if (Window_App == null)
                Window_App = New_Window;
        }

        public static void SetFrame(Frame New_Frame)
        {
            if (Window_Frame == null)
                Window_Frame = New_Frame;
        }
    }
}
