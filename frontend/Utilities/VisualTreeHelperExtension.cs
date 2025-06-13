// C:\Users\p.eminiant\Desktop\PROJETS\ExcelFlow\frontend\Helpers\VisualTreeHelperExtensions.cs


using System.Collections.Generic; // For IEnumerable
using System.Windows;             // For DependencyObject
using System.Windows.Media;       // For VisualTreeHelper

namespace ExcelFlow.Utilities
{


    /// <summary>
    /// Provides extension methods for the VisualTreeHelper to find visual children of a specific type.
    public static class VisualTreeHelperExtensions
    {
        public static IEnumerable<T> FindVisualChildren<T>(this DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null) // <--- Add this null check for 'child'
                    {
                        if (child is T)
                        {
                            yield return (T)child;
                        }

                        // Recursively search children of children
                        // Now we pass 'child' only if it's not null
                        foreach (T childOfChild in FindVisualChildren<T>(child))
                        {
                            yield return childOfChild;
                        }
                    }
                }
            }
        }
    }
}