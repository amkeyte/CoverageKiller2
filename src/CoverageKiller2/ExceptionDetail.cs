/*
Example Usage:

using ExceptionNerdery;

try
{
    var col = wordTable.Columns[1]; // YOLO access that blows up on merged cell widths
}
catch (Exception ex)
{
    if (ExceptionDetail.ExceptionIs(ex, KnownExceptions.WordInterop.MixedCellWidths))
    {
        Console.WriteLine("Yep, it's that Word 'mixed cell widths' COMException again. Chill.");
    }

    // OR: enum-based fancy way
    if (ExceptionDetail.ExceptionIs(ex, ExceptionRegistry.Get(KnownExceptionType.WordMixedCellWidths)))
    {
        Console.WriteLine("Caught via enum-based lookup. This codebase is now 12% more elegant.");
    }
}
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace CoverageKiller2
{
    /// <summary>
    /// Types of known exception patterns for optional enum-based access.
    /// </summary>
    public enum KnownExceptionType
    {
        WordMixedCellWidths,
        // Add more if your app finds new ways to explode.
    }

    /// <summary>
    /// Represents a filter for matching known exceptions by their characteristics.
    /// </summary>
    public class ExceptionDetail
    {
        public Type ExceptionType { get; set; }
        public string MessageContains { get; set; }
        public string HelpLinkEndsWith { get; set; }
        public string TargetSiteName { get; set; }
        public List<int> HResult { get; set; }

        /// <summary>
        /// Checks whether a given exception matches the provided ExceptionDetail definition.
        /// </summary>
        public static bool Is(Exception ex, ExceptionDetail detail)
        {
            if (ex == null || detail == null)
                return false;

            if (detail.ExceptionType != null && !detail.ExceptionType.IsAssignableFrom(ex.GetType()))
                return false;

            if (detail.MessageContains != null && (ex.Message == null || !ex.Message.Contains(detail.MessageContains)))
                return false;

            if (detail.HelpLinkEndsWith != null && (ex.HelpLink == null || !ex.HelpLink.EndsWith(detail.HelpLinkEndsWith)))
                return false;

            if (detail.TargetSiteName != null && (ex.TargetSite?.Name != detail.TargetSiteName))
                return false;

            if (detail.HResult != null && detail.HResult.Count > 0)
            {
                if (!detail.HResult.Contains(ex.HResult))
                {
                    //DEBUG
                    var message = $"HResult mismatch: Actual={ex.HResult}; " +
                        $"Listed: {string.Join(", ", detail.HResult.Select(hr => $"[{hr}]"))}";

                    if (Debugger.IsAttached) Debugger.Break();
                    throw new Exception(message, ex);
                    //return false;
                }
            }


            return true;
        }
    }

    /// <summary>
    /// Known grouped exceptions by context (like WordInterop, ExcelInterop, etc).
    /// </summary>
    public static class KnownExceptions
    {
        public static class VSTO
        {
            public static readonly ExceptionDetail MixedCellWidths = new ExceptionDetail
            {
                ExceptionType = typeof(COMException),
                MessageContains = "mixed cell widths",
                HelpLinkEndsWith = "#25472",
                TargetSiteName = "get_Item",
                HResult = new List<int>
                {
                    -2146822296  // what you're actually seeing
                }
            };
            public static readonly ExceptionDetail ObjectDeleted = new ExceptionDetail
            {
                ExceptionType = typeof(ArgumentException),
                MessageContains = "Object has been deleted",
                HelpLinkEndsWith = "#25305",
                TargetSiteName = "split",
                HResult = new List<int>
                {
                    -2107024809  // what you're actually seeing
                }
            };

        }

        public static class ExcelInterop
        {
            // Add Excel-specific exception patterns here
        }

        public static class PowerPointInterop
        {
            // Add PowerPoint-specific exception patterns here
        }
    }

    /// <summary>
    /// Optional: Maps enum identifiers to ExceptionDetail instances.
    /// </summary>
    public static class ExceptionRegistry
    {
        private static readonly Dictionary<KnownExceptionType, ExceptionDetail> _map = new Dictionary<KnownExceptionType, ExceptionDetail>
        {
            [KnownExceptionType.WordMixedCellWidths] = KnownExceptions.VSTO.MixedCellWidths
        };

        public static ExceptionDetail Get(KnownExceptionType type) => _map[type];
    }
}
