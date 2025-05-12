using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace CoverageKiller2.Logging
{
    /// <summary>
    /// Marker attribute indicating that a property is unsafe for tracing.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class UnsafeTraceAttribute : Attribute { }

    /// <summary>
    /// Provides structured tracing and stashing of values for debugging.
    /// </summary>
    /// <remarks>CK2.00.00.0000</remarks>
    public class Tracer
    {
        public bool Enabled { get; set; } = true;
        public int IndentTabs { get; }

        private readonly Dictionary<string, string> _dict = new Dictionary<string, string>();
        private readonly Type _ownerType;

        public Tracer(Type ownerType, int indentTabs = 1)
        {
            _ownerType = ownerType;
            IndentTabs = indentTabs;
        }

        /// <summary>
        /// Stores a named value (converted to string).
        /// </summary>
        public void Stash(string name, object value)
        {
            _dict[name] = value?.ToString() ?? "[NULL]";
        }

        /// <summary>
        /// Retrieves a previously stashed value by name.
        /// </summary>
        public string Recall(string name)
        {
            return _dict.TryGetValue(name, out var value)
                ? (string.IsNullOrEmpty(value) ? " -- " : value)
                : "[NO RECALL]";
        }

        /// <summary>
        /// Tries to stash the value of a specific property from an object.
        /// Honors UnsafeTraceAttribute.
        /// </summary>
        public void StashProperty<T>(string propertyName, T target)
        {
            try
            {
                var prop = typeof(T).GetProperty(propertyName);

                if (prop == null)
                {
                    Serilog.Log.Debug("[WARNING] Trace => Could not Stash. Property {Property} is not a member of {Type}", propertyName, target?.GetType());
                    return;
                }

                if (Attribute.IsDefined(prop, typeof(UnsafeTraceAttribute)))
                {
                    Stash(propertyName, "[UNSAFE]");
                }
                else
                {
                    var val = prop.GetValue(target)?.ToString() ?? "[NULL]";
                    Stash(propertyName, val);
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Debug("Trace [ERROR] => Failed to stash property {Property} on {Type}: {Error}", propertyName, target?.GetType(), ex.Message);
                Debugger.Break();
            }
        }

        /// <summary>
        /// Stashes a value and returns it, using the caller's method name by default.
        /// </summary>
        public T Trace<T>(T operation, string name = "", [CallerMemberName] string callerName = "")
        {
            var finalName = string.IsNullOrWhiteSpace(name) ? callerName : name;
            Stash(finalName, operation);
            return operation;
        }

        [Flags]
        public enum LogOptions
        {
            None = 0,
            Ignore = 1,
            Force = 1 << 1
        }

        public void Log(string message, LogOptions options = LogOptions.None, [CallerMemberName] string memberName = "")
            => Log(message, string.Empty, new DataPoints(), options, memberName);

        public void Log(string message, IEnumerable<(string, object)> dataPoints, LogOptions options = LogOptions.None, [CallerMemberName] string memberName = "")
            => Log(message, string.Empty, dataPoints, options, memberName);

        public void Log(string message, string tag, LogOptions options = LogOptions.None, [CallerMemberName] string memberName = "")
            => Log(message, tag, new DataPoints(), options, memberName);

        public void Log(string message, string tag, IEnumerable<(string, object)> dataPoints, LogOptions options = LogOptions.None, [CallerMemberName] string memberName = "")
        {
            if (!options.HasFlag(LogOptions.Force) && (!Enabled || options.HasFlag(LogOptions.Ignore)))
                return;

            try
            {
                var formatted = $"Trace {tag ?? string.Empty} => {_ownerType.Name}.{memberName}";
                formatted += message is null ? "\n" : $" :: {message}";

                foreach (var (key, val) in dataPoints)
                {
                    if (val is DataPoints.Actions action && action == DataPoints.Actions.RecallValue)
                    {
                        formatted += $"\n{new string('\t', IndentTabs)}{key} = {Recall(key)}";
                    }
                    else
                    {
                        Stash(key, val);
                        formatted += $"\n{new string('\t', IndentTabs)}{key} = {val}";
                    }
                }

                Serilog.Log.Verbose(formatted);
            }
            catch (Exception ex)
            {
                Serilog.Log.Verbose($"Trace [ERROR] => Unexpected error in {nameof(Tracer)}.{nameof(Log)}: {message} :: {ex.Message}");
                Debugger.Break();
            }
        }
    }

    /// <summary>
    /// A collection of named data points used in logging and tracing.
    /// </summary>
    /// <remarks>CK2.00.00.0000</remarks>
    public class DataPoints : IEnumerable<(string, object)>
    {
        private readonly List<(string, object)> _items = new List<(string, object)>();

        public enum Actions { RecallValue }

        public DataPoints() { }

        public DataPoints(string recallName) => Add(recallName);

        public DataPoints(string name, object value) => Add(name, value);

        public DataPoints Add(string name, object value)
        {
            _items.Add((name, value));
            return this;
        }

        public DataPoints Add(string name)
        {
            _items.Add((name, (object)Actions.RecallValue));
            return this;
        }

        public IEnumerator<(string, object)> GetEnumerator() => _items.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// Extension method for enums to test flag presence (mainly for LogOptions).
    /// </summary>
    /// <remarks>CK2.00.00.0000</remarks>
    public static class TracerEnumExtensions
    {
        public static bool HasFlag(this Enum flags, Enum flag)
        {
            int flagsValue = Convert.ToInt32(flags);
            int flagValue = Convert.ToInt32(flag);
            return (flagsValue & flagValue) == flagValue;
        }
    }

    public class TracerHelpers
    {
        public enum PP
        {
            Enter,
            Result,
            TestPoint,
            PropertyGet,
            PropertySet,
        }
        public static string TraceCaller(params object[] paramPairs)
        {
            // Validate the length of pairs (must be even)
            if (paramPairs.Length < 4 || paramPairs.Length % 2 != 0)
            {
                throw new ArgumentException("The number of parameters must be even and at least 4 (traceType, msg, className, methodName).");
            }

            // Validate traceType
            if (paramPairs[0] == null || paramPairs[0].GetType() != typeof(PP))
            {
                throw new ArgumentException("Bad caller type flag at param1.");
            }

            PP traceType = (PP)paramPairs[0];
            string msg = paramPairs[1] as string ?? string.Empty;

            // Shared logic to format trace msg
            return FormatTracemsg(traceType, msg, paramPairs);
        }

        private static string FormatTracemsg(PP traceType, string msg, object[] paramPairs)
        {
            string defaultmsg;

            // Standard switch statement instead of switch expression
            switch (traceType)
            {
                case PP.Enter:
                    defaultmsg = "Entering member:";
                    break;
                case PP.Result:
                    defaultmsg = "Member returned:";
                    break;
                case PP.TestPoint:
                    defaultmsg = "Test point:";
                    break;
                case PP.PropertyGet:
                    defaultmsg = "Property returned:";
                    break;
                case PP.PropertySet:
                    defaultmsg = "Property set to:";
                    break;
                default:
                    throw new ArgumentException("Invalid trace type.");
            }

            msg = string.IsNullOrEmpty(msg) ? defaultmsg : msg;

            string className = paramPairs[2].ToString();
            string methodName = paramPairs[3].ToString();
            string formattedPairs = string.Empty;

            // Formatting name-value pairs
            for (int i = 4; i < paramPairs.Length; i += 2)
            {
                string name = paramPairs[i].ToString();
                string value = paramPairs[i + 1].ToString();
                formattedPairs += $"\t\t[{name} = {value}]";
            }
            formattedPairs = string.IsNullOrEmpty(formattedPairs) ? "" : "\n" + formattedPairs;

            return $"TRACE => {className}.{methodName} :: {msg}{formattedPairs}";
        }

    }
}
