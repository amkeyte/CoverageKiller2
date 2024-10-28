using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace CoverageKiller2.Logging
{
    [AttributeUsage(AttributeTargets.Property)]
    public class UnsafeTraceAttribute : Attribute
    {
        // This attribute class doesn't need to store any values.
        // It serves as a marker to identify properties as unsafe for tracing.
    }
    public class Tracer
    {
        public bool Enabled { get; set; } = true;
        private readonly Dictionary<string, string> _dict = new Dictionary<string, string>();
        private readonly Type _ownerType;
        public Tracer(Type ownerType)
        {
            _ownerType = ownerType;
        }


        public void Stash(string name, object value)
        {
            _dict[name] = value?.ToString() ?? "[NULL]";
        }
        public string Recall(string name)
        {
            var recallSuccess = _dict.TryGetValue(name, out var value);
            if (!recallSuccess)
                value = "[NO RECALL]";
            else if (value == string.Empty)
                value = " -- ";

            return value;
        }
        public void StashProperty<T>(string propertyName, T valueToStash)
        {
            try
            {
                var propertyInfo = typeof(T).GetProperty(propertyName);

                if (propertyInfo == null)
                {
                    Serilog.Log.Debug("[WARNING] Trace => Could not Stash. Property {propertyName} is not a member of {valueToStashType}",
                        propertyName,
                        valueToStash.GetType());
                    return; // Exit early if property is not found
                }

                // Check if the property has the UnsafeTraceAttribute
                if (Attribute.IsDefined(propertyInfo, typeof(UnsafeTraceAttribute)))
                {
                    Stash(propertyName, "[UNSAFE]");
                }
                else
                {
                    // Safely access and stash property value if no UnsafeTrace attribute is present
                    try
                    {
                        var value = propertyInfo.GetValue(valueToStash)?.ToString() ?? "[NULL]";
                        Stash(propertyName, value);
                    }
                    catch (Exception ex)
                    {
                        Serilog.Log.Debug("Trace [ERROR] => Failed to access property {propertyName} on {valueToStashType}: {message}",
                            propertyName,
                            valueToStash.GetType(),
                            ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Debug("Trace [ERROR] => Unexpected error in StashProperty: {message}", ex.Message);
                Debugger.Break();
            }
        }

        public T Trace<T>(T operation, string name = "", [CallerMemberName] string callerName = "")
        {
            name = !string.IsNullOrEmpty(name) ? name : callerName;
            {
                Stash(name, operation);
                return operation;
            }
        }

        [Flags]
        public enum LogOptions
        {
            None = 0,
            Ignore = 1,
            Force = 1 << 1
        }

        public void Log(
            string message,
            LogOptions logOptions = LogOptions.None,
            [CallerMemberName] string memberName = "")
        {
            Log(message,
                string.Empty,
                new DataPoints(),
                logOptions,
                memberName);
        }

        public void Log(
            string message,
            IEnumerable<(string, object)> dataPoints,
            LogOptions logOptions = LogOptions.None,
            [CallerMemberName] string memberName = "")
        {
            Log(message,
                string.Empty,
                dataPoints,
                logOptions,
                memberName);
        }

        public void Log(
            string message,
            string tag,
            LogOptions logOptions = LogOptions.None,
            [CallerMemberName] string memberName = "")
        {
            Log(message,
                tag,
                new DataPoints(),
                logOptions,
                memberName);
        }



        public void Log(
            string message,
            string tag,
            IEnumerable<(string, object)> dataPoints,
            LogOptions options = LogOptions.None,
            [CallerMemberName] string memberName = "")
        {
            if (!options.HasFlag(LogOptions.Force))
            {
                if (options.HasFlag(LogOptions.Ignore) || !Enabled) return;
            }

            try
            {
                // Prepare the base of the message template
                var formattedMessage = $"Trace {tag ?? string.Empty} => {_ownerType.Name}.{memberName}";

                // Check if the message is null
                formattedMessage += message is null ? "\n" : $" :: {message}";




                // Add recalled values to the message
                foreach (var dataPoint in dataPoints)
                {
                    string stashedName = string.Empty;

                    if (dataPoint.Item2 is DataPoints.Actions action && action == DataPoints.Actions.RecallValue)
                    {
                        stashedName = dataPoint.Item1;
                    }
                    else
                    {
                        stashedName = dataPoint.Item1;
                        Stash(stashedName, dataPoint.Item2);
                    }

                    formattedMessage += $"\n\t\t{stashedName} = {Recall(stashedName)}";
                }

                // Perform the logging
                Serilog.Log.Debug(formattedMessage);
            }
            catch (Exception ex)
            {
                Serilog.Log.Debug($"Trace [ERROR] => Unexpected error in {nameof(Tracer)}.{nameof(Log)}: {message}", ex.Message);
                Debugger.Break();
            }
        }

        //internal void Log(string message, string tag, Type className, string memberName)
        //{
        //    if (!Enabled) return;

        //    var formattedMessage = $"Trace {tag ?? string.Empty} => {className.Name}.{memberName}";

        //    // Check if the message is null
        //    formattedMessage += message is null ? "\n" : $" :: {message}";

        //    Serilog.Log.Debug(formattedMessage);
        //}


    }
    public class DataPoints : IEnumerable<(string, object)>
    {
        private readonly List<(string, object)> _items = new List<(string, object)>();
        public enum Actions
        {
            RecallValue
        }

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

        public DataPoints()
        {

        }
        public DataPoints(string recallName)
        {
            _items.Add((recallName, (object)Actions.RecallValue));
        }

        public DataPoints(string dataPointName, object dataPoint)
        {
            _items.Add((dataPointName, dataPoint));
        }
        //public DataPoints(params object[] dataPoints)
        //{
        //    foreach (object item in dataPoints)
        //    {
        //        (string, object) dataPoint;

        //        bool isTwoParamTuple = item != null
        //            && item.GetType().IsGenericType
        //            && item.GetType().GetGenericTypeDefinition() == typeof(ValueTuple<,>);
        //        if (isTwoParamTuple)
        //        {
        //            dataPoint = //what to do here?


        //        }
        //        else if (item is string)
        //        {
        //            dataPoint = ((string)item, Actions.RecallValue);
        //        }
        //        else
        //        {
        //            dataPoint = ("[ERROR]",
        //                $"Attempt to log invalid Datapoint: {item.ToString()}. Are you casting your Tuple item 2 to (object)?");
        //        }

        //        _items.Add(dataPoint);
        //    }
        //}

        public IEnumerator<(string, object)> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
    public static class TracerEnumExtensions
    {
        public static bool HasFlag(this Enum flags, Enum flag)
        {
            // Convert enums to integers and perform a bitwise AND comparison
            int flagsValue = Convert.ToInt32(flags);
            int flagValue = Convert.ToInt32(flag);
            return (flagsValue & flagValue) == flagValue;
        }
    }
}
