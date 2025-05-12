using Serilog;
using System;
using System.Linq;
using System.Runtime.CompilerServices;

namespace CoverageKiller2.Logging

{
    public static class PingService
    {


        public static void Checkpoint(string msg, Type caller = null, [CallerMemberName] string callerName = "")
        {

            Log.Verbose($"{caller?.Name ?? _UNKNOWN_}::{callerName} --- {msg}");
        }



        private static int _pingDepth = 0;
        private const string _UNKNOWN_ = "???";

        private static string IndentBar() => string.Concat(Enumerable.Repeat("│  ", _pingDepth++));

        private static string IndentBarDecremented()
        {
            if (_pingDepth > 0) _pingDepth--;
            return string.Concat(Enumerable.Repeat("│  ", _pingDepth));
        }

        //public static void Ping(Type caller, [CallerMemberName] string callerMemberName = "")
        //{
        //    Log.Verbose($"{IndentBar()}-> Ping from {caller?.Name ?? _UNKNOWN_}::{callerMemberName}");
        //}
        // --- Ping Methods ---

        public static void Ping<T>([CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}");
        }

        public static void Ping<T>(string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Ping<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}");
        }

        public static void Ping<T>(this T caller, string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Ping<T>(this T caller, Type[] genericParams, [CallerMemberName] string callerName = "")
        {
            string genericParamsString = $"<{string.Join(",", genericParams.Select(p => p.Name))}>";
            Log.Verbose($"{IndentBar()}-> Ping from {typeof(T).Name}::{callerName}{genericParamsString}");
        }

        // --- Pong Methods ---

        public static void Pong<T>([CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}");
        }

        public static void Pong<T>(string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Pong<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}");
        }

        public static void Pong<T>(this T caller, string msg, [CallerMemberName] string callerName = "")
        {
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName} --- {msg}");
        }

        public static void Pong<T>(this T caller, Type[] genericParams, [CallerMemberName] string callerName = "")
        {
            string genericParamsString = $"<{string.Join(",", genericParams.Select(p => p.Name))}>";
            Log.Verbose($"{IndentBarDecremented()}<- Pong from {typeof(T).Name}::{callerName}{genericParamsString}");
        }

        public static void Pong<T>(this T caller, Type genericParam, [CallerMemberName] string callerName = "")
        {
            caller.Pong(new[] { genericParam }, callerName);
        }

        // --- PingPong Helpers ---

        public static void PingPong<T>(this T caller, [CallerMemberName] string callerName = "")
        {
            caller.Ping(callerName);
            caller.Pong(callerName);
        }

        public static void PingPong<T>(this T caller, string msg, [CallerMemberName] string callerName = "")
        {
            caller.Ping(msg, callerName);
            caller.Pong(msg, callerName);
        }
        public static TResult PingPong<T, TResult>(Func<TResult> action, string msg = null, [CallerMemberName] string callerName = "")
        {
            return PingPong(typeof(T), action, msg, callerName);
        }
        public static TResult PingPong<T, TResult>(this T caller, Func<TResult> action, string msg = null, [CallerMemberName] string callerName = "")
        {
            if (msg == null)
                caller.Ping(callerName);
            else
                caller.Ping(msg, callerName);

            var result = action();

            if (msg == null)
                caller.Pong(callerName);
            else
                caller.Pong(msg, callerName);

            return result;
        }

        public static TResult Pong<T, TResult>(this T caller, Func<TResult> action, string msg = null, [CallerMemberName] string callerName = "")
        {
            var result = action();

            if (msg == null)
                caller.Pong(callerName);
            else
                caller.Pong(msg, callerName);

            return result;
        }

    }
}
