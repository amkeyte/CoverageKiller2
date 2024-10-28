namespace CoverageKiller2.Logging
{
    public class ComplexObject
    {
        [UnsafeTrace]
        public string DoUnsafeStuff { get; } = "Logger: don't call this.";
    }
    public class SomeTestClass
    {
        private readonly Tracer Tracer = new Tracer(typeof(SomeTestClass));
        private int _A;

        public int PropA
        {
            get => Tracer.Trace(_A);
            set => _A = Tracer.Trace(value);
        }

        public void SomeMethod(ComplexObject co)
        {
            Tracer.StashProperty("co.DoUnsafeStuff_Before", co.DoUnsafeStuff);
            // ... do stuff with complex object
            Tracer.StashProperty("co.DoUnsafeStuff_After", co.DoUnsafeStuff);
        }

        public void SomeOtherMethod()
        {
            Tracer.Log("SomeOtherMethod was called.", "[LOGTEST]", new DataPoints(nameof(PropA)));
        }
    }

}
