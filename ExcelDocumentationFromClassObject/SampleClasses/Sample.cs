using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;

namespace ExcelDocumentationFromClassObject.SampleClasses
{
    internal sealed class SampleClass
    {
        public string sc_field_1 { get; set; }

        public string sc_field_2 { get; set; }

        public string sc_field_3 { get; set; }

        public string sc_field_4 { get; set; }

        public string sc_field_5 { get; set; }

        public DemoClass demoClass { get; set; }

        public Collection<RandomClass> randomClasses { get; set; }

    }



    internal sealed class DemoClass

    {

        public string dc_field_1 { get; set; }

        public string dc_field_2 { get; set; }

        public string dc_field_3 { get; set; }

        public string dc_field_4 { get; set; }

        public string dc_field_5 { get; set; }

        public AdhocClass adhocClass { get; set; }

    }

    internal sealed class RandomClass

    {

        public List<string> rc_field_1 { get; set; }

        public string rc_field_2 { get; set; }

        public string rc_field_3 { get; set; }

        public string rc_field_4 { get; set; }

        public string rc_field_5 { get; set; }

    }

    internal sealed class AdhocClass

    {

        public string ac_field_1 { get; set; }

        public string ac_field_2 { get; set; }

        public string ac_field_3 { get; set; }

        public string ac_field_4 { get; set; }

        public List<string> ac_field_5 { get; set; }

    }


}
