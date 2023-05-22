using System;
using System.Collections.Generic;

namespace Diser
{
    internal class Model
    {
        public Model()
        {

            DCLinksCenterRight = new List<DCLink>();
            DCLinksCenterLeft = new List<DCLink>();
            GenerationCenter = new List<Generation>();
            GenerationRight = new List<Generation>();
            GenerationLeft = new List<Generation>();
            LoadCenter = new LoadCollection();
            LoadRight = new LoadCollection();
            LoadLeft = new LoadCollection();
        }
        public double FrequencyCenter { get; set; }
        public double FrequencyRight { get; set; }
        public double FrequencyLeft { get; set; }
        public List<DCLink> DCLinksCenterRight { get; set; }
        public List<DCLink> DCLinksCenterLeft { get; set; }
        public List<Generation> GenerationCenter { get; set; }
        public List<Generation> GenerationRight { get; set; }
        public List<Generation> GenerationLeft { get; set; }
        public LoadCollection LoadCenter { get; set; }
        public LoadCollection LoadRight { get; set; }
        public LoadCollection LoadLeft { get; set; }
        public Random Random { get; set; }
        public double Kload { get; set; }
        public double StartingFrequencyCenter { get; set; }
        public double StartingFrequencyRight { get; set; }
        public double StartingFrequencyLeft { get; set; }
        public double AngleFrequencyCenter { get; set; }
        public double AngleFrequencyRight { get; set; }
        public double AngleFrequencyLeft { get; set; }

        public void SetPowerDCLinks(double S)
        {
            foreach (var link in DCLinksCenterLeft)
                link.S += S;
            foreach (var link in DCLinksCenterRight)
                link.S += S;
        }
        public void SetDefaultS()
        {
            foreach (var link in DCLinksCenterLeft)
                link.S = link.DefaultS;
            foreach (var link in DCLinksCenterRight)
                link.S = link.DefaultS;
        }
        public void SetS()
        {
            foreach (var link in DCLinksCenterLeft)
                link.DefaultS = link.S;
            foreach (var link in DCLinksCenterRight)
                link.DefaultS = link.S;
        }
    }
}
