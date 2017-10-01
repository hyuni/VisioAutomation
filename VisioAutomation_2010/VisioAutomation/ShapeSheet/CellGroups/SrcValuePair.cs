namespace VisioAutomation.ShapeSheet.CellGroups
{
    public struct SrcValuePair
    {
        public readonly ShapeSheet.Src Src;
        public readonly string Value;
        public readonly string SrcName;
        public readonly string Name;

        public SrcValuePair(ShapeSheet.Src src, string value)
        {
            this.Src = src;
            this.Value = value;
            this.Name = null;
            this.SrcName = null;
        }

        public SrcValuePair(ShapeSheet.Src src, string value, string srcname, string name)
        {
            this.Src = src;
            this.Value = value;
            this.Name = srcname;
            this.SrcName = name;
        }

        public static SrcValuePair Create(ShapeSheet.Src src, string value)
        {
            return new SrcValuePair(src,value,null,null);
        }

        public static SrcValuePair Create(ShapeSheet.Src src, CellValueLiteral cvf)
        {
            return new SrcValuePair(src, cvf.Value, null, null);
        }

        public static SrcValuePair Create(ShapeSheet.Src src, string value, string srcname, string name)
        {
            return new SrcValuePair(src, value, srcname, name);
        }

        public static SrcValuePair Create(ShapeSheet.Src src, CellValueLiteral cvf, string srcname, string name)
        {
            return new SrcValuePair(src, cvf.Value, srcname, name);
        }
    }
}