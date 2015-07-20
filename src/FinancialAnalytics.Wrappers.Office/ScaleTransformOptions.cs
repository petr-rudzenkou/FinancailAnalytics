namespace FinancialAnalytics.Wrappers.Office
{
    public class ScaleTransformOptions
    {
        public const float EmptyValue = -1;

        public float Top { get; set; }
        public float Left { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public bool LockAspectRatio { get; set; }
		public int ZOrder { get; set; }
		public float Rotation { get; set; }

		public ScaleTransformOptions(float top, float left, float width, float height, bool lockAspectRatio, int zOrder, float rotation)
		{
			Top = top;
			Left = left;
			Width = width;
			Height = height;
			LockAspectRatio = lockAspectRatio;
			ZOrder = zOrder;
			Rotation = rotation;
		}

		public ScaleTransformOptions(float top, float left, float width, float height, bool lockAspectRatio)
			: this(top, left, width, height, lockAspectRatio, 0, float.NaN)
		{
			
		}

        public ScaleTransformOptions()
            : this(EmptyValue, EmptyValue, EmptyValue, EmptyValue, false) {  }
    }
}
