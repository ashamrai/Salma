using System;

namespace WordToTFS
{
    public static class TextUtil
    {
        public enum TextStyle
        {
            Normal,
            Title,
            Heading1,
            Heading2,
            Heading3,
            
        }

        public static String GetStyleName(TextUtil.TextStyle style)
        {
            string heading = ResourceHelper.GetResourceString("WORD_HEADING_RESOURCE");
            string title = ResourceHelper.GetResourceString("WORD_TITLE_RESOURCE");
            string normal = ResourceHelper.GetResourceString("WORD_NORMAL_RESOURCE");


            switch (style)
            {
                case TextStyle.Heading1:
                return heading + " 1";

                case TextStyle.Heading2:
                return heading + " 2";

                case TextStyle.Heading3:
                return heading + " 3";

                case TextStyle.Title:
                return title;

                case TextStyle.Normal:
                default:
                return normal;
            }
        }
    }
}
