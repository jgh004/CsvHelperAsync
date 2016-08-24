namespace CsvHelperAsync
{
    /// <summary>
    /// CSV 字段分隔符与限定符
    /// </summary>
    public class CsvFlag
    {
        private string qualifier;
        private string doubleQualifier;
        private char fieldQualifier;

        /// <summary>
        /// 获取字段限定符的字符串表示, 用于在序列化 CSV 字段时快速读取此值, 避免过于频繁的转换类型, 提高效率.
        /// </summary>
        internal string Qualifier
        {
            get
            {
                return this.qualifier;
            }
        }

        /// <summary>
        /// 获取连续两个字段限定符的字符串表示, 用于在序列化 CSV 字段时快速读取此值, 避免过于频繁的转换类型, 提高效率.
        /// </summary>
        internal string DoubleQualifier
        {
            get
            {
                return this.doubleQualifier;
            }
        }


        /// <summary>
        /// RFC4180 指定的 CSV 字段分隔符与限定符(即半角逗号与半角双引号)
        /// </summary>
        public static CsvFlag FlagForRFC4180
        {
            get; private set;
        }

        /// <summary>
        /// 获取字段分隔符
        /// </summary>
        public char FieldSeparator
        {
            get; private set;
        }

        /// <summary>
        /// 获取字段含有特殊字符时使用的限定字符
        /// </summary>
        public char FieldQualifier
        {
            get
            {
                return this.fieldQualifier;
            }
            private set
            {
                this.fieldQualifier = value;
                this.qualifier = value.ToString();
                this.doubleQualifier = new string( value, 2 );
            }
        }


        /// <summary>
        /// 使用指定的分隔符与限定符创建实例, 如果使用 RFC4180 标准(即半角逗号与半角双引号)可直接使用 CsvFlag.FlagForRFC4180 静态属性.
        /// </summary>
        /// <param name="separator">字段分隔符</param>
        /// <param name="enclosed">字段限定符</param>
        public CsvFlag( char separator, char enclosed )
        {
            this.FieldQualifier = enclosed;
            this.FieldSeparator = separator;
        }

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static CsvFlag()
        {
            FlagForRFC4180 = new CsvFlag( ',', '"' );
        }
    }
}
