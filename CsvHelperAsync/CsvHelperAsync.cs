/* Description:	标准CSV文件存取类，字段以","分隔，以‘"’为限定符,也可自定义分隔符与限定符.采用边读写边处理方式，减少内存占用, 可用于超大文件读写。
 *				使用方法: //读csv, ReadAsync 方法只能调用一次, 通过进度事件返回读取的数据.
 *				        var csvR = new CsvReadHelper( stream or fileName, ... ).ReadAsync<T>( IProgress<CsvReadProgressInfo<T>> 类型的参数, ... );
 *				        csvR.Close();
 *				         
 *				        //写csv, WriteAsync 方法可调用多次.
 *				        var csvW = new CsvWriteHelper( stream or fileName, ... ).WriteAsync( ... );
 *				        csvW.Flush();
 *				        csvW.Close();
 *				IETF标准	https://tools.ietf.org/html/rfc4180
 * Creator:		IT农民工
 * Home:		www.ITnmg.net
 * Create date:	2011.12.01
 * Modified date:	2016.08.24
 * Version:		0.9.0
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

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

    /// <summary>
    /// 读取 csv 文件时, 进度报告中的数据.
    /// </summary>
    /// <typeparam name="T">每行数据要转换成的实体类</typeparam>
    public class CsvReadProgressInfo<T> where T : new()
    {
        /// <summary>
        /// 获取是否读取完毕
        /// </summary>
        public bool IsComplete
        {
            get;
            internal set;
        }

        /// <summary>
        /// 获取列标题集合, 如果指定了将 csv 第一行做为列标题, 则返回第一行的数据; 如果没有指定, 则返回各列的索引.
        /// </summary>
        public List<string> ColumnNames
        {
            get;
            internal set;
        }

        /// <summary>
        /// 获取当前批次的数据行集合
        /// </summary>
        public List<T> RowsData
        {
            get;
            internal set;
        }

        /// <summary>
        /// 获取已读取的字节数
        /// </summary>
        public long CurrentBytes
        {
            get;
            internal set;
        }

        /// <summary>
        /// 获取总字节数
        /// </summary>
        public long TotalBytes
        {
            get;
            internal set;
        }

        /// <summary>
        /// 获取当前进度(已读字节数 / 总字节数)
        /// </summary>
        public decimal ProgressValue
        {
            get
            {
                return this.TotalBytes <= 0 || this.CurrentBytes < 0 ? 0 : this.CurrentBytes / (decimal)this.TotalBytes * 100;
            }
        }


        /// <summary>
        /// 默认构造函数
        /// </summary>
        public CsvReadProgressInfo()
        {
            this.IsComplete = false;
            this.ColumnNames = new List<string>();
            this.RowsData = new List<T>();
            this.CurrentBytes = 0;
            this.TotalBytes = 0;
        }
    }

    /// <summary>
    /// CSV 读取类
    /// </summary>
    public class CsvReadHelper
    {
        /// <summary>
        /// 读取或写入 CSV 数据的流
        /// </summary>
        private StreamReader CsvStream;

        /// <summary>
        /// 获取或设置是否可读取流, 用于防止多次调用 ReadAsync 方法.
        /// </summary>
        private volatile bool CanRead;


        /// <summary>
        /// 获取读取流时的内存缓冲字节数, 默认为 40960.
        /// </summary>
        public int ReadStreamBufferLength
        {
            get; private set;
        }

        /// <summary>
        /// 获取是否将第一行数据做为列标题
        /// </summary>
        public bool FirstRowIsHead
        {
            get; private set;
        }

        /// <summary>
        /// 获取 csv 列数
        /// </summary>
        public int ColumnCount
        {
            get; private set;
        }

        /// <summary>
        /// 获取写入的总行数
        /// </summary>
        public long TotalRowCount
        {
            get; private set;
        }

        /// <summary>
        /// 获取 csv 字段分隔符与限定符
        /// </summary>
        public CsvFlag Flag
        {
            get; private set;
        }

        /// <summary>
        /// 获取字符编码
        /// </summary>
        public Encoding DataEncoding
        {
            get; private set;
        }


        /// <summary>
        /// 初始化读取流
        /// </summary>
        /// <param name="stream">要读取的流</param>
        /// <param name="dataEncoding">字符编码</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="firstRowIsHead">是否将第一行数据做为标题行</param>
        /// <param name="readStreamBufferLength">流读取缓冲大小, 默认为 40960 字节.</param>
        public CsvReadHelper( Stream stream, Encoding dataEncoding, CsvFlag flag, bool firstRowIsHead = true, int readStreamBufferLength = 40960 )
        {
            this.CanRead = true;
            this.ColumnCount = 0;
            this.TotalRowCount = 0L;
            this.DataEncoding = dataEncoding;
            this.Flag = flag;
            this.FirstRowIsHead = firstRowIsHead;
            this.ReadStreamBufferLength = readStreamBufferLength;
            this.CsvStream = new StreamReader( stream, this.DataEncoding, false, this.ReadStreamBufferLength );
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码初始化读取流
        /// </summary>
        /// <param name="stream">要读取的流</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="firstRowIsHead">是否将第一行数据做为标题行</param>
        /// <param name="readStreamBufferLength">流读取缓冲大小, 默认为 40960 字节.</param>
        public CsvReadHelper( Stream stream, CsvFlag flag, bool firstRowIsHead = true, int readStreamBufferLength = 40960 )
            : this( stream, Encoding.UTF8, flag, firstRowIsHead, readStreamBufferLength )
        {
        }

        /// <summary>
        /// 初始化读取文件
        /// </summary>
        /// <param name="csvFileName">要读取的文件路径及名称</param>
        /// <param name="dataEncoding">字符编码</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="firstRowIsHead">是否将第一行数据做为标题行</param>
        /// <param name="readStreamBufferLength">流读取缓冲大小, 默认为 40960 字节.</param>
        public CsvReadHelper( string csvFileName, Encoding dataEncoding, CsvFlag flag, bool firstRowIsHead = true, int readStreamBufferLength = 40960 )
            : this( File.Open( csvFileName, FileMode.Open ), dataEncoding, flag, firstRowIsHead, readStreamBufferLength )
        {
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码初始化读取文件
        /// </summary>
        /// <param name="csvFileName">要读取的文件路径及名称</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="firstRowIsHead">是否将第一行数据做为标题行</param>
        /// <param name="readStreamBufferLength">流读取缓冲大小, 默认为 40960 字节.</param>
        public CsvReadHelper( string csvFileName, CsvFlag flag, bool firstRowIsHead = true, int readStreamBufferLength = 40960 )
            : this( csvFileName, Encoding.UTF8, flag, firstRowIsHead, readStreamBufferLength )
        {
        }



        /// <summary>
        /// 异步读取, 每读取 readProgressSize 条记录或到文件末尾触发通知事件. 此方法只能调用一次, 如果多次调用会产生异常.
        /// </summary>
        /// <typeparam name="T">数据行转换时对应的实体类型</typeparam>
        /// <param name="progress">通知方法</param>
        /// <param name="expression">数据行转换为 T 实例的方法</param>
        /// <param name="cancelToken">取消参数</param>
        /// <param name="readProgressSize">每读取多少行数据触发通知事件, 默认为 1000.</param>
        /// <returns></returns>
        public async Task ReadAsync<T>( IProgress<CsvReadProgressInfo<T>> progress, Func<List<string>, T> expression, CancellationToken cancelToken
            , int readProgressSize = 1000 ) where T : new()
        {
            #region Check params

            if ( CanRead )
            {
                CanRead = false;
            }
            else
            {
                throw new Exception( "ReadAsync method allows only one call" );
            }

            if ( cancelToken.IsCancellationRequested )
            {
                cancelToken.ThrowIfCancellationRequested();
            }

            if ( progress == null )
            {
                throw new ArgumentNullException( "progress" );
            }

            if ( expression == null )
            {
                throw new ArgumentNullException( "expression" );
            }

            if ( readProgressSize <= 0 )
            {
                throw new ArgumentException( "The property 'readProgressSize' must be greater than 0" );
            }

            #endregion

            //标题行
            List<string> columnNames = new List<string>();
            //通过通知事件返回的数据形式, 每次通知后将清空
            List<T> rowsData = null;
            //用于临时存放不足一行的数据
            List<char> subLine = new List<char>();

            //获得数据流总字节数
            long totalBytes = this.CsvStream.BaseStream.Length;
            //当前读取字节数
            long currentBytes = 0;
            //每次读取字节缓冲区
            char[] buffer = new char[ReadStreamBufferLength];
            //开始循环读取数据
            while ( !this.CsvStream.EndOfStream )
            {
                if ( cancelToken.IsCancellationRequested )
                {
                    cancelToken.ThrowIfCancellationRequested();
                }

                //读取一块数据
                int count = await CsvStream.ReadBlockAsync( buffer, 0, buffer.Length );
                currentBytes = CsvStream.BaseStream.Position;
                //这块数据的字节数组
                char[] input = null;

                //如果读满数组
                if ( count == buffer.Length )
                {
                    //直接复制
                    input = buffer;
                }
                else if ( count < buffer.Length ) //如果填不满数组
                {
                    //缩小数据到实际大小
                    input = new char[count];
                    Array.Copy( buffer, 0, input, 0, count );
                }

                //取出完整行数据和剩余不满一行数据
                List<List<string>> rows = this.GetRows( input, ref subLine );

                //如果到了文件流末尾且还有未处理的字符,用不检查末尾换行符方式处理余下字符.
                if ( CsvStream.EndOfStream && subLine.Count > 0 )
                {
                    List<char> tSubline = new List<char>();
                    //值复制,不能直接用等于, 否则是引用类型.
                    List<char> tInput = new List<char>( subLine );
                    tInput.AddRange( new char[] { '\r', '\n' } );//在末尾添加换行符
                    List<List<string>> lastRows = this.GetRows( tInput.ToArray(), ref tSubline );

                    if ( lastRows != null )
                    {
                        if ( rows == null )
                        {
                            rows = lastRows;
                        }
                        else
                        {
                            rows.AddRange( lastRows );
                        }

                        //如果还有剩余字符,说明格式错误
                        if ( tSubline.Count > 0 )
                        {
                            throw new Exception( "The csv file format error!" );
                        }

                        //为下一块数据使用准备
                        subLine.Clear();
                    }
                }

                Progress( currentBytes, totalBytes, progress, expression, readProgressSize, rows, ref columnNames, ref rowsData );
            }

            //资料有不完整的行，或者读取错位导致剩余。
            if ( subLine.Count > 0 )
            {
                throw new Exception( "The csv file format error!" );
            }
        }

        /// <summary>
        /// 关闭流
        /// </summary>
        public void Close()
        {
            if ( this.CsvStream != null )
            {
                this.CsvStream.Close();
            }
        }


        /// <summary>
        /// 组织数据, 发送通知.
        /// </summary>
        /// <typeparam name="T">通知中数据行的实体类型</typeparam>
        /// <param name="currentBytes">当前字节数</param>
        /// <param name="totalBytes">流字节总数</param>
        /// <param name="progress">通知方法</param>
        /// <param name="expression">转换类型方法</param>
        /// <param name="readProgressSize">多少条数据触发通知</param>
        /// <param name="rows">原始的字符串数据集合</param>
        /// <param name="columnNames">标题行</param>
        /// <param name="rowsData">转换后的数据</param>
        private void Progress<T>( long currentBytes, long totalBytes
            , IProgress<CsvReadProgressInfo<T>> progress, Func<List<string>, T> expression, int readProgressSize
            , List<List<string>> rows, ref List<string> columnNames, ref List<T> rowsData ) where T : new()
        {//生成通知数据
            if ( rows != null && rows.Count > 0 )
            {
                //当返回第一批数据时,将首行设为标题行.
                if ( rowsData == null )
                {
                    rowsData = new List<T>();
                    List<string> firstRow = rows[0];

                    //如果第一行做为标题
                    if ( this.FirstRowIsHead )
                    {
                        columnNames = firstRow;
                        //从数据中移除第一行
                        rows.Remove( firstRow );
                    }
                    else
                    {
                        //否则用字段索引做标题行
                        for ( int i = 0; i < firstRow.Count; i++ )
                        {
                            columnNames.Add( i.ToString() );
                        }
                    }

                    this.ColumnCount = columnNames.Count;
                }

                //加入数据行
                for ( int i = 0; i < rows.Count; i++ )
                {
                    rowsData.Add( expression.Invoke( rows[i] ) );
                    this.TotalRowCount++; //读到通知数据里才算读取

                    //当读取批次满足指定返回行数时
                    if ( rowsData.Count == readProgressSize )
                    {
                        CsvReadProgressInfo<T> info = new CsvReadProgressInfo<T>();
                        info.ColumnNames = columnNames;
                        info.RowsData = rowsData;
                        info.IsComplete = CsvStream.EndOfStream && i + 1 == rowsData.Count;
                        info.CurrentBytes = currentBytes;
                        info.TotalBytes = totalBytes;
                        //重置通知数据
                        rowsData = new List<T>();
                        progress.Report( info );//异步触发事件
                    }
                    else if ( CsvStream.EndOfStream && i + 1 == rows.Count ) //当读取批次不足指定行数且到了流末尾时
                    {
                        CsvReadProgressInfo<T> info = new CsvReadProgressInfo<T>();
                        info.ColumnNames = columnNames;
                        info.RowsData = rowsData;
                        info.IsComplete = true;
                        info.CurrentBytes = currentBytes;
                        info.TotalBytes = totalBytes;
                        //重置通知数据
                        rowsData = new List<T>();
                        progress.Report( info );
                    }
                }
            }
        }

        /// <summary>
        /// 取出完整数据行, 不满一行的数据保存在 subLine 中.
        /// </summary>
        /// <param name="input">原始字符数组</param>
        /// <param name="subLine">不满一行的数据</param>
        /// <returns>解析后的数据行集合</returns>
        private List<List<string>> GetRows( char[] input, ref List<char> subLine )
        {
            List<List<string>> result = null;

            if ( input != null )
            {
                //将上一部分不足一行的数据与新数据合并
                subLine.AddRange( input );
                List<List<char>> charLines = new List<List<char>>();//得到的所有行
                int qualifierQty = 0;//每行限定符数量
                List<char> line = new List<char>();//每行内容

                //找出完整的行备用
                for ( int i = 0; i < subLine.Count; i++ )
                {
                    char c = subLine[i];

                    if ( c == this.Flag.FieldQualifier )//如果是限定符,统计数量
                    {
                        qualifierQty++;
                    }

                    if ( c == '\n' )//如果是换行符
                    {
                        if ( qualifierQty % 2 == 0 )//如果限定符成对,说明是字段后换行, 即一行结束.
                        {
                            //一行结束时, 如果前一个字符是\r, 移除掉.
                            if ( i - 1 >= 0 && subLine[i - 1] == '\r' && line.Count > 0 )
                            {
                                line.RemoveAt( line.Count - 1 );//移除之前加入的\r
                            }

                            if ( line.Count > 0 )//空行不加入
                            {
                                charLines.Add( line );
                                line = new List<char>();
                            }

                            qualifierQty = 0;
                        }
                        else//如果限定符不成对,说明是字段中的换行符,加入字段中.
                        {
                            line.Add( c );
                        }
                    }
                    else//如果不是换行符,直接加入字段中.
                    {
                        line.Add( c );
                    }
                }

                //将最后不足一行的数据传出去
                subLine = line;
                result = this.DeserializeRows( charLines.ToArray() );
            }

            return result;
        }

        /// <summary>
        /// 找出每行的字段并还原转义字符.
        /// </summary>
        /// <param name="charLines">未还原的数据行集合</param>
        /// <returns>还原后的数据行集合</returns>
        private List<List<string>> DeserializeRows( params IList<char>[] charLines )
        {
            List<List<string>> result = null;

            if ( charLines != null )
            {
                //遍历每行，找出每个字段。
                foreach ( var line in charLines )
                {
                    int enclosedQty = 0;//统计限定符数量
                    List<string> sline = new List<string>();//一行
                    List<char> field = new List<char>();//一个字段

                    //遍历一行数据的所有字符
                    for ( int i = 0; i < line.Count; i++ )
                    {
                        char c = line[i];

                        if ( c == this.Flag.FieldQualifier )//如果是限定符,统计数量
                        {
                            enclosedQty++;
                        }

                        if ( c == this.Flag.FieldSeparator )//如果是分隔符
                        {
                            if ( enclosedQty % 2 == 0 )//双引号成对,说明字段完整,加入到行中.
                            {
                                sline.Add( DeserializeField( new string( field.ToArray() ) ) );
                                field.Clear();//重新收集字段
                                enclosedQty = 0;
                            }
                            else//限定符不成对,属于字段内的字符,加入字段中.
                            {
                                field.Add( c );
                            }
                        }
                        else//不是分隔符,直接加入字段中.
                        {
                            field.Add( c );
                        }

                        if ( i + 1 == line.Count )//到了一行结尾,将最后一个字段加入行中.
                        {
                            sline.Add( DeserializeField( new string( field.ToArray() ) ) );
                        }
                    }

                    //初始化返回数据
                    if ( result == null )
                    {
                        //List<string> columns = new List<string>();

                        //for ( int i = 0; i < sline.Count; i++ )
                        //{
                        //    columns.Add( i.ToString() );
                        //}

                        result = new List<List<string>>();
                    }

                    result.Add( sline );
                }
            }

            return result;
        }

        /// <summary>
        /// 还原 CSV 字段值,将两个相邻限定符替换为一个限定符,去掉两边的限定符.
        /// </summary>
        /// <param name="field">待还原的字段</param>
        /// <param name="trimField">是否去除无限定符包围字段两端的空格，默认保留.</param>
        /// <returns>还原转义符后的字段</returns>
        private string DeserializeField( string field, bool trimField = false )
        {
            string result = field;

            if ( trimField )
            {
                result = result.Trim();
            }

            //当字段包含限定符时
            if ( result.IndexOf( this.Flag.FieldQualifier ) >= 0 )
            {
                result = result.Trim();//先去除左右空格

                //当字符数小于2个(限定符不成对)或第1个和最后1个字符不是限定符时,说明字段格式错误,引发异常.
                if ( result.Length < 2 || (result[0] != this.Flag.FieldQualifier || result[result.Length - 1] != this.Flag.FieldQualifier) )
                {
                    throw new Exception( "The input field '" + field + "' is invalid!" );
                }
                else
                {
                    //result = result.Substring( 1, result.Length - 2 ).Replace( new string( this.Flag.FieldEnclosed, 2 ), this.Flag.FieldEnclosed.ToString() );
                    result = result.Substring( 1, result.Length - 2 ).Replace( this.Flag.DoubleQualifier, this.Flag.Qualifier );
                }
            }

            return result;
        }
    }

    /// <summary>
    /// CSV 写入类
    /// </summary>
    public class CsvWriteHelper
    {
        /// <summary>
        /// 读取或写入 CSV 数据的流
        /// </summary>
        private StreamWriter CsvStream;

        /// <summary>
        /// 获取或设置通知事件参数
        /// </summary>
        private IProgress<int> Progress;


        /// <summary>
        /// 获取每行的字段数
        /// </summary>
        public int ColumnCount
        {
            get; private set;
        }

        /// <summary>
        /// 获取写入的总行数
        /// </summary>
        public long TotalRowCount
        {
            get; private set;
        }

        /// <summary>
        /// 获取字段分隔符与限定符, 默认为 CsvFlag.FlagForRFC4180
        /// </summary>
        public CsvFlag Flag
        {
            get; private set;
        }

        /// <summary>
        /// 获取当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0
        /// </summary>
        public int WriteProgressSize
        {
            get; private set;
        }

        /// <summary>
        /// 获取字符编码,默认为 Encoding.UTF8
        /// </summary>
        public Encoding DataEncoding
        {
            get; private set;
        }

        /// <summary>
        /// 获取取消操作的token, 默认为 CancellationToken.None
        /// </summary>
        public CancellationToken CancelToken
        {
            get; private set;
        }


        /// <summary>
        /// 初始化写入流
        /// </summary>
        /// <param name="stream">要写入的流</param>
        /// <param name="dataEncoding">字符编码</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="cancelToken">取消操作的token</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( Stream stream, Encoding dataEncoding, CsvFlag flag
            , CancellationToken cancelToken, IProgress<int> progress = null, int writeProgressSize = 1000 )
        {
            this.ColumnCount = 0;
            this.TotalRowCount = 0L;
            this.DataEncoding = dataEncoding;
            this.Flag = flag;
            this.CancelToken = cancelToken;
            this.Progress = progress;
            this.WriteProgressSize = writeProgressSize;
            this.CsvStream = new StreamWriter( stream, this.DataEncoding );
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码初始化写入流
        /// </summary>
        /// <param name="stream">要写入的流</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="cancelToken">取消操作的token</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( Stream stream, CsvFlag flag, CancellationToken cancelToken, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( stream, Encoding.UTF8, flag, cancelToken, progress, writeProgressSize )
        {
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码, CsvFlag.FlagForRFC4180 分隔符与限定符 初始化写入流.
        /// </summary>
        /// <param name="stream">要写入的流</param>
        /// <param name="cancelToken">取消操作的token</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( Stream stream, CancellationToken cancelToken, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( stream, Encoding.UTF8, CsvFlag.FlagForRFC4180, cancelToken, progress, writeProgressSize )
        {
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码, CsvFlag.FlagForRFC4180 分隔符与限定符 初始化写入流, 且不允许取消操作.
        /// </summary>
        /// <param name="stream">要写入的流</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( Stream stream, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( stream, CancellationToken.None, progress, writeProgressSize )
        {
        }


        /// <summary>
        /// 初始化写入文件
        /// </summary>
        /// <param name="csvFileName">csv 文件路径及名称</param>
        /// <param name="dataEncoding">字符编码</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="cancelToken">取消操作的token</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( string csvFileName, Encoding dataEncoding, CsvFlag flag
            , CancellationToken cancelToken, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( File.Open( csvFileName, FileMode.Create ), dataEncoding, flag, cancelToken, progress, writeProgressSize )
        {
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码初始化写入文件
        /// </summary>
        /// <param name="csvFileName">csv 文件路径及名称</param>
        /// <param name="flag">csv 字段分隔符与限定符</param>
        /// <param name="cancelToken">取消操作的token</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( string csvFileName, CsvFlag flag, CancellationToken cancelToken, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( csvFileName, Encoding.UTF8, flag, cancelToken, progress, writeProgressSize )
        {
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码, CsvFlag.FlagForRFC4180 分隔符与限定符 初始化写入文件.
        /// </summary>
        /// <param name="csvFileName">csv 文件路径及名称</param>
        /// <param name="cancelToken">取消操作的token</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( string csvFileName, CancellationToken cancelToken, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( csvFileName, Encoding.UTF8, CsvFlag.FlagForRFC4180, cancelToken, progress, writeProgressSize )
        {
        }

        /// <summary>
        /// 使用 Encoding.UTF8 编码, CsvFlag.FlagForRFC4180 分隔符与限定符 初始化写入文件, 且不允许取消操作.
        /// </summary>
        /// <param name="csvFileName">csv 文件路径及名称</param>
        /// <param name="progress">通知事件参数, 每次通返回自上次通知以来写入的行数. 默认为 null, 表示不通知.</param>
        /// <param name="writeProgressSize">当写入多少条数据时应触发进度通知事件, 默认为 1000, 此值应大于 0.</param>
        public CsvWriteHelper( string csvFileName, IProgress<int> progress = null, int writeProgressSize = 1000 )
            : this( csvFileName, CancellationToken.None, progress, writeProgressSize )
        {
        }



        /// <summary>
        /// 异步写入单行数据, 可多次执行, 之后执行 Close 方法关闭写入流.
        /// </summary>
        /// <param name="rowData">一行数据,由字段集合组成.</param>
        /// <returns>Task</returns>
        public async Task WriteLineAsync( IList<string> rowData )
        {
            if ( this.CancelToken.IsCancellationRequested )
            {
                CancelToken.ThrowIfCancellationRequested();
            }

            if ( this.CsvStream == null )
            {
                throw new NullReferenceException( "the csv stream is null" );
            }

            if ( rowData == null )
            {
                throw new ArgumentNullException( "rowData" );
            }

            //如果写入过一条数据, 则字段数固定. 如果再次写入的字段数不同, 报异常.
            if ( this.ColumnCount > 0 && this.ColumnCount != rowData.Count )
            {
                throw new ArgumentException( "the rowData count must be equal to " + ColumnCount.ToString() );
            }

            if ( Progress != null )
            {
                if ( this.WriteProgressSize <= 0 )
                {
                    throw new ArgumentException( "The property 'WriteProgressSize' must be greater than 0" );
                }
            }

            List<string> rows = this.SerializeRows( rowData );
            await this.CsvStream.WriteLineAsync( rows[0] );
            this.TotalRowCount++;

            //设置字段数
            if ( this.ColumnCount == 0 )
            {
                this.ColumnCount = rowData.Count;
            }

            //发送通知
            if ( Progress != null )
            {
                //如果取余数=0, 发送通知.
                if ( TotalRowCount % this.WriteProgressSize == 0L )
                {
                    Progress.Report( this.WriteProgressSize );
                }
            }
        }

        /// <summary>
        /// 异步写入单行数据, 可多次执行, 之后执行 Close 方法关闭写入流.
        /// </summary>
        /// <typeparam name="T">要写入的数据对象类型</typeparam>
        /// <param name="rowData">要写入的数据对象实例</param>
        /// <param name="expression">处理对象实例,返回字段集合的方法.</param>
        /// <returns>Task</returns>
        public async Task WriteLineAsync<T>( T rowData, Func<T, IList<string>> expression ) where T : new()
        {
            if ( rowData == null )
            {
                throw new ArgumentNullException( "rowData" );
            }

            if ( expression == null )
            {
                throw new ArgumentNullException( "expression" );
            }

            await WriteLineAsync( expression.Invoke( rowData ) );
        }

        /// <summary>
        /// 异步写入多行数据, 可多次执行, 之后执行 Close 方法关闭写入流.
        /// </summary>
        /// <param name="rowDataList">行数据集合</param>
        /// <returns>Task</returns>
        public async Task WriteAsync( IList<IList<string>> rowDataList )
        {
            if ( rowDataList == null )
            {
                throw new ArgumentNullException( "rowDataList" );
            }

            foreach ( var row in rowDataList )
            {
                await WriteLineAsync( row );
            }
        }

        /// <summary>
        /// 异步写入多行数据, 可多次执行, 之后执行 Close 方法关闭写入流.
        /// </summary>
        /// <typeparam name="T">要写入的数据对象类型</typeparam>
        /// <param name="rowDataList">要写入的数据对象实例集合</param>
        /// <param name="expression">处理对象实例集合,返回包含字段集合的行集合方法.</param>
        /// <returns>Task</returns>
        public async Task WriteAsync<T>( IList<T> rowDataList, Func<T, IList<string>> expression ) where T : new()
        {
            if ( rowDataList == null )
            {
                throw new ArgumentNullException( "rowDataList" );
            }

            foreach ( var row in rowDataList )
            {
                await WriteLineAsync( row, expression );
            }
        }

        /// <summary>
        /// 异步清除缓冲区,将数据写入流.
        /// </summary>
        /// <returns></returns>
        public async Task FlushAsync()
        {
            if ( this.CsvStream != null )
            {
                await this.CsvStream.FlushAsync();
            }
        }

        /// <summary>
        /// 关闭写入流, 并引发可能的最后一次通知事件.
        /// </summary>
        public void Close()
        {
            if ( this.CsvStream != null )
            {
                this.CsvStream.Close();

                if ( Progress != null )
                {
                    //如果记录总数等于通知设定总数, 说明写入结束时刚好是要通知的数量, 但在 WriteLineAsync 方法中已经通知, 所以在这不再通知.
                    if ( TotalRowCount != WriteProgressSize )
                    {
                        Progress.Report( (int)(TotalRowCount % this.WriteProgressSize) );
                    }
                }
            }
        }


        /// <summary>
        /// 转义 CSV 多行内容, 返回转义后的行集合内容.
        /// </summary>
        /// <param name="lines">转义前的行集合</param>
        /// <returns>转义后的行集合</returns>
        private List<string> SerializeRows( params IList<string>[] lines )
        {
            List<string> result = new List<string>();

            if ( lines != null )
            {
                foreach ( var ss in lines )
                {
                    StringBuilder sb = new StringBuilder( 2048 );

                    for ( int i = 0; i < ss.Count; i++ )
                    {
                        sb.Append( SerializeField( ss[i] ) );

                        if ( i + 1 < ss.Count )
                        {
                            sb.Append( this.Flag.FieldSeparator );
                        }
                    }

                    if ( sb.Length > 0 )
                    {
                        result.Add( sb.ToString() );
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 转义字段值。如果字段包含分隔符或 '"' 或 ‘\n’ 或 "\r\n" 或字段分隔符，用双引号将字段包围起来，再將字段中的每个双引号替换为两个双引号。
        /// </summary>
        /// <param name="field">输入字符串</param>
        /// <returns>加上转义符后的字段</returns>
        private string SerializeField( string field )
        {
            string result = field;

            if ( string.IsNullOrEmpty( field ) )
            {
                result = "";
            }
            else
            {
                if ( result.IndexOf( Flag.FieldSeparator ) >= 0 || result.IndexOf( Flag.FieldQualifier ) >= 0 || result.IndexOf( '\r' ) >= 0 || result.IndexOf( "\n" ) >= 0 )
                {
                    result = this.Flag.Qualifier + result.Replace( this.Flag.Qualifier, this.Flag.DoubleQualifier ) + this.Flag.Qualifier;
                }
            }

            return result;
        }
    }
}
