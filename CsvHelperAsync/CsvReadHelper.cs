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
}
