# CsvHelperAsync
A library for async reading and writing csv file.
简单易用的 csv 异步读写类库. 

# Test Form
![实现效果](https://raw.githubusercontent.com/jgh004/CsvHelperAsync/master/Solution%20Items/test.png)

# Install

Run the following command in the Package Manager Console.
在 nuget 包管理器控制台输入以下命令

    PM> Install-Package CsvHelperAsync

# Getting Started
### Csv File IETF Standard
[IETF RFC4180](https://tools.ietf.org/html/rfc4180)
### Reading csv
    public async void ReadCsv(...)
    {
        var csvReader = new CsvReadHelper( fileName or stream, encoding, flag, firstRowIsHead, readStreamBufferLength );
        
        //using delegate to get rows data. 使用委托获取数据
        Progress<CsvReadProgressInfo<T>> progress = new Progress<CsvReadProgressInfo<T>>( e =>
        {
            //Update ui should use SynchronizationContext. 更新 ui 时应使用 SynchronizationContext 相关方法.
            ShowData( e.CurrentRowsData );
            SetProgress( Convert.ToInt32( e.ProgressValue ) );
        } );
        
        //prevent ui thread blocking. 防止 ui 线程阻塞
        await Task.Run( async () =>
        {
            await csvReader.ReadAsync( progress, f =>
            {
                return ConvertCsvRowToCustomerModel( f );
            }, cancelToken, 1000 );
        }, cancelToken );
    }
    
