﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CsvHelperAsync.Test
{
    public partial class Main : Form
    {
        SynchronizationContext sync = null;
        CancellationTokenSource cancelSource = new CancellationTokenSource();
        //保存csv数据
        List<string> columnsName = null;
        List<TestProductModel> modelCsvData = null;

        public Main()
        {
            InitializeComponent();


            var properInfo = this.dgv_Data.GetType().GetProperty( "DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic );
            properInfo.SetValue( this.dgv_Data, true, null );

            this.dgv_Data.RowPostPaint += ExtLibrary_DataGridView_RowPostPaint;

            this.dgv_Data.VirtualMode = true;
            this.dgv_Data.AutoGenerateColumns = false;
            this.dgv_Data.CellValueNeeded += dgv_Data_CellValueNeeded;
        }


        private void Main_Load( object sender, EventArgs e )
        {
            this.cob_separator.SelectedIndex = 0;
            this.cob_FieldEnclosed.SelectedIndex = 0;
            this.cob_FirstIsHead.SelectedIndex = 0;
            this.sync = SynchronizationContext.Current;
            this.cancelSource = new CancellationTokenSource();
        }

        private void bt_Open_Click( object sender, EventArgs e )
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                this.bt_Open.Enabled = false;
                this.bt_Save.Enabled = false;
                this.bt_GenerateTestData.Enabled = false;
                this.pb_Progress.Value = 0;

                OpenFileDialog f = new OpenFileDialog();
                f.Filter = "CSV Files|*.csv|TxtFile|*.txt";
                f.InitialDirectory = Environment.GetFolderPath( Environment.SpecialFolder.DesktopDirectory );

                if ( f.ShowDialog() == System.Windows.Forms.DialogResult.OK )
                {
                    this.tb_FileName.Text = f.FileName;

                    ReadData().ContinueWith( k =>
                    {
                        this.Cursor = Cursors.Default;
                        this.bt_Open.Enabled = true;
                        this.bt_Save.Enabled = true;
                        this.bt_GenerateTestData.Enabled = true;

                        if ( k.IsCanceled )
                        {
                            MessageBox.Show( "Read operation has been canceled." );
                        }

                        if ( k.Exception != null )
                        {
                            MessageBox.Show( k.Exception.InnerException.Message );
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext() );
                }
                else
                {
                    this.Cursor = Cursors.Default;
                    this.bt_Open.Enabled = true;
                    this.bt_Save.Enabled = true;
                    this.bt_GenerateTestData.Enabled = true;
                }
            }
            catch ( Exception ex )
            {
                MessageBox.Show( ex.Message );
            }
            finally
            {
            }
        }

        private void bt_Save_Click( object sender, EventArgs e )
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                this.bt_Open.Enabled = false;
                this.bt_Save.Enabled = false;
                this.bt_GenerateTestData.Enabled = false;
                this.pb_Progress.Value = 0;
                SaveFileDialog f = new SaveFileDialog();
                f.Filter = "CSV File|*.csv";
                f.FileName = string.IsNullOrWhiteSpace( this.tb_FileName.Text ) ? "test" : Path.GetFileNameWithoutExtension( this.tb_FileName.Text.Trim() ) + "-after.csv";
                f.InitialDirectory = Environment.GetFolderPath( Environment.SpecialFolder.Desktop );

                if ( f.ShowDialog( this ) == System.Windows.Forms.DialogResult.OK )
                {
                    WriteData( f.FileName ).ContinueWith( k =>
                    {
                        this.Cursor = Cursors.Default;
                        this.bt_Open.Enabled = true;
                        this.bt_Save.Enabled = true;
                        this.bt_GenerateTestData.Enabled = true;

                        if ( k.IsCanceled )
                        {
                            MessageBox.Show( "Write operation has been canceled." );
                        }

                        if ( k.Exception != null )
                        {
                            MessageBox.Show( k.Exception.InnerException.Message );
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext() );
                }
                else
                {
                    this.Cursor = Cursors.Default;
                    this.bt_Open.Enabled = true;
                    this.bt_Save.Enabled = true;
                    this.bt_GenerateTestData.Enabled = true;
                }
            }
            catch ( Exception ex )
            {
                MessageBox.Show( ex.Message );
            }
            finally
            {
            }
        }

        private void bt_Cancel_Click( object sender, EventArgs e )
        {
            if ( this.cancelSource != null && !this.cancelSource.IsCancellationRequested )
            {
                this.cancelSource.Cancel();
                this.cancelSource = new CancellationTokenSource();
            }
        }

        private void bt_GenerateTestData_Click( object sender, EventArgs e )
        {
            this.GenerateTestData();
            this.InitDataGridView( this.columnsName );
            this.RefreshDataGridView( this.modelCsvData.Count );
        }

        /// <summary>
        /// 显示行号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExtLibrary_DataGridView_RowPostPaint( object sender, DataGridViewRowPostPaintEventArgs e )
        {
            DataGridView dgv = sender as DataGridView;
            Rectangle rectangle = new Rectangle( e.RowBounds.Location.X
                , e.RowBounds.Location.Y
                , dgv.RowHeadersWidth - 4
                , e.RowBounds.Height );

            TextRenderer.DrawText( e.Graphics, (e.RowIndex + 1).ToString(),
                dgv.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dgv.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right );
        }

        /// <summary>
        /// 虚拟模式绑定数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_Data_CellValueNeeded( object sender, DataGridViewCellValueEventArgs e )
        {
            if ( modelCsvData != null && e.RowIndex < modelCsvData.Count && e.ColumnIndex > -1 && e.RowIndex > -1 )
            {
                List<string> row = ConvertModelToRowData( modelCsvData[e.RowIndex] );
                e.Value = row[e.ColumnIndex];
            }
        }


        /// <summary>
        /// 异步读取csv
        /// </summary>
        /// <returns></returns>
        private async Task ReadData()
        {
            if ( !string.IsNullOrWhiteSpace( this.tb_FileName.Text ) )
            {
                Stopwatch sc = null;
                CsvReadHelper csv = null;
                
                try
                {
                    this.modelCsvData = null;
                    CsvFlag flag = new CsvFlag( Convert.ToChar( this.cob_separator.Text ), Convert.ToChar( this.cob_FieldEnclosed.Text ) );
                    csv = new CsvReadHelper( this.tb_FileName.Text, Encoding.UTF8, flag, !Convert.ToBoolean( this.cob_FirstIsHead.SelectedIndex ), 40960 );

                    Progress<CsvReadProgressInfo<TestProductModel>> prog = new Progress<CsvReadProgressInfo<TestProductModel>>( e =>
                    {
                        this.sync.Post( f =>
                        {
                            var eve = f as CsvReadProgressInfo<TestProductModel>;

                            if ( eve.CurrentRowsData != null )
                            {
                                if ( this.modelCsvData == null )
                                {
                                    InitDataGridView( eve.ColumnNames );
                                    this.modelCsvData = eve.CurrentRowsData;
                                }
                                else
                                {
                                    this.modelCsvData.AddRange( eve.CurrentRowsData );
                                }

                                this.RefreshDataGridView( this.modelCsvData.Count );
                            }

                            this.pb_Progress.Value = Convert.ToInt32( eve.ProgressValue );
                        }, e );
                    } );

                    sc = Stopwatch.StartNew();
                    //因为 ui 线程同步执行 ReadAsync 中的部分代码, 所以用 Task.Run 在其它线程中执行, 避免 ui 阻塞.
                    await Task.Run( async () =>
                    {
                        await csv.ReadAsync( prog, f =>
                        {
                            return ConvertCsvRowToTestProductData( f );
                        }, this.cancelSource.Token, 1000 );
                    }, this.cancelSource.Token );
                }
                finally
                {
                    if ( csv != null )
                    {
                        csv.Close();
                    }

                    if ( sc != null )
                    {
                        sc.Stop();
                    }

                    this.sync.Post( k =>
                    {
                        this.tb_Times.Text = k.ToString();
                    }, sc.Elapsed.TotalSeconds.ToString() );
                }
            }
        }

        /// <summary>
        /// 异步写入csv
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private async Task WriteData( string fileName )
        {
            if ( this.modelCsvData != null )
            {
                Stopwatch sc = null;
                CsvWriteHelper csv = null;

                try
                {
                    CsvFlag flag = new CsvFlag( Convert.ToChar( this.cob_separator.Text ), Convert.ToChar( this.cob_FieldEnclosed.Text ) );
                    Progress<CsvWriteProgressInfo> p = new Progress<CsvWriteProgressInfo>( r =>
                    {
                        this.sync.Post( t =>
                        {
                            double val = (t as CsvWriteProgressInfo).WirteRowCount / (double)(this.modelCsvData.Count + (this.cob_FirstIsHead.SelectedIndex == 0 ? 1 : 0));
                            this.pb_Progress.Value = Convert.ToInt32( val * 100 );
                        }, r );
                    } );

                    csv = new CsvWriteHelper( fileName, Encoding.UTF8, flag, cancelSource.Token, p, 1000 );

                    sc = Stopwatch.StartNew();

                    //因为 ui 线程同步执行 WriteLineAsync 中的部分代码, 所以用 Task.Run 在其它线程中执行, 避免 ui 阻塞.
                    await Task.Run( async () =>
                    {
                        if ( columnsName != null )
                        {
                            await csv.WriteLineAsync( columnsName );
                        }

                        await csv.WriteAsync( this.modelCsvData, f =>
                        {
                            return ConvertModelToRowData( f );
                        } );

                        await csv.FlushAsync();
                    }, this.cancelSource.Token );
                }
                finally
                {
                    if ( csv != null )
                    {
                        csv.Close();
                    }

                    if ( sc != null )
                    {
                        sc.Stop();
                    }

                    this.sync.Post( k =>
                    {
                        this.tb_Times.Text = k.ToString();
                    }, sc.Elapsed.TotalSeconds.ToString() );
                }
            }
        }

        /// <summary>
        /// 生成测试csv数据
        /// </summary>
        private void GenerateTestData()
        {
            this.columnsName = new List<string>();
            this.columnsName.Add( "ID" );
            this.columnsName.Add( "Name" );
            this.columnsName.Add( "Description" );
            this.columnsName.Add( "Size" );
            this.columnsName.Add( "Price" );
            this.columnsName.Add( "Html" );
            this.columnsName.Add( "Url" );
            this.columnsName.Add( "CreateDate" );

            this.modelCsvData = new List<TestProductModel>();

            for ( int i = 0; i < 100000; i++ )
            {
                TestProductModel model = new Test.TestProductModel();
                model.ID = i + 1;
                model.Name = "name " + (i + 1).ToString();
                model.Price = i / 3m;
                model.Url = "http://www.test.com/product/" + i.ToString();
                model.Html = @"<div class=""description-div""><table class=""description-table""><tr>
            <td colspan=""2"" class=""description-td-title"">Details:</td>
        </tr>
        <tr>
            <td  style=""color:#ffff"">Color Type:</td>
            <td  class=""description-right-td"">Black/Brown/Red</td>
        </tr>
        <tr>
            <td  class=""description-left-td"">MATERIAL:</td>
            <td  class=""description-right-td"">Vinyl</td>
        </tr>
        <tr>
            <td  class=""description-left-td"">Fabric:</td>
            <td  class=""description-right-td"">Vinyl</td>
        </tr>
        <tr>
            <td  class=""description-left-td"">Height:</td>
            <td  class=""description-right-td"">-</td>
        </tr>
        <tr>
            <td  class=""description-left-td"">Length:</td>
            <td  class=""description-right-td"">5</td>
        </tr>
        <tr>
            <td  class=""description-left-td"">Width:</td>
            <td  class=""description-right-td"">3</td>
        </tr>
        <tr>
            <td  class=""description-left-td"">Weight:</td>
            <td  class=""description-right-td"">3.40g</td>
        </tr>
    
</table></div>";
                model.CreateDate = DateTime.Now;

                modelCsvData.Add( model );
            }
        }

        /// <summary>
        /// 将模型转为一行csv数据
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private List<string> ConvertModelToRowData( TestProductModel model )
        {
            List<string> result = new List<string>();

            if ( model != null )
            {
                result.Add( model.ID.ToString() );
                result.Add( model.Name );
                result.Add( model.Description );
                result.Add( model.Size );
                result.Add( model.Price.ToString() );
                result.Add( model.Html );
                result.Add( model.Url );
                result.Add( model.CreateDate.GetDateTimeFormats('r')[0].ToString() );
            }

            return result;
        }

        /// <summary>
        /// 将一行csv数据转为模型实例
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private TestProductModel ConvertCsvRowToTestProductData(List<string> data)
        {
            TestProductModel result = new TestProductModel();

            if ( data != null )
            {
                for ( int i = 0; i < data.Count; i++ )
                {
                    switch ( i )
                    {
                        case 0:
                            result.ID = Convert.ToInt32( data[i] );
                            break;
                        case 1:
                            result.Name = data[i];
                            break;
                        case 2:
                            result.Description = data[i];
                            break;
                        case 3:
                            result.Size = data[i];
                            break;
                        case 4:
                            result.Price = Convert.ToDecimal( data[i] );
                            break;
                        case 5:
                            result.Html = data[i];
                            break;
                        case 6:
                            result.Url = data[i];
                            break;
                        case 7:
                            result.CreateDate = DateTime.Parse( data[i] );
                            break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 初始化表格
        /// </summary>
        /// <param name="columns"></param>
        private void InitDataGridView( List<string> columns )
        {
            this.dgv_Data.Rows.Clear();
            this.dgv_Data.Columns.Clear();

            if ( columns != null )
            {
                foreach ( var c in columns )
                {
                    var column = new DataGridViewTextBoxColumn()
                    {
                        Name = c,
                        HeaderText = c,
                        DataPropertyName = c
                    };

                    this.dgv_Data.Columns.Add( column );
                }
            }
        }

        /// <summary>
        /// 刷新表格显示
        /// </summary>
        /// <param name="rowCount"></param>
        private void RefreshDataGridView( int rowCount )
        {
            this.dgv_Data.RowCount = rowCount;

            if ( this.dgv_Data.RowCount > 0 )
            {
                this.dgv_Data.FirstDisplayedScrollingRowIndex = this.dgv_Data.RowCount - 1;
            }
        }
    }
}
