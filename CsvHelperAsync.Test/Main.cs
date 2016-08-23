using System;
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
        //保存csv数据
        List<List<string>> csvData = null;

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

            sync = SynchronizationContext.Current;
        }

        private void bt_Open_Click( object sender, EventArgs e )
        {
            try
            {
                this.bt_Open.Enabled = false;
                this.bt_Save.Enabled = false;
                this.Cursor = Cursors.WaitCursor;
                this.pb_Progress.Value = 0;

                OpenFileDialog f = new OpenFileDialog();
                f.Filter = "CSV Files|*.csv|TxtFile|*.txt";
                f.InitialDirectory = Environment.GetFolderPath( Environment.SpecialFolder.DesktopDirectory );

                if ( f.ShowDialog() == System.Windows.Forms.DialogResult.OK )
                {
                    this.tb_FileName.Text = f.FileName;

                    ReadData().ContinueWith( k =>
                    {
                        this.bt_Open.Enabled = true;
                        this.bt_Save.Enabled = true;
                        this.Cursor = Cursors.Default;

                        if ( k.Exception != null )
                        {
                            MessageBox.Show( k.Exception.InnerException.Message );
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext() );
                }
                else
                {
                    this.bt_Open.Enabled = true;
                    this.bt_Save.Enabled = true;
                    this.Cursor = Cursors.Default;
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
                this.bt_Open.Enabled = false;
                this.bt_Save.Enabled = false;
                this.Cursor = Cursors.WaitCursor;
                this.pb_Progress.Value = 0;
                SaveFileDialog f = new SaveFileDialog();
                f.Filter = "CSV File|*.csv";
                f.FileName = string.IsNullOrWhiteSpace( this.tb_FileName.Text ) ? "test" : Path.GetFileNameWithoutExtension( this.tb_FileName.Text.Trim() ) + "-after.csv";
                f.InitialDirectory = Environment.GetFolderPath( Environment.SpecialFolder.Desktop );

                if ( f.ShowDialog( this ) == System.Windows.Forms.DialogResult.OK )
                {
                    WriteData( f.FileName ).ContinueWith( k =>
                    {
                        this.bt_Open.Enabled = true;
                        this.bt_Save.Enabled = true;
                        this.Cursor = Cursors.Default;

                        if ( k.Exception != null )
                        {
                            MessageBox.Show( k.Exception.InnerException.Message );
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext() );
                }
                else
                {
                    this.bt_Open.Enabled = true;
                    this.bt_Save.Enabled = true;
                    this.Cursor = Cursors.Default;
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

        //虚拟模式绑定数据
        private void dgv_Data_CellValueNeeded( object sender, DataGridViewCellValueEventArgs e )
        {
            if ( csvData != null && e.RowIndex < csvData.Count && e.ColumnIndex > -1 && e.RowIndex > -1 )
            {
                e.Value = csvData[e.RowIndex][e.ColumnIndex];
            }
        }


        private async Task ReadData()
        {
            if ( !string.IsNullOrWhiteSpace( this.tb_FileName.Text ) )
            {
                Stopwatch sc = Stopwatch.StartNew();

                try
                {
                    this.csvData = null;
                    CsvFlag flag = new CsvFlag( Convert.ToChar( this.cob_separator.Text ), Convert.ToChar( this.cob_FieldEnclosed.Text ) );
                    CsvReadHelper csv = new CsvReadHelper( this.tb_FileName.Text, Encoding.UTF8, flag, !Convert.ToBoolean( this.cob_FirstIsHead.SelectedIndex ) );

                    Progress<CsvReadProgressInfo<List<string>>> prog = new Progress<CsvReadProgressInfo<List<string>>>( e =>
                    {
                        this.sync.Post( f =>
                        {
                            var eve = f as CsvReadProgressInfo<List<string>>;

                            if ( eve.RowsData != null )
                            {
                                if ( this.csvData == null )
                                {
                                    //初始化之前确保 this.Data 是 null, 否则出错.
                                    InitDataGridView( eve.ColumnNames );
                                    this.csvData = eve.RowsData;
                                }
                                else
                                {
                                    this.csvData.AddRange( eve.RowsData );
                                }

                                this.RefreshDataGridView( this.csvData );
                            }

                            this.pb_Progress.Value = Convert.ToInt32( eve.ProgressValue );
                        }, e );
                    } );

                    CancellationTokenSource source = new CancellationTokenSource();
                    //取消测试
                    //source.CancelAfter( 300 );
                    await csv.ReadAsync( prog, f =>
                    {
                        return f;
                    }, CancellationToken.None, 1000 );

                    csv.Close();
                }
                catch ( Exception ex )
                {
                    sc.Stop();
                    throw ex;
                }
                finally
                {
                    sc.Stop();

                    this.sync.Post( k =>
                    {
                        this.label1.Text = "读取用时: " + k.ToString() + " 秒";
                    }, sc.Elapsed.TotalSeconds.ToString() );
                }
            }
        }

        private async Task WriteData( string fileName )
        {
            if ( this.csvData != null )
            {
                Stopwatch sc = null;
                CsvWriteHelper csv = null;

                try
                {
                    CsvFlag flag = new CsvFlag( Convert.ToChar( this.cob_separator.Text ), Convert.ToChar( this.cob_FieldEnclosed.Text ) );

                    //将数据分两批,为测试多次写入
                    int index = this.csvData.Count / 2;
                    List<List<string>> data1 = new List<List<string>>();

                    for ( int i = 0; i < index; i++ )
                    {
                        data1.Add( this.csvData[i] );
                    }

                    List<List<string>> data2 = new List<List<string>>();

                    for ( int i = index; i < this.csvData.Count; i++ )
                    {
                        data2.Add( this.csvData[i] );
                    }

                    Progress<int> p = new Progress<int>( r =>
                    {
                        double val = csv.TotalRowCount / (double)(this.csvData.Count + (this.cob_FirstIsHead.SelectedIndex == 0 ? 1 : 0));

                        this.sync.Post( t =>
                        {
                            this.pb_Progress.Value = (int)t;
                        }, Convert.ToInt32( val * 100 ) );
                    } );

                    CancellationTokenSource cs = new CancellationTokenSource();
                    csv = new CsvWriteHelper( fileName, Encoding.UTF8, flag, cs.Token, p, 1000 );
                    List<string> columnNames = null;
                    //标题行
                    if ( !Convert.ToBoolean( this.cob_FirstIsHead.SelectedIndex ) )
                    {
                        columnNames = new List<string>();

                        foreach ( DataGridViewColumn c in this.dgv_Data.Columns )
                        {
                            columnNames.Add( c.Name );
                        }
                    }

                    sc = Stopwatch.StartNew();
                    //测试取消
                    //cs.CancelAfter( 300 );

                    if ( columnNames != null )
                    {
                        await csv.WriteLineAsync( columnNames );
                    }

                    await csv.WriteAsync( data1, f =>
                    {
                        return f;
                    } );

                    await csv.WriteAsync( data2, f =>
                    {
                        return f;
                    } );

                    await csv.FlushAsync();
                }
                catch ( Exception ex )
                {
                    sc.Stop();
                    throw ex;
                }
                finally
                {
                    csv.Close();
                    sc.Stop();

                    this.sync.Post( k =>
                    {
                        this.label1.Text = "写入用时: " + k.ToString() + " 秒";
                    }, sc.Elapsed.TotalSeconds.ToString() );
                }
            }
        }

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

        private void RefreshDataGridView( List<List<string>> data )
        {
            this.dgv_Data.RowCount = data == null ? 0 : data.Count;

            if ( this.dgv_Data.RowCount > 0 )
            {
                this.dgv_Data.FirstDisplayedScrollingRowIndex = this.dgv_Data.RowCount - 1;
            }
        }
    }
}
