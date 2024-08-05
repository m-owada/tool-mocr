using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion ("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]
[assembly: AssemblyTitle("")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyProduct("MOCR")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyCopyright("Copyright (C) 2024 m-owada.")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += (sender, e) =>
            {
                throw new Exception(e.Exception.Message);
            };
            Application.Run(new MainForm());
        }
        catch(Exception e)
        {
            MessageBox.Show(e.Message, e.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
        }
    }
}

class MainForm : Form
{
    private PictureBox picImage = new PictureBox();
    private TextBox txtMemory = new TextBox();
    private TextBox txtResult = new TextBox();
    private CheckBox chkEnlarge = new CheckBox();
    private CheckBox chkGrayscale = new CheckBox();
    private CheckBox chk1bppImage = new CheckBox();
    private CheckBox chkNoiseOff = new CheckBox();
    private CheckBox chkEdge = new CheckBox();
    private CheckBox chkNegative = new CheckBox();
    private Button btnApply = new Button();
    private Button btnGc = new Button();
    private Button btnFile = new Button();
    private Button btnClip = new Button();
    private Button btnRemove = new Button();
    private Button btnCopy = new Button();
    private CheckBox chkTop = new CheckBox();
    private Timer timer = new Timer();
    private string initialDirectory = Environment.CurrentDirectory;
    private Image importImage = null;
    
    public MainForm()
    {
        // フォーム
        this.Text = "MOCR";
        this.Size = new Size(435, 400);
        this.MinimumSize = this.Size;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
        
        // パネル
        var panel = new Panel();
        panel.Location = new Point(10, 10);
        panel.Size = new Size(260, 240);
        panel.BorderStyle = BorderStyle.FixedSingle;
        panel.AutoScroll = true;
        panel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(panel);
        
        // ピクチャーボックス
        picImage.Location = new Point(0, 0);
        picImage.SizeMode = PictureBoxSizeMode.AutoSize;
        panel.Controls.Add(picImage);
        
        // グループボックス
        var groupBox = new GroupBox();
        groupBox.Location = new Point(280, 10);
        groupBox.Size = new Size(130, 210);
        groupBox.Text = "オプション";
        groupBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(groupBox);
        
        // 画像拡大
        chkEnlarge.Location = new Point(10, 20);
        chkEnlarge.AutoSize = true;
        chkEnlarge.Text = "画像拡大";
        chkEnlarge.Checked = false;
        groupBox.Controls.Add(chkEnlarge);
        
        // グレースケール
        chkGrayscale.Location = new Point(10, 45);
        chkGrayscale.AutoSize = true;
        chkGrayscale.Text = "グレースケール";
        chkGrayscale.Checked = false;
        groupBox.Controls.Add(chkGrayscale);
        
        // 2値化
        chk1bppImage.Location = new Point(10, 70);
        chk1bppImage.AutoSize = true;
        chk1bppImage.Text = "2値化";
        chk1bppImage.Checked = false;
        groupBox.Controls.Add(chk1bppImage);
        
        // ノイズ除去
        chkNoiseOff.Location = new Point(10, 95);
        chkNoiseOff.AutoSize = true;
        chkNoiseOff.Text = "ノイズ除去";
        chkNoiseOff.Checked = false;
        groupBox.Controls.Add(chkNoiseOff);
        
        // エッジ検出
        chkEdge.Location = new Point(10, 120);
        chkEdge.AutoSize = true;
        chkEdge.Text = "エッジ検出";
        chkEdge.Checked = false;
        groupBox.Controls.Add(chkEdge);
        
        // 色反転
        chkNegative.Location = new Point(10, 145);
        chkNegative.AutoSize = true;
        chkNegative.Text = "色反転";
        chkNegative.Checked = false;
        groupBox.Controls.Add(chkNegative);
        
        // 適用
        btnApply.Location = new Point(10, 170);
        btnApply.Size = new Size(48, 24);
        btnApply.Text = "適用";
        btnApply.Click += Click_btnApply;
        groupBox.Controls.Add(btnApply);
        
        // ラベル
        var label = new Label();
        label.Text = "メモリ";
        label.Location = new Point(280, 230);
        label.Size = new Size(label.PreferredWidth, 20);
        label.TextAlign = ContentAlignment.MiddleLeft;
        label.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(label);
        
        // メモリ
        txtMemory.Location = new Point(310, 230);
        txtMemory.Size = new Size(55, 20);
        txtMemory.Text = string.Empty;
        txtMemory.ReadOnly = true;
        txtMemory.TabStop = false;
        txtMemory.TextAlign = HorizontalAlignment.Right;
        txtMemory.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(txtMemory);
        
        // GC
        btnGc.Location = new Point(370, 230);
        btnGc.Size = new Size(40, 20);
        btnGc.Text = "GC";
        btnGc.Click += (sender, e) => GC.Collect();
        btnGc.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        this.Controls.Add(btnGc);
        
        // テキストボックス
        txtResult.Location = new Point(10, 260);
        txtResult.Size = new Size(400, 54);
        txtResult.Text = string.Empty;
        txtResult.ReadOnly = true;
        txtResult.TabStop = false;
        txtResult.Multiline = true;
        txtResult.ScrollBars = ScrollBars.Vertical;
        txtResult.WordWrap = true;
        txtResult.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        this.Controls.Add(txtResult);
        
        // ファイル選択
        btnFile.Location = new Point(10, 325);
        btnFile.Size = new Size(80, 24);
        btnFile.Text = "ファイル選択";
        btnFile.Click += Click_btnFile;
        btnFile.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(btnFile);
        
        // クリップボード
        btnClip.Location = new Point(95, 325);
        btnClip.Size = new Size(80, 24);
        btnClip.Text = "クリップボード";
        btnClip.Click += Click_btnClip;
        btnClip.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(btnClip);
        
        // 最前面
        chkTop.Location = new Point(185, 325);
        chkTop.Size = new Size(60, 24);
        chkTop.Text = "最前面";
        chkTop.CheckedChanged += (sender, e) => this.TopMost = chkTop.Checked;
        chkTop.Checked = false;
        chkTop.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(chkTop);
        
        // 空白除去
        btnRemove.Location = new Point(245, 325);
        btnRemove.Size = new Size(80, 24);
        btnRemove.Text = "空白除去";
        btnRemove.Click += Click_btnRemove;
        btnRemove.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(btnRemove);
        
        // 文字列コピー
        btnCopy.Location = new Point(330, 325);
        btnCopy.Size = new Size(80, 24);
        btnCopy.Text = "文字列コピー";
        btnCopy.Click += Click_btnCopy;
        btnCopy.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        this.Controls.Add(btnCopy);
        
        // タイマー
        timer.Interval = 100;
        timer.Enabled = true;
        timer.Tick += (sender, e) => txtMemory.Text = (Environment.WorkingSet / 1024 / 1024).ToString("N0") + "MB";
        
        btnFile.Select();
    }
    
    private void Click_btnApply(object sender, EventArgs e)
    {
        if(importImage != null)
        {
            ExecuteOcr(importImage);
            btnApply.Select();
        }
    }
    
    private void Click_btnFile(object sender, EventArgs e)
    {
        using(var dialog = new OpenFileDialog())
        {
            dialog.Title = "画像ファイルを選択してください。";
            dialog.Filter = "すべてのファイル (*.*)|*.*|画像ファイル (*.bmp;*.jpg;*.jpeg;*.gif;*.tif;*.tiff;*.png;*.ico)|*.bmp;*.jpg;*.jpeg;*.gif;*.tif;*.tiff;*.png;*.ico";
            dialog.FilterIndex = 2;
            if(Directory.Exists(initialDirectory))
            {
                dialog.InitialDirectory = initialDirectory;
            }
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                initialDirectory = Path.GetDirectoryName(dialog.FileName);
                try
                {
                    importImage = Image.FromFile(dialog.FileName);
                    ExecuteOcr(importImage);
                    btnFile.Select();
                }
                catch
                {
                    MessageBox.Show("選択されたファイルの種類はサポートされていません。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
    }
    
    private void Click_btnClip(object sender, EventArgs e)
    {
        var image = Clipboard.GetImage();
        if(image != null)
        {
            importImage = image;
            ExecuteOcr(importImage);
            btnClip.Select();
        }
        else
        {
            MessageBox.Show("画像をクリップボードにコピーしてください。", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
    
    private void Click_btnRemove(object sender, EventArgs e)
    {
        if(!String.IsNullOrEmpty(txtResult.Text))
        {
            txtResult.Text = txtResult.Text.Replace(" ", string.Empty);
        }
    }
    
    private void Click_btnCopy(object sender, EventArgs e)
    {
        if(!String.IsNullOrEmpty(txtResult.Text))
        {
            Clipboard.SetText(txtResult.Text);
        }
    }
    
    private async void ExecuteOcr(Image image)
    {
        picImage.Image = null;
        txtResult.Text = string.Empty;
        SetControlsEnabled(false);
        
        await Task.Run(() =>
        {
            if(chkEnlarge.Checked)
            {
                image = ImageEnlarger(image);
            }
            if(chkGrayscale.Checked)
            {
                image = ConvertToGrayscale(image);
            }
            if(chk1bppImage.Checked)
            {
                image = ConvertTo1bppImage(image);
            }
            if(chkNoiseOff.Checked)
            {
                image = MedianFilter(image);
            }
            if(chkEdge.Checked)
            {
                image = LaplacianFilter(image);
            }
            if(chkNegative.Checked)
            {
                image = ConvertToNegativeImage(image);
            }
        });
        
        picImage.Image = image;
        var imageBytes = ConvertToByteArray(image);
        var bitmap = await ConvertToSoftwareBitmap(imageBytes);
        var result = await Detect(bitmap);
        txtResult.Text = result.Text;
        SetControlsEnabled(true);
    }
    
    private void SetControlsEnabled(bool enabled)
    {
        chkEnlarge.Enabled = enabled;
        chkGrayscale.Enabled = enabled;
        chk1bppImage.Enabled = enabled;
        chkNoiseOff.Enabled = enabled;
        chkEdge.Enabled = enabled;
        chkNegative.Enabled = enabled;
        btnApply.Enabled = enabled;
        btnGc.Enabled = enabled;
        txtResult.Enabled = enabled;
        btnFile.Enabled = enabled;
        btnClip.Enabled = enabled;
        btnRemove.Enabled = enabled;
        btnCopy.Enabled = enabled;
    }
    
    private byte[] ConvertToByteArray(Image image)
    {
        var imageConverter = new ImageConverter();
        return (byte[])imageConverter.ConvertTo(image, typeof(byte[]));
    }
    
    private async Task<SoftwareBitmap> ConvertToSoftwareBitmap(byte[] image)
    {
        using(var memoryStream = new MemoryStream(image))
        {
            using(var randomAccessStream = new InMemoryRandomAccessStream())
            {
                var outputStream = randomAccessStream.GetOutputStreamAt(0);
                var dataWriter = new DataWriter(outputStream);
                await Task.Run(() => dataWriter.WriteBytes(memoryStream.ToArray()));
                await dataWriter.StoreAsync();
                await outputStream.FlushAsync();
                var decoder = await BitmapDecoder.CreateAsync(randomAccessStream);
                var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
                return bitmap;
            }
        }
    }
    
    private async Task<OcrResult> Detect(SoftwareBitmap bitmap)
    {
        var ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
        var ocrResult = await ocrEngine.RecognizeAsync(bitmap);
        return ocrResult;
    }
    
    private Image ImageEnlarger(Image image)
    {
        var bitmap = new Bitmap(image.Width * 2, image.Height * 2);
        using(var graphics = Graphics.FromImage(bitmap))
        {
            graphics.InterpolationMode = InterpolationMode.Bilinear;
            graphics.DrawImage(image, 0, 0, bitmap.Width, bitmap.Height);
        }
        return bitmap;
    }
    
    private Image ConvertToGrayscale(Image image)
    {
        var bitmap = new Bitmap(image.Width, image.Height);
        using(var graphics = Graphics.FromImage(bitmap))
        {
            var colorMatrix = new ColorMatrix
            (
                new float[][]
                {
                    new float[]{0.299f, 0.299f, 0.299f, 0, 0},
                    new float[]{0.587f, 0.587f, 0.587f, 0, 0},
                    new float[]{0.114f, 0.114f, 0.114f, 0, 0},
                    new float[]{0, 0, 0, 1, 0},
                    new float[]{0, 0, 0, 0, 1}
                }
            );
            var imageAttributes = new ImageAttributes();
            imageAttributes.SetColorMatrix(colorMatrix);
            graphics.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, imageAttributes);
        }
        return bitmap;
    }
    
    private Image ConvertTo1bppImage(Image image)
    {
        var bitmap = new Bitmap(image);
        var pixelSize = (int)(Image.GetPixelFormatSize(bitmap.PixelFormat) / 8);
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, bitmap.PixelFormat);
        var pixels = new byte[bitmapData.Stride * bitmapData.Height];
        Marshal.Copy(bitmapData.Scan0, pixels, 0, pixels.Length);
        for(var x = 0; x < bitmapData.Width; x++)
        {
            for(var y = 0; y < bitmapData.Height; y++)
            {
                var pos = pixelSize * x + bitmapData.Stride * y;
                var brightness = (int)(pixels[pos + 2] * 0.299f + pixels[pos + 1] * 0.587f + pixels[pos] * 0.114f);
                if(brightness >= 128)
                {
                    brightness = 255;
                }
                else
                {
                    brightness = 0;
                }
                pixels[pos + 2] = (byte)brightness;
                pixels[pos + 1] = (byte)brightness;
                pixels[pos] = (byte)brightness;
            }
        }
        Marshal.Copy(pixels, 0, bitmapData.Scan0, pixels.Length);
        bitmap.UnlockBits(bitmapData);
        return bitmap;
    }
    
    private Image MedianFilter(Image image)
    {
        var bitmap = new Bitmap(image);
        var pixelSize = (int)(Image.GetPixelFormatSize(bitmap.PixelFormat) / 8);
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, bitmap.PixelFormat);
        var pixels = new byte[bitmapData.Stride * bitmapData.Height];
        Marshal.Copy(bitmapData.Scan0, pixels, 0, pixels.Length);
        var writer = new byte[bitmapData.Stride * bitmapData.Height];
        Marshal.Copy(bitmapData.Scan0, writer, 0, writer.Length);
        const int maskSize = 3;
        var mask = new byte[maskSize * maskSize];
        for(var x = 0; x < bitmapData.Width; x++)
        {
            for(var y = 0; y < bitmapData.Height; y++)
            {
                for(var c = 0; c < 3; c++)
                {
                    var len = 0;
                    for(var mx = -maskSize / 2; mx <= maskSize / 2; mx++)
                    {
                        for(var my = -maskSize / 2; my <= maskSize / 2; my++)
                        {
                            if(x + mx >= 0 && x + mx < bitmapData.Width && y + my >= 0 && y + my < bitmapData.Height)
                            {
                                mask[len] = pixels[(pixelSize * (x + mx) + bitmapData.Stride * (y + my)) + c];
                                len++;
                            }
                        }
                    }
                    Array.Sort(mask);
                    Array.Reverse(mask);
                    writer[(pixelSize * x + bitmapData.Stride * y) + c] = mask[len / 2];
                }
            }
        }
        Marshal.Copy(writer, 0, bitmapData.Scan0, writer.Length);
        bitmap.UnlockBits(bitmapData);
        return bitmap;
    }
    
    private Image LaplacianFilter(Image image)
    {
        var bitmap = new Bitmap(image);
        var pixelSize = (int)(Image.GetPixelFormatSize(bitmap.PixelFormat) / 8);
        var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, bitmap.PixelFormat);
        var pixels = new byte[bitmapData.Stride * bitmapData.Height];
        Marshal.Copy(bitmapData.Scan0, pixels, 0, pixels.Length);
        var writer = new byte[bitmapData.Stride * bitmapData.Height];
        Marshal.Copy(bitmapData.Scan0, writer, 0, writer.Length);
        const int maskSize = 3;
        var mask = new double[maskSize, maskSize] { {0, 1, 0}, {1, -4, 1}, {0, 1, 0} };
        for(var x = 0; x < bitmapData.Width; x++)
        {
            for(var y = 0; y < bitmapData.Height; y++)
            {
                for(var c = 0; c < 3; c++)
                {
                    var sum = 0d;
                    for(var mx = -maskSize / 2; mx <= maskSize / 2; mx++)
                    {
                        for(var my = -maskSize / 2; my <= maskSize / 2; my++)
                        {
                            if(x + mx >= 0 && x + mx < bitmapData.Width && y + my >= 0 && y + my < bitmapData.Height)
                            {
                                sum += pixels[(pixelSize * (x + mx) + bitmapData.Stride * (y + my)) + c] * mask[mx + maskSize / 2, my + maskSize / 2];
                            }
                        }
                    }
                    if(sum > 255.0)
                    {
                        writer[(pixelSize * x + bitmapData.Stride * y) + c] = 255;
                    }
                    else if(sum < 0)
                    {
                        writer[(pixelSize * x + bitmapData.Stride * y) + c] = 0;
                    }
                    else
                    {
                        writer[(pixelSize * x + bitmapData.Stride * y) + c] = (byte)sum;
                    }
                }
            }
        }
        Marshal.Copy(writer, 0, bitmapData.Scan0, writer.Length);
        bitmap.UnlockBits(bitmapData);
        return bitmap;
    }
    
    private Image ConvertToNegativeImage(Image image)
    {
        var bitmap = new Bitmap(image.Width, image.Height);
        using(var graphics = Graphics.FromImage(bitmap))
        {
            var colorMatrix = new ColorMatrix
            {
                Matrix00 = -1,
                Matrix11 = -1,
                Matrix22 = -1,
                Matrix33 = 1,
                Matrix40 = 1,
                Matrix41 = 1,
                Matrix42 = 1,
                Matrix44 = 1
            };
            var imageAttributes = new ImageAttributes();
            imageAttributes.SetColorMatrix(colorMatrix);
            graphics.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, imageAttributes);
        }
        return bitmap;
    }
}

static class TaskEx
{
    public static Task<T> AsTask<T>(this Windows.Foundation.IAsyncOperation<T> operation)
    {
        var tcs = new TaskCompletionSource<T>();
        operation.Completed = delegate
        {
            switch(operation.Status)
            {
                case Windows.Foundation.AsyncStatus.Completed:
                    tcs.SetResult(operation.GetResults());
                    break;
                case Windows.Foundation.AsyncStatus.Error:
                    tcs.SetException(operation.ErrorCode);
                    break;
                case Windows.Foundation.AsyncStatus.Canceled:
                    tcs.SetCanceled();
                    break;
            }
        };
        return tcs.Task;
    }
    
    public static TaskAwaiter<T> GetAwaiter<T>(this Windows.Foundation.IAsyncOperation<T> operation)
    {
        return operation.AsTask().GetAwaiter();
    }
}
