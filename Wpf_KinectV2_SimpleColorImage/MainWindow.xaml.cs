using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Windows.Media.Media3D;

using AForge.Video.VFW;
using AForge.Video.FFMPEG;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

using Microsoft.Kinect;
using Microsoft.Kinect.Face;

using NAudio.Wave;

namespace Wpf_KinectV2_SimpleColorImage
{

    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Kinect 本体への参照。
        /// </summary>
        KinectSensor kinect;
		
		private BodyFrameSource _bodySource = null;

		private BodyFrameReader _bodyReader = null;

		private HighDefinitionFaceFrameSource _faceSource = null;

		private HighDefinitionFaceFrameReader _faceReader = null;

		private FaceAlignment _faceAlignment = null;

		private FaceModel _faceModel = null;

		private List<Ellipse> _points = new List<Ellipse>();

		WaveIn waveIn;
		BufferedWaveProvider wavProvider;
		WaveFileWriter waveWriter;
		int cnt = 0;

		public StreamWriter sw = new StreamWriter(@"c:\users\abelab\desktop\log\log_audio_ver.csv", false, Encoding.GetEncoding("utf-8"));
		/*
		private MeshGeometry3D _Geometry3d;
		public MeshGeometry3D Geometry3d
		{
			get { return this._Geometry3d; }
			set
			{
				this._Geometry3d = value;
			}
		}
		*/
		/// <summary>
		/// 取得するカラー画像の詳細情報。
		/// </summary>
		FrameDescription colorFrameDescription;

        /// <summary>
        /// 取得するカラー画像のフォーマット。
        /// </summary>
        ColorImageFormat colorImageFormat;

        /// <summary>
        /// カラー画像を継続的に読み込むためのリーダ。
        /// </summary>
        ColorFrameReader colorFrameReader;

		int count = 0;
		
        /// <summary>
        /// コンストラクタ。実行時に一度だけ実行される。
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            //Kinect 本体への参照を確保する。
            this.kinect = KinectSensor.GetDefault();

            //読み込む画像のフォーマットとリーダを設定する。
            this.colorImageFormat = ColorImageFormat.Bgra;
            this.colorFrameDescription
                = this.kinect.ColorFrameSource.CreateFrameDescription(this.colorImageFormat);
            this.colorFrameReader = this.kinect.ColorFrameSource.OpenReader();
            this.colorFrameReader.FrameArrived += ColorFrameReader_FrameArrived;

			// 顔回転検出用
			//this.faceFrameSource = new FaceFrameSource(this.kinect, 0, );


			if (this.kinect != null)
			{
				_bodySource = kinect.BodyFrameSource;
				_bodyReader = _bodySource.OpenReader();
				_bodyReader.FrameArrived += BodyReader_FrameArrived;

				_faceSource = new HighDefinitionFaceFrameSource(kinect);

				_faceReader = _faceSource.OpenReader();
				_faceReader.FrameArrived += FaceReader_FrameArrived;

				_faceModel = new FaceModel();
				_faceAlignment = new FaceAlignment();
			}

			//Kinect の動作を開始する。
			//aviWriter.FrameRate = 30;
			//aviWriter.Open(@"c:\users\abelab\desktop\log\test.avi", 1920, 1080);
			//writer.Open(@"c:\users\abelab\desktop\log\test.avi", 1920, 1080, 30, VideoCodec.MPEG4);
			/*
			for (int i=0; i<1347; i++)
			{
				sw.Write(i + ",,,,,,");
			}
			sw.WriteLine();
			for(int i=0; i<1347; i++)
			{
				sw.Write("X(m),Y(m),Z(m),X(pixel),Y(pixel),,");
			}
			sw.WriteLine();
			*/
			this.kinect.Open();
        }

		/*
		public static Bitmap toBitmap(byte[] pixel, int width, int height)
		{
			if (pixel == null) return null;

			var bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
			var data = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, bitmap.PixelFormat);
			Marshal.Copy(pixel, 0, data.Scan0, pixel.Length);
			bitmap.UnlockBits(data);
			return bitmap;
		}
		*/


        /// <summary>
        /// Kinect がカラー画像を取得したとき実行されるメソッド(イベントハンドラ)。
        /// </summary>
        /// <param name="sender">
        /// イベントを通知したオブジェクト。ここでは Kinect になる。
        /// </param>
        /// <param name="e">
        /// イベントの発生時に渡されるデータ。ここではカラー画像の情報が含まれる。
        /// </param>
		/// 
        void ColorFrameReader_FrameArrived(object sender, ColorFrameArrivedEventArgs e)
        {
            //通知されたフレームを取得する。
            ColorFrame colorFrame = e.FrameReference.AcquireFrame();

            //フレームが上手く取得できない場合がある。
            if (colorFrame == null)
            {
                return;
            }

            //画素情報を確保する領域(バッファ)を用意する。
            //"高さ * 幅 * 画素あたりのデータ量"だけ保存できれば良い。
            byte[] colors = new byte[this.colorFrameDescription.Width
                                     * this.colorFrameDescription.Height
                                     * this.colorFrameDescription.BytesPerPixel];

            //用意した領域に画素情報を複製する。
            colorFrame.CopyConvertedFrameDataToArray(colors, this.colorImageFormat);

            //画素情報をビットマップとして扱う。
            BitmapSource bitmapSource
                = BitmapSource.Create(this.colorFrameDescription.Width,
                                      this.colorFrameDescription.Height,
                                      96,
                                      96,
                                      PixelFormats.Bgra32,
                                      null,
                                      colors,
                                      this.colorFrameDescription.Width * (int)this.colorFrameDescription.BytesPerPixel);
			//Bitmap image = new Bitmap(colorFrameDescription.Width, colorFrameDescription.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
			//writer.WriteVideoFrame(image);
			//imageBitmap = toBitmap(colors, imageBitmap.Width, imageBitmap.Height);
			//imageBitmap.Save(@"c:\users\abelab\desktop\log\picture\image_" + count + ".bmp", ImageFormat.Bmp);
			//count++;
			//aviWriter.AddFrame(imageBitmap);
            //キャンバスに表示する。
            this.canvas.Background = new ImageBrush(bitmapSource);
			
            //取得したフレームを破棄する。
            colorFrame.Dispose();
        }

        /// <summary>
        /// この WPF アプリケーションが終了するときに実行されるメソッド。
        /// </summary>
        /// <param name="e">
        /// イベントの発生時に渡されるデータ。
        /// </param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
			sw.Close();

			//カラー画像の取得を中止して、関連するリソースを破棄する。
			if (this.colorFrameReader != null) 
            {
				//aviWriter.Dispose();
                this.colorFrameReader.Dispose();
                this.colorFrameReader = null;
            }

            //Kinect を停止して、関連するリソースを破棄する。
            if (this.kinect != null)
            {
				//writer.Close();
				//aviWriter.Close();
                this.kinect.Close();
                this.kinect = null;
            }
        }

		private void BodyReader_FrameArrived(object sender, BodyFrameArrivedEventArgs e)
		{
			using (var frame = e.FrameReference.AcquireFrame())
			{
				if (frame != null)
				{
					Body[] bodies = new Body[frame.BodyCount];
					frame.GetAndRefreshBodyData(bodies);

					Body body = bodies.Where(b => b.IsTracked).FirstOrDefault();

					if (!_faceSource.IsTrackingIdValid)
					{
						if (body != null)
						{
							_faceSource.TrackingId = body.TrackingId;
						}
					}
				}
			}
		}

		private void FaceReader_FrameArrived(object sender, HighDefinitionFaceFrameArrivedEventArgs e)
		{
			using (var frame = e.FrameReference.AcquireFrame())
			{
				if (frame != null && frame.IsFaceTracked)
				{
					frame.GetAndRefreshFaceAlignmentResult(_faceAlignment);

					UpdateFacePoints();
				}
			}
		}

		private void UpdateFacePoints()
		{
			if (_faceModel == null) return;

			var vertices = _faceModel.CalculateVerticesForAlignment(_faceAlignment);
			if (vertices.Count > 0)
			{
				if (_points.Count == 0)
				{
					for (int index = 0; index < vertices.Count; index++)
					{
						Ellipse ellipse_mouth = new Ellipse
						{
							Width = 2.0,
							Height = 2.0,
							Fill = new SolidColorBrush(Colors.Blue)
						};

						_points.Add(ellipse_mouth);
					}
					foreach (Ellipse ellipse in _points)
					{
						canvas.Children.Add(ellipse);
					}
				}

				for (int index = 0; index < vertices.Count; index++)
				{
					var vert = vertices[index];
					/*
					vert.X = vert.X / vert.Z;
					vert.Y = vert.Y / vert.Z;
					vert.Z = vert.Z / vert.Z;
					*/
					//this._Geometry3d.Positions[index] = new Point3D(vert.X, vert.Y, -vert.Z);
					sw.Write(vert.X + "," + vert.Y + "," + vert.Z + ",");

					CameraSpacePoint vertice = vert;
					ColorSpacePoint point = kinect.CoordinateMapper.MapCameraPointToColorSpace(vertice);
					//sw.Write(point.X + "," + point.Y + ",");

					if (float.IsInfinity(point.X) || float.IsInfinity(point.Y)) return;

					Ellipse ellipse = _points[index];

					Canvas.SetLeft(ellipse, point.X);
					Canvas.SetTop(ellipse, point.Y);
					
				}
				sw.WriteLine();
			}

		}

		private void Rec_Button_Click(object sender, RoutedEventArgs e)
		{
			waveIn = new WaveIn();
			waveIn.DeviceNumber = 0;
			waveIn.WaveFormat = new WaveFormat(44100, 16, WaveIn.GetCapabilities(0).Channels);

			waveIn.DataAvailable += new EventHandler<WaveInEventArgs>(sorceStream_DataAvailable);

			waveWriter = new WaveFileWriter(@"c:\users\abelab\desktop\log\out"+cnt+".wav", waveIn.WaveFormat);

			waveIn.StartRecording();
		}

		private void Stop_Button_Click(object sender, RoutedEventArgs e)
		{
			waveIn.StopRecording();
			cnt++;
		}

		private void sorceStream_DataAvailable(object sender, WaveInEventArgs e)
		{
			if (waveWriter == null) return;

			waveWriter.Write(e.Buffer, 0, e.BytesRecorded);
			waveWriter.Flush();
		}
	}
}