using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace GitHub.secile.Audio
{
    public class Wave : IDisposable
    {
        #region IDisposable メンバ
        /// <summary>明示的にDisposeされた場合。</summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this); // デストラクタ不要
        }

        /// <summary>Disposeされなかった場合はデストラクタから呼ぶ。</summary>
        ~Wave()
        { Dispose(false); }

        private event EventHandler Disposing = (s, e) => { };
        protected virtual void Dispose(bool disposing)
        { if (disposing) Disposing(this, EventArgs.Empty); }
        #endregion

        /// <summary>
        /// サンプリングレート。44100など。
        /// </summary>
        public int SamplingRate { get; private set; }

        /// <summary>
        /// 量子化ビット。8か16。
        /// </summary>
        public int BitPerSample { get; private set; }

        /// <summary>
        /// チャンネル数。基本的に1か2。
        /// </summary>
        public int Channels { get; private set; }

        /// <summary>
        /// Waveデータ本体。
        /// </summary>
        public IEnumerable<short> WaveData { get; private set; }

        /// <summary>
        /// Waveデータ数。
        /// </summary>
        public int WaveDataCount { get; private set; }

        /// <summary>
        /// 再生時間をmsで。
        /// </summary>
        public int Length { get { return WaveDataCount / SamplingRate * 1000; } }

        /// <summary>
        /// ストリームから作成。
        /// </summary>
        public Wave(System.IO.Stream stream)
        {
            FromStream(stream);
            Disposing += (s, e) => stream.Close();
        }

        /// <summary>
        /// ファイルから作成。
        /// </summary>
        public Wave(string path) : this(System.IO.File.OpenRead(path)) { }

        /// <summary>
        /// データから作成。
        /// </summary>
        /// <param name="wave_data">-32768～32767範囲で0が無音であるデータ。</param>
        public Wave(IEnumerable<short> wave_data, int rate, int channels)
        {
            FromData(wave_data, rate, 16, channels);
        }

        /// <summary>
        /// データから作成。
        /// </summary>
        /// <param name="wave_data">0～255範囲で128が無音であるデータ。</param>
        /// <remarks>8ビットデータから作成しても、内部的には16ビットで管理される。</remarks>
        public Wave(IEnumerable<byte> wave_data, int rate, int channels)
        {
            var wave_data16 = WaveDataByteToShort(wave_data);
            FromData(wave_data16, rate, 8, channels);
        }

        private void FromData(IEnumerable<short> wave_data, int rate, int bps, int channels)
        {
            int count = wave_data.Count();
            if (count % channels != 0)
            {
                throw new ArgumentOutOfRangeException("count", "チャンネル数に対してデータが不足しています。");
            }
            WaveDataCount = count / channels;

            WaveData = wave_data;

            SamplingRate = rate;
            BitPerSample = bps;
            Channels = channels;
        }

        [StructLayout( LayoutKind.Sequential, Pack=1)]
        private struct RiffHeader
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[] riff_head;
            public UInt32 file_size;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[] riff_type;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct RiffFormatChunk
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[] chunkID;
            public UInt32 chunkSize;
            public UInt16 wFormatTag;
            public UInt16 wChannels;
            public UInt32 dwSamplesPerSec;
            public UInt32 dwAvgBytesPerSec;
            public UInt16 wBlockAlign;
            public UInt16 wBitsPerSample;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct RiffDataChunk
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[] head;
            public UInt32 size;
        }

        private void FromStream(System.IO.Stream stream)
        {
            var br = new System.IO.BinaryReader(stream);

            // RIFFヘッダを読み込む
            var riff_size = Marshal.SizeOf(typeof(RiffHeader));
            var riff_byte = new byte[riff_size];
            br.Read(riff_byte, 0, riff_size);
            var riff_ptr = Marshal.AllocHGlobal(riff_size);
            Marshal.Copy(riff_byte, 0, riff_ptr, riff_size);
            var riff = (RiffHeader)Marshal.PtrToStructure(riff_ptr, typeof(RiffHeader));
            Marshal.FreeHGlobal(riff_ptr);

            // fmtチャンクを読む
            var fmt_size = Marshal.SizeOf(typeof(RiffFormatChunk));
            var fmt_byte = new byte[fmt_size];
            br.Read(fmt_byte, 0, fmt_size);
            var fmt_ptr = Marshal.AllocHGlobal(fmt_size);
            Marshal.Copy(fmt_byte, 0, fmt_ptr, fmt_size);
            var fmt = (RiffFormatChunk)Marshal.PtrToStructure(fmt_ptr, typeof(RiffFormatChunk));
            Marshal.FreeHGlobal(fmt_ptr);

            // fmtチャンクの拡張部分を読む（通常無いことが多い）
            var fmt_extend_bytes = fmt.chunkSize - 16; // fmtチャンクの拡張部分のバイト数
            if (fmt_extend_bytes > 0)
            {
                br.ReadBytes((int)fmt_extend_bytes);  // 読み捨てる
            }

            // dataチャンクが現れるまでチャンクを読み捨てる
            // サウンドレコーダーで録音したファイルはここにfactチャンクがあった
            while (true)
            {
                var fourcc = br.ReadChars(4);
                var type = new string(fourcc);
                if (type == "data") break;
                var chunk_size = br.ReadInt32();
                br.ReadBytes(chunk_size);
            }

            var channels = fmt.wChannels;
            var rate = fmt.dwSamplesPerSec;
            var bps = fmt.wBitsPerSample;
            var wave_size = br.ReadInt32();
            var data_count = wave_size / (bps / 8) / channels;

            Channels = channels;
            SamplingRate = (int)rate;
            BitPerSample = bps;
            WaveDataCount = data_count;

            var DataChunkLength = data_count * channels;
            var DataChunkPosition = stream.Position;

            // Waveデータの読み出し
            if (BitPerSample == 16)
            {
                // 16ビットはそのまま読む
                WaveData = WaveDataAsShort(br, DataChunkPosition, DataChunkLength);
            }
            else
            {
                // 8ビットは16ビットとして読み込む
                var data = WaveDataAsByte(br, DataChunkPosition, DataChunkLength);
                WaveData = WaveDataByteToShort(data);
            }
        }

        /// <summary>
        /// チャンネルごとに分割する。
        /// </summary>
        public Wave[] Split()
        {
            short[][] wave_data = new short[Channels][];
            for (int i = 0; i < Channels; i++) wave_data[i] = new short[WaveDataCount];

            int idx = 0;
            using( var e = WaveData.GetEnumerator())
            {
                while(e.MoveNext())
                {
                    wave_data[idx % Channels][idx / Channels] = e.Current;
                    idx++;
                }
            }

            var result = new Wave[Channels];
            for (int i = 0; i < Channels; i++) result[i] = new Wave(wave_data[i], SamplingRate, 1);
            return result;
        }

        /// <summary>
        /// 再生する。
        /// </summary>
        /// <returns>停止する関数。</returns>
        public Action Play(bool loop)
        {
            var ms = new System.IO.MemoryStream();
            ToStream(ms);
            ms.Position = 0;
            
            var player = new System.Media.SoundPlayer(ms);
            if (loop) { player.PlayLooping(); }
            else      { player.Play(); }

            return () =>
            {
                player.Stop();
                ms.Close();
            };
        }

        /// <summary>
        /// ファイルに保存する。
        /// </summary>
        public void Save(string path)
        {
            Save(System.IO.File.OpenWrite(path));
        }

        /// <summary>
        /// ストリームに保存する。
        /// </summary>
        public void Save(System.IO.Stream stream)
        {
            ToStream(stream);
            stream.Close();
        }

        private void ToStream(System.IO.Stream stream)
        {
            WriteHeader(stream, WaveDataCount, Channels, SamplingRate, BitPerSample);

            if (BitPerSample == 8)
            {
                // 8ビットに戻してから保存
                var wave_data = WaveDataByteToShort(WaveData);
                WriteBody(stream, wave_data);
            }
            else
            {
                // そのまま保存
                WriteBody(stream, WaveData);
            }
        }

        /// <summary>
        /// 時間を指定して抽出する。
        /// </summary>
        public Wave Extract(int start_ms, int length_ms)
        {
            var data_idx = start_ms * SamplingRate / 1000;
            var data_len = length_ms * SamplingRate / 1000;
            return ExtractData(data_idx, data_len);
        }

        private Wave ExtractData(int data_idx, int data_len)
        {
            // 範囲を超えていたら調整
            if (data_idx + data_len > WaveDataCount)
            {
                data_len = WaveDataCount - data_idx;
            }

            data_idx *= Channels;
            data_len *= Channels;

            var wave_data = WaveData.Skip(data_idx).Take(data_len);
            var result = new Wave(wave_data, SamplingRate, Channels);
            return result;
        }

        public Wave ToMonaural()
        {
            // チャンネルの平均値のデータにする
            var wave_data = WaveData.Buffer(Channels).Select(x => x.Average(y => y)).Cast<short>();
            return new Wave(wave_data, SamplingRate, 1);
        }

        private void WriteHeader(System.IO.Stream stream, int data_count, int channels, int rate, int bps)
        {
            var br = new System.IO.BinaryWriter(stream);
            var header = MakeWaveHader(data_count, channels, rate, bps);
            br.Write(header);
        }
        private byte[] MakeWaveHader(int data_count, int channels, int rate, int bps)
        {
            var data_size = data_count * (bps / 8) * channels;
            var file_size = data_size + 36;

            var riff = new RiffHeader();
            riff.riff_head = GetBytes("RIFF");
            riff.file_size = (uint)file_size;
            riff.riff_type = GetBytes("WAVE");

            var fmt = new RiffFormatChunk();
            fmt.chunkID = GetBytes("fmt ");
            fmt.chunkSize = 16;
            fmt.wFormatTag = 1;
            fmt.wChannels = (ushort)channels;
            fmt.dwSamplesPerSec = (uint)rate;
            fmt.dwAvgBytesPerSec = (uint)(rate * (bps / 8) * channels);
            fmt.wBlockAlign = (ushort)((bps / 8) * channels);
            fmt.wBitsPerSample = (ushort)bps;

            var data = new RiffDataChunk();
            data.head = GetBytes("data");
            data.size = (uint)data_size;

            var size = Marshal.SizeOf(riff) + Marshal.SizeOf(fmt) + Marshal.SizeOf(data);

            var riff_ptr  = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(riff, riff_ptr, false);

            var fmt_ptr = new IntPtr(riff_ptr.ToInt32() + Marshal.SizeOf(riff));
            Marshal.StructureToPtr(fmt, fmt_ptr, false);

            var data_ptr  = new IntPtr(fmt_ptr.ToInt32() + Marshal.SizeOf(fmt));
            Marshal.StructureToPtr(data, data_ptr, false);

            var header = new byte[size];
            Marshal.Copy(riff_ptr, header, 0, size);

            Marshal.FreeHGlobal(riff_ptr);
            // Marshal.FreeHGlobal(fmt_ptr);  //riff_ptrを開放すると同時に開放されるので不要
            // Marshal.FreeHGlobal(data_ptr); //riff_ptrを開放すると同時に開放されるので不要

            return header;
        }

        private static byte[] GetBytes(string s)
        {
            return System.Text.Encoding.ASCII.GetBytes(s);
        }

        private void WriteBody(System.IO.Stream stream, IEnumerable<short> wave_data)
        {
            var bw = new System.IO.BinaryWriter(stream);
            foreach (var item in wave_data)
            {
                bw.Write(item);
            }
        }
        private void WriteBody(System.IO.Stream stream, IEnumerable<byte> wave_data)
        {
            var bw = new System.IO.BinaryWriter(stream);
            foreach (var item in wave_data)
            {
                bw.Write(item);
            }
        }

        /// <summary>
        /// Waveデータを16ビットとして読む
        /// </summary>
        private IEnumerable<short> WaveDataAsShort(System.IO.BinaryReader br, long start, int length)
        {
            br.BaseStream.Position = start;
            for (long i = 0; i < length; i++)
            {
                yield return br.ReadInt16();
            }
        }

        /// <summary>
        /// Waveデータを8ビットとして読む
        /// </summary>
        private IEnumerable<byte> WaveDataAsByte(System.IO.BinaryReader br, long start, int length)
        {
            br.BaseStream.Position = start;
            for (long i = 0; i < length; i++)
            {
                yield return br.ReadByte();
            }
        }

        /// <summary>
        /// Waveのデータ形式を8ビットから16ビットに変換する。
        /// </summary>
        private static IEnumerable<short> WaveDataByteToShort(IEnumerable<byte> wave_data)
        {
            foreach (var item in wave_data)
            {
                yield return WaveDataByteToShort(item);
            }
        }

        /// <summary>
        /// Waveのデータ形式を8ビットから16ビットに変換する。
        /// </summary>
        private static short WaveDataByteToShort(byte wave_data)
        {
            // 0-255を0-65535に対応させる
            var tmp = ((double)wave_data / (byte.MaxValue + 1) * (ushort.MaxValue + 1)) - 1;

            // 128 = 32767に対応させる
            return (short)(tmp - short.MaxValue);
        }

        private static IEnumerable<byte> WaveDataByteToShort(IEnumerable<short> wave_data)
        {
            foreach (var item in wave_data)
            {
                yield return WaveDataShortToByte(item);
            }
        }
        private static byte WaveDataShortToByte(short wave_data)
        {
            // 0～65535を0～255に対応させる
            var tmp = (double)wave_data / (ushort.MaxValue + 1) * (byte.MaxValue + 1);
            
            // 0=128に対応させる
            return (byte)(tmp + ((Byte.MaxValue + 1) / 2));
        }

        /// <summary>
        /// データをshort(-32768～32767)に変換する
        /// </summary>
        public static IEnumerable<short> NormalizeDataToShort(IEnumerable<float> wave_data, float range_max)
        {
            foreach (var item in wave_data)
            {
                yield return NormalizeDataToShort(item, range_max);
            }
        }

        /// <summary>
        /// データをshort(-32768～32767)に変換する
        /// </summary>
        public static short NormalizeDataToShort(float wave_data, float range_max)
        {
            return (short)(wave_data / range_max * short.MaxValue);
        }
    }

    public static class LinqEx
    {
        public static IEnumerable<IEnumerable<T>> Buffer<T>(this IEnumerable<T> source, int count)
        {
            var result = new List<T>(count);
            foreach (var item in source)
            {
                result.Add(item);
                if (result.Count == count)
                {
                    yield return result;
                    result = new List<T>(count);
                }
            }
            if (result.Count != 0)
                yield return result.ToArray();
        }
    }
}
