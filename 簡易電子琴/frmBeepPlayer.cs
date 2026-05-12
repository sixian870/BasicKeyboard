using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace 簡易電子琴
{

    public partial class frmBeepPlayer : Form
    {
        //[DllImport("kernel32.dll")]

        [DllImport("winmm.dll")]

        private static extern int midiOutOpen(out IntPtr lphmo, int uDeviceID, int dwCallback, int dwInstance, int dwFlags);

        [DllImport("winmm.dll")]

        private static extern int midiOutShortMsg(IntPtr hmo, int dwMsg);

        [DllImport("winmm.dll")]

        private static extern int midiOutClose(IntPtr hmo);

        IntPtr midiHandle;
        int offset = 0;
        private Dictionary<Keys, int> keyMap = new Dictionary<Keys, int>();
        private Dictionary<int, Button> buttonMap = new Dictionary<int, Button>();
        private Dictionary<Button, Color> originalColors = new Dictionary<Button, Color>();
        private HashSet<Keys> pressedKeys = new HashSet<Keys>(); //防止重複按鍵

        //public static extern bool Beep(int frequency, int duration);
        //int[] freq = { 523, 554, 587, 622, 659, 698, 740, 784, 831, 880, 932, 988, 1046 };

        int initWidth = 0;
        int initHeight = 0;
        Dictionary<string, Rect> initControl = new Dictionary<string, Rect>();

        public frmBeepPlayer()
        {
            InitializeComponent();
            InitializeKeyMap();
        }

        private void InitializeKeyMap()
        {
            // 第一八度白鍵
            keyMap[Keys.Z] = 0; // Do
            keyMap[Keys.X] = 2; // Re
            keyMap[Keys.C] = 4; // Mi
            keyMap[Keys.V] = 5; // Fa
            keyMap[Keys.B] = 7; // So
            keyMap[Keys.N] = 9; // La
            keyMap[Keys.M] = 11; // Si
            // 第一八度黑鍵
            keyMap[Keys.S] = 1; // Do#
            keyMap[Keys.D] = 3; // Re#
            keyMap[Keys.G] = 6; // Fa#
            keyMap[Keys.H] = 8; // So#
            keyMap[Keys.J] = 10; // La#
            // 第二八度白鍵
            keyMap[Keys.Q] = 12; // Do (高)
            keyMap[Keys.W] = 14; // Re (高)
            keyMap[Keys.E] = 16; // Mi (高)
            keyMap[Keys.R] = 17; // Fa (高)
            keyMap[Keys.T] = 19; // So (高)
            keyMap[Keys.Y] = 21; // La (高)
            keyMap[Keys.U] = 23; // Si (高)
            keyMap[Keys.I] = 24; // Do (更高)
            // 第二八度黑鍵
            keyMap[Keys.D2] = 13; // Do# (高)
            keyMap[Keys.D3] = 15; // Re# (高)
            keyMap[Keys.D5] = 18; // Fa# (高)
            keyMap[Keys.D6] = 20; // So# (高)
            keyMap[Keys.D7] = 22; // La# (高)

        }

        private void InitializeInstruments()
        {
            List<InstrumentItem> list = new List<InstrumentItem>
            {
                // 0-7 鋼琴
                new InstrumentItem { Id = 0, Name = "[鋼琴] 大鋼琴 (Acoustic Grand Piano)" },
                new InstrumentItem { Id = 1, Name = "[鋼琴] 亮音鋼琴 (Bright Acoustic Piano)" },
                new InstrumentItem { Id = 2, Name = "[鋼琴] 電鋼琴 (Electric Grand Piano)" },
                new InstrumentItem { Id = 3, Name = "[鋼琴] 酒吧鋼琴 (Honky-tonk Piano)" },
                new InstrumentItem { Id = 4, Name = "[鋼琴] 柔和電鋼琴 (Rhodes Piano)" },
                new InstrumentItem { Id = 5, Name = "[鋼琴] 合唱效果電鋼琴 (Chorused Piano)" },
                new InstrumentItem { Id = 6, Name = "[鋼琴] 羽管鍵琴 (Harpsichord)" },
                new InstrumentItem { Id = 7, Name = "[鋼琴] 擊弦古鋼琴 (Clavinet)" },

                // 8-15 擊弦與鐵琴類
                new InstrumentItem { Id = 8, Name = "[鐵琴] 鋼片琴 (Celesta)" },
                new InstrumentItem { Id = 9, Name = "[鐵琴] 鐘琴 (Glockenspiel)" },
                new InstrumentItem { Id = 10, Name = "[鐵琴] 音樂盒 (Music Box)" },
                new InstrumentItem { Id = 11, Name = "[鐵琴] 顫音琴 (Vibraphone)" },
                new InstrumentItem { Id = 12, Name = "[鐵琴] 馬林巴琴 (Marimba)" },
                new InstrumentItem { Id = 13, Name = "[鐵琴] 木琴 (Xylophone)" },
                new InstrumentItem { Id = 14, Name = "[鐵琴] 管鐘 (Tubular Bells)" },
                new InstrumentItem { Id = 15, Name = "[鐵琴] 揚琴 (Dulcimer)" },

                // 16-23 風琴
                new InstrumentItem { Id = 16, Name = "[風琴] 擊桿風琴 (Drawbar Organ)" },
                new InstrumentItem { Id = 17, Name = "[風琴] 敲擊風琴 (Percussive Organ)" },
                new InstrumentItem { Id = 18, Name = "[風琴] 搖滾風琴 (Rock Organ)" },
                new InstrumentItem { Id = 19, Name = "[風琴] 教堂管風琴 (Church Organ)" },
                new InstrumentItem { Id = 20, Name = "[風琴] 簧管風琴 (Reed Organ)" },
                new InstrumentItem { Id = 21, Name = "[風琴] 手風琴 (Accordion)" },
                new InstrumentItem { Id = 22, Name = "[風琴] 口琴 (Harmonica)" },
                new InstrumentItem { Id = 23, Name = "[風琴] 探戈手風琴 (Tango Accordion)" },

                // 24-31 吉他
                new InstrumentItem { Id = 24, Name = "[吉他] 尼龍吉他 (Acoustic Guitar - Nylon)" },
                new InstrumentItem { Id = 25, Name = "[吉他] 鋼弦吉他 (Acoustic Guitar - Steel)" },
                new InstrumentItem { Id = 26, Name = "[吉他] 爵士電吉他 (Electric Guitar - Jazz)" },
                new InstrumentItem { Id = 27, Name = "[吉他] 清音電吉他 (Electric Guitar - Clean)" },
                new InstrumentItem { Id = 28, Name = "[吉他] 悶音電吉他 (Electric Guitar - Muted)" },
                new InstrumentItem { Id = 29, Name = "[吉他] 破音電吉他 (Overdriven Guitar)" },
                new InstrumentItem { Id = 30, Name = "[吉他] 失真電吉他 (Distortion Guitar)" },
                new InstrumentItem { Id = 31, Name = "[吉他] 吉他和音 (Guitar Harmonics)" },

                // 32-39 Bass
                new InstrumentItem { Id = 32, Name = "[貝斯] 大貝斯 (Acoustic Bass)" },
                new InstrumentItem { Id = 33, Name = "[貝斯] 電貝斯-指彈 (Electric Bass - Finger)" },
                new InstrumentItem { Id = 34, Name = "[貝斯] 電貝斯-撥片 (Electric Bass - Pick)" },
                new InstrumentItem { Id = 35, Name = "[貝斯] 無琴格貝斯 (Fretless Bass)" },
                new InstrumentItem { Id = 36, Name = "[貝斯] 拍弦貝斯 1 (Slap Bass 1)" },
                new InstrumentItem { Id = 37, Name = "[貝斯] 拍弦貝斯 2 (Slap Bass 2)" },
                new InstrumentItem { Id = 38, Name = "[貝斯] 合成貝斯 1 (Synth Bass 1)" },
                new InstrumentItem { Id = 39, Name = "[貝斯] 合成貝斯 2 (Synth Bass 2)" },

                // 40-47 弦樂
                new InstrumentItem { Id = 40, Name = "[弦樂] 小提琴 (Violin)" },
                new InstrumentItem { Id = 41, Name = "[弦樂] 中提琴 (Viola)" },
                new InstrumentItem { Id = 42, Name = "[弦樂] 大提琴 (Cello)" },
                new InstrumentItem { Id = 43, Name = "[弦樂] 低音大提琴 (Contrabass)" },
                new InstrumentItem { Id = 44, Name = "[弦樂] 顫音弦樂 (Tremolo Strings)" },
                new InstrumentItem { Id = 45, Name = "[弦樂] 撥奏弦樂 (Pizzicato Strings)" },
                new InstrumentItem { Id = 46, Name = "[弦樂] 豎琴 (Orchestral Harp)" },
                new InstrumentItem { Id = 47, Name = "[弦樂] 定音鼓 (Timpani)" },

                // 48-55 合奏
                new InstrumentItem { Id = 48, Name = "[合奏] 弦樂合奏 1 (String Ensemble 1)" },
                new InstrumentItem { Id = 49, Name = "[合奏] 弦樂合奏 2 (String Ensemble 2)" },
                new InstrumentItem { Id = 50, Name = "[合奏] 合成弦樂 1 (Synth Strings 1)" },
                new InstrumentItem { Id = 51, Name = "[合奏] 合成弦樂 2 (Synth Strings 2)" },
                new InstrumentItem { Id = 52, Name = "[合奏] 合唱人聲 (Choir Aahs)" },
                new InstrumentItem { Id = 53, Name = "[合奏] 嘟嘟人聲 (Voice Oohs)" },
                new InstrumentItem { Id = 54, Name = "[合奏] 合成人聲 (Synth Voice)" },
                new InstrumentItem { Id = 55, Name = "[合奏] 交響打擊樂 (Orchestra Hit)" },

                // 56-63 銅管
                new InstrumentItem { Id = 56, Name = "[銅管] 小號 (Trumpet)" },
                new InstrumentItem { Id = 57, Name = "[銅管] 長號 (Trombone)" },
                new InstrumentItem { Id = 58, Name = "[銅管] 大號/低音號 (Tuba)" },
                new InstrumentItem { Id = 59, Name = "[銅管] 弱音小號 (Muted Trumpet)" },
                new InstrumentItem { Id = 60, Name = "[銅管] 法國號 (French Horn)" },
                new InstrumentItem { Id = 61, Name = "[銅管] 銅管合奏 (Brass Section)" },
                new InstrumentItem { Id = 62, Name = "[銅管] 合成銅管 1 (SynthBrass 1)" },
                new InstrumentItem { Id = 63, Name = "[銅管] 合成銅管 2 (SynthBrass 2)" },

                // 64-71 簧管
                new InstrumentItem { Id = 64, Name = "[簧管] 高音薩克斯風 (Soprano Sax)" },
                new InstrumentItem { Id = 65, Name = "[簧管] 次中音薩克斯風 (Alto Sax)" },
                new InstrumentItem { Id = 66, Name = "[簧管] 中音薩克斯風 (Tenor Sax)" },
                new InstrumentItem { Id = 67, Name = "[簧管] 低音薩克斯風 (Baritone Sax)" },
                new InstrumentItem { Id = 68, Name = "[簧管] 雙簧管 (Oboe)" },
                new InstrumentItem { Id = 69, Name = "[簧管] 英國管 (English Horn)" },
                new InstrumentItem { Id = 70, Name = "[簧管] 巴松管 (Bassoon)" },
                new InstrumentItem { Id = 71, Name = "[簧管] 單簧管/黑管 (Clarinet)" },

                // 72-79 笛類
                new InstrumentItem { Id = 72, Name = "[笛類] 短笛 (Piccolo)" },
                new InstrumentItem { Id = 73, Name = "[笛類] 長笛 (Flute)" },
                new InstrumentItem { Id = 74, Name = "[笛類] 豎笛/直笛 (Recorder)" },
                new InstrumentItem { Id = 75, Name = "[笛類] 排蕭 (Pan Flute)" },
                new InstrumentItem { Id = 76, Name = "[笛類] 蘆笛 (Blown Bottle)" },
                new InstrumentItem { Id = 77, Name = "[笛類] 日本尺八 (Shakuhachi)" },
                new InstrumentItem { Id = 78, Name = "[笛類] 哨笛 (Whistle)" },
                new InstrumentItem { Id = 79, Name = "[笛類] 陶笛 (Ocarina)" },

                // 80-87 合成主音
                new InstrumentItem { Id = 80, Name = "[合成主音] 方波 (Lead 1 - Square)" },
                new InstrumentItem { Id = 81, Name = "[合成主音] 鋸齒波 (Lead 2 - Sawtooth)" },
                new InstrumentItem { Id = 82, Name = "[合成主音] 汽笛風琴 (Lead 3 - Calliope)" },
                new InstrumentItem { Id = 83, Name = "[合成主音] 吹管 (Lead 4 - Chiff)" },
                new InstrumentItem { Id = 84, Name = "[合成主音] 電吉他 (Lead 5 - Charang)" },
                new InstrumentItem { Id = 85, Name = "[合成主音] 人聲 (Lead 6 - Voice)" },
                new InstrumentItem { Id = 86, Name = "[合成主音] 五度音 (Lead 7 - Fifths)" },
                new InstrumentItem { Id = 87, Name = "[合成主音] 貝斯加主音 (Lead 8 - Bass + Lead)" },

                // 88-95 合成柔音
                new InstrumentItem { Id = 88, Name = "[合成柔音] 新時代 (Pad 1 - New age)" },
                new InstrumentItem { Id = 89, Name = "[合成柔音] 暖音 (Pad 2 - Warm)" },
                new InstrumentItem { Id = 90, Name = "[合成柔音] 複音 (Pad 3 - Polysynth)" },
                new InstrumentItem { Id = 91, Name = "[合成柔音] 人聲合唱 (Pad 4 - Choir)" },
                new InstrumentItem { Id = 92, Name = "[合成柔音] 弓弦 (Pad 5 - Bowed)" },
                new InstrumentItem { Id = 93, Name = "[合成柔音] 金屬 (Pad 6 - Metallic)" },
                new InstrumentItem { Id = 94, Name = "[合成柔音] 光環 (Pad 7 - Halo)" },
                new InstrumentItem { Id = 95, Name = "[合成柔音] 掃掠 (Pad 8 - Sweep)" },

                // 96-103 合成特效
                new InstrumentItem { Id = 96, Name = "[合成特效] 雨滴 (FX 1 - Rain)" },
                new InstrumentItem { Id = 97, Name = "[合成特效] 音軌 (FX 2 - Soundtrack)" },
                new InstrumentItem { Id = 98, Name = "[合成特效] 水晶 (FX 3 - Crystal)" },
                new InstrumentItem { Id = 99, Name = "[合成特效] 氣氛 (FX 4 - Atmosphere)" },
                new InstrumentItem { Id = 100, Name = "[合成特效] 明亮 (FX 5 - Brightness)" },
                new InstrumentItem { Id = 101, Name = "[合成特效] 鬼怪 (FX 6 - Goblins)" },
                new InstrumentItem { Id = 102, Name = "[合成特效] 回聲 (FX 7 - Echoes)" },
                new InstrumentItem { Id = 103, Name = "[合成特效] 科幻 (FX 8 - Sci-fi)" },

                // 104-111 民族樂器
                new InstrumentItem { Id = 104, Name = "[民族] 西塔琴 (Sitar)" },
                new InstrumentItem { Id = 105, Name = "[民族] 斑鳩琴 (Banjo)" },
                new InstrumentItem { Id = 106, Name = "[民族] 三味線 (Shamisen)" },
                new InstrumentItem { Id = 107, Name = "[民族] 箏 (Koto)" },
                new InstrumentItem { Id = 108, Name = "[民族] 拇指琴 (Kalimba)" },
                new InstrumentItem { Id = 109, Name = "[民族] 風笛 (Bagpipe)" },
                new InstrumentItem { Id = 110, Name = "[民族] 提琴/古提琴 (Fiddle)" },
                new InstrumentItem { Id = 111, Name = "[民族] 嗩吶 (Shanai)" },

                // 112-119 打擊樂器
                new InstrumentItem { Id = 112, Name = "[打擊] 叮噹鈴 (Tinkle Bell)" },
                new InstrumentItem { Id = 113, Name = "[打擊] 牛鈴 (Agogo)" },
                new InstrumentItem { Id = 114, Name = "[打擊] 鋼鼓 (Steel Drums)" },
                new InstrumentItem { Id = 115, Name = "[打擊] 木魚 (Woodblock)" },
                new InstrumentItem { Id = 116, Name = "[打擊] 太鼓 (Taiko Drum)" },
                new InstrumentItem { Id = 117, Name = "[打擊] 旋律筒鼓 (Melodic Tom)" },
                new InstrumentItem { Id = 118, Name = "[打擊] 合成鼓 (Synth Drum)" },
                new InstrumentItem { Id = 119, Name = "[打擊] 銅鈸 (Reverse Cymbal)" },

                // 120-127 音效
                new InstrumentItem { Id = 120, Name = "[音效] 吉他品格雜音 (Guitar Fret Noise)" },
                new InstrumentItem { Id = 121, Name = "[音效] 呼吸聲 (Breath Noise)" },
                new InstrumentItem { Id = 122, Name = "[音效] 海浪聲 (Seashore)" },
                new InstrumentItem { Id = 123, Name = "[音效] 鳥鳴聲 (Bird Tweet)" },
                new InstrumentItem { Id = 124, Name = "[音效] 電話鈴聲 (Telephone Ring)" },
                new InstrumentItem { Id = 125, Name = "[音效] 直升機 (Helicopter)" },
                new InstrumentItem { Id = 126, Name = "[音效] 掌聲 (Applause)" },
                new InstrumentItem { Id = 127, Name = "[音效] 槍聲 (Gunshot)" }
            };

            cboInstrument.DataSource = list;
            cboInstrument.DisplayMember = "DisplayName";
            cboInstrument.ValueMember = "Id";
        }

        private void ChangeInstrument(int program)
        {
            if (program < 0 || program > 127) return;
            midiOutShortMsg(midiHandle, 0x0000C0 | (program << 8));
        }

        private void PlayNote(int index)
        {
            // 中央Do:60
            int midiNote = 60 + index + offset;
            if(midiNote < 0 || midiNote > 127) return; // 0-127

            int msg = 0x007F0090 | (midiNote << 8);
            midiOutShortMsg(midiHandle, msg);

            // 延時一小段時間自動放開 (不寫聲音會持續)
            Task.Delay(300).ContinueWith(_ => {
                int stopMsg = 0x00000080 | (midiNote << 8);
                midiOutShortMsg(midiHandle, stopMsg);
            });
        } 

        private void frmBeepPlayer_Load(object sender, EventArgs e)
        {
            midiOutOpen(out midiHandle, 0, 0, 0, 0); // 打開預設合成器

            InitializeInstruments();
            cboInstrument.SelectedValue = 0;
            ChangeInstrument(0);

            this.initWidth = this.palMain.Width;
            this.initHeight = this.palMain.Height;
            foreach (Control ctl in this.palMain.Controls)
            {
                if (!string.IsNullOrEmpty(ctl.Name))
                {
                    this.initControl.Add(ctl.Name, new Rect(ctl.Left, ctl.Top, ctl.Width, ctl.Height));
                }
                /*if (ctl is Button )
                {
                    ctl.Click += (s, args) => {
                        // 利用按鈕的 Tag 來對應頻率索引
                        if (ctl.Tag != null && int.TryParse(ctl.Tag.ToString(), out int idx))
                        {
                            PlayNote(idx);
                        }
                    };
                }*/
                if (ctl is Button btn)
                {
                    if (btn.Tag != null && int.TryParse(btn.Tag.ToString(), out int idx))
                    {
                        // 按鈕顏色與對應索引
                        buttonMap[idx] = btn;
                        originalColors[btn] = btn.BackColor;

                        // MouseDown 變色 + 發聲
                        btn.MouseDown += (s, args) => {
                            HighlightButton(idx, true);
                            PlayNote(idx);
                        };

                        // MouseUp 恢復原色
                        btn.MouseUp += (s, args) => {
                            HighlightButton(idx, false);
                        };

                        btn.PreviewKeyDown += (s, args) => { this.ActiveControl = null; };
                    }
                }
            }
        }

        private void CboInstrument_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboInstrument.SelectedItem is InstrumentItem item)
            {
                ChangeInstrument(item.Id);
                this.ActiveControl = null; // 歸還焦點
            }
        }

        private void CboInstrument_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // 先取得目前的文字內容，並去掉前後空白
                string inputText = cboInstrument.Text.Trim();

                // 如果文字中包含 " - "（例如 "024 - [吉他]..."），就把它切開，只保留前面的數字部分
                if (inputText.Contains("-"))
                {
                    inputText = inputText.Split('-')[0].Trim();
                }

                if (int.TryParse(inputText, out int id) && id >= 0 && id <= 127)
                {
                    cboInstrument.SelectedValue = id; // 自動匹配選單中的完整名稱

                    e.Handled = true;
                    e.SuppressKeyPress = true; // 消除系統警告聲
                    this.ActiveControl = null; // 歸還焦點給主視窗
                    return; // 成功就離開
                }
                MessageBox.Show("請輸入 0 ~ 127 之間的有效數字！");
                e.Handled = true;
            }
        }

        private void TkbOctave_ValueChanged(object sender, EventArgs e)
        {
            offset = tkbOctave.Value * 12;
            this.ActiveControl = null; // 歸還焦點給主視窗，避免上下鍵被滑桿吃掉
        }

        private void frmBeepPlayer_KeyDown(object sender, KeyEventArgs e)
        {
            if (cboInstrument.Focused) return; // 避免在輸入樂器編號時觸發KeyDown
            // 鋼琴
            if (keyMap.ContainsKey(e.KeyCode))
            {
                if (!pressedKeys.Contains(e.KeyCode))
                {
                    pressedKeys.Add(e.KeyCode);
                    int idx = keyMap[e.KeyCode];

                    HighlightButton(idx, true); // 變色
                    PlayNote(idx); // 發聲
                }
                e.Handled = true;
            }

            // 上下方向鍵移調(與 TrackBar 同步且不超過限制)
            if (e.KeyCode == Keys.Up)
            {
                if (tkbOctave.Value < tkbOctave.Maximum) tkbOctave.Value++;
                e.Handled = true;
            }
            if (e.KeyCode == Keys.Down)
            {
                if (tkbOctave.Value > tkbOctave.Minimum) tkbOctave.Value--;
                e.Handled = true;
            }

            // F1~F4 快速切換音色
            if (e.KeyCode == Keys.F1 || e.KeyCode == Keys.F2 || e.KeyCode == Keys.F3 || e.KeyCode == Keys.F4)
            {
                int targetId = 0;
                if (e.KeyCode == Keys.F1) targetId = 0;  // 鋼琴
                if (e.KeyCode == Keys.F2) targetId = 24; // 吉他
                if (e.KeyCode == Keys.F3) targetId = 40; // 提琴
                if (e.KeyCode == Keys.F4) targetId = 56; // 小號

                cboInstrument.SelectedValue = targetId; // 自動觸發 SelectedIndexChanged
                e.Handled = true;
            }
        }

        private void frmBeepPlayer_KeyUp(object sender, KeyEventArgs e)
        {
            if (keyMap.ContainsKey(e.KeyCode))
            {
                // 標記該按鍵已放開
                pressedKeys.Remove(e.KeyCode);

                // 恢復顏色
                int idx = keyMap[e.KeyCode];
                HighlightButton(idx, false);
            }
        }

        private void frmBeepPlayer_SizeChanged(object sender, EventArgs e)
        {
            if (initWidth == 0) return;
            double iRatioWith = (double)this.palMain.Width / this.initWidth;
            double iRatioHeight = (double)this.palMain.Height / this.initHeight;

            foreach (Control ctl in this.palMain.Controls)
            {
                if (initControl.ContainsKey(ctl.Name))
                {
                    ctl.Left = (int)(initControl[ctl.Name].Left * iRatioWith);
                    ctl.Top = (int)(initControl[ctl.Name].Top * iRatioHeight);
                    ctl.Width = (int)(initControl[ctl.Name].Width * iRatioWith);
                    ctl.Height = (int)(initControl[ctl.Name].Height * iRatioHeight);
                }
            }
        }

        private void HighlightButton(int index, bool isPressed)
        {
            if (!buttonMap.ContainsKey(index)) return;
            Button btn = buttonMap[index];

            if (isPressed)
            {
                btn.BackColor = Color.Lavender;
            }
            else
            {
                if (originalColors.ContainsKey(btn))
                {
                    btn.BackColor = originalColors[btn];
                }
            }
        }

        private void frmBeepPlayer_FormClosing(object sender, FormClosingEventArgs e)
        {
            var result = MessageBox.Show("確定要關閉應用程式嗎？", "關閉確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                // 確認關閉釋放資源
                midiOutClose(midiHandle);
            }
        }
    }

    public class InstrumentItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string DisplayName => $"{Id:D3} - {Name}"; // 000-[鋼琴] 大鋼琴 (Acoustic Grand Piano)
    }
}