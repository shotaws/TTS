// オープンモード
var FORREADING      = 1;    // 読み取り専用
var FORWRITING      = 2;    // 書き込み専用
var FORAPPENDING    = 8;    // 追加書き込み

// 開くファイルの形式
var TRISTATE_TRUE       = -1;   // Unicode
var TRISTATE_FALSE      =  0;   // ASCII
var TRISTATE_USEDEFAULT = -2;   // システムデフォルト

// ファイル出力準備
var stream = WScript.CreateObject("SAPI.SpFileStream");
var fs     = new ActiveXObject("Scripting.FileSystemObject");
var wavName = "text.wav";

// ファイル読み込み
var file = fs.OpenTextFile("text.txt", FORREADING, true, TRISTATE_FALSE);
var str = file.ReadAll();
// WScript.Echo(str);

// TTSの設定
var tts    = WScript.CreateObject("SAPI.SpVoice");
tts.Rate = -3; // (-10,10)

// ストリームを開いて
stream.open(wavName, 3);

// TTSの読み上げ先をストリームにセット
tts.AudioOutputStream = stream;

// しゃべる（WAVに書き込みされる）
tts.Speak(str);

// ストリームを閉じる
stream.close();

WScript.Echo("Done!");