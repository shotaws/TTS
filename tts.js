// �I�[�v�����[�h
var FORREADING      = 1;    // �ǂݎ���p
var FORWRITING      = 2;    // �������ݐ�p
var FORAPPENDING    = 8;    // �ǉ���������

// �J���t�@�C���̌`��
var TRISTATE_TRUE       = -1;   // Unicode
var TRISTATE_FALSE      =  0;   // ASCII
var TRISTATE_USEDEFAULT = -2;   // �V�X�e���f�t�H���g

// �t�@�C���o�͏���
var stream = WScript.CreateObject("SAPI.SpFileStream");
var fs     = new ActiveXObject("Scripting.FileSystemObject");
var wavName = "text.wav";

// �t�@�C���ǂݍ���
var file = fs.OpenTextFile("text.txt", FORREADING, true, TRISTATE_FALSE);
var str = file.ReadAll();
// WScript.Echo(str);

// TTS�̐ݒ�
var tts    = WScript.CreateObject("SAPI.SpVoice");
tts.Rate = -3; // (-10,10)

// �X�g���[�����J����
stream.open(wavName, 3);

// TTS�̓ǂݏグ����X�g���[���ɃZ�b�g
tts.AudioOutputStream = stream;

// ����ׂ�iWAV�ɏ������݂����j
tts.Speak(str);

// �X�g���[�������
stream.close();

WScript.Echo("Done!");