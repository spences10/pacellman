Attribute VB_Name = "Module1"

'*********************************************************
'  PACELL-MAN
'          for Excel97,2000
'                                By N.Chikada
'Ver001  2003/03/16
'Ver004　2003/03/21
'Ver006  2003/03/22
'Ver008 モンスターの移動ルーチン反転対応　2H
'Ver009　モンスター移動時のドット非消去　１H
'Ver010  パッセルのドット消去判定、反転時１ドット欠問題対応　2H
'Ver011  モンスターの状態、出現、出動処理,ゲーム終了判定　2H
'Ver012　 いじけモードのモンスター移動、反転可能処理　2H
'Ver013  全体の速度調整、モード切替パラメーター化、いじけ時減速　1H
'Ver014　目玉処理　2H
'Ver015  パワーエサ表示、食べたときのいじけ反応　1H
'Ver017　ワープトンネル作成、通過時モンスター速度半減　2H
'Ver018　経過時間処理（移動スピード、面ごとの加速、いじけの有無、モード切替　2H
'Ver019　反転処理見直し、強制反転、自動反転　2H
'Ver020　面ごとのモード変更時間、いじけ時の移動方向見直し　1H
'Ver020　当たり判定、食い点数表示　1H
'Ver021　モンスター食い時の目玉以外凍結処理　2H
'Ver024　パッセル消去処理、残機表示、ゲームオーバー判定　2H
'Ver025　面クリアー処理　1H
'Ver026　スコア高速表示、ハイスコア表示、表示有無OPTION 1H
'Ver027  文字列描画関数、Ready!の消去、GAMEOVERの表示 1H
'Ver028　効果音、同時再生、ループ再生　3H
'Ver029　OPENINGデモ,デモ１、デモ２、その他微調整　4H
'Ver030　恒常ループ化、 Wait導入、最適値設定　2H
'Ver031　そしてリリースへ。
'Ver032　説明追加、恒常ループ見直し
'Ver033　2003/8/21　自機数バグ修正、スコアバグ修正
'*********************************************************
Option Explicit
Option Base 1
Declare Function GetTickCount Lib "kernel32" () As Long     'Windows起動後経過時間取得API
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand _
      As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Const VK_LEFT As Long = 37
Const VK_DOWN As Long = 40
Const VK_RIGHT As Long = 39
Const VK_UP As Long = 38
Const VK_Z As Long = 90
Const VK_SHIFT As Long = 16
'　←：37　↑：38　→：39　↓：40
Public Game_Flag As Boolean
Const SpeedB As Long = 256 'スピードの基準値（通常）
'ゲームカウンタ
Private Count As Long
Private Timing1(1 To 21, 1 To 9) As Long   '縄張り追跡モードの切り替え時間データ
Private Timing2(1 To 16) As Long  'スピードの切替経過時間　データ
Private TimingNo(1 To 2) As Long
Private Men(1 To 2) As Long '1:面数 2：21面以降は２１
Private i As Long
Private i2 As Long
Private i3 As Long
Private i4 As Long
Private i5 As Long

Private Pac_Y As Long  '位置座標
Private Pac_X As Long
Private Pac_S As Long '状態
Private Mon_Y(1 To 4) As Long  '位置座標
Private Mon_X(1 To 4) As Long
Private Mon_S(1 To 4) As Long '状態
Private Mon_JT(1 To 4) As Integer 'モンスターの状態（0:登場　1:出動　2:追跡　3:縄張　4：目玉　５:点数　）
Private Black(0 To 1) As Range
Private DotCount As Long 'ドットの数
Private PowCount As Long 'パワーエサの数
Private MonSyutsugenC(1 To 4) As Integer  'モンスター出現時の往復回数カウント
Private Pac_KS As Long  'パックのキー入力方向
Private Mon_SPDBuf(1 To 4) As Long  'モンスターの移動速度バッファ
Private Mon_IdoF(1 To 4) As Boolean 'モンスター移動可否フラグ
Private HantenF(1 To 4) As Boolean     'モンスターの反転可能サイン（追跡と待機や、いじけと通常の切替時にTrue）
Private Score(0 To 1) As Long '1は表示のスコア
Private HiScore As Long
Private WaitNo As Long
Private Pac_No As Long

Private Mon_TenC0(1 To 4) As Long 'モンスターの食いの点数表示カウントスタート
Private Mon_TenC1(1 To 4) As Long 'モンスターの食いの点数表示カウント
Private Mon_TenC2(1 To 4) As Boolean 'モンスターの食いの点数表示カウント最終黒表示
Private Mon_TenF As Boolean '点数表示中フラグ
Private Score_data As Range
Private Temp As Variant
Private strStatus As String 'ループ再生判定用
Private CoinF As Boolean 'コインF
'画像データ
Private Map 'As Variant   '移動判定用MAPデータ（ColorIndex）
Private Map2 'As Variant   'ドットデータ作業用配列
Private Pow_S(1 To 4) As Boolean  'パワーエサ用ステータス
Private PowC As Long '１：表示　２：黒
Private PowCN As Long  '表示判定用カウンタ
Private Mon_Ijike(1 To 4) As Boolean 'いじけモードの時True
Private Mon_IjikeSt As Long 'いじけスタートタイム
Private Mon_IjikeS As Long 'いじけの状態　１：通常　２：点滅青　３：点滅白
Private Moji As Range
Private SoundF As Boolean
Private WaitStart As Long
Private Up(1 To 3) As Boolean

Sub Start()
   Dim TempStartTime As Long
   SyokiSettei
   
   On Error GoTo ErrH   'エラー発生時のエラーハンドル
   Application.EnableCancelKey = xlErrorHandler 'ESC押下時にエラーハンドルへ
   
   '音再生エイリアス作成
   For i = 0 To 11
      FileOpen LTrim(i) + ".wav", i
   Next
   
   strStatus = String$(256, " ")  'ループ判定用文字列に空白を代入
  
   Randomize
   Game_Flag = True
   
   TimingNo(1) = 1 '追跡縄張りのモード切替タイミング
   TimingNo(2) = 1  'スピードの切り替えタイミング
   Men(1) = 1
   Men(2) = 1
   Up(1) = False

   Set Black(0) = Cells(827, 98)
   Set Black(1) = Cells(827, 114)

   Dim DotData As Range
   Set DotData = Range(Cells(720, 16), Cells(721, 17))

   'パッセルマンのデータ
   Dim Pac_data0 As Range
   Dim Pac_data1 As Range
   Dim Pac_data2 As Range
   Set Pac_data0 = Cells(602, 2)
   Dim Pow_data(1 To 2) As Range
   Set Pow_data(1) = Range(Cells(826, 5), Cells(833, 12))
   Set Pow_data(2) = Range(Cells(826, 5), Cells(833, 12)).Offset(0, 16)
   Dim Pow_Ichi(1 To 4) As Range
   Dim Pac_Syokyo As Range
   Set Pac_Syokyo = Range(Cells(877, 1), Cells(891, 15))

   Dim Pow_Y(1 To 4) As Long  'あたり判定用パワーエサ座標
   Dim Pow_X(1 To 4) As Long  'あたり判定用パワーエサ座標

   Dim Pac_SY As Long
   Dim Pac_SX As Long
   Dim Pac_SP As Long  'パックのパターン
   Dim Pac_SPP As Long  'パックのパターン子
   Dim Pac_SPD As Long 'パッセルの移動速度
   Dim Pac_SPDBuf As Long  'パッセルの移動速度バッファ
   Dim Pac_IdoF As Boolean '移動速度による移動可否フラグ
   Dim Pac_dotZ As Boolean 'ドットを食べたときTrue
   

  'モンスターデータ
   Dim Mon_data0(1 To 6) As Range   '5は、いじけ、目玉モード用　6は点数
   Dim Mon_data1(1 To 4) As Range
   Dim Mon_data2(1 To 4) As Range
   For i = 1 To 5
      Set Mon_data0(i) = Cells((i - 1) * 16 + 618, 2)
   Next i
   Set Mon_data0(6) = Cells(826, 34)

   Dim Mon_SY(1 To 4) As Long
   Dim Mon_SX(1 To 4) As Long
   Dim Mon_SP(1 To 4) As Long  'パックのパターン
   Dim Mon_SPP(1 To 4) As Long  'パックのパターン子
   Dim Mon_KS(1 To 4) As Long  'パックのキー入力方向

   Dim Mon_SPD(1 To 4) As Long 'モンスターの移動速度
   
   Dim Mon_TempY(1 To 4) As Long
   Dim Mon_TempX(1 To 4) As Long
   Dim Mon_KabeF(1 To 4, 1 To 4) As Boolean  'モンスター方向の壁有無（番号、方向）
   Dim Mon_Kabe(1 To 4) As Long  'モンスターの移動可能方向数
   Dim Mon_Kyori(1 To 4, 1 To 4) As Long 'パッセルマンとモンスターの距離（番号、方向）
   Dim Mon_KyoriL(1 To 4) As Long  '判定収納用

   Dim Mon_IjikeC As Long 'いじけ時間経過タイム
   Dim Mon_IjikeD(1 To 21, 1 To 3) As Long  '1:第一期（青）2:第二期（点滅）3:点滅間隔
   Dim Renzoku_Ten(1 To 4) As Long
   Renzoku_Ten(1) = 200
   Renzoku_Ten(2) = 400
   Renzoku_Ten(3) = 800
   Renzoku_Ten(4) = 1600
   
   
   Dim Mon_NY(1 To 4) As Integer 'モンスターの縄張り座標=待機中の基準値
   Dim Mon_NX(1 To 4) As Integer 'モンスターの縄張り座標=待機中の基準値
   Mon_NY(1) = 45
   Mon_NY(2) = 45
   Mon_NY(3) = 205
   Mon_NY(4) = 205
   Mon_NX(1) = 221
   Mon_NX(2) = 21
   Mon_NX(3) = 221
   Mon_NX(4) = 21
   
   Dim Mon_NY2(1 To 4) As Integer 'モンスターの縄張り座標=待機中の基準値
   Dim Mon_NX2(1 To 4) As Integer 'モンスターの縄張り座標=待機中の基準値

   Mon_NY2(1) = 17
   Mon_NY2(2) = 17
   Mon_NY2(3) = 208
   Mon_NY2(4) = 208
   Mon_NX2(1) = 224
   Mon_NX2(2) = 17
   Mon_NX2(3) = 224
   Mon_NX2(4) = 22
   
   'あらかじめパワーエサ表示座標を取得しておく
   For i = 1 To 4
      Set Pow_Ichi(i) = Range(Cells(Mon_NY(i) + 4, Mon_NX(i) + 4), Cells(Mon_NY(i) + 11, Mon_NX(i) + 11))
      Pow_Y(i) = Mon_NY(i) + 1
      Pow_X(i) = Mon_NX(i) + 1
   Next i
   
   'Dim Mon_OtobokeF As Boolean  'おとぼけ橙用乱数決定フラグ
   
   Set Score_data = Range(Cells(893, 1), Cells(899, 7))

   Dim IdoH_Y1(0 To 4) As Long  '移動判定Y開始座標
   Dim IdoH_Y2(0 To 4) As Long  '移動判定Y終了座標
   Dim IdoH_X1(0 To 4) As Long  '移動判定X開始座標
   Dim IdoH_X2(0 To 4) As Long  '移動判定Y終了座標
      IdoH_Y1(1) = -1
      IdoH_Y2(1) = 14
      IdoH_X1(1) = -2
      IdoH_X2(1) = -2
      IdoH_Y1(2) = -1
      IdoH_Y2(2) = 14
      IdoH_X1(2) = 15
      IdoH_X2(2) = 15
      IdoH_Y1(3) = -2
      IdoH_Y2(3) = -2
      IdoH_X1(3) = -1
      IdoH_X2(3) = 14
      IdoH_Y1(4) = 15
      IdoH_Y2(4) = 15
      IdoH_X1(4) = -1
      IdoH_X2(4) = 14
      
   '移動方向による＋座標
   Dim Ido_Y(0 To 4) As Long
   Dim Ido_X(0 To 4) As Long
      Ido_Y(0) = 0
      Ido_Y(1) = 0
      Ido_Y(2) = 0
      Ido_Y(3) = -1
      Ido_Y(4) = 1
      Ido_X(0) = 0
      Ido_X(1) = -1
      Ido_X(2) = 1
      Ido_X(3) = 0
      Ido_X(4) = 0
      
   Dim Hantai(1 To 4) As Long '進行逆方向
   Hantai(1) = 2
   Hantai(2) = 1
   Hantai(3) = 4
   Hantai(4) = 3
   
   Dim KijunY As Integer  'モンスター移動判定時の基準座標Ｙ
   Dim KijunX As Integer  'モンスター移動判定時の基準座標Ｘ
   
   'モンスター移動時のドットの有無フラグ
   Dim MonDotF(1 To 4) As Boolean
   For i = 1 To 4
      MonDotF(i) = True
   Next i
   '追跡と待機の切替タイミング（パッセルマンの移動カウントを基準）
   'Dim HantenTime As Long
   'HantenTime = 3000
   
   'モンスタードットの有無判定相対座標（１左移動時２右３上４下）
   Dim DotHanteiY(1 To 4) As Integer
   Dim DotHanteiX(1 To 4) As Integer
   DotHanteiY(1) = 7
   DotHanteiY(2) = 7
   DotHanteiY(3) = 14
   DotHanteiY(4) = -1
   DotHanteiX(1) = 14
   DotHanteiX(2) = -1
   DotHanteiX(3) = 7
   DotHanteiX(4) = 7
      
   'パッセルドットイートの有無判定相対座標（１左移動時２右３上４下、1前進２後退）
   Dim P_DotHanteiY(0 To 4, 1 To 2) As Integer  'パッセル停止中は０もありえるので便宜上0も入れる
   Dim P_DotHanteiX(0 To 4, 1 To 2) As Integer
   P_DotHanteiY(1, 1) = 6
   P_DotHanteiY(2, 1) = 6
   P_DotHanteiY(3, 1) = 0
   P_DotHanteiY(4, 1) = 12
   P_DotHanteiX(1, 1) = 0
   P_DotHanteiX(2, 1) = 12
   P_DotHanteiX(3, 1) = 6
   P_DotHanteiX(4, 1) = 6
   P_DotHanteiY(1, 2) = 6
   P_DotHanteiY(2, 2) = 6
   P_DotHanteiY(3, 2) = 13
   P_DotHanteiY(4, 2) = -1
   P_DotHanteiX(1, 2) = 13
   P_DotHanteiX(2, 2) = -1
   P_DotHanteiX(3, 2) = 6
   P_DotHanteiX(4, 2) = 6
   'ドット判定作業用座標
   Dim DotY As Long
   Dim DotX As Long
   Dim P_Dot As Boolean '方向転換時後ろに1ドット描き加えるフラグ
   Dim SoundSF As Boolean
   Dim Speed(1 To 16, 1 To 4) As Long  '経過時間によるモンスタースピード
   Dim Speed2(1 To 21, 1 To 4) As Long  '面ごとに加算する速度
   Dim Mode(1 To 16, 1 To 4) As Long   '経過毎の追跡縄張データ
   Dim ModeUM(1 To 21, 1 To 4) As Long 'モード変更の有無
   Dim IjikeUM(1 To 21) As Long '　　　　’いじけの有無
   Dim MonSyutsugen(1 To 4) As Integer 'モンスター出現時の往復回数
   Dim Mon_S0(1 To 4) As Long '   直前の移動方向  反転時に使用
   Dim HantenUM(1 To 21, 1 To 4) As Long  '面ごとの反転有無設定
   Dim Renzoku(0 To 4) As Long 'モンスターの連続食い回数
   Dim TenTime As Long
   Dim HiScoreF As Long 'ハイスコア表示オプション

   'ワークシートよりデータ取得
   Dim MySheet As Worksheet
   Set MySheet = Worksheets("Config")
      Pac_SPD = MySheet.Cells(5, 3)  'パッセルのスピード
      MonSyutsugen(1) = MySheet.Cells(14, 3)
      MonSyutsugen(2) = MySheet.Cells(14, 4)
      MonSyutsugen(3) = MySheet.Cells(14, 5)
      MonSyutsugen(4) = MySheet.Cells(14, 6)
      TenTime = MySheet.Cells(9, 7)
      Pac_No = MySheet.Cells(8, 17)
      HiScoreF = MySheet.Cells(9, 17)
      If MySheet.Cells(10, 17) = 1 Then
         SoundF = True
      Else
         SoundF = False
      End If
      WaitNo = MySheet.Cells(11, 17)
      
      For i = 1 To 21
         For i2 = 1 To 3
            Mon_IjikeD(i, i2) = MySheet.Cells(i + 63, i2 + 17)
         Next i2
      Next i

      For i = 1 To 16
         For i2 = 1 To 4
            Mode(i, i2) = MySheet.Cells(i + 39, i2 + 2)
         Next i2
      Next i
      
      For i = 1 To 16
         Timing2(i) = MySheet.Cells(i + 20, 1)
         For i2 = 1 To 4
            Speed(i, i2) = MySheet.Cells(20 + i, i2 + 2)
         Next i2
      Next i
            
      For i = 1 To 21
         For i2 = 1 To 9
            Timing1(i, i2) = MySheet.Cells(i + 63, i2 + 38)
         Next i2
      Next i
      
      For i = 1 To 21
         IjikeUM(i) = MySheet.Cells(i + 63, 22)
         For i2 = 1 To 4
            HantenUM(i, i2) = MySheet.Cells(i + 63, i2 + 7)
            Speed2(i, i2) = MySheet.Cells(i + 63, i2 + 2)
         Next i2
      Next i

   Do While CoinF = False
      Demo2
      Opening
   Loop

   'ゲームスタート
   Sound 0 '開始の音
   MenKaishi
   ZankiHyoji (Pac_No)
   '速度計測実験用
   TempStartTime = GetTickCount
   '////////MainLoop///////
   CoinF = False
   
   Do While Game_Flag = True
      '速度による移動判定
      Pac_SPDBuf = Pac_SPDBuf + Pac_SPD
      If Pac_SPDBuf >= SpeedB Then
         Pac_SPDBuf = Pac_SPDBuf - SpeedB
         Pac_IdoF = True
      Else
         Pac_IdoF = False
      End If
      
      'DoEvents
      
      '----------------------
      'パッセルマンの移動判定
      '----------------------
      If Pac_IdoF = True Then
         'ゲームの基本となるカウントを加算
         Count = Count + 1 'パッセルマンが１移動する度に＋１
         Do While GetTickCount - WaitStart < WaitNo
         Loop
         WaitStart = GetTickCount
         '速度計測実験
         'If Count = 300 Then
         '   MsgBox GetTickCount - TempStartTime
         'End If
         
         'スピード切り替えタイミングの判定
         If Count = Timing2(TimingNo(2) + 1) Then
            TimingNo(2) = TimingNo(2) + 1
         End If
         
         '縄張り追跡モードの判定
         If Count = Timing1(Men(2), TimingNo(1) + 1) Then
            TimingNo(1) = TimingNo(1) + 1
            'モンスターのモード切替（追いかけと待機）
            For i = 1 To 4
               If Mon_JT(i) = 2 Or Mon_JT(i) = 3 Then '状態が追跡、縄張りの時、
                  Mon_JT(i) = Mode(TimingNo(1), i)
                  If Not (Mon_Ijike(i)) Then 'かつ、いじけてないとき反転可
                     If HantenUM(Men(2), i) = 1 Then  '面ごとモンスターごとの反転有無
                        HantenF(i) = True
                     End If
                  End If
               End If
            Next i
         End If
         
         'パワーエサの表示
         PowCN = Count Mod 10
            
         If PowCN = 0 Then
            If PowC = 1 Then
               PowC = 2
            Else
               PowC = 1
            End If
         ElseIf PowCN <= 4 Then
            If Pow_S(PowCN) Then 'パワーエサが存在すれば
               Pow_data(PowC).Copy Destination:=Pow_Ichi(PowCN)
            End If
         End If
           
         If Not (Mon_TenF) Then '点数表示中は凍結
 
                
            If GetAsyncKeyState(VK_LEFT) <> 0 Then '左
               Pac_KS = 1
              ElseIf GetAsyncKeyState(VK_RIGHT) <> 0 Then   '右
               Pac_KS = 2
              ElseIf GetAsyncKeyState(VK_UP) <> 0 Then   '上
               Pac_KS = 3
              ElseIf GetAsyncKeyState(VK_DOWN) <> 0 Then   '下
               Pac_KS = 4
              Else
              Pac_KS = 0
            End If
            
            '進行方向の決定
            If Pac_KS <> 0 Then 'キー入力があった場合
               If Pac_KS <> Pac_S Then 'かつ、キー入力が現在の移動方向と違う場合
                  '移動可否判定
                  For i = IdoH_Y1(Pac_KS) To IdoH_Y2(Pac_KS)
                     For i2 = IdoH_X1(Pac_KS) To IdoH_X2(Pac_KS)
                     If Map(Pac_Y + i, Pac_X + i2) = 5 Then
                        Pac_KS = Pac_S '壁がある場合は現在の進行方向
                        Exit For
                     End If
                     Next i2
                  Next i
               End If
            End If
            
            '移動方向の決定
            If Pac_KS <> 0 Then 'キー入力方向に壁がない場合
               Pac_S = Pac_KS
            End If
            
            '移動可否チェック
            If Pac_S <> 0 Then '移動する場合
               '移動可否判定
               For i = IdoH_Y1(Pac_S) To IdoH_Y2(Pac_S)
                  For i2 = IdoH_X1(Pac_S) To IdoH_X2(Pac_S)
                  If Map(Pac_Y + i, Pac_X + i2) = 5 Then
                     Pac_S = 0 '壁がある場合は移動不可
                     Exit For
                  End If
                  Next i2
               Next i
            End If
            
            '決定した移動方向により座標を収納
            Pac_SY = Ido_Y(Pac_S)
            Pac_SX = Ido_X(Pac_S)
               
            'パターン判定
            If Pac_S <> 0 Then
               Pac_SPP = Pac_SPP + 1
               If Pac_SPP > 4 Then
                  Pac_SPP = 1
               End If
               
               If Pac_SPP = 1 Then
                  Pac_SP = Pac_SP + 1
                  If Pac_SP > 4 Then
                     Pac_SP = 1
                  End If
               End If
            End If
            
            '移動描写
            Pac_Y = Pac_Y + Pac_SY
            Pac_X = Pac_X + Pac_SX
                     
            'ワープトンネル判定
            If Pac_X < 3 Then
               Pac_X = 241
            ElseIf Pac_X > 241 Then
               Pac_X = 3
            End If
            
            'ドット消去処理
            DotY = Pac_Y + P_DotHanteiY(Pac_S, 1)
            DotX = Pac_X + P_DotHanteiX(Pac_S, 1)
            
            If Map(DotY, DotX) = 40 Then
               'ドットあり→ドットの塊を走査開始
               If Pac_S = 1 Or Pac_S = 2 Then '横移動のとき
                  For i = -1 To 1 '上下1ドットを０に
                     Map2(DotY + i, DotX) = 0
                  Next i
                  'ドット左右のMAPデータが０の場合
                  If Map2(DotY, DotX + 1) <> 40 And Map2(DotY, DotX - 1) <> 40 Then
                     'ドットイートと決定、ドットデータ消去
                     DotCount = DotCount - 1
                     Pac_dotZ = True

                     'MsgBox Map2(DotY, DotX - 1) & Map(DotY, DotX + 1)       'ドットカウンター減
                     For i = -1 To 1
                        For i2 = -1 To 1
                           Map(DotY + i, DotX + i2) = 0
                        Next i2
                     Next i
                  End If
              
              ElseIf Pac_S = 3 Or Pac_S = 4 Then '縦移動のとき
                  For i = -1 To 1
                     Map2(DotY, DotX + i) = 0
                  Next i
                  'ドット上下のMAPデータが０の場合
                  If Map2(DotY + 1, DotX) <> 40 And Map2(DotY - 1, DotX) <> 40 Then
                     'ドットイートと決定、ドットデータ消去
                     DotCount = DotCount - 1 'ドットカウンター減
                     Pac_dotZ = True
                     For i = -1 To 1
                        For i2 = -1 To 1
                           Map(DotY + i, DotX + i2) = 0
                        Next i2
                     Next i
                  End If
               End If
      
            End If
            
            '後ろに1ドット加えるか判定(前進して１ドット消去後方向転換したときに、ドットを復活させる)
            'ドット消去処理
            DotY = Pac_Y + P_DotHanteiY(Pac_S, 2)
            DotX = Pac_X + P_DotHanteiX(Pac_S, 2)
            If Map(DotY, DotX) = 40 Then
               'ドットあり→ドットの塊を走査開始
               P_Dot = True
            Else
               P_Dot = False
            End If
            
            '移動書き換え
            If Pac_S <> 0 Then  'パックが移動できる時
               Set Pac_data1 = Pac_data0.Offset((Pac_S = 4) - (P_Dot * 209), (Pac_S = 2) + (Pac_S - 1) * 64 + (Pac_SP - 1) * 16). _
                  Resize(13 + Abs(Pac_SY), 13 + Abs(Pac_SX))
               Set Pac_data2 = Cells(Pac_Y, Pac_X). _
                  Offset((Pac_S = 4), (Pac_S = 2)).Resize(13 + Abs(Pac_SY), 13 + Abs(Pac_SX))
            End If
            Pac_data1.Copy Destination:=Pac_data2
           
            'ドット食いのスコア、音処理
            If Pac_dotZ Then
               Pac_dotZ = False
               Score(0) = Score(0) + 10
               Score_hyoji
               '音
               'PlaySound ActiveWorkbook.Path & "\" & DotCount Mod 2 + 2 & ".wav", 0, &H11
               Sound DotCount Mod 2 + 2
            End If
          
            '　いじけ ここにパワー餌判定を入れる
                  'パワーエサ喰い判定
            For i = 1 To 4
               If Pac_Y = Pow_Y(i) Then
                  If Pac_X = Pow_X(i) Then
                     If Pow_S(i) Then
                        'パワーエサ喰い
                        '反転フラグをTrue
                        For i2 = 1 To 4
                           If Mon_JT(i2) < 4 Then
                              HantenF(i2) = True
                           End If
                        Next i2
                        Pow_S(i) = False
                        PowCount = PowCount - 1
                        Pow_data(2).Copy Destination:=Pow_Ichi(i)
                        Renzoku(0) = 0
                        If IjikeUM(Men(2)) = 1 Then 'いじけありの面の場合
                           For i2 = 1 To 4
                              'いじけフラグをTrue
                              If Mon_JT(i2) < 4 Then
                                  Mon_Ijike(i2) = True
                              End If
                           Next i2
                           Mon_IjikeSt = Count 'いじけスタートタイム
                           Mon_IjikeS = 1
                           Exit For
                        End If
                     End If
                  End If
               End If
            Next i
         End If
'         If GetAsyncKeyState(16) <> 0 Then 'Shift仮置き実験用
'            For i = 1 To 4
'               'いじけフラグをTrue
'                Mon_Ijike(i) = True
'               '反転フラグをTrue
'               HantenF(i) = True
'
'            Next i
'            Mon_IjikeSt = Count 'いじけスタートタイム
'            Mon_IjikeS = 1
'         End If
         
         'いじけ経過時間処理
         If Mon_IjikeS = 1 Then
            If Count - Mon_IjikeSt > Mon_IjikeD(Men(2), 1) Then
               Mon_IjikeS = 2
               Mon_IjikeC = 0
            End If
         End If
         If Mon_IjikeS >= 2 Then
            Mon_IjikeC = Mon_IjikeC + 1
            If Mon_IjikeC Mod Mon_IjikeD(Men(2), 3) = 0 Then
               Select Case Mon_IjikeS
                  Case 2
                     Mon_IjikeS = 3
                  Case 3
                     Mon_IjikeS = 2
               End Select
            End If
            
            If Count - Mon_IjikeSt > Mon_IjikeD(Men(2), 2) Then
               'いじけの終了
               For i = 1 To 4
                  'If Mon_Ijike(i) = True Then  'まだ、いじけ中のモンスターのみ
                     '反転フラグをTrue
                     'HantenF(i) = True ’いじけ終了時には反転できない気がする
                     'いじけフラグをFalse
                     Mon_Ijike(i) = False
                  'End If

               Next i
               For i = 1 To 4
                  'モードを該当の追跡or待ち伏せに変更
                  If Mon_JT(i) > 0 And Mon_JT(i) < 4 Then '目玉､点数､出現はそのまま
                     Mon_JT(i) = Mode(TimingNo(1), i)
                  End If
               
               Next i
               
               Mon_IjikeSt = 0
               Mon_IjikeS = 0
            End If
         End If
      End If
     
     'モンスターの移動処理
     
      For i = 1 To 4
         
         'いじけ中はモンスターのスピードを半減
         If Mon_Ijike(i) Then
            Mon_SPD(i) = (Speed(TimingNo(2), i) + Speed2(Men(2), i)) / 2
         ElseIf Mon_JT(i) = 4 Then '目玉の時はＭａｘスピード
            Mon_SPD(i) = SpeedB

         ElseIf Mon_JT(i) = 5 Then  '点数のときはPACと同じ
             'モンスターのスピード収納
            Mon_SPD(i) = Pac_SPD
         Else
            Mon_SPD(i) = Speed(TimingNo(2), i) + Speed2(Men(2), i)
         End If
         
         'ワープトンネル中の例外処理
         If Mon_JT(i) < 4 Then  '追跡、縄張りのとき
            If Mon_S(i) = 1 Then 'ワープトンネル中は半減
               If Mon_Y(i) = 134 Then
                  If Mon_X(i) < 48 Or Mon_X(i) > 184 Then
                     Mon_SPD(i) = (Speed(TimingNo(2), i) + Speed2(Men(2), i)) / 2
                  End If
               End If
            ElseIf Mon_S(i) = 2 Then
               If Mon_Y(i) = 134 Then
                  If Mon_X(i) < 62 Or Mon_X(i) > 196 Then
                     Mon_SPD(i) = (Speed(TimingNo(2), i) + Speed2(Men(2), i)) / 2
                  End If
               End If
            End If
         End If
         
         '速度による移動判定
         Mon_SPDBuf(i) = Mon_SPDBuf(i) + Mon_SPD(i)
         If Mon_SPDBuf(i) >= SpeedB Then
            Mon_SPDBuf(i) = Mon_SPDBuf(i) - SpeedB
            Mon_IdoF(i) = True
         Else
            Mon_IdoF(i) = False
         End If
      
         If Mon_IdoF(i) = True Then
            If Mon_JT(i) = 5 Then '点数のときは動かず、この処理のみ
               If Mon_TenC0(i) = 0 Then
                  Mon_TenC0(i) = Count
                  Mon_TenC1(i) = Count
               End If
               Mon_TenC1(i) = Mon_TenC1(i) + 1
               
               If Mon_TenC1(i) - Mon_TenC0(i) + 1 > TenTime Then
                  Mon_TenC2(i) = True
                  If Mon_TenC1(i) - Mon_TenC0(i) > TenTime Then
                     Mon_TenC1(i) = 0
                     Mon_TenC0(i) = 0
                     Mon_TenC2(i) = False
                     Mon_JT(i) = 4
                     Mon_TenF = False
                  End If
               End If
            Else
            
              'パターンの決定
              'パターン判定
               If Not (Mon_TenF) Then '点数表示中は凍結
                  If Mon_S(i) <> 0 Then
                     Mon_SPP(i) = Mon_SPP(i) + 1
                     If Mon_SPP(i) > 4 Then
                        Mon_SPP(i) = 1
                     End If
                     
                     If Mon_SPP(i) = 1 Then
                        Mon_SP(i) = Mon_SP(i) + 1
                        If Mon_SP(i) > 2 Then
                           Mon_SP(i) = 1
                        End If
                     End If
                  End If
               End If
              'テスト移動ルーチン
           '   If i = 1 Then  '赤
                 'If Mon_X(i) < 71 Then
                 '   Mon_S(i) = 2
                 'ElseIf Mon_X(i) > 140 Then
                 '   Mon_S(i) = 1
                 'End If
                 'パッセルマンを最短距離で追いかける
              
               If Not (Mon_TenF) Or Mon_JT(i) = 4 Then '点数表示中かつ目玉でないとき凍結
                  If Mon_JT(i) = 0 Then '登場時
                     If Mon_Y(i) < 131 Then
                       MonSyutsugenC(i) = MonSyutsugenC(i) + 1
                         If MonSyutsugenC(i) > MonSyutsugen(i) Then
                            Mon_JT(i) = 1
                         Else
                            Mon_S(i) = 4
                         End If
                      ElseIf Mon_Y(i) > 138 Then
                       Mon_S(i) = 3
                      End If
             
                   ElseIf Mon_JT(i) >= 1 And Mon_JT(i) <= 4 Then  '出動、追跡、縄張
                       'まず現在の状況を確認
                       For i2 = 1 To 4  '進行方向反対以外の道を走査
                          If i2 = Hantai(Mon_S(i)) Then
                             Mon_KabeF(i, i2) = False  '通常時はバック出来ないのでFALSEに
                          Else
                             Mon_KabeF(i, i2) = True  '一旦壁フラグをTrueに
                          End If
                       Next i2
                       
                       For i2 = 1 To 4  '進行方向反対以外の道を走査
                          If i2 <> Hantai(Mon_S(i)) Then
                             For i3 = IdoH_Y1(i2) To IdoH_Y2(i2)
                                For i4 = IdoH_X1(i2) To IdoH_X2(i2)
                                   If Map(Mon_Y(i) + i3, Mon_X(i) + i4) = 5 Then
                                      Mon_KabeF(i, i2) = False 'Trueが壁なし
                                      Exit For
                                   End If
                                Next i4
                             Next i3
                          End If
                       Next i2
                                        
                       'モンスターが目玉でゲート真上の時はゲートを開く
                       If Mon_JT(i) = 4 And Mon_JT(i) <> 0 Then 'モンスターが目玉　かつ待機中でない
                          If Mon_Y(i) = 110 And Mon_X(i) = 122 Then 'ゲート真上
                             Mon_KabeF(i, 4) = -1
                          End If
                       End If
                       'モンスターが出動中でゲート真下にいるときはをゲートを開く
                       If Mon_JT(i) = 1 Then 'モンスターが出勤中
                          If Mon_Y(i) < 131 And Mon_Y(i) > 126 And Mon_X(i) = 122 Then 'ゲート真上
'                          If Mon_Y(i) < 107 And Mon_X(i) = 122 Then 'ゲート真上
                             Mon_KabeF(i, 3) = -1
                          End If
                       End If
      
                       For i2 = 1 To 4
                          Mon_Kabe(i) = Mon_Kabe(i) + Mon_KabeF(i, i2)
                       Next i2
                       
                       If Mon_Kabe(i) = -1 Then   '一本道
                          For i2 = 1 To 4
                             If Mon_KabeF(i, i2) = True Then
                                Mon_S(i) = i2  '一つしかない方向に決定
                             End If
                          Next i2
                       ElseIf Mon_Kabe(i) < -1 Then    '分かれ道 ここで初めて移動思考判定
                          '0: 登場 1: 出動 2: 追跡 3: 縄張 4: 目玉
                       
                          'いじけ中はパッセルと逆方向
                          If Mon_Ijike(i) Then
                             KijunY = Mon_Y(i) - Pac_Y + Mon_Y(i)
                             KijunX = Mon_X(i) - Pac_X + Mon_X(i)
                          
                          Else
                             'モンスター毎、モード毎の移動方向決定判定を入れる
                             Select Case Mon_JT(i)   'モンスターの状態
                                Case 1  '出動
                                   KijunY = 10
                                   KijunX = 121
                                   If Mon_Y(i) < 112 Then  '通路に出きったら追跡へ
        
                                      Mon_JT(i) = Mode(TimingNo(1), 2)  'パラメータ設定に戻す
                                   End If
                                Case 2  '追跡
                                   Select Case i  'モンスターの種類
                                      Case 1 'おいかけ赤(Pの座標)
                                         KijunY = Pac_Y
                                         KijunX = Pac_X
                                      Case 2  'まちぶせ桃（P進行方向3キャラ前座標）
                                         KijunY = Pac_Y + Ido_Y(Pac_S) * 48
                                         KijunX = Pac_X + Ido_X(Pac_S) * 48
                                      Case 3  'きまぐれ青（Pを中心とした赤との点対称座標）
                                         KijunY = Pac_Y + (Pac_Y - Mon_Y(1))
                                         KijunX = Pac_X + (Pac_X - Mon_X(1))
                                      Case 4  'おとぼけ橙（Pとの距離が130以上のとき赤、それ以外ランダム）
                                         If Abs(Mon_Y(i) - Pac_Y) ^ 2 + Abs(Mon_X(i) - Pac_X) ^ 2 >= 130 ^ 2 Then
                                            KijunY = Pac_Y
                                            KijunX = Pac_X
                                         Else
                                            KijunY = Mon_Y(i) + Int(Rnd * 2) * 2 - 1
                                            KijunX = Mon_X(i) + Int(Rnd * 2) * 2 - 1
                                         End If
                                         
                                    End Select
                                
                                Case 3 '縄張　（橙に縄張はない）
                                   KijunY = Mon_NY2(i)
                                   KijunX = Mon_NX2(i)
                                
                                Case 4 '目玉
                                   KijunY = 133
                                   KijunX = 122
                                   If Mon_Y(i) = 132 And Mon_X(i) = 122 Then '巣に到着したらモードを出動へ
                                      If Mon_JT(i) <> 0 Then  '登場中でない場合
                                         Mon_JT(i) = 1
                                         'If i <> 4 Then
                                         '   Mon_SP(i) = mod(TimingNo(1), 2)
                                         'End If
                                      End If
                                   End If
                                   
                                'Case 5 '復活
                                
                             End Select
                          End If
                          
                          Mon_KyoriL(i) = 133120   'とりあえずＭａｘ収納
                          For i2 = 1 To 4
                             If Mon_KabeF(i, i2) = True Then
                                Mon_Kyori(i, i2) = Abs(Mon_Y(i) + Ido_Y(i2) - KijunY) ^ 2 + Abs(Mon_X(i) + Ido_X(i2) - KijunX) ^ 2
                                If Mon_KyoriL(i) > Mon_Kyori(i, i2) Then '今までの距離より近ければ
                                   Mon_KyoriL(i) = Mon_Kyori(i, i2)
                                   Mon_S(i) = i2  '一番近い方向に決定
                                End If
                                'If i = 4 And Mon_JT(i) <> 1 And Mon_JT(i) <> 4 Then 'おとぼけ橙はPとの距離が130以下のときランダム
                                '   If Abs(Mon_Y(i) - Pac_Y) ^ 2 + Abs(Mon_X(i) - Pac_X) ^ 2 <= 130 ^ 2 Then
                                '      Do Until Mon_OtobokeF = True
                                '         Mon_S(i) = Int(Rnd * 4) + 1
                                '         If Mon_KabeF(i, Mon_S(i)) Then
                                '            Mon_OtobokeF = True
                                '         End If
                                '      Loop
                                '      Mon_OtobokeF = False
                                '   End If
                                'End If
                             End If
                          Next i2
                       End If
                    
                    
                    ElseIf Mon_JT(i) = 4 Then
                    
                    End If
                 
                 Mon_Kabe(i) = 0
               
                  '初期の赤の左進行例外処理
                  'If i = 1 And Count < 22 And Mon_JT(1) = 3 Then Mon_S(1) = 1
                  'If i = 1 And Count = 59 And Mon_JT(1) = 3 Then Mon_S(1) = 1
                  
                  If Mon_S0(i) <> 0 Then '直前が反転可だった場合
                     HantenF(i) = False
                     Mon_S(i) = Mon_S0(i) ' 直前の方向に設定（反転）
                     Mon_S0(i) = 0
                  End If
               
                  If HantenF(i) Then   '反転可の際は、現在の進行方向を保存
                     Mon_S0(i) = Hantai(Mon_S(i))
                  End If
               
                  '決定した移動方向により座標を収納
                  Mon_SY(i) = Ido_Y(Mon_S(i))
                  Mon_SX(i) = Ido_X(Mon_S(i))
                  Mon_Y(i) = Mon_Y(i) + Mon_SY(i)
                  Mon_X(i) = Mon_X(i) + Mon_SX(i)
               
               End If
               
               'ドットの有無判定
               If Map(Mon_Y(i) + DotHanteiY(Mon_S(i)), Mon_X(i) + DotHanteiX(Mon_S(i))) = 40 Then
                  MonDotF(i) = True
               Else
                  MonDotF(i) = False
               End If
               
               'ワープトンネル判定
               If Mon_X(i) < 3 Then
                  Mon_X(i) = 241
               ElseIf Mon_X(i) > 241 Then
                  Mon_X(i) = 3
               End If
            
            End If
            
            '書換画像データの収納
            If Mon_Ijike(i) Then  'いじけモードのときは
               Set Mon_data1(i) = Mon_data0(5). _
                  Offset((Mon_S(i) = 4) + (MonDotF(i) * (-96)), (Mon_S(i) = 2) + (Mon_SP(i) - 1) * 16 + (Mon_IjikeS = 3) * (-32)). _
                  Resize(14 + Abs(Mon_SY(i)), 14 + Abs(Mon_SX(i)))
                  Set Mon_data2(i) = Cells(Mon_Y(i), Mon_X(i)).Offset((Mon_S(i) = 4), (Mon_S(i) = 2)). _
                  Resize(14 + Abs(Mon_SY(i)), 14 + Abs(Mon_SX(i)))

            ElseIf Mon_JT(i) = 4 Then '目玉の時
               Set Mon_data1(i) = Mon_data0(5). _
                  Offset((Mon_S(i) = 4) + (MonDotF(i) * (-96)), (Mon_S(i) = 2) + (Mon_S(i) - 1) * 16 + 64). _
                  Resize(14 + Abs(Mon_SY(i)), 14 + Abs(Mon_SX(i)))
                  Set Mon_data2(i) = Cells(Mon_Y(i), Mon_X(i)).Offset((Mon_S(i) = 4), (Mon_S(i) = 2)). _
                  Resize(14 + Abs(Mon_SY(i)), 14 + Abs(Mon_SX(i)))
            ElseIf Mon_JT(i) = 5 Then  '点数の時
               If Mon_TenC2(i) Then
                  Set Mon_data1(i) = Mon_data0(6). _
                     Offset(0, 64). _
                     Resize(14, 15)
               Else
                  Set Mon_data1(i) = Mon_data0(6). _
                     Offset(0, (Renzoku(i) - 1) * 16). _
                     Resize(14, 15)
               End If
                  Set Mon_data2(i) = Cells(Mon_Y(i), Mon_X(i)). _
                  Resize(14, 15)
            Else 'いじけ目玉以外のとき
               Set Mon_data1(i) = Mon_data0(i). _
                  Offset((Mon_S(i) = 4) + (MonDotF(i) * (-96)), (Mon_S(i) = 2) + (Mon_SP(i) - 1) * 16 + (Mon_S(i) - 1) * 32). _
                  Resize(14 + Abs(Mon_SY(i)), 14 + Abs(Mon_SX(i)))
                  Set Mon_data2(i) = Cells(Mon_Y(i), Mon_X(i)).Offset((Mon_S(i) = 4), (Mon_S(i) = 2)). _
                  Resize(14 + Abs(Mon_SY(i)), 14 + Abs(Mon_SX(i)))
            End If
            
            '移動描写
            Mon_data1(i).Copy Destination:=Mon_data2(i)
            
         End If
      Next i

      'モンスター当たり判定
      For i = 1 To 4
         If Atari_Hantei(Pac_Y, Pac_X, Mon_Y(i), Mon_X(i)) Then
            
            If Mon_Ijike(i) Then 'いじけ食い
               '音入れ
               Sound 7
               Mon_Ijike(i) = False
               Mon_JT(i) = 5
               Renzoku(0) = Renzoku(0) + 1
               Renzoku(i) = Renzoku(0)
               Mon_TenF = True
               Score(0) = Score(0) + Renzoku_Ten(Renzoku(0))
               Score_hyoji
            ElseIf Mon_JT(i) >= 1 And Mon_JT(i) <= 3 Then '食われ
               Sleep 1000
               'モンスターを消去
               For i3 = 1 To 4
                  Set Mon_data1(i3) = Black(Abs(MonDotF(i3))). _
                  Offset((Mon_S(i3) = 4), (Mon_S(i3) = 2)). _
                  Resize(14 + Abs(Mon_SY(i3)), 14 + Abs(Mon_SX(i3)))
                  Mon_data1(i3).Copy Destination:=Mon_data2(i3)
               Next i3
               
               'ドットエサの復活描写
               For i3 = 1 To 4
                  For i4 = Mon_Y(i3) To Mon_Y(i3) + 15
                     For i5 = Mon_X(i3) To Mon_X(i3) + 15
                        If Map(i4, i5) = 40 And Map(i4 + 1, i5 + 1) = 40 Then
                           DotData.Copy Destination:=Range(Cells(i4, i5), Cells(i4 + 1, i5 + 1))
                        End If
                     Next i5
                  Next i4
               Next i3
               
               'パッセルの消去
               Pac_Syokyo.Copy Destination:= _
                  Cells(Pac_Y, Pac_X).Resize(15, 15)
               Sleep 300
               '音入れ
               Sound 8
               For i3 = 1 To 10
                  Sleep 100
                  Pac_Syokyo.Offset(0, i3 * 16).Copy Destination:= _
                     Cells(Pac_Y, Pac_X).Resize(15, 15)
               Next i3
               Sleep 500
               Black(0).Resize(15, 15).Copy Destination:= _
               Cells(Pac_Y, Pac_X).Resize(15, 15)
               'ドットエサの復活描写
               For i4 = Pac_Y To Pac_Y + 15
                  For i5 = Pac_X To Pac_X + 15
                     If Map(i4, i5) = 40 And Map(i4 + 1, i5 + 1) = 40 Then
                        DotData.Copy Destination:=Range(Cells(i4, i5), Cells(i4 + 1, i5 + 1))
                     End If
                  Next i5
               Next i4

               Sleep 500
               Pac_No = Pac_No - 1
               SyokiSettei
               ZankiHyoji (Pac_No)
            End If
         End If
      Next i
   
      '目玉モード切替仮置き
      If GetAsyncKeyState(17) <> 0 Then 'Ctrl仮置き実験用
         For i = 1 To 4
            Mon_JT(i) = 4
         
         Next i
      End If
      
      'ハイスコアの判定
      If HiScoreF = 1 Then
         If HiScore < Score(0) Then
            HiScore_hyoji
            HiScore = Score(0)
         End If
      End If
         
      '１UP判定
      If Score(0) > 10000 And Not (Up(1)) Then
         Up(1) = True
         Pac_No = Pac_No + 1
         Sound 6
         ZankiHyoji (Pac_No)
      End If
      
      '音
      For i = 1 To 4
         '点数中は音カット
         If Mon_JT(i) = 5 Then
            SoundSF = True
            Exit For
         ElseIf Mon_JT(i) = 4 Then '帰巣中
            Sound2 11  'ループ再生
            SoundSF = True
            Exit For
         End If
      Next i
      If Not (SoundSF) Then
         If Mon_IjikeS <> 0 Then
            If Count Mod 16 = 0 Then
               Sound 10
            End If
         Else
            Sound2 9
         End If
      End If
      SoundSF = False
      
      If DotCount = 0 And PowCount = 0 Then
      '面クリアー
         'パッセルマンを反転データに描写
         Pac_data1.Copy Destination:=Pac_data2.Offset(1800, 0)
         Pac_data1.Copy Destination:=Pac_data2.Offset(2048, 0)
         'モンスターの消去
         For i3 = 1 To 4
            Set Mon_data1(i3) = Black(Abs(MonDotF(i3))). _
            Offset((Mon_S(i3) = 4), (Mon_S(i3) = 2)). _
            Resize(14 + Abs(Mon_SY(i3)), 14 + Abs(Mon_SX(i3)))
            Mon_data1(i3).Copy Destination:=Mon_data2(i3)
         Next i3

         GamenHanten
         
         '反転データからパッセルマンを消去
         Set Pac_data1 = Black(0).Offset((Pac_S = 4), (Pac_S = 2)). _
                  Resize(13 + Abs(Pac_SY), 13 + Abs(Pac_SX))
         Pac_data1.Copy Destination:=Pac_data2.Offset(1800, 0)
         Pac_data1.Copy Destination:=Pac_data2.Offset(2048, 0)
         MenClear

      End If
   
      'ゲーム終了判定
      If Pac_No = 0 Then
      'GAMEOVERの表示
         Temp = Moji_Hyoji("GAME", 161, 90, 1) '第4引数の１は赤
         Temp = Moji_Hyoji("OVER", 161, 138, 1) '第4引数の１は赤
         Game_Flag = False
         Score(0) = 0
         Score(1) = 0
      End If
      
   Loop

Exit Sub

ErrH:

   'GAMEOVERの表示
   Temp = Moji_Hyoji("GAME", 161, 90, 1) '第4引数の１は赤
   Temp = Moji_Hyoji("OVER", 161, 138, 1) '第4引数の１は赤
   For i = 0 To 11   '音用エイリアスを開放
      FileClose LTrim(i) + ".wav", i
   Next
   Cells(1, 1).Select
   CoinF = False

   Temp = MsgBox("終了しますか？", vbYesNo, "PACELLMAN")
   Score(0) = 0
   Score(1) = 0
   Sleep 1000
   If Temp = vbNo Then Restart

End Sub

Sub GamenHanten()
   Dim Gamen1 As Range  '元画面
   Dim Gamen2 As Range  'フラッシュデータ
   Dim Gamen3 As Range  '元画面保管用
   Dim FlashC As Integer
   Set Gamen1 = Range(Cells(25, 17), Cells(272, 240))
   Set Gamen2 = Range(Cells(1825, 17), Cells(2072, 240))
   Set Gamen3 = Range(Cells(2073, 17), Cells(2320, 240))
   
   Gamen1.Copy Destination:=Gamen3

   For FlashC = 1 To 5
      Do While GetTickCount - WaitStart < 300
      Loop
      Gamen2.Copy Destination:=Gamen1
      WaitStart = GetTickCount
      Do While GetTickCount - WaitStart < 300
      Loop
      Gamen3.Copy Destination:=Gamen1
      WaitStart = GetTickCount
   Next FlashC
End Sub

Function Atari_Hantei(ByVal PY As Long, ByVal PX As Long, ByVal MY As Long, ByVal MX As Long) As Boolean
    Atari_Hantei = False
    If Abs(PY - MY) < 6 Then
      If Abs(PX - MX) < 6 Then
         Atari_Hantei = True
      End If
    End If
End Function

    
Sub WaitS(Start As Long, Finish As Long)
   Do While GetTickCount - Start < Finish
   Loop
End Sub

Sub SyokiSettei()
DoEvents
   Dim b As Integer
   For b = 1 To 4
      Mon_Ijike(b) = False
   Next b
      
   Mon_IjikeSt = 0
   Mon_IjikeS = 0
   Mon_JT(1) = 3
   Mon_JT(2) = 0
   Mon_JT(3) = 0
   Mon_JT(4) = 0
   Mon_S(1) = 1
   Mon_S(2) = 3
   Mon_S(3) = 4
   Mon_S(4) = 3
   
   Mon_Y(1) = 110
   Mon_X(1) = 122
   Mon_Y(2) = 134
   Mon_X(2) = 122
   Mon_Y(3) = 134
   Mon_X(3) = 106
   Mon_Y(4) = 134
   Mon_X(4) = 138
      Pac_Y = 206
   Pac_X = 123
   Pac_S = 1

   Set Moji = Cells(901, 1)

   
End Sub

Sub ZankiHyoji(a As Long)
   Dim b As Integer
   Dim Pac_Zanki As Range
   Set Pac_Zanki = Range(Cells(826, 129), Cells(836, 139))
   
   If a > 5 Then
      a = 5
   End If
   If a = 1 Then
      Black(1).Resize(11, 11).Copy Destination:=Cells(275, 30).Resize(11, 11)
   Else
      For b = 2 To a
         Pac_Zanki.Copy Destination:=Cells(275, 30).Offset(0, (b - 2) * 16).Resize(11, 11)
      Next b
   End If
   If a < 5 Then
      Black(1).Resize(11, 11).Copy Destination:=Cells(275, 30).Offset(0, (a - 1) * 16).Resize(11, 11)
   End If
   
End Sub

Sub MenClear()
   Men(1) = Men(1) + 1
   Men(2) = Men(2) + 1
   If Men(2) > 21 Then
      Men(2) = 21
   End If
   
   If Men(2) = 3 Then
      Demo1
   End If
   
   MenKaishi
   SyokiSettei

End Sub

Sub MenKaishi()
DoEvents
   Dim b As Integer
    '画面描写
   Range(Cells(301, 17), Cells(588, 240)).Copy Destination:=Range(Cells(1, 17), Cells(288, 240))
   Map = Range(Cells(1113, 1), Cells(1400, 256))
   Map2 = Range(Cells(1113, 1), Cells(1400, 256))
   Count = 0
   PowCN = 0
   PowC = 1
   DotCount = 240
   PowCount = 4
   Pac_KS = 1
   For b = 1 To 4
      Pow_S(b) = True
      MonSyutsugenC(b) = 0
      Mon_SPDBuf(b) = 0
      Mon_IdoF(b) = False
      HantenF(b) = False
      Mon_TenC0(b) = 0
      Mon_TenC1(b) = 0
      Mon_TenC2(b) = 0
      Mon_TenF = False
   Next b
   ZankiHyoji (Pac_No)
   Score(1) = 0
   Score_hyoji
   
   'readyの消去
   Sleep 4000
   Cells(893, 201).Resize(8, 48).Copy Destination:=Cells(161, 106).Resize(8, 48)

End Sub

'=========================================
'   スコア表示
'     ～引数に指定した数字を該当セルに表示
'=========================================
Sub Score_hyoji()
   Dim Sb0 As Long
   Dim Sb1 As Long
   Dim Si As Long
   Dim S0 As Long
   Dim S1 As Long
   
   S0 = Score(0)
   S1 = Score(1)
   For Si = 5 To 1 Step -1
                  
      Sb0 = Int(S0 / (10 ^ Si))
      Sb1 = Int(S1 / (10 ^ Si))
      If Sb0 <> Sb1 Then
         Score_data.Offset(0, Sb0 * 8).Copy Destination:=Range(Cells(10, 26 + (5 - Si) * 8), Cells(16, 32 + (5 - Si) * 8))
      End If
   
      S0 = S0 - Sb0 * (10 ^ Si)
      S1 = S1 - Sb1 * (10 ^ Si)
   Next Si
   Score(1) = Score(0)
End Sub

Sub HiScore_hyoji()
    Range(Cells(10, 26), Cells(16, 72)).Copy Destination:=Range(Cells(10, 106), Cells(16, 152))
End Sub

Function Moji_Hyoji(a As String, Y As Long, X As Long, C As Long)
'座標（Y,X）に、アルファベット文字列ａを描画
'"A"は６５
Dim b As Long
For b = 1 To Len(a)
   Moji.Offset(C * 8, (Asc(Mid(a, b, 1)) - 65) * 8).Resize(8, 8).Copy Destination:= _
   Cells(Y, X).Offset(0, (b - 1) * 8).Resize(8, 8)
Next b

End Function

Function FileOpen(FileName As String, FileNo As Long)

    mciSendString "stop audiofile" & LTrim(FileNo), vbNullString, 0, 0
    mciSendString "close audiofile" & LTrim(FileNo), vbNullString, 0, 0
    mciSendString "open """ & ActiveWorkbook.Path & "\" & FileName & """ alias audiofile" & LTrim(FileNo), vbNullString, 0, 0

End Function
Function Sound(FileNo As Long)  '再生
   If SoundF Then
    mciSendString "play audiofile" & LTrim(FileNo) & " from 0", vbNullString, 0, 0
   End If
End Function
Function Sound2(FileNo As Long)  'ループ再生、再生が終了しているときのみ再生する
   If SoundF Then
      mciSendString "status audiofile" & LTrim(FileNo) & " mode", strStatus, 256, 0
      If Left$(strStatus, 7) = "stopped" Then
         mciSendString "play audiofile" & LTrim(FileNo) & " from 0", vbNullString, 0, 0
      End If
   End If
End Function
Function Sound3(FileNo As Long)  '再生終了後制御を返す
   If SoundF Then
    mciSendString "play audiofile" & LTrim(FileNo) & " wait" & " from 0", vbNullString, 0, 0
   End If
End Function

Function FileClose(FileName As String, FileNo As Long)
    mciSendString "stop audiofile" & LTrim(FileNo), vbNullString, 0, 0
    mciSendString "close audiofile" & LTrim(FileNo), vbNullString, 0, 0
End Function

Sub Opening()
If CoinF Then Exit Sub
DoEvents
   Dim DemoF As Boolean
   DemoF = True
   Dim a(1 To 20, 0 To 4) As Long '0:Wait 1:Y 2:X 3:YS 4:XS)
   Dim Data0 As Range
   Set Data0 = Cells(1, 17)
   Dim Data1 As Range
   Set Data1 = Cells(2331, 17)
   Dim Data2 As Range
   Set Data2 = Cells(2631, 17)
   Dim Esa As Range
   Set Esa = Cells(826, 5)
   '画面を黒く
   Data2.Resize(288, 224).Copy Destination:=Data0.Resize(288, 224)
   Sleep 200
   '書き換えデータの収納
   For i = 1 To 17
      For i2 = 0 To 4
         a(i, i2) = Worksheets("Demo").Cells(2 + i, 2 + i2)
      Next i2
   Next i
   Range(Cells(2331, 43), Cells(2377, 232)).Copy Destination:= _
      Range(Cells(1, 43), Cells(47, 232))
   
   WaitStart = GetTickCount
   Do While GetTickCount - WaitStart < 1000
      If GetAsyncKeyState(13) <> 0 Then 'リターン
         Sound 4
         CoinF = True
         DemoF = False
      End If
   Loop
      
   'キャラクター紹介表示
   For i = 3 To 17
         
      Data1.Offset(a(i, 1), a(i, 2)).Resize(a(i, 3), a(i, 4)).Copy Destination:= _
        Data0.Offset(a(i, 1), a(i, 2)).Resize(a(i, 3), a(i, 4))
     ' For i2 = 1 To a(i, 0) / 100
         WaitStart = GetTickCount
      Do While GetTickCount - WaitStart < a(i, 0)
         If GetAsyncKeyState(13) <> 0 Then 'リターン
            Sound 4
            CoinF = True
            DemoF = False
            Exit Do
         End If
      Loop
      If CoinF = True Then Exit For
      
   Next i
     'モンスターデータ
   Dim Mon_data0(1 To 6) As Range   '5は、いじけ、目玉モード用　6は点数
   For i = 1 To 5
      Set Mon_data0(i) = Cells((i - 1) * 16 + 618, 2)
   Next i
   Set Mon_data0(6) = Cells(826, 34)
      'パッセルマンのデータ
   Dim Pac_data0 As Range
   Set Pac_data0 = Cells(602, 2)
   Dim Pow_data(1 To 2) As Range
   Set Pow_data(1) = Range(Cells(826, 5), Cells(833, 12))
   Set Pow_data(2) = Range(Cells(826, 5), Cells(833, 12)).Offset(0, 16)
   Dim Y(0 To 4) As Long
   Dim X(0 To 4) As Long
   Dim Speed(0 To 4) As Long
      Speed(0) = 128
      Speed(1) = 132
      Speed(2) = 132
      Speed(3) = 132
      Speed(4) = 132
   Dim SpeedBuf(0 To 4) As Long

   Dim DemoC As Long
   Dim S(0 To 4) As Long  '１：左　２：右
   Dim P(0 To 4) As Long   '描画パターン
   Dim J(0 To 4) As Long   '１：通常　２：いじけ　３：目玉
   For i = 0 To 4
      Y(i) = 158 + (i > 0)
      X(i) = 240 + 16 * i - (i > 0) * 8
      S(i) = 1
      P(i) = 1
      J(i) = 1
   Next i
   
   Do While DemoF
      If GetAsyncKeyState(13) <> 0 Then 'リターン
         Sound 4
         CoinF = True
         DemoF = False
      End If
      'パックマン
      SpeedBuf(0) = SpeedBuf(0) + Speed(0)
      If SpeedBuf(0) > 256 Then
         DemoC = DemoC + 1
         Do While GetTickCount - WaitStart < WaitNo
         Loop
         WaitStart = GetTickCount
         SpeedBuf(0) = SpeedBuf(0) - 256
         P(0) = P(0) + 1
         If P(0) > 4 Then P(0) = 1
         If X(0) < 50 Then
            S(0) = 2
            Cells(1, 1).Copy Destination:=Cells(209, 97).Resize(8, 8)
         End If
         X(0) = X(0) + (S(0) - 1) * 2 - 1
         Pac_data0.Offset(0, -(S(0) = 2) * 64 + (P(0) - 1) * 16 + (S(0) = 2)).Resize(13, 14).Copy Destination:= _
         Cells(Y(0), X(0) + (S(0) = 2)).Resize(13, 14)
      
         'パワー餌
         Select Case DemoC Mod 16
            Case 8
               If S(0) = 1 Then
                  Esa.Resize(8, 8).Copy Destination:=Cells(209, 97).Resize(8, 8)
                  Esa.Resize(8, 8).Copy Destination:=Cells(161, 49).Resize(8, 8)
               End If
            Case 0
               Cells(1, 1).Copy Destination:=Cells(209, 97).Resize(8, 8)
               Cells(1, 1).Copy Destination:=Cells(161, 49).Resize(8, 8)
         End Select
      End If
      
      'モンスターの座標
      For i = 1 To 4
         SpeedBuf(i) = SpeedBuf(i) + Speed(i)
         If SpeedBuf(i) > 256 Then
            SpeedBuf(i) = SpeedBuf(i) - 256
            P(i) = P(i) + 1
            If P(i) > 2 Then P(i) = 1
            If X(0) < 54 Then
               If S(i) = 1 Then Speed(i) = Speed(i) / 2
               S(i) = 2
               J(i) = 2
            End If
            X(i) = X(i) + (S(i) - 1) * 2 - 1
           'あたり判定
            If X(0) = X(i) - 7 And J(i) = 2 Then J(i) = 3
            If X(i) < 240 Then
               Select Case J(i)
                  Case 1 '通常
                     Mon_data0(i).Offset(0, -(S(i) = 2) * 32 + (P(i) - 1) * 16 + (S(i) = 2)).Resize(14, 15).Copy Destination:= _
                     Cells(Y(i), X(i) + (S(i) = 2)).Resize(14, 15)
                  Case 2  'いじけ
                     Mon_data0(5).Offset(0, (P(i) - 1) * 16 + (S(i) = 2)).Resize(14, 15).Copy Destination:= _
                     Cells(Y(i), X(i) + (S(i) = 2)).Resize(14, 15)
                  Case 3  '点数
                  Cells(1, 17).Resize(13, 14).Copy Destination:= _
                  Cells(Y(0), X(0) + (S(0) = 2)).Resize(13, 14)
                  Mon_data0(6).Offset(0, (i - 1) * 16).Resize(14, 15).Copy Destination:= _
                  Cells(Y(i), X(i)).Resize(14, 15)
                  Sleep 1000
                  J(i) = 4
                  Cells(1, 17).Resize(14, 15).Copy Destination:= _
                  Cells(Y(i), X(i)).Resize(14, 15)
                  If i = 4 Then DemoF = False
               End Select
            End If
         End If
      Next i
   Loop
   Cells(1, 1).Copy Destination:=Cells(161, 49).Resize(8, 8)
   Sleep 1000
   
End Sub

Sub Demo1()

DoEvents
   Dim MusicF As Boolean

   Dim Data0 As Range
   Set Data0 = Cells(1, 17)
   Dim Data2 As Range
   Set Data2 = Cells(2631, 17)
   '画面を黒く
   Data2.Resize(288, 224).Copy Destination:=Data0.Resize(288, 224)
   Sound 1
   Sleep 1000
   
   '書き換えデータの収納
     'モンスターデータ
   Dim Mon_data0(1 To 6) As Range   '5は、いじけ、目玉モード用　6は点数
   For i = 1 To 5
      Set Mon_data0(i) = Cells((i - 1) * 16 + 618, 2)
   Next i
   Set Mon_data0(6) = Cells(826, 34)
      'パッセルマンのデータ
   Dim Pac_data0 As Range
   Set Pac_data0 = Cells(602, 2)
   Dim Pow_data(1 To 2) As Range
   Set Pow_data(1) = Range(Cells(826, 5), Cells(833, 12))
   Set Pow_data(2) = Range(Cells(826, 5), Cells(833, 12)).Offset(0, 16)
   Dim Y(0 To 1) As Long
   Dim X(0 To 1) As Long
   Dim Speed(0 To 4) As Long
      Speed(0) = 128
      Speed(1) = 136
   Dim SpeedBuf(0 To 1) As Long
   Dim DemoF As Boolean
   DemoF = True
   Dim DemoC As Long
   Dim Taiki(0 To 1) As Long
   Dim S(0 To 1) As Long  '１：左　２：右
   Dim P(0 To 1) As Long   '描画パターン
   Dim J(0 To 1) As Long   '１：通常　２：いじけ   For i = 0 To 4
   For i = 0 To 1
      Y(i) = 158 + (i > 0)
      X(i) = 240 + 16 * i - (i > 0) * 12
      S(i) = 1
      P(i) = 1
      J(i) = 1
   Next i

   Do While DemoF
      'パックマン
      SpeedBuf(0) = SpeedBuf(0) + Speed(0)
      If SpeedBuf(0) > 256 Then
         DemoC = DemoC + 1
         Do While GetTickCount - WaitStart < WaitNo
         Loop
         WaitStart = GetTickCount
         SpeedBuf(0) = SpeedBuf(0) - 256
         P(0) = P(0) + 1
         If P(0) > 4 Then P(0) = 1
         If X(0) < 4 Then
            S(0) = 2
         End If
         
         X(0) = X(0) + (S(0) - 1) * 2 - 1
         
         If S(0) = 1 Then
            Pac_data0.Offset(0, -(S(0) = 2) * 64 + (P(0) - 1) * 16 + (S(0) = 2)).Resize(13, 14).Copy Destination:= _
            Cells(Y(0), X(0) + (S(0) = 2)).Resize(13, 14)

         Else
            Taiki(0) = Taiki(0) + 1
            If Taiki(0) < 256 Then
               X(0) = X(0) - 1
            Pac_data0.Offset(0, -(S(0) = 2) * 64 + (P(0) - 1) * 16 + (S(0) = 2)).Resize(13, 14).Copy Destination:= _
            Cells(Y(0), X(0) + (S(0) = 2)).Resize(13, 14)
            Else
               Cells(842, 1).Offset(0, (P(0) - 1) * 32).Resize(32, 32).Copy Destination:= _
               Cells(Y(0) - 19, X(0)).Resize(32, 32)
            
             End If
         End If
      End If

      'モンスターの座標
      For i = 1 To 1
         SpeedBuf(i) = SpeedBuf(i) + Speed(i)
         If SpeedBuf(i) > 304 Then
            SpeedBuf(i) = SpeedBuf(i) - 256
            P(i) = P(i) + 1
            If P(i) > 2 Then P(i) = 1
            If X(i) < 3 Then
               If S(i) = 1 Then Speed(i) = Speed(i) / 2
               S(i) = 2
               J(i) = 2
            End If
            X(i) = X(i) + (S(i) - 1) * 2 - 1
            
            If X(i) < 240 Then
               Select Case J(i)
                  Case 1 '通常
                     If Taiki(1) = 0 Then
                        Mon_data0(i).Offset(0, -(S(i) = 2) * 32 + (P(i) - 1) * 16 + (S(i) = 2)).Resize(14, 15).Copy Destination:= _
                        Cells(Y(i), X(i) + (S(i) = 2)).Resize(14, 15)
                     End If
                  Case 2  'いじけ
                     Taiki(1) = Taiki(1) + 1
                     If Taiki(1) < 32 Then
                        X(i) = X(i) - 1
                     End If
                                          
                     If Taiki(1) = 32 Then MusicF = True
                                          
                     If MusicF Then
                        Sound 1
                        MusicF = False
                     End If
                     
                     Mon_data0(5).Offset(0, (P(i) - 1) * 16 + (S(i) = 2)).Resize(14, 15).Copy Destination:= _
                     Cells(Y(i), X(i) + (S(i) = 2)).Resize(14, 15)
               End Select
            End If
         End If
      Next i
      
      If DemoC > 629 Then
         DemoF = False
      End If
   Loop

   Sleep 1000
End Sub

Sub Demo2()
   DoEvents
   SyokiSettei
   Dim Data0 As Range
   Set Data0 = Cells(1, 1)
   Dim Data2 As Range
   Set Data2 = Cells(2631, 1)

   '画面を黒く
   Data2.Resize(288, 256).Copy Destination:=Data0.Resize(288, 256)

   For i = 1 To 100
      Cells(826, 193).Resize(51, 52).Copy Destination:= _
      Cells(i, 100).Resize(51, 52)
      If GetAsyncKeyState(13) <> 0 Then 'リターン
         Sound 4
         CoinF = True
         'DemoF = False
         Exit For
      End If
   Next i
   
   If CoinF Then GoTo Game
   
   Temp = Moji_Hyoji("VBA", 160, 72, 1)
   Temp = Moji_Hyoji("ACTIONGAME", 160, 72 + 32, 1)
   Sleep 1000
   
   'Data2.Offset(1000, 16).Resize(288, 224).Copy Destination:=Data0.Resize(288, 224)
   'Data2.Resize(288, 256).Copy Destination:=Data0.Resize(288, 256)
   
   Data2.Resize(51, 52).Copy Destination:= _
      Cells(100, 100).Resize(51, 52)
   Cells(826, 193).Resize(51, 52).Copy Destination:= _
      Cells(1, 1).Resize(51, 52)
   Sleep 200
   Cells(1, 1).Select
   For i = 10 To 400 Step 5
       ActiveWindow.Zoom = i
      If GetAsyncKeyState(13) <> 0 Then 'リターン
         Sound 4
         CoinF = True
         'DemoF = False
      Exit For
      End If
   Next i
   
   'For i = 400 To 10 Step -5
   '    ActiveWindow.Zoom = i
   'Next i
Game:
   Cells(1, 200).Resize(52, 52).Copy Destination:= _
      Cells(1, 1).Resize(52, 52)
   ActiveWindow.Zoom = 10

End Sub

Sub Book_Open()
   CoinF = False
   SyokiSettei

   Start
End Sub

Sub KakushiStart()
   CoinF = True
   SyokiSettei
   Start
End Sub


Sub Config_Def()
   Worksheets("Config").Cells(1, 101).Resize(100, 100).Copy Destination:= _
   Worksheets("Config").Cells(1, 1).Resize(100, 100)
End Sub

Sub Config_Slow()
   Worksheets("Config").Cells(101, 101).Resize(100, 100).Copy Destination:= _
   Worksheets("Config").Cells(1, 1).Resize(100, 100)
End Sub

Sub Settei()
   Worksheets("Config").Select
   Range("a1").Select
End Sub

Sub Restart()
   Worksheets("Main").Select
   Range("a1").Select
   Book_Open
End Sub
