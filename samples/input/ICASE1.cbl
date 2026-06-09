999999 IDENTIFICATION                         DIVISION.
999999 PROGRAM-ID.       ICASE1.
000100**********************************************************************
000200* PROGRAM-ID   : ICASE1
000300* PROGRAM-NAME : 分岐構文 網羅サンプル (デモ)
000400* DATA-WRITTEN : 26.05.18
000500**********************************************************************
999999 ENVIRONMENT                            DIVISION.
999999 CONFIGURATION                          SECTION.
999999     SPECIAL-NAMES.
999999     UPSI-7  ON  STATUS  IS  TRACE-ON.
999999 DATA                                   DIVISION.
999999 FILE                                   SECTION.
999999 WORKING-STORAGE                        SECTION.
999999*
000700*  パラメータエリア
000900* LOCAL DATA-AREA
000900 01  ICA-WORK.
000900     COPY  ICAWORK     PREFIXING  ICA-WORK-.
001000 01  WK-AREA.
001000     03  WK-MODE         PIC X(01).
001000     03  WK-FLG          PIC X(01).
001000     03  WK-RANK         PIC X(01).
001000     03  WK-IDX          PIC 9(02).
001000     03  WK-TBL.
001000         05  WK-ENT  OCCURS 5 INDEXED BY WK-I.
001000             10  WK-ENT-CD   PIC X(02).
001000             10  WK-ENT-NM   PIC X(10).
999999**********************************************************************
999999*        L I N K   A R E A                                           *
999999**********************************************************************
999999 LINKAGE                                SECTION.
999999 01  ICA-PARAM.
999999     COPY  ICAPARAM    PREFIXING  ICA-PARAM-.
999999 PROCEDURE                              DIVISION
999999     USING  ICA-PARAM.
001100*
001200**********************************************************************
001300*   ICASE1   （ メイン処理 ）                                         *
001400**********************************************************************
999999 ICASE1-PROC                            SECTION.
999999 ICASE1-000.
001401* 初期処理
001402         PERFORM  ICASE1-INIT.
001501* 主処理
001502         PERFORM  ICASE1-MAIN.
001601* 終了処理
001602         PERFORM  ICASE1-EXIT THRU  ICASE1-EXIT-OUT.
999999 ICASE1-999.
999999     GOBACK.
001700**********************************************************************
001800*   ICASE1-INIT  （ 初期処理 ）                                       *
001900**********************************************************************
002000 ICASE1-INIT                            SECTION.
999999 ICASE1-INIT-000.
002100* SUB-ID SET
999999         MOVE  'ICASE1'  TO  ICA-PARAM-F300.
999999         MOVE  '0000'    TO  ICA-PARAM-F400.
999999         MOVE  0         TO  ICA-PARAM-F200.
002200* 入力パラメータ妥当性チェック
999999         IF    ICA-PARAM-F100  =  SPACE
999999             MOVE  'PARM'   TO  ICA-PARAM-F400
999999         ELSE
999999             MOVE  ICA-PARAM-F100  TO  WK-MODE
999999         END-IF.
002300* 業務区分判定
999999         EVALUATE  WK-MODE
999999             WHEN  '1'
999999                 PERFORM  INIT-NORMAL
999999             WHEN  '2'
999999                 PERFORM  INIT-URGENT
999999             WHEN  OTHER
999999                 PERFORM  INIT-DEFAULT
999999         END-EVALUATE.
002400* エリア・コード検索
999999         SET   WK-I       TO  1.
999999         SEARCH  WK-ENT
999999             AT END
999999                 MOVE  'AREA' TO  ICA-PARAM-F400
999999             WHEN  WK-ENT-CD (WK-I)  =  ICA-PARAM-F110
999999                 MOVE  WK-ENT-NM (WK-I)  TO  ICA-WORK-F120
999999         END-SEARCH.
002500 ICASE1-INIT-999.
999999     EXIT.
002600 INIT-NORMAL.
999999         MOVE  '1'   TO  WK-FLG.
002700 INIT-URGENT.
999999         MOVE  '2'   TO  WK-FLG.
002800 INIT-DEFAULT.
999999         MOVE  '9'   TO  WK-FLG.
002900**********************************************************************
003000*   ICASE1-MAIN  （ 主処理：7段ネスト IF による支払判定 ）          *
003100**********************************************************************
003200 ICASE1-MAIN                            SECTION.
999999 ICASE1-MAIN-000.
999999         MOVE  'D2101330'  TO  ICA-WORK-F110.
999999         CALL  'ICASUB'    USING  ICA-WORK.
003300* L1 種別
999999         IF    ICA-WORK-F600  =  '1'
003400* L2 種類
999999             IF    ICA-WORK-F610  =  'A'
003500* L3 カバレッジ有無
999999                 IF    ICA-WORK-F620  >  0
003600* L4 支払額
999999                     IF    ICA-WORK-F630  >  1000000
003700* L5 契約期間
999999                         IF    ICA-WORK-F640  =  '12'
003800* L6 年齢下限
999999                             IF    ICA-WORK-F650  >=  20
003900* L7 年齢上限
999999                                 IF    ICA-WORK-F650  <=  60
999999                                     PERFORM  MAIN-PREMIUM
999999                                 ELSE
999999                                     PERFORM  MAIN-AGE-OVER
999999                                 END-IF
999999                             ELSE
999999                                 PERFORM  MAIN-AGE-UNDER
999999                             END-IF
999999                         ELSE
999999                             PERFORM  MAIN-PERIOD-OTHER
999999                         END-IF
999999                     ELSE
999999                         PERFORM  MAIN-AMOUNT-LOW
999999                     END-IF
999999                 ELSE
999999                     PERFORM  MAIN-NO-COVERAGE
999999                 END-IF
999999             ELSE
999999                 PERFORM  MAIN-POLICY-OTHER
999999             END-IF
999999         ELSE
999999             PERFORM  MAIN-NON-PREMIUM
999999         END-IF.
004000 ICASE1-MAIN-999.
999999     EXIT.
004100 MAIN-PREMIUM.
999999         MOVE  'P'   TO  WK-RANK.
004200 MAIN-AGE-OVER.
999999         MOVE  'O'   TO  WK-RANK.
004300 MAIN-AGE-UNDER.
999999         MOVE  'U'   TO  WK-RANK.
004400 MAIN-PERIOD-OTHER.
999999         MOVE  'T'   TO  WK-RANK.
004500 MAIN-AMOUNT-LOW.
999999         MOVE  'L'   TO  WK-RANK.
004600 MAIN-NO-COVERAGE.
999999         MOVE  'N'   TO  WK-RANK.
004700 MAIN-POLICY-OTHER.
999999         MOVE  'X'   TO  WK-RANK.
004800 MAIN-NON-PREMIUM.
999999         MOVE  'S'   TO  WK-RANK.
004900**********************************************************************
005000*   ICASE1-EXIT  （ 終了処理 ）                                       *
005100**********************************************************************
005200 ICASE1-EXIT                            SECTION.
999999 ICASE1-EXIT-000.
999999         MOVE  WK-RANK   TO  ICA-PARAM-F500.
999999         MOVE  '9999'    TO  ICA-PARAM-F999.
999999 ICASE1-EXIT-OUT.
999999     EXIT.
