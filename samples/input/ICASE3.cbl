999999 IDENTIFICATION                           DIVISION.
999999 PROGRAM-ID.                              ICASE3.
000100**********************************************************************
000200* PROGRAM-ID   : ICASE3
000300* PROGRAM-NAME : カーソル検索・結果編集サンプル (デモ)
000400* DATA-WRITTEN : 26.06.10
000500**********************************************************************
999999 ENVIRONMENT                           DIVISION.
999999 CONFIGURATION                              SECTION.
999999 DATA                                       DIVISION.
999999 WORKING-STORAGE                            SECTION.
999999*
999999 01  CNS-AREA.
999999     03  CNS-PROGRAM-ID                 PIC X(08) VALUE 'ICASE3'.
999999*
999999 01  MSG-AREA.
999999     03  MSG-ERR-K77-OPEN               PIC X(20)
999999         VALUE 'OPEN ERROR (K77)'.
999999     03  MSG-ERR-K77-FETCH              PIC X(20)
999999         VALUE 'FETCH ERROR (K77)'.
999999     03  MSG-ERR-K77-CLOSE              PIC X(20)
999999         VALUE 'CLOSE ERROR (K77)'.
999999*
999999 01  W-AREA.
999999     03  W-RESULT-CD                    PIC X(04).
999999     03  W-SEARCH-KEY.
999999         05  W-KEY-YEAR                 PIC X(04).
999999         05  W-KEY-CODE                 PIC X(05).
999999     03  W-YEAR-N                       PIC S9(04).
999999     03  W-YEAR-N-R REDEFINES W-YEAR-N.
999999         05  W-YEAR-A                   PIC X(04).
999999*
999999 01  T77-VIEW.
999999     03  T7715200                       PIC X(05).
999999     03  T7716700                       PIC X(08).
999999     03  T7717200                       PIC X(10).
999999*
999999 01  WK-ABEND-AREA.
999999     03  WK-RTN-CD                      PIC X(03).
999999     03  WK-MSG.
999999         05  WK-USER-MSG                PIC X(20).
999999         05  WK-KEY                     PIC X(10).
999999*
999999* 日付サブ (DATESUB) インターフェース
999999 01  DT1-PARAM.
999999     03  DT1-MODE                       PIC X(01).
999999 01  DT2-PARAM.
999999     03  DT2-YEAR                       PIC 9(04).
999999     03  DT2-MONTH                      PIC 9(02).
999999*
999999* 業務サブ (SUBX) インターフェース
999999 01  SBX-PARAM.
999999     03  SBX-PA100                      PIC 9(01).
999999     03  SBX-PA200                      PIC X(08).
999999     03  SBX-PA400                      PIC X(04).
999999     03  SBX-PA500                      PIC X(01).
999999*
888888     EXEC ADABAS
888888          BEGIN DECLARE SECTION
888888     END-EXEC.
999999     EXEC ADABAS
999999          FIND
999999          DECLARE K77 CURSOR FOR
999999          SELECT T7715200,
999999                 T7716700,
999999                 T7717200
999999          FROM   TBL7700 TD1
999999          WHERE  TKEY0100 = :W-SEARCH-KEY
999999     END-EXEC.
999999*
999999 LINKAGE                                                 SECTION.
999999 01  IC3-PARAM.
999999     03  IC3-PA100                      PIC X(05).
999999     03  IC3-PA200                      PIC X(18).
999999     03  IC3-PA300                      PIC X(08).
999999     03  IC3-PA400                      PIC X(04).
999999*
999999 PROCEDURE                                               DIVISION
999999     USING IC3-PARAM.
003100**********************************************************************
003200*   ICASE3      （メイン処理）                                      *
003300**********************************************************************
999999 ICASE3-PROC                                             SECTION.
999999 ICASE3-000.
999999     INITIALIZE WK-ABEND-AREA.
003300* 初期処理
999999     PERFORM ICASE360-PROC.
003400* 主処理
999999     PERFORM ICASE361-PROC.
999999*
999999 ICASE3-999.
999999     GOBACK.
003500**********************************************************************
003600*   ICASE360    （初期処理）                                        *
003700**********************************************************************
999999 ICASE360-PROC                                           SECTION.
999999 ICASE360-000.
999999*カーソルＯＰＥＮ
999999     PERFORM K77-OPEN-PROC.
999999* SET AREA CLEAR
999999     MOVE ' '        TO IC3-PA200.
999999* SUB-ID SET
999999     MOVE 'ICASE3'   TO IC3-PA300.
999999* ERR-FLG CLEAR
999999     MOVE '0000'     TO IC3-PA400.
999999     MOVE '0000'     TO W-RESULT-CD.
999999 ICASE360-999.
999999     EXIT.
004400**********************************************************************
004500*   ICASE361    （主処理）                                          *
004600**********************************************************************
999999 ICASE361-PROC                                           SECTION.
999999 ICASE361-000.
999999*日付サブ（年月取得）
999999     MOVE '2'            TO DT1-MODE.
999999     CALL 'DATESUB'      USING DT1-PARAM DT2-PARAM.
999999     IF DT2-MONTH = 01 OR
999999        DT2-MONTH = 02 OR
999999        DT2-MONTH = 03
999999     THEN
999999         COMPUTE W-YEAR-N = DT2-YEAR - 1
999999         MOVE W-YEAR-A TO W-KEY-YEAR
999999     ELSE
999999         MOVE DT2-YEAR TO W-KEY-YEAR
999999     END-IF.
999999     MOVE IC3-PA100 TO W-KEY-CODE.
999999*カーソルＦＩＮＤ
999999     PERFORM K77-FETCH-PROC.
999999*カーソルＣＬＯＳＥ
999999     PERFORM K77-CLOSE-PROC.
999999*
999999     IF W-RESULT-CD = '0001'
999999     THEN
999999         MOVE 1 TO SBX-PA100
999999         MOVE 'T7700300' TO SBX-PA200
999999         CALL 'SUBX' USING SBX-PARAM
999999         IF SBX-PA500 = '1'
999999         THEN
999999             MOVE SBX-PA400 TO IC3-PA400
999999         ELSE
999999             MOVE '9999' TO IC3-PA400
999999         END-IF
999999     ELSE
999999         CONTINUE
999999     END-IF.
999999 ICASE361-999.
999999     EXIT.
999999**********************************************************************
999999*   OPEN: カーソルＯＰＥＮ    TBL7700 TD1                           *
999999**********************************************************************
999999 K77-OPEN-PROC                                           SECTION.
999999*
999999     EXEC ADABAS
999999          OPEN K77
999999     END-EXEC.
999999*異常処理
999999     IF ADACODE NOT = ZERO
999999     THEN
999999         MOVE MSG-ERR-K77-OPEN TO WK-USER-MSG
999999         MOVE 001              TO WK-RTN-CD
999999         PERFORM S99-ABEND-PROC
999999     ELSE
999999         CONTINUE
999999     END-IF.
999999 K77-OPEN-PROC-EX.
999999     EXIT.
999999**********************************************************************
999999*   FETCH: カーソルＦＩＮＤ   TBL7700 TD1                           *
999999**********************************************************************
999999 K77-FETCH-PROC                                          SECTION.
999999*
999999     EXEC ADABAS
999999          FETCH K77
999999     END-EXEC.
999999*
999999     IF ADACODE = 3
999999     THEN
999999         MOVE '0001' TO W-RESULT-CD
999999     ELSE
999999         IF ADACODE = 0
999999         THEN
999999             STRING T7716700 OF TD1  T7717200 OF TD1
999999                    DELIMITED BY SIZE INTO IC3-PA200
999999             MOVE '0000' TO W-RESULT-CD
999999         ELSE
999999             MOVE MSG-ERR-K77-FETCH TO WK-USER-MSG
999999             MOVE IC3-PA100         TO WK-KEY
999999             MOVE 002               TO WK-RTN-CD
999999             PERFORM S99-ABEND-PROC
999999         END-IF
999999     END-IF.
999999 K77-FETCH-PROC-EX.
999999     EXIT.
999999**********************************************************************
999999*   CLOSE: カーソルＣＬＯＳＥ  TBL7700 TD1                          *
999999**********************************************************************
999999 K77-CLOSE-PROC                                          SECTION.
999999*
999999     EXEC ADABAS
999999          CLOSE K77
999999     END-EXEC.
999999*異常処理
999999     IF ADACODE NOT = ZERO
999999     THEN
999999         MOVE MSG-ERR-K77-CLOSE TO WK-USER-MSG
999999         MOVE 003               TO WK-RTN-CD
999999         PERFORM S99-ABEND-PROC
999999     ELSE
999999         CONTINUE
999999     END-IF.
999999 K77-CLOSE-PROC-EX.
999999     EXIT.
999999**********************************************************************
999999*   S99: ＤＢ異常終了処理                                           *
999999**********************************************************************
999999 S99-ABEND-PROC                                          SECTION.
999999 S99-ABEND-000.
999999*パラメータの設定
999999     MOVE CNS-PROGRAM-ID TO WK-KEY.
999999*
999999     CALL 'ABSUB'        USING WK-ABEND-AREA
999999                               DT1-PARAM
999999                               DT2-PARAM.
999999*
999999 S99-ABEND-999.
999999     EXIT.
