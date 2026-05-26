999999 IDENTIFICATION                         DIVISION.
999999 PROGRAM-ID.       ICASE2.
000100**********************************************************************
000200* PROGRAM-ID   : ICASE2
000300* PROGRAM-NAME : 分岐構文 網羅サンプル (デモ)
000400* NOTE         : IF / EVALUATE / SEARCH / SEARCH ALL / PERFORM THRU /
000500*                CALL / GO TO / NEXT SENTENCE / 複合条件 を網羅
000600**********************************************************************
999999 ENVIRONMENT                            DIVISION.
999999 CONFIGURATION                          SECTION.
999999 DATA                                   DIVISION.
999999 WORKING-STORAGE                        SECTION.
999999 01  WK-AREA.
999999     03  WK-MODE         PIC X(01).
999999     03  WK-CODE         PIC X(02).
999999     03  WK-FLG          PIC X(01).
999999     03  WK-RANK         PIC X(01).
999999     03  WK-VAL          PIC 9(05).
999999     03  WK-IDX          PIC 9(02).
999999 01  WK-TBL.
999999     03  WK-ENT  OCCURS 10 ASCENDING KEY IS WK-ENT-CD INDEXED BY WK-I.
999999         05  WK-ENT-CD   PIC X(02).
999999         05  WK-ENT-NM   PIC X(10).
999999 01  IN-REC.
999999     03  IN-CODE         PIC X(02).
999999     03  IN-TYPE         PIC X(01).
999999     03  IN-QTY          PIC 9(05).
999999 PROCEDURE                              DIVISION.
000700**********************************************************************
000800*   ICASE2  （ メイン処理 ）
000900**********************************************************************
999999 ICASE2-PROC                            SECTION.
999999 ICASE2-000.
001000* 初期処理
999999         PERFORM  ICASE2-INIT.
001100* 主処理
999999         PERFORM  ICASE2-MAIN.
001200* 終了処理
999999         PERFORM  ICASE2-TERM THRU  ICASE2-TERM-EXIT.
999999 ICASE2-999.
999999     GOBACK.
001300**********************************************************************
001400*   ICASE2-INIT  （ 初期処理：IF / EVALUATE / SEARCH ALL ）
001500**********************************************************************
999999 ICASE2-INIT                            SECTION.
999999 ICASE2-INIT-000.
999999         MOVE  'A'    TO  WK-MODE.
999999         MOVE  ZERO   TO  WK-VAL.
001600* 入力チェック
999999         IF    IN-CODE  =  SPACE
999999             MOVE  'ER'  TO  WK-CODE
999999         ELSE
999999             MOVE  IN-CODE  TO  WK-CODE
999999         END-IF.
001700* 区分判定
999999         EVALUATE  WK-MODE
999999             WHEN  'A'
999999                 PERFORM  INIT-TYPE-A
999999             WHEN  'B'
999999                 PERFORM  INIT-TYPE-B
999999             WHEN  OTHER
999999                 PERFORM  INIT-TYPE-OTHER
999999         END-EVALUATE.
001800* コード検索 (二分探索)
999999         SEARCH ALL  WK-ENT
999999             AT END
999999                 MOVE  'NF'  TO  WK-CODE
999999             WHEN  WK-ENT-CD (WK-I)  =  IN-CODE
999999                 MOVE  WK-ENT-NM (WK-I)  TO  WK-RANK
999999         END-SEARCH.
999999 ICASE2-INIT-999.
999999     EXIT.
999999 INIT-TYPE-A.
999999         MOVE  '1'   TO  WK-FLG.
999999 INIT-TYPE-B.
999999         MOVE  '2'   TO  WK-FLG.
999999 INIT-TYPE-OTHER.
999999         MOVE  '9'   TO  WK-FLG.
001900**********************************************************************
002000*   ICASE2-MAIN  （ 主処理：6段ネスト + EVALUATE + SEARCH + GO TO ）
002100**********************************************************************
999999 ICASE2-MAIN                            SECTION.
999999 ICASE2-MAIN-000.
999999         CALL  'SUBCALC'  USING  WK-AREA.
002200* L1 複合条件 (AND)
999999         IF    WK-FLG  =  '1'  AND  WK-VAL  >  0
002300* L2
999999             IF    WK-CODE  =  'AA'
002400* L3
999999                 IF    WK-VAL  >  100
002500* L4 区分 (EVALUATE)
999999                     EVALUATE  WK-RANK
999999                         WHEN  'P'
002600* L5
999999                             IF    WK-VAL  >  1000
002700* L6
999999                                 IF    WK-VAL  >  10000
999999                                     PERFORM  MAIN-TOP
999999                                 ELSE
999999                                     PERFORM  MAIN-HIGH
999999                                 END-IF
999999                             ELSE
999999                                 PERFORM  MAIN-MID
999999                             END-IF
999999                         WHEN  'Q'
999999                             PERFORM  MAIN-Q
999999                         WHEN  OTHER
999999                             PERFORM  MAIN-OTHER
999999                     END-EVALUATE
999999                 ELSE
999999                     COMPUTE  WK-VAL  =  WK-VAL  +  1
999999                 END-IF
999999             ELSE
999999                 PERFORM  MAIN-CODE-OTHER
999999             END-IF
999999         ELSE
999999             GO TO  ICASE2-MAIN-EXIT
999999         END-IF.
002800* 単純 SEARCH + 各種命令 + NEXT SENTENCE + 複合条件 (OR)
999999         SET   WK-I  TO  1.
999999         SEARCH  WK-ENT
999999             AT END
999999                 NEXT SENTENCE
999999             WHEN  WK-ENT-CD (WK-I)  =  WK-CODE  OR  WK-IDX  >  5
999999                 READ     IN-REC
999999                 WRITE    IN-REC
999999                 REWRITE  IN-REC
999999         END-SEARCH.
999999 ICASE2-MAIN-EXIT.
999999     EXIT.
999999 MAIN-TOP.
999999         MOVE  'T'   TO  WK-RANK.
999999 MAIN-HIGH.
999999         MOVE  'H'   TO  WK-RANK.
999999 MAIN-MID.
999999         MOVE  'M'   TO  WK-RANK.
999999 MAIN-Q.
999999         MOVE  'Q'   TO  WK-RANK.
999999 MAIN-OTHER.
999999         MOVE  'X'   TO  WK-RANK.
999999 MAIN-CODE-OTHER.
999999         MOVE  'C'   TO  WK-RANK.
002900**********************************************************************
003000*   ICASE2-TERM  （ 終了処理 ）
003100**********************************************************************
999999 ICASE2-TERM                            SECTION.
999999 ICASE2-TERM-000.
999999         IF    WK-CODE  =  'NF'
999999             DELETE  IN-REC
999999         END-IF.
999999         MOVE  WK-RANK   TO  WK-FLG.
999999 ICASE2-TERM-EXIT.
999999     EXIT.
