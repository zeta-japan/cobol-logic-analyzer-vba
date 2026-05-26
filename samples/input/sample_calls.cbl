999999 IDENTIFICATION                         DIVISION.
999999 PROGRAM-ID.       ABC100.
000100**********************************************************************
000200* PROGRAM-ID   : ABC100
000300* PROGRAM-NAME : 呼出/利用関係 デモ (サブ3本 + ファイル)
000400**********************************************************************
999999 ENVIRONMENT                            DIVISION.
999999 INPUT-OUTPUT                           SECTION.
999999 FILE-CONTROL.
999999     SELECT TBL0005  ASSIGN TO  "TBL0005".
999999 DATA                                   DIVISION.
999999 FILE                                   SECTION.
999999 FD  TBL0005.
999999 01  TBL0005-REC         PIC X(80).
999999 WORKING-STORAGE                        SECTION.
999999 01  AABB210             PIC X(10).
999999 01  AABB220             PIC 9(05).
999999 01  AABB310             PIC X(02).
999999 01  AABB320             PIC 9(03).
999999 01  AABB410             PIC X(01).
999999 01  RTNCODE             PIC X(02).
999999 01  RTN050              PIC 9(04).
999999 01  RTN060              PIC 9(04).
999999 PROCEDURE                              DIVISION.
000500**********************************************************************
000600*   ABC100  （ メイン処理 ）
000700**********************************************************************
999999 ABC100-PROC                            SECTION.
999999 ABC100-000.
000800* ファイル読込
999999         OPEN  INPUT  TBL0005.
999999         READ  TBL0005.
000900* サブ呼出 1
999999         CALL  'SUB001'  USING  AABB210  AABB220.
001000* サブ呼出 2
999999         CALL  'SUB002'  USING  AABB310  AABB320.
001100* サブ呼出 3
999999         CALL  'SUB003'  USING  AABB410.
001200* ファイル更新
999999         REWRITE  TBL0005-REC.
999999         CLOSE  TBL0005.
999999 ABC100-999.
999999     GOBACK.
