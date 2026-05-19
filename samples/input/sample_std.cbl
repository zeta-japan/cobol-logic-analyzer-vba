      *================================================================
      * SAMPLE01 - COBOLロジック階層可視化ツール 標準形式サンプル
      *================================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SAMPLE01.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WK-AGE         PIC 9(03).
       01 WK-KBN         PIC X(01).
       01 WK-FLG         PIC X(01).
       01 IN-CD          PIC X(02).
       01 WK-TBL.
          05 WK-ENT OCCURS 5 INDEXED BY WK-I.
             10 WK-ENT-CD   PIC X(02).
             10 WK-ENT-NM   PIC X(10).
       PROCEDURE DIVISION.
       MAIN-RTN SECTION.
       MAIN-START.
           PERFORM INIT-RTN.
           PERFORM CHECK-RTN.
           PERFORM SEARCH-RTN.
           PERFORM EDIT-RTN THRU EDIT-EXIT.
           CALL 'SUBPGM01' USING WK-KBN.
           STOP RUN.
       INIT-RTN.
           MOVE ZERO TO WK-AGE.
           MOVE SPACE TO WK-FLG.
       CHECK-RTN.
           IF WK-AGE = 0
               IF WK-KBN = '1'
                   MOVE 'A' TO WK-FLG
                   PERFORM UPDATE-RTN
               ELSE
                   MOVE 'B' TO WK-FLG
               END-IF
           ELSE
               PERFORM SKIP-RTN
           END-IF.
           EVALUATE WK-KBN
               WHEN '1'
                   PERFORM PROC-A
               WHEN '2'
                   PERFORM PROC-B
               WHEN OTHER
                   PERFORM PROC-OTHER
           END-EVALUATE.
       SEARCH-RTN.
           SET WK-I TO 1.
           SEARCH WK-ENT
               AT END
                   MOVE 'N' TO WK-FLG
               WHEN WK-ENT-CD (WK-I) = IN-CD
                   MOVE 'Y' TO WK-FLG
           END-SEARCH.
       UPDATE-RTN.
           MOVE '9' TO WK-KBN.
       SKIP-RTN.
           CONTINUE.
       PROC-A.
           MOVE 'A' TO WK-KBN.
       PROC-B.
           MOVE 'B' TO WK-KBN.
       PROC-OTHER.
           MOVE 'X' TO WK-KBN.
       EDIT-RTN.
           MOVE SPACE TO WK-FLG.
       EDIT-EXIT.
           EXIT.
