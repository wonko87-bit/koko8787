;; CoordinateCSV.lsp
;; 기준점 대비 상대좌표를 화면에 출력하는 AutoCAD LISP 프로그램
;; 사용법: 명령어 "COORDCSV" 입력 후 기준점 클릭, 이후 목표점들을 연속 클릭
;; ESC를 누르면 좌표 데이터가 화면에 출력됨 (복사해서 사용 가능)

(defun C:COORDCSV (/ base-pt target-pt point-list rel-x rel-y counter)
  (setq point-list '())

  ;; 기준점 입력
  (princ "\n기준점을 클릭하세요: ")
  (setq base-pt (getpoint))

  (if base-pt
    (progn
      (princ "\n기준점 설정 완료!")
      (princ (strcat "\n기준점 좌표: X=" (rtos (car base-pt) 2 3) ", Y=" (rtos (cadr base-pt) 2 3)))

      ;; 목표점 연속 입력
      (princ "\n\n목표점들을 클릭하세요 (ESC로 종료): ")
      (while (setq target-pt (getpoint base-pt))
        ;; 상대좌표 계산
        (setq rel-x (- (car target-pt) (car base-pt)))
        (setq rel-y (- (cadr target-pt) (cadr base-pt)))

        ;; 리스트에 추가
        (setq point-list (append point-list (list (list rel-x rel-y))))

        ;; 화면에 표시
        (princ (strcat "\n  상대좌표: X=" (rtos rel-x 2 3) ", Y=" (rtos rel-y 2 3)))
        (princ "\n다음 목표점 클릭 (ESC로 종료): ")
      )

      ;; 결과 출력
      (if (> (length point-list) 0)
        (progn
          ;; 텍스트 창 열기
          (textscr)

          (princ "\n\n")
          (princ "========================================")
          (princ "\n좌표 데이터 (아래 내용을 복사하세요)")
          (princ "\n========================================")
          (princ "\nNo,X,Y")

          ;; 데이터 출력
          (setq counter 1)
          (foreach pt point-list
            (princ
              (strcat
                "\n" (itoa counter) ","
                (rtos (car pt) 2 6) ","
                (rtos (cadr pt) 2 6)
              )
            )
            (setq counter (1+ counter))
          )

          (princ "\n========================================")
          (princ (strcat "\n총 " (itoa (length point-list)) "개의 좌표"))
          (princ "\n========================================")
          (princ "\n\n위 데이터를 Ctrl+C로 복사한 후 Excel 등에 붙여넣기 하세요.")
        )
        (princ "\n입력된 목표점이 없습니다.")
      )
    )
    (princ "\n취소되었습니다.")
  )

  (princ)
)

(princ "\n*** CoordinateCSV 로드됨 ***")
(princ "\n명령어: COORDCSV")
(princ "\n출력: 화면에 CSV 형식으로 출력 (복사 가능)")
(princ)
