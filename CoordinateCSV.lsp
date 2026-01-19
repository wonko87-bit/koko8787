;; CoordinateCSV.lsp
;; 기준점 대비 상대좌표를 CSV 파일로 저장하는 AutoCAD LISP 프로그램
;; 사용법: 명령어 "COORDCSV" 입력 후 기준점 클릭, 이후 목표점들을 연속 클릭
;; ESC를 누르면 CSV 파일이 생성됨

;; 타임스탬프를 문자열로 변환하는 헬퍼 함수
(defun get-timestamp (/ julian-date date-str time-str)
  (setq julian-date (getvar "DATE"))
  ;; Julian date를 문자열로 변환 (소수점 제거)
  (setq date-str (rtos julian-date 2 9))
  ;; 소수점과 공백 제거하여 고유한 문자열 생성
  (setq date-str (vl-string-translate ". " "" date-str))
  date-str
)

(defun C:COORDCSV (/ base-pt target-pt point-list rel-x rel-y csv-path csv-file counter file-handle timestamp)
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

      ;; CSV 파일 저장
      (if (> (length point-list) 0)
        (progn
          ;; CSV 저장 경로 설정
          (setq csv-path (strcat (getenv "USERPROFILE") "\\Desktop\\temp_csv\\"))

          ;; 디렉토리가 없으면 생성
          (if (not (vl-file-directory-p csv-path))
            (vl-mkdir csv-path)
          )

          ;; 파일명 생성 (타임스탬프 사용)
          (setq timestamp (get-timestamp))
          (setq csv-file (strcat csv-path "coordinates_" timestamp ".csv"))

          ;; CSV 파일 쓰기
          (setq file-handle (open csv-file "w"))
          (if file-handle
            (progn
              ;; 헤더 작성
              (write-line "No,X,Y" file-handle)

              ;; 데이터 작성
              (setq counter 1)
              (foreach pt point-list
                (write-line
                  (strcat
                    (itoa counter) ","
                    (rtos (car pt) 2 6) ","
                    (rtos (cadr pt) 2 6)
                  )
                  file-handle
                )
                (setq counter (1+ counter))
              )

              (close file-handle)
              (princ (strcat "\n\nCSV 파일 저장 완료: " csv-file))
              (princ (strcat "\n총 " (itoa (length point-list)) "개의 좌표가 저장되었습니다."))
            )
            (princ "\n오류: CSV 파일을 생성할 수 없습니다.")
          )
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
(princ)
