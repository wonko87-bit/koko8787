;; 점 좌표를 누적 저장하고 CSV로 출력하는 명령어
;; 사용법: SAVEPOINTS 명령 입력 후 점들을 계속 클릭, Enter로 종료

(defun c:SAVEPOINTS (/ pt_list pt counter file_path file_handle)
  (setq pt_list '())
  (setq counter 0)

  (princ "\n========================================")
  (princ "\n점 좌표 저장 시작")
  (princ "\n점을 계속 클릭하세요 (Enter로 종료)")
  (princ "\n========================================")

  ;; 점들을 계속 입력받기
  (while (setq pt (getpoint "\n점을 클릭하세요 (Enter로 종료): "))
    (setq counter (1+ counter))
    (setq pt_list (append pt_list (list pt)))
    (princ (strcat "\n점 " (itoa counter) " 저장됨: "
                   "X=" (rtos (car pt) 2 4) ", "
                   "Y=" (rtos (cadr pt) 2 4) ", "
                   "Z=" (rtos (caddr pt) 2 4)))
  )

  ;; 저장된 점이 있으면 CSV 파일로 저장
  (if (> counter 0)
    (progn
      (princ (strcat "\n\n총 " (itoa counter) "개의 점이 입력되었습니다."))

      ;; 저장 경로 설정 (사용자가 선택)
      (setq file_path (getfiled "CSV 파일로 저장" "" "csv" 1))

      (if file_path
        (progn
          ;; 파일 열기
          (setq file_handle (open file_path "w"))

          (if file_handle
            (progn
              ;; 헤더 작성
              (write-line "번호,X,Y,Z" file_handle)

              ;; 각 점의 좌표 작성
              (setq counter 0)
              (foreach pt pt_list
                (setq counter (1+ counter))
                (write-line
                  (strcat
                    (itoa counter) ","
                    (rtos (car pt) 2 4) ","
                    (rtos (cadr pt) 2 4) ","
                    (rtos (caddr pt) 2 4)
                  )
                  file_handle
                )
              )

              ;; 파일 닫기
              (close file_handle)
              (princ (strcat "\n파일 저장 완료: " file_path))
            )
            (princ "\n오류: 파일을 열 수 없습니다.")
          )
        )
        (princ "\n파일 저장이 취소되었습니다.")
      )
    )
    (princ "\n입력된 점이 없습니다.")
  )

  (princ "\n========================================")
  (princ)
)

;; 상대좌표를 누적 저장하는 명령어 (기준점 대비)
(defun c:SAVERELPOINTS (/ base_pt pt_list pt counter dx dy dz file_path file_handle)
  ;; 기준점 선택
  (setq base_pt (getpoint "\n기준점을 클릭하세요: "))

  (if base_pt
    (progn
      (setq pt_list '())
      (setq counter 0)

      (princ "\n========================================")
      (princ "\n상대좌표 저장 시작")
      (princ (strcat "\n기준점: X=" (rtos (car base_pt) 2 4)
                     ", Y=" (rtos (cadr base_pt) 2 4)
                     ", Z=" (rtos (caddr base_pt) 2 4)))
      (princ "\n점을 계속 클릭하세요 (Enter로 종료)")
      (princ "\n========================================")

      ;; 점들을 계속 입력받기
      (while (setq pt (getpoint base_pt "\n점을 클릭하세요 (Enter로 종료): "))
        (setq counter (1+ counter))
        (setq pt_list (append pt_list (list pt)))

        ;; 상대좌표 계산
        (setq dx (- (car pt) (car base_pt)))
        (setq dy (- (cadr pt) (cadr base_pt)))
        (setq dz (- (caddr pt) (caddr base_pt)))

        (princ (strcat "\n점 " (itoa counter) " 저장됨 (상대): "
                       "ΔX=" (rtos dx 2 4) ", "
                       "ΔY=" (rtos dy 2 4) ", "
                       "ΔZ=" (rtos dz 2 4)))
      )

      ;; 저장된 점이 있으면 CSV 파일로 저장
      (if (> counter 0)
        (progn
          (princ (strcat "\n\n총 " (itoa counter) "개의 점이 입력되었습니다."))

          ;; 저장 경로 설정
          (setq file_path (getfiled "CSV 파일로 저장" "" "csv" 1))

          (if file_path
            (progn
              ;; 파일 열기
              (setq file_handle (open file_path "w"))

              (if file_handle
                (progn
                  ;; 기준점 정보 작성
                  (write-line "기준점 정보" file_handle)
                  (write-line
                    (strcat "기준점,X=" (rtos (car base_pt) 2 4)
                            ",Y=" (rtos (cadr base_pt) 2 4)
                            ",Z=" (rtos (caddr base_pt) 2 4))
                    file_handle
                  )
                  (write-line "" file_handle)

                  ;; 헤더 작성
                  (write-line "번호,절대X,절대Y,절대Z,상대ΔX,상대ΔY,상대ΔZ" file_handle)

                  ;; 각 점의 좌표 작성
                  (setq counter 0)
                  (foreach pt pt_list
                    (setq counter (1+ counter))
                    (setq dx (- (car pt) (car base_pt)))
                    (setq dy (- (cadr pt) (cadr base_pt)))
                    (setq dz (- (caddr pt) (caddr base_pt)))

                    (write-line
                      (strcat
                        (itoa counter) ","
                        (rtos (car pt) 2 4) ","
                        (rtos (cadr pt) 2 4) ","
                        (rtos (caddr pt) 2 4) ","
                        (rtos dx 2 4) ","
                        (rtos dy 2 4) ","
                        (rtos dz 2 4)
                      )
                      file_handle
                    )
                  )

                  ;; 파일 닫기
                  (close file_handle)
                  (princ (strcat "\n파일 저장 완료: " file_path))
                )
                (princ "\n오류: 파일을 열 수 없습니다.")
              )
            )
            (princ "\n파일 저장이 취소되었습니다.")
          )
        )
        (princ "\n입력된 점이 없습니다.")
      )
    )
    (princ "\n기준점 선택이 취소되었습니다.")
  )

  (princ "\n========================================")
  (princ)
)

;; 명령어 로드 메시지
(princ "\n점 저장 명령어 로드 완료!")
(princ "\n사용 가능 명령어:")
(princ "\n  SAVEPOINTS - 절대 좌표로 점 저장")
(princ "\n  SAVERELPOINTS - 상대 좌표로 점 저장")
(princ)
