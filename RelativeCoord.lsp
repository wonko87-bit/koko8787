;; 상대좌표 출력 명령어
;; 사용법: RELCOORD 명령 입력 후 기준점 클릭, 대상점 클릭

(defun c:RELCOORD (/ base_pt target_pt dx dy dz)
  ;; 기준점 선택
  (setq base_pt (getpoint "\n기준점을 클릭하세요: "))

  (if base_pt
    (progn
      ;; 대상점 선택 (기준점 기준으로 고무줄 표시)
      (setq target_pt (getpoint base_pt "\n대상점을 클릭하세요: "))

      (if target_pt
        (progn
          ;; 상대좌표 계산
          (setq dx (- (car target_pt) (car base_pt)))
          (setq dy (- (cadr target_pt) (cadr base_pt)))
          (setq dz (- (caddr target_pt) (caddr base_pt)))

          ;; 결과 출력
          (princ "\n========================================")
          (princ "\n기준점 좌표:")
          (princ (strcat "\n  X: " (rtos (car base_pt) 2 4)))
          (princ (strcat "\n  Y: " (rtos (cadr base_pt) 2 4)))
          (princ (strcat "\n  Z: " (rtos (caddr base_pt) 2 4)))

          (princ "\n\n대상점 좌표:")
          (princ (strcat "\n  X: " (rtos (car target_pt) 2 4)))
          (princ (strcat "\n  Y: " (rtos (cadr target_pt) 2 4)))
          (princ (strcat "\n  Z: " (rtos (caddr target_pt) 2 4)))

          (princ "\n\n상대좌표 (대상점 - 기준점):")
          (princ (strcat "\n  ΔX: " (rtos dx 2 4)))
          (princ (strcat "\n  ΔY: " (rtos dy 2 4)))
          (princ (strcat "\n  ΔZ: " (rtos dz 2 4)))
          (princ "\n========================================")
        )
        (princ "\n대상점 선택이 취소되었습니다.")
      )
    )
    (princ "\n기준점 선택이 취소되었습니다.")
  )
  (princ)
)

;; 명령어 로드 메시지
(princ "\n상대좌표 명령어 로드 완료!")
(princ "\n사용법: RELCOORD 명령 입력")
(princ)
