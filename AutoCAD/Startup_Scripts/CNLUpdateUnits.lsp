(defun C:CNLUpdateUnits (/ ss1)
  (setq ss1 nil)
  (mapcar 'setvar (list 'aunits 'auprec 'angdir 'lunits 'luprec 'snapmode) (list 0 2 0 2 4 0))
  (command-s "._pspace")
  ;(setq ss1 (ssget "X"'((0 . "VIEWPORT"))))
  ;(command-s "mview" "LOCK" "ON" ss1 "")
  ;(PRINC)
)

;(defun vplock (/ ss lock x obj)
;  (vl-load-com)
;  (setq ss (ssget "X" '((0 . "VIEWPORT"))))
;  (progn
;    (setq lock :vlax-true)
;    (setq x 0)
;    (while (< x (sslength ss))
;      (setq obj (vlax-ename->vla-object (ssname ss x)))
;      (vla-put-displaylocked obj lock)
;      (setq x (1+ x))
;    )
;  )
;)
;  (princ)
;  )
;(vplock "y")

;(defun C:CNLLockViewports (/ ss n vp); = Lock All Viewports
;  (repeat (setq n (sslength (setq ss (ssget "_X" '((0 . "VIEWPORT"))))))
;    (if (> (cdr (assoc 69 (entget (setq vp (ssname ss (setq n (1- n))))))) 1); not the Paper Space Viewport of its Layout
;      (vla-put-DisplayLocked (vlax-ename->vla-object vp) -1); lock it
;
;    ); if
;  ); repeat
;); defun
;(C:CNLLockViewports)

(defun C:CNLUpdateUnits (/ ss1)
  (setq ss1 nil)
  (mapcar 'setvar (list 'aunits 'auprec 'angdir 'lunits 'luprec 'snapmode) (list 0 2 0 2 4 0))
  (command-s "._pspace")
  ;(setq ss1 (ssget "X"'((0 . "VIEWPORT"))))
  ;(command-s "mview" "LOCK" "ON" ss1 "")
  ;(PRINC)
)