;Acaddoc.lsp
; On every document open, defines a unit update command, locks and hides the NAV Cube in all viewports, and loads the annotation scripts
(defun-q MYSTARTUP () 
  (defun C:CNLUpdateUnits ()
    (mapcar 'setvar (list 'aunits 'auprec 'angdir 'lunits 'luprec 'snapmode) (list 0 2 0 2 4 0))
    (PRINC)
  )
  (defun C:BlockLayerZero ( / att_blk_ss blk att $att)
	(setq att_blk_ss (ssget "X" '((0 . "INSERT") (66 . 1))))
	(while (/= (ssname att_blk_ss 0) nil)
		(setq blk (ssname att_blk_ss 0)
			att (entnext blk)
		)
		(while (and (/= att nil) (/= (cdr (assoc 0 (entget att))) "SEQEND"))
			(setq $att (vlax-ename->vla-object att))
				(if (= (vla-get-layer $att) oldlayername)
					(vla-put-layer $att newlayername)
			)
			(setq att (entnext att))
		)
		(setq att_blk_ss (ssdel (ssname att_blk_ss 0) att_blk_ss))
	)
	(princ)
  )
  (defun C:CNLLockViewports (/ ss n vp); = Lock All Viewports
    ;(vl-load-com)
    (command-s "-layer" "unlock" "DEFPOINTS" "")
    (repeat (setq n (sslength (setq ss (ssget "_X" '((0 . "VIEWPORT"))))))
      (if (> (cdr (assoc 69 (entget (setq vp (ssname ss (setq n (1- n))))))) 1); not the Paper Space Viewport of its Layout
        (progn
          (vla-put-DisplayLocked (vlax-ename->vla-object vp) -1); lock it
          ;(if (/= (getvar "navbardisplay") nil)
          ;  (setvar 'navbardisplay 0)
          ;)
          (if (/= (getvar "navvcubedisplay") nil)
            (setvar 'navvcubedisplay 0)
          )
        )
      ); if
    ); repeat
    (command-s "-layer" "lock" "DEFPOINTS" "")
    (command)
  ); defun
  (C:CNLLockViewports)
  (load "Plant3DAnnotations.lsp")
   
  
)
(setq S::STARTUP (append S::STARTUP MYSTARTUP))