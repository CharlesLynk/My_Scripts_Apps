; Loaded from acaddoc.lsp to apply CNL settings whenever a drawing opens.
(vl-load-com)

(defun C:CNLUpdateUnits ()
  (mapcar 'setvar
    (list 'aunits 'auprec 'angdir 'lunits 'luprec 'snapmode)
    (list 0 2 0 2 4 0)
  )
  (princ)
)

(defun C:CNLBlockLayerZero (/ att_blk_ss blk att $att)
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

(defun CNL-LockDocumentViewports (doc / layers layer defpoints-layer layouts layout block entity)
  (setq layers (vla-get-Layers doc))
  (vlax-for layer layers
    (if (= (strcase (vla-get-Name layer)) "DEFPOINTS")
      (setq defpoints-layer layer)
    )
  )
  (if defpoints-layer
    (vla-put-Lock defpoints-layer :vlax-false)
  )
  (setq layouts (vla-get-Layouts doc))
  (vlax-for layout layouts
    (setq block (vla-get-Block layout))
    (vlax-for entity block
      (if (= (vla-get-ObjectName entity) "AcDbViewport")
        (vla-put-DisplayLocked entity :vlax-true)
      )
    )
  )
  (if (getvar "NAVVCUBEDISPLAY")
    (setvar "NAVVCUBEDISPLAY" 0)
  )
  (if defpoints-layer
    (vla-put-Lock defpoints-layer :vlax-true)
  )
  (princ)
)

(defun C:CNLLockViewports ()
  (CNL-LockDocumentViewports
    (vla-get-ActiveDocument (vlax-get-acad-object))
  )
  (princ)
)

(load "CNL_Plant3DAnnotations.lsp")

(CNL-LockDocumentViewports
  (vla-get-ActiveDocument (vlax-get-acad-object))
)
(princ)
