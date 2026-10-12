; Optional reactor-based equivalent for reference; load after CNL_Acaddoc_Open.lsp.
(defun CNL_OnDrawingOpen (reactor params / doc)
  (setq doc
    (vla-get-ActiveDocument
      (vlax-get-acad-object)
    )
  )
  (CNL-LockDocumentViewports doc)
  (princ)
)

(setq *CNL-DOC-REACTOR*
  (vlr-docmanager-reactor
    nil
    '(
      (:vlr-documentCreated . CNL_OnDrawingOpen)
    )
  )
)

(setq *CNL-DWG-REACTOR*
  (vlr-dwg-reactor
    nil
    '(
      (:vlr-endDwgOpen . CNL_OnDrawingOpen)
    )
  )
)
