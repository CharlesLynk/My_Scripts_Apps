(defun batch:fail (message)
  (princ (strcat "\nBatch error: " message))
  (exit)
)

(defun batch:ensure-xref (/ record flags inserts)
  (if (not (findfile *batch-xref-path*))
    (batch:fail (strcat "Border XREF file cannot be found: " *batch-xref-path*))
  )

  (setq record (tblsearch "BLOCK" *batch-xref-name*))
  (if record
    (if (/= 0 (logand 4 (cdr (assoc 70 record))))
      (command "_.-XREF" "_Detach" *batch-xref-name*)
      (batch:fail (strcat "XREF name is already used by a non-XREF block: " *batch-xref-name*))
    )
  )

  (command "_.PSPACE")
  (command "_.-XREF" "_Overlay" *batch-xref-path* "0,0" "1" "1" "0")
  (command "_.REGEN")

  (setq record (tblsearch "BLOCK" *batch-xref-name*)
        flags (if record (cdr (assoc 70 record)) 0)
  )
  (if (= 0 (logand 4 flags))
    (batch:fail (strcat "Border XREF was not attached: " *batch-xref-name*))
  )
  (if (= 0 (logand 32 flags))
    (batch:fail (strcat "Border XREF did not resolve: " *batch-xref-name* " -> " *batch-xref-path*))
  )
  (setq inserts
    (ssget "_X"
      (list '(0 . "INSERT") (cons 2 *batch-xref-name*) '(67 . 1))
    )
  )
  (if (null inserts)
    (batch:fail (strcat "Border XREF was not inserted in paper space: " *batch-xref-name*))
  )
  (princ)
)

(defun batch:validate-xrefs (/ record flags xref-name xref-path xref-count)
  (setq record (tblnext "BLOCK" T)
        xref-count 0
  )
  (while record
    (setq flags (cdr (assoc 70 record)))
    (if (/= 0 (logand 4 flags))
      (progn
        (setq xref-count (1+ xref-count)
              xref-name (cdr (assoc 2 record))
              xref-path (cdr (assoc 1 record))
        )
        (if (= 0 (logand 32 flags))
          (batch:fail
            (strcat
              "Template XREF is unresolved: "
              xref-name
              (if xref-path (strcat " -> " xref-path) "")
            )
          )
        )
      )
    )
    (setq record (tblnext "BLOCK"))
  )
  (if (= xref-count 0)
    (batch:fail "No attached XREFs were found in the template.")
  )
  (princ)
)

(defun batch:find-paper-space-inserts (block-name / selection)
  (setq selection
    (ssget "_X"
      (list '(0 . "INSERT") (cons 2 block-name) '(67 . 1))
    )
  )
  selection
)

(defun batch:validate-mappings (/ mapping block-name attributes selection index insert attribute tag found found-all pair successes failures missing-handles handle handles-text failure-summary mapping-name)
  (setq successes nil
        failures nil
  )
  (foreach mapping *batch-mappings*
    (setq block-name (car mapping)
          attributes (cdr mapping)
          selection (batch:find-paper-space-inserts block-name)
    )
    (if (null selection)
      (foreach pair attributes
        (setq failures
          (cons
            (strcat block-name "." (car pair) " (no top-level paper-space insert for block)")
            failures
          )
        )
      )
      (progn
        (foreach pair attributes
          (setq found-all T
                missing-handles nil
                index 0
          )
          (while (< index (sslength selection))
            (setq insert (ssname selection index)
                  attribute (entnext insert)
                  found nil
                  handle (cdr (assoc 5 (entget insert)))
            )
            (while (and attribute (/= (cdr (assoc 0 (entget attribute))) "SEQEND"))
              (setq tag (strcase (cdr (assoc 2 (entget attribute)))))
              (if (= tag (strcase (car pair)))
                (setq found T)
              )
              (setq attribute (entnext attribute))
            )
            (if (not found)
              (progn
                (setq found-all nil)
                (setq missing-handles (cons handle missing-handles))
              )
            )
            (setq index (1+ index))
          )
          (setq mapping-name (strcat block-name "." (car pair)))
          (if found-all
            (setq successes (cons mapping-name successes))
            (progn
              (setq handles-text "")
              (foreach handle (reverse missing-handles)
                (setq handles-text
                  (strcat
                    handles-text
                    (if (= handles-text "") "" ", ")
                    (if handle handle "<unknown>")
                  )
                )
              )
              (setq failures
                (cons
                  (strcat mapping-name " (missing on insert handle(s): " handles-text ")")
                  failures
                )
              )
            )
          )
        )
      )
    )
  )

  (princ "\nCSV attribute mapping results:")
  (foreach mapping-name (reverse successes)
    (princ (strcat "\n  MAPPED: " mapping-name))
  )
  (foreach mapping-name (reverse failures)
    (princ (strcat "\n  MISSING: " mapping-name))
  )

  (if failures
    (progn
      (setq failure-summary "")
      (foreach mapping-name (reverse failures)
        (setq failure-summary
          (strcat
            failure-summary
            (if (= failure-summary "") "" "; ")
            mapping-name
          )
        )
      )
      (batch:fail
        (strcat
          "CSV attribute mapping failed. Missing mappings: "
          failure-summary
        )
      )
    )
  )
  (princ)
)

(defun batch:set-attributes (/ mapping block-name attributes selection index insert attribute data tag pair changed)
  (foreach mapping *batch-mappings*
    (setq block-name (car mapping)
          attributes (cdr mapping)
          selection (batch:find-paper-space-inserts block-name)
    )
    (if (null selection)
      (batch:fail (strcat "No top-level paper-space insert found for block: " block-name))
    )

    (setq index 0)
    (while (< index (sslength selection))
      (setq insert (ssname selection index)
            attribute (entnext insert)
      )
      (while (and attribute (/= (cdr (assoc 0 (entget attribute))) "SEQEND"))
        (setq data (entget attribute)
              tag (strcase (cdr (assoc 2 data)))
        )
        (foreach pair attributes
          (if (= tag (strcase (car pair)))
            (if (assoc 1 data)
              (if (not (entmod (subst (cons 1 (cdr pair)) (assoc 1 data) data)))
                (batch:fail (strcat "Could not update attribute " (car pair) " in block " block-name))
              )
              (batch:fail (strcat "Attribute has no value field: " (car pair)))
            )
          )
        )
        (setq attribute (entnext attribute))
      )
      (entupd insert)
      (setq index (1+ index))
    )
  )
  (princ)
)

(defun batch:write-success (/ result-file)
  (setq result-file (open *batch-result-file* "w"))
  (if (null result-file)
    (batch:fail (strcat "Could not write job result file: " *batch-result-file*))
  )
  (write-line "SUCCESS" result-file)
  (close result-file)
  (princ)
)

(defun batch:save-as (drawing-path / old-filedia)
  (setq old-filedia (getvar "FILEDIA"))
  (setvar "FILEDIA" 0)
  (command "_.SAVEAS" "" drawing-path)
  (setvar "FILEDIA" old-filedia)
  (if (not (findfile drawing-path))
    (batch:fail (strcat "SAVEAS did not create drawing: " drawing-path))
  )
  (princ)
)

(defun c:BATCHPREFLIGHT ()
  (batch:ensure-xref)
  (batch:validate-xrefs)
  (batch:validate-mappings)
  (command "_.QSAVE")
  (batch:write-success)
  (princ "\nTemplate and border XREF preflight succeeded.")
  (princ)
)

(defun c:CREATEFROMCSV ()
  (batch:validate-mappings)
  (batch:set-attributes)
  (command "_.QSAVE")
  (batch:write-success)
  (princ "\nBatch drawing creation succeeded.")
  (princ)
)

(defun c:CREATEBATCH (/ job)
  (if (null *batch-jobs*)
    (batch:fail "No drawing jobs were supplied.")
  )
  (setq *batch-mappings* (caddr (car *batch-jobs*)))
  (batch:validate-mappings)
  (foreach job *batch-jobs*
    (setq *batch-mappings* (caddr job)
          *batch-result-file* (cadr job)
    )
    (princ (strcat "\nCreating drawing: " (car job)))
    (batch:set-attributes)
    (batch:save-as (car job))
    (batch:write-success)
  )
  (princ "\nBatch drawing generation succeeded.")
  (princ)
)
