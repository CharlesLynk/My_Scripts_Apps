
(vl-load-com)

(setq *ARCH_D_CANON* "ARCH_full_bleed_D_(36.00_x_24.00_Inches)")
(setq *ANSI_D_CANON* "ANSI_full_bleed_D_(34.00_x_22.00_Inches)")

(defun _basename (p) (vl-filename-base p))
(defun _ensure-dir (p) (if (and p (not (vl-file-directory-p p))) (vl-mkdir p)))
(defun _norm (p) (vl-string-translate "/" "\\" p))
(defun _ci= (a b) (= (strcase a) (strcase b)))
(defun _truep (x)
  (cond ((null x) nil)
        ((= (type x) 'INT) (/= x 0))
        ((= (type x) 'REAL) (/= x 0.0))
        ((= (type x) 'STR) (not (wcmatch (strcase x) "0,FALSE,NO,N,OFF,\"\"")))
        ((= x :vlax-true) T)
        ((= x :vlax-false) nil)
        (T T)))

(setq *log-h* nil)
(defun _log-open (outPdf / outDir tryPath tempPath)
  (setq outDir (vl-filename-directory outPdf))
  (setq tempPath (strcat (getenv "TEMP") "\\_publish_log.txt"))
  (_ensure-dir outDir)
  (setq tryPath (strcat outDir "\\_publish_log.txt"))
  (setq *log-h* (open tryPath "w"))
  (if (null *log-h*) (setq *log-h* (open tempPath "w")))
  (if *log-h*
    (progn
      (write-line (strcat "Publish log - " (menucmd "M=$(edtime,$(getvar,date),YYYY-MM-DD HH:MM:SS)")) *log-h*)
      (write-line "------------------------------------------------------------" *log-h*)
    )
  )
)
(defun _log (s)
  (if *log-h* (write-line s *log-h*))
  (if (getenv "PUBLISH_VERBOSE") (princ (strcat "\nLOG>> " s)))
)
(defun _log-close () (if *log-h* (close *log-h*)) (setq *log-h* nil))

(defun _create-dbx ( / app obj rc progids done)
  (vl-load-com) (setq obj nil done nil)
  (setq rc (vl-catch-all-apply 'vlax-get-acad-object '()))
  (if (not (vl-catch-all-error-p rc)) (setq app rc))
  (if (and app (vlax-method-applicable-p app 'Release))
    (progn
      (setq rc (vl-catch-all-apply 'vla-GetInterfaceObject (list app "ObjectDBX.AxDbDocument")))
      (if (not (vl-catch-all-error-p rc)) (setq obj rc done t))))
  (if (not done)
    (progn
      (setq progids '("ObjectDBX.AxDbDocument" "ObjectDBX.AxDbDocument.25" "ObjectDBX.AxDbDocument.25.1"
                      "ObjectDBX.AxDbDocument.24.2" "ObjectDBX.AxDbDocument.24"))
      (while (and progids (null obj))
        (setq rc (vl-catch-all-apply 'vlax-create-object (list (car progids))))
        (if (not (vl-catch-all-error-p rc)) (setq obj rc))
        (setq progids (cdr progids)))))
  (if (not (and obj (vlax-method-applicable-p obj 'Release)))
    (progn (setq obj nil) (_log "ERROR: Could not obtain a valid ObjectDBX instance.")))
  obj)

(defun _open-dbx (dwg / dbx rc)
  (setq dbx (_create-dbx))
  (if (null dbx) (progn (_log (strcat "ERROR: Unable to create ObjectDBX for: " dwg)) nil)
    (progn
      (setq rc (vl-catch-all-apply 'vla-Open (list dbx dwg)))
      (if (vl-catch-all-error-p rc)
        (progn (_log (strcat "ERROR: DBX open failed: " dwg "\n " (vl-catch-all-error-message rc)))
               (vlax-release-object dbx) nil)
        dbx))))

(defun _paper-layouts-of (dwg / dbx lays out lst media rc)
  (setq lst '())
  (setq dbx (_open-dbx dwg))
  (if (null dbx) (progn (_log (strcat "SKIP: Could not open (paper layouts not read): " dwg)) lst)
    (progn
      (setq rc (vl-catch-all-apply 'vla-get-Layouts (list dbx)))
      (if (vl-catch-all-error-p rc) (progn (_log (strcat "ERROR: No Layouts collection in: " dwg)) (setq lays nil))
        (setq lays rc))
      (if lays (vlax-for out lays
        (if (not (= :vlax-true (vla-get-ModelType out)))
          (progn
            (setq media (vl-catch-all-apply 'vla-get-CanonicalMediaName (list out)))
            (if (vl-catch-all-error-p media) (setq media ""))
            (setq lst (cons (list (vla-get-Name out) media) lst))))))
      (if lays (vlax-release-object lays))
      (vlax-release-object dbx)
      (reverse lst))))

(defun _plotcfg-name (pcfg) (vlax-get-property pcfg 'Name))
(defun _plotcfg-device (pcfg) (vlax-get-property pcfg 'ConfigName))
(defun _plotcfg-stylesheet (pcfg)
  (if (vl-catch-all-error-p (vl-catch-all-apply 'vlax-get-property (list pcfg 'StyleSheet))) ""
    (vlax-get-property pcfg 'StyleSheet)))

(defun _find-ps-by-name-in-dwg (dwg name / dbx cfgs hit rc)
  (setq hit nil)
  (setq dbx (_open-dbx dwg))
  (if (null dbx) nil
    (progn
      (setq rc (vl-catch-all-apply 'vla-get-PlotConfigurations (list dbx)))
      (if (vl-catch-all-error-p rc) (setq cfgs nil) (setq cfgs rc))
      (if cfgs (vlax-for pc cfgs (if (_ci= (_plotcfg-name pc) name) (setq hit name))))
      (if cfgs (vlax-release-object cfgs))
      (vlax-release-object dbx)
      hit)))

(defun _find-ps-by-name-in-folder (folder name / files full nm)
  (setq files (vl-directory-files folder "*.dwg" 1))
  (setq full (mapcar '(lambda (x) (strcat folder "\\" x)) files))
  (setq nm nil)
  (while (and full (null nm))
    (if (_find-ps-by-name-in-dwg (car full) name)
      (setq nm (cons (car full) name))
      (setq full (cdr full))))
  nm)

(defun _find-ps-by-criteria-in-dwg (dwg dev ctb / dbx cfgs pick rc)
  (setq pick nil)
  (setq dbx (_open-dbx dwg))
  (if (null dbx) nil
    (progn
      (setq rc (vl-catch-all-apply 'vla-get-PlotConfigurations (list dbx)))
      (if (vl-catch-all-error-p rc) (setq cfgs nil) (setq cfgs rc))
      (if cfgs (vlax-for pc cfgs
        (if (and (or (null dev) (= dev "") (_ci= (_plotcfg-device pc) dev))
                 (or (null ctb) (= ctb "") (_ci= (_plotcfg-stylesheet pc) ctb)))
          (if (null pick) (setq pick (cons dwg (_plotcfg-name pc)))))))
      (if cfgs (vlax-release-object cfgs))
      (vlax-release-object dbx)
      pick)))

(defun _find-ps-by-criteria-in-folder (folder dev ctb / files full pick)
  (setq files (vl-directory-files folder "*.dwg" 1))
  (setq full (mapcar '(lambda (x) (strcat folder "\\" x)) files))
  (setq pick nil)
  (while (and full (null pick))
    (setq pick (_find-ps-by-criteria-in-dwg (car full) dev ctb))
    (setq full (cdr full)))
  pick)

(defun _first-some (a b c) (cond (a a) (b b) (c c) (t nil)))

(defun _choose-nps-for-media (media nps2234 nps2436 npsGeneric)
  (cond
    ((and media (_ci= media *ARCH_D_CANON*)) (_first-some nps2436 npsGeneric nps2234))
    ((and media (_ci= media *ANSI_D_CANON*)) (_first-some nps2234 npsGeneric nps2436))
    (t (_first-some npsGeneric nps2234 nps2436))))

(defun _count-target-sizes (sheets / ansi arch med)
  (setq ansi 0 arch 0)
  (foreach s sheets
    (setq med (nth 2 s))
    (if (and med (_ci= med *ANSI_D_CANON*)) (setq ansi (+ ansi 1)))
    (if (and med (_ci= med *ARCH_D_CANON*)) (setq arch (+ arch 1))))
  (list ansi arch))

(defun _pref-normalize (s)
  (cond ((and s (_ci= s "ANSI")) "ANSI")
        ((and s (_ci= s "ARCH")) "ARCH")
        (t nil)))

(defun _select-by-pref (pref nps2234 nps2436 npsGeneric)
  (cond
    ((and pref (_ci= pref "ANSI"))
      (cond (nps2234 nps2234)
            (nps2436 (_log "Chosen ANSI not found; falling back to ARCH.") nps2436)
            (npsGeneric (_log "Chosen ANSI not found; using generic fallback.") npsGeneric)))
    ((and pref (_ci= pref "ARCH"))
      (cond (nps2436 nps2436)
            (nps2234 (_log "Chosen ARCH not found; falling back to ANSI.") nps2234)
            (npsGeneric (_log "Chosen ARCH not found; using generic fallback.") npsGeneric)))
    (t nil)))
(defun _canonical-from-override (ovr)
  (cond
    ((null ovr) "")
    ((= (strcase ovr) "ANSI") *ANSI_D_CANON*)
    ((= (strcase ovr) "ARCH") *ARCH_D_CANON*)
    (t ovr))) ;; pass-through for non-canonical labels from Probe

(defun _make-ps-lib (deviceName styleTable mediaLabel / dbx coll pc tmp path name rc med)
  ;; Create a tiny DWG in %TEMP% that contains a single page setup we can reference in DSD.
  (setq path (strcat (getenv "TEMP") "\\_ps_lib_" (menucmd "M=$(edtime,$(getvar,date),YYYYMMDD-HHMMSS)") ".dwg"))
  (setq name "_PublishTemp")
  (setq med (_canonical-from-override mediaLabel))
  (setq dbx (_create-dbx))
  (if (null dbx)
    (progn (_log "ERROR: DBX not available for PS library.") nil)
    (progn
      ;; Write an empty DWG file we can modify
      (setq rc (vl-catch-all-apply 'vla-SaveAs (list dbx path)))
      (if (vl-catch-all-error-p rc)
        (progn (_log (strcat "ERROR: SaveAs failed for PS library: " path)) (vlax-release-object dbx) nil)
        (progn
          (setq coll (vl-catch-all-apply 'vla-get-PlotConfigurations (list dbx)))
          (if (vl-catch-all-error-p coll)
            (progn (_log "ERROR: PlotConfigurations not accessible in PS library.") (vlax-release-object dbx) nil)
            (progn
              ;; paperspace type = :vlax-false
              (setq pc (vl-catch-all-apply 'vla-Add (list coll name :vlax-false)))
              (if (vl-catch-all-error-p pc)
                (progn (_log "ERROR: Could not add PlotConfiguration to PS library.") (vlax-release-object coll) (vlax-release-object dbx) nil)
                (progn
                  ;; Device + CTB
                  (vl-catch-all-apply 'vlax-put-property (list pc 'ConfigName deviceName))
                  (if (and styleTable (/= styleTable "")) (vl-catch-all-apply 'vlax-put-property (list pc 'StyleSheet styleTable)))
                  ;; Media (optional; use when provided)
                  (if (and med (/= med "")) (vl-catch-all-apply 'vlax-put-property (list pc 'CanonicalMediaName med)))
                  ;; Light defaults (avoid enums to keep CC-safe)
                  (vl-catch-all-apply 'vlax-put-property (list pc 'CenterPlot :vlax-true))
                  ;; Save and release
                  (vl-catch-all-apply 'vla-Save (list dbx))
                  (if coll (vlax-release-object coll))
                  (vlax-release-object dbx)
                  (cons path name)
                )
              )
            )
          )
        )
      )
    )
  )
)
(setq *strict_fail* nil)
(defun _build-dsd (dsdPath sheets outPdf nps2234 nps2436 npsGeneric / fh outDir selPair dwg lay med npsName npsSrc
                           _cntPS _cntLayout strictOk)
  (setq outDir (vl-filename-directory outPdf))
  (_ensure-dir outDir)
  (setq fh (open dsdPath "w"))
  (if (null fh) (progn (princ "\nERROR: Cannot write DSD.") (exit)))
  (setq _cntPS 0 _cntLayout 0)
  (write-line "[DWF6Version]" fh) (write-line "Ver=1" fh)
  (write-line "[DWF6MinorVersion]" fh) (write-line "MinorVer=1" fh)
  (foreach s sheets
    (setq dwg (nth 0 s) lay (nth 1 s) med (nth 2 s))
    (setq selPair (_choose-nps-for-media med nps2234 nps2436 npsGeneric))
    (setq npsName (if selPair (cdr selPair) ""))
    (setq npsSrc  (if selPair (car selPair) ""))
    (if selPair
      (progn
        (setq _cntPS (+ _cntPS 1))
        (_log (strcat "Sheet: " (_basename dwg) " " lay " Media=" med " -> Nps=" npsName " (src=" npsSrc ")")))
      (progn
        (setq _cntLayout (+ _cntLayout 1))
        (_log (strcat "Sheet: " (_basename dwg) " " lay " Media=" med " -> Nps=<none> (using layout)"))))
    (write-line (strcat "[DWF6Sheet:" (_basename dwg) "-" lay "]") fh)
    (write-line (strcat "DWG=" dwg) fh)
    (write-line (strcat "Layout=" lay) fh)
    (if selPair
      (write-line (strcat "Setup=" npsName "|" npsSrc) fh)
      (write-line "Setup=" fh))
    (write-line (strcat "OriginalSheetPath=" dwg) fh)
    (write-line "Has Plot Port=0" fh)
    (write-line "Has3DDWF=0" fh))
  (write-line "[Target]" fh)
  (write-line "Type=6" fh)
  (write-line (strcat "DWF=" outPdf) fh)
  (write-line (strcat "OUT=" outDir) fh)
  (write-line "PWD=" fh)
  (write-line "[PdfOptions]" fh)
  (write-line "IncludeHyperlinks=TRUE" fh)
  (write-line "CreateBookmarks=TRUE" fh)
  (write-line "CaptureFontsInDrawing=TRUE" fh)
  (write-line "ConvertTextToGeometry=FALSE" fh)
  (write-line "VectorResolution=1200" fh)
  (write-line "RasterResolution=400" fh)
  (close fh)
  (_log (strcat "Summary: explicit page setups=" (itoa _cntPS) ", using layout=" (itoa _cntLayout)))
  (if (and (_truep (getenv "PUBLISH_STRICT")) (> _cntLayout 0))
    (progn
      (setq *strict_fail* T)
      (_log "STRICT: one or more sheets have no resolvable page setup; will abort before publish.")))
  dsdPath)

(defun _publish-dsd (dsdPath outPdf / ok)
  (setvar "BACKGROUNDPLOT" 0)
  (setvar "PUBLISHCOLLATE" 1)
  (setvar "FILEDIA" 0)
  (command "-PUBLISH" dsdPath)
  (if (and (setq ok (findfile outPdf)) (> (vl-file-size outPdf) 0) (null (getenv "PUBLISH_KEEP_DSD")))
    (progn (vl-file-delete dsdPath) (_log (strcat "Deleted temporary DSD: " dsdPath)))
    (_log (strcat "Keeping DSD for troubleshooting: " dsdPath)))
  (princ))

;; ---- Font preflight helpers (unchanged) ----
(defun _file-exists (p)
  (if (and p (/= p "")) (not (null (vl-catch-all-apply 'vl-file-systime (list p)))) nil))

(defun _font-file-exists-p (fname / up winfonts ttfPath)
  (cond
    ((or (null fname) (= fname "")) T)
    ((findfile fname) T)
    (T
     (setq up (strcase fname)
           winfonts (strcat (getenv "WINDIR") "\\Fonts\\"))
     (or
       (and (not (wcmatch up "*.TTF"))
            (or (findfile fname)
                (findfile (strcat fname ".shx"))))
       (progn
         (setq ttfPath (if (wcmatch up "*.TTF") fname (strcat fname ".ttf")))
         (_file-exists (strcat winfonts ttfPath)))))))

(defun _collect_styles_fonts (dbx / tsColl rec lst sname pfont bfont)
  (setq lst '())
  (setq tsColl (vl-catch-all-apply 'vla-get-TextStyles (list dbx)))
  (if (not (vl-catch-all-error-p tsColl))
    (vlax-for rec tsColl
      (setq sname (vla-get-Name rec)
            pfont (vla-get-FontFile rec)
            bfont (vla-get-BigFontFile rec)
            lst   (cons (list sname pfont bfont) lst))))
  (reverse lst))

(defun _collect_mtext_inline_fonts (dbx / fonts blocks blk ent txt hit rx)
  (setq rx (vlax-create-object "VBScript.RegExp"))
  (vlax-put rx 'Pattern "\\\\f\\([^;]\\+\\);")
  (vlax-put rx 'Global :vlax-true)
  (vlax-put rx 'IgnoreCase :vlax-true)
  (setq fonts '())
  (setq blocks (vl-catch-all-apply 'vla-get-Blocks (list dbx)))
  (if (not (vl-catch-all-error-p blocks))
    (vlax-for blk blocks
      (vlax-for ent blk
        (if (and (= (vlax-get ent 'ObjectName) "AcDbMText")
                 (setq txt (vlax-get ent 'TextString))
                 (setq hit (vlax-invoke rx 'Execute txt)))
          (vlax-for m hit
            (setq fonts (cons (vlax-get (vlax-get m 'SubMatches) 0) fonts)))))))
  (vlax-release-object rx)
  (setq fonts (reverse fonts))
  (setq fonts (mapcar 'strcase fonts))
  fonts)

(defun _uniq-ci (lst / acc) (setq acc '())
  (foreach x lst (if (not (member (strcase x) (mapcar 'strcase acc))) (setq acc (cons x acc))))
  (reverse acc))

(defun _compact (lst / acc) (setq acc '())
  (foreach x lst (if x (setq acc (cons x acc))))
  (reverse acc))

(defun _scan-fonts-in-dwg (dwg deepScan / dbx styles rows missing it sty pfont bfont pmiss bmiss status inlineFonts)
  (setq rows '() missing '())
  (setq dbx (_open-dbx dwg))
  (if (null dbx)
    (setq rows (cons (list dwg "<open-failed>" "" "" "OPEN_ERROR") rows))
    (progn
      (setq styles (_collect_styles_fonts dbx))
      (foreach it styles
        (setq sty (car it)
              pfont (cadr it)
              bfont (caddr it)
              pmiss (and pfont (/= pfont "") (not (_font-file-exists-p pfont)))
              bmiss (and bfont (/= bfont "") (not (_font-file-exists-p bfont)))
              status (if (or pmiss bmiss) "MISSING" "OK"))
        (setq rows (cons (list dwg sty pfont bfont status) rows))
        (if pmiss (setq missing (cons (strcase pfont) missing)))
        (if bmiss (setq missing (cons (strcase bfont) missing))))
      (if deepScan
        (progn
          (setq inlineFonts (_collect_mtext_inline_fonts dbx))
          (foreach f inlineFonts
            (if (and (/= f "") (not (_font-file-exists-p f)))
              (progn
                (setq rows (cons (list dwg "<MTEXT inline>" f "" "MISSING") rows))
                (setq missing (cons (strcase f) missing)))))))
      (vlax-release-object dbx)))
  (setq missing (_uniq-ci (_compact missing)))
  (cons (reverse rows) missing))

(defun _write-csv (csvPath rows missing / fh)
  (setq fh (open csvPath "w"))
  (if (null fh) (progn (princ (strcat "\nERROR: Cannot write CSV: " csvPath)) (exit)))
  (write-line "Drawing,Style,PrimaryFont,BigFont,Status" fh)
  (foreach r rows (write-line (strcat (nth 0 r) "," (nth 1 r) "," (nth 2 r) "," (nth 3 r) "," (nth 4 r)) fh))
  (write-line "" fh)
  (write-line "MissingFontsUnique" fh)
  (foreach m (_uniq-ci missing) (write-line m fh))
  (close fh)
  csvPath)

(defun c:ScanFonts (srcFolder outCsv deepScanFlag / files accRows accMissing R)
  (if (or (null srcFolder) (null outCsv))
    (progn (princ "\nUsage: (c:ScanFonts <sourceFolder> <outCsvPath> [deepScanFlag])") (princ) (exit)))
  (setq files (vl-directory-files srcFolder "*.dwg" 1))
  (if (null files) (progn (princ (strcat "\nNo DWGs in: " srcFolder)) (princ) (exit)))
  (setq files (mapcar '(lambda (f) (strcat srcFolder "\\" f)) files))
  (setq accRows '() accMissing '())
  (foreach dwg files
    (setq R (_scan-fonts-in-dwg dwg (not (null deepScanFlag))))
    (setq accRows (append accRows (car R)))
    (setq accMissing (append accMissing (cdr R))))
  (setq accMissing (_uniq-ci (mapcar 'strcase accMissing)))
  (_write-csv outCsv accRows accMissing)
  (princ (strcat "\nFont report written: " outCsv))
  (princ))

;; ---- Probe media across layouts (for prompt control) ----
(defun c:ProbeMedia (srcFolder outIni / files sheets cnts ansiCnt archCnt medias fh i n)
  ;; Normalize paths to Windows style
  (setq srcFolder (_norm srcFolder))
  (setq outIni    (_norm outIni))

  (if (or (null srcFolder) (null outIni))
    (progn (princ "\nERR: c:ProbeMedia requires srcFolder and outIni.") (princ) (exit)))

  (setq files (vl-directory-files srcFolder "*.dwg" 1))
  (if (null files) (progn (princ (strcat "\nNo DWGs in: " srcFolder)) (princ) (exit)))

  (setq files (mapcar '(lambda (f) (strcat srcFolder "\\" f)) files))
  (setq n (length files))
  (princ (strcat "\nProbe: " (itoa n) " DWG(s)"))

  ;; Build list of (dwg layout media)
  (setq sheets '() i 0)
  (foreach dwg files
    (setq i (+ i 1))
    (princ ".") ;; lightweight progress
    (foreach ent (_paper-layouts-of dwg)
      (setq sheets (cons (list dwg (car ent) (cadr ent)) sheets))
    )
  )
  (setq sheets (reverse sheets))

  ;; Count canonical ANSI/ARCH
  (setq cnts (_count-target-sizes sheets))
  (setq ansiCnt (car cnts) archCnt (cadr cnts))

  ;; Unique media list
  (setq medias '())
  (foreach s sheets
    (if (and (nth 2 s) (not (member (strcase (nth 2 s)) (mapcar 'strcase medias))))
      (setq medias (cons (nth 2 s) medias))
    )
  )
  (setq medias (reverse medias))

  ;; Write INI
  (setq fh (open outIni "w"))
  (if (null fh) (progn (princ (strcat "\nERR: Cannot write probe file: " outIni)) (princ) (exit)))
  (write-line "[COUNTS]" fh)
  (write-line (strcat "ANSI_D=" (itoa ansiCnt)) fh)
  (write-line (strcat "ARCH_D=" (itoa archCnt)) fh)
  (write-line "" fh)
  (write-line "[MEDIAS]" fh)
  (setq i 0)
  (foreach m medias
    (write-line (strcat (itoa i) "=" m) fh)
    (setq i (+ i 1)))
  (close fh)
  (princ (strcat "\nProbe written: " outIni))
  (princ)
)

;; ---- Main publish entry point (8 args; last is optional overrideMedia) ----
(defun c:PublishFolder ( srcFolder outPdf npsName deviceName styleTable forceCtb tiePrefArg overrideMedia
  / files sheets dsdPath nps2234 nps2436 npsForced npsGeneric
    cnts ansiCnt archCnt npsMajor tiePrefEnv tiePref forceIt noCanon )
  (if (or (null srcFolder) (null outPdf) (null npsName) (null deviceName) (null styleTable))
    (progn (princ "\nERROR: c:PublishFolder requires 5 arguments (plus optional 6th/7th/8th).") (princ) (exit)))
  (setq srcFolder (_norm srcFolder))
  (setq outPdf (_norm outPdf))
  (if (= deviceName "") (setq deviceName "MasterPwPdf.pc3"))
  (setq forceIt (_truep forceCtb))

  (setq files (vl-directory-files srcFolder "*.dwg" 1))
  (if (null files) (progn (princ (strcat "\nNo DWGs in: " srcFolder)) (princ) (exit)))
  (setq files (mapcar '(lambda (f) (strcat srcFolder "\\" f)) files))
  (setq sheets '())
  (foreach dwg files (foreach ent (_paper-layouts-of dwg)
    (setq sheets (cons (list dwg (car ent) (cadr ent)) sheets))))
  (setq sheets (reverse sheets))

  (if (null sheets) (progn (princ "\nNo paper-space layouts found in the set.") (princ) (exit)))

  (_log-open outPdf)
  (_log (strcat "Source folder: " srcFolder))
  (_log (strcat "Output PDF : " outPdf))
  (_log (strcat "Device/CTB : " deviceName " / " styleTable))
  (if (> (strlen npsName) 0) (_log (strcat "Forced Nps : " npsName)))

  ;; Normalize tiePref from arg ONLY (PS passes it only on tie)
  (setq tiePref (_pref-normalize tiePrefArg))

  ;; Forced named page setup by exact name (if provided)
  (if (> (strlen npsName) 0) (setq npsForced (_find-ps-by-name-in-folder srcFolder npsName)))
  (if npsForced
    (progn
      (setq dsdPath (strcat (vl-filename-directory outPdf) "\\_publish_set.dsd"))
      (setq nps2234 npsForced nps2436 npsForced npsGeneric npsForced)
      (_log (strcat "Using FORCED page setup: " (cdr npsForced) " (src=" (car npsForced) ")"))
      (_build-dsd dsdPath sheets outPdf nps2234 nps2436 npsGeneric)
      (if *strict_fail*
        (progn (vl-file-delete dsdPath) (_log "STRICT abort: forced PS but some sheets fell back.") (_log-close)
               (princ "\nERROR: STRICT abort (some sheets would use layout).") (princ) (exit))
        (progn (_publish-dsd dsdPath outPdf) (_log "Publish complete (forced PS).") (_log-close)
               (princ (strcat "\nPublished: " outPdf "\nUsing forced page setup " (chr 34) (cdr npsForced) (chr 34) " from: " (car npsForced))) (princ) (exit)))))
  ;; Force CTB override via matching device/CTB
  (if forceIt
    (progn
      (setq nps2234 nil nps2436 nil)
      (setq npsGeneric (_find-ps-by-criteria-in-folder srcFolder deviceName styleTable))
      (if npsGeneric
        (progn
          (setq dsdPath (strcat (vl-filename-directory outPdf) "\\_publish_set.dsd"))
          (setq npsMajor npsGeneric
                nps2234 npsMajor
                nps2436 npsMajor
                npsGeneric npsMajor)
          (_log (strcat "Using CTB override via page setup: " (cdr npsMajor) " (src=" (car npsMajor) ")"))
          (_build-dsd dsdPath sheets outPdf nps2234 nps2436 npsGeneric)
          (if *strict_fail*
            (progn (vl-file-delete dsdPath) (_log "STRICT abort: CTB override but some sheets fell back.") (_log-close)
                   (princ "\nERROR: STRICT abort (some sheets would use layout).") (princ) (exit))
            (progn (_publish-dsd dsdPath outPdf) (_log "Publish complete (CTB override).") (_log-close)
                   (princ (strcat "\nPublished: " outPdf "\n(CTB override = " styleTable ")")) (princ) (exit))))
        (progn
          (_log "CTB override requested; no matching page setup found. Creating temp PS library...")
          (setq npsMajor (_make-ps-lib deviceName styleTable tiePrefArg)) ;; tiePrefArg may be "" here; okay
          (if (null npsMajor)
            (progn
              (_log "ERROR: Could not create temp PS library.")
              (_log-close)
              (princ "\nERROR: Failed to create temporary page setup for CTB override.") (princ) (exit))
            (progn
              (setq dsdPath (strcat (vl-filename-directory outPdf) "\\_publish_set.dsd"))
              (setq nps2234 npsMajor nps2436 npsMajor npsGeneric npsMajor)
              (_log (strcat "Using generated page setup: " (cdr npsMajor) " (src=" (car npsMajor) ")"))
              (_build-dsd dsdPath sheets outPdf nps2234 nps2436 npsGeneric)
              (if *strict_fail*
                (progn (vl-file-delete dsdPath) (_log "STRICT abort: generated PS but some sheets fell back.") (_log-close)
                      (princ "\nERROR: STRICT abort (some sheets would use layout).") (princ) (exit))
                (progn (_publish-dsd dsdPath outPdf) (_log "Publish complete (generated PS).") (_log-close)
                      (princ (strcat "\nPublished: " outPdf "\n(Generated PS with CTB=" styleTable ")")) (princ) (exit))
              )
            )
          )
        )

      ;; Try known names, then majority/tie preference (tiePref only if provided)
      (setq nps2234 (_find-ps-by-name-in-folder srcFolder "CDMS_Pdf_2234_Mono"))
      (setq nps2436 (_find-ps-by-name-in-folder srcFolder "CDMS_Pdf_2436_Mono"))
      (setq npsGeneric (_find-ps-by-criteria-in-folder srcFolder deviceName styleTable))
      (setq cnts (_count-target-sizes sheets))
      (setq ansiCnt (car cnts) archCnt (cadr cnts))
      (_log (strcat "Detected target media counts: ANSI_D=" (itoa ansiCnt) ", ARCH_D=" (itoa archCnt)))

      ;; If no canonical media exist, allow override from caller (ANSI/ARCH or specific media label)
      (setq noCanon (and (= ansiCnt 0) (= archCnt 0)))
      (if (and noCanon overrideMedia (> (strlen overrideMedia) 0))
        (progn
          (cond
            (( _ci= overrideMedia "ANSI")
              (setq npsMajor (_first-some nps2234 npsGeneric nps2436)))
            (( _ci= overrideMedia "ARCH")
              (setq npsMajor (_first-some nps2436 npsGeneric nps2234)))
            (T
              ;; Specific non-canonical media: use generic (device+CTB) if available
              (setq npsMajor (_first-some npsGeneric nps2234 nps2436))))
          (if (npsMajor) 
            (progn
              (setq nps2234 npsMajor nps2436 npsMajor npsGeneric npsMajor)
              (_log (strcat "No canonical media; using override '" overrideMedia
                            "' with page setup " (cdr npsMajor) " (src=" (car npsMajor) ")")))
            (progn
              (_log (strcat "No suitable page setup found for override '" overrideMedia "'. Creating temp PS library..."))
              (setq npsMajor (_make-ps-lib deviceName styleTable overrideMedia))
              (if (null npsMajor)
                (progn (_log "ERROR: Could not create temp PS library.") (_log-close)
                      (princ "\nERROR: Failed to create temporary page setup for override media.") (princ) (exit))
                (progn
                  (setq dsdPath (strcat (vl-filename-directory outPdf) "\\_publish_set.dsd"))
                  (setq nps2234 npsMajor nps2436 npsMajor npsGeneric npsMajor)
                  (_log (strcat "Using generated page setup (override='" overrideMedia "'): "
                                (cdr npsMajor) " (src=" (car npsMajor) ")"))
                  (_build-dsd dsdPath sheets outPdf nps2234 nps2436 npsGeneric)
                  (if *strict_fail*
                    (progn (vl-file-delete dsdPath) (_log "STRICT abort: generated PS but some sheets fell back.") (_log-close)
                          (princ "\nERROR: STRICT abort (some sheets would use layout).") (princ) (exit))
                    (progn (_publish-dsd dsdPath outPdf) (_log "Publish complete (generated PS / override).") (_log-close)
                          (princ (strcat "\nPublished: " outPdf " (OverrideMedia=" overrideMedia ")")) (princ) (exit))
                  )
                )
              )
            )
          )
          )
        ))

      ;; Only perform majority/tie logic if we DO have canonical counts case
      (if (not noCanon)
        (progn
          (setq npsMajor nil)
          (cond
            ((> ansiCnt archCnt) (setq npsMajor (_first-some nps2234 nps2436 npsGeneric)))
            ((> archCnt ansiCnt) (setq npsMajor (_first-some nps2436 nps2234 npsGeneric)))
            (T (setq npsMajor (_select-by-pref tiePref nps2234 nps2436 npsGeneric)))))
      )

      (if npsMajor
        (progn
          (setq dsdPath (strcat (vl-filename-directory outPdf) "\\_publish_set.dsd"))
          (setq nps2234 npsMajor nps2436 npsMajor npsGeneric npsMajor)
          (_log (strcat "Using selected page setup: " (cdr npsMajor) " (src=" (car npsMajor) ")"))
          (_build-dsd dsdPath sheets outPdf nps2234 nps2436 npsGeneric)
          (if *strict_fail*
            (progn (vl-file-delete dsdPath) (_log "STRICT abort: selected PS but some sheets fell back.") (_log-close)
                  (princ "\nERROR: STRICT abort (some sheets would use layout).") (princ) (exit))
            (progn (_publish-dsd dsdPath outPdf) (_log "Publish complete (selected/majority PS).") (_log-close)
                  (princ (strcat "\nPublished: " outPdf
                                  "\nUsing selected page setup " (chr 34) (cdr npsMajor) (chr 34) " from: " (car npsMajor)
                                  " (ANSI=" (itoa ansiCnt) ", ARCH=" (itoa archCnt) ")"))
                  (princ) (exit)))))

      (if (and (null nps2234) (null nps2436) (null npsGeneric))
        (progn
          (_log "ERROR: No suitable page setup found (names or device/ctb).")
          (_log-close)
          (princ (strcat "\nERROR: No suitable page setup found in folder.\n"
                        "Looked for names " (chr 34) "CDMS_Pdf_2234_Mono" (chr 34) " / " (chr 34) "CDMS_Pdf_2436_Mono" (chr 34) ",\n"
                        "then for any page setup using device " (chr 34) deviceName (chr 34) " and style " (chr 34) styleTable (chr 34) "."))
          (princ) (exit)))

      ;; Fallback build (should be rare)
      (setq dsdPath (strcat (vl-filename-directory outPdf) "\\_publish_set.dsd"))
      (_build-dsd dsdPath sheets outPdf nps2234 nps2436 npsGeneric)
      (if *strict_fail*
        (progn (vl-file-delete dsdPath) (_log "STRICT abort: fallback would use layout for some sheets.") (_log-close)
              (princ "\nERROR: STRICT abort (some sheets would use layout).") (princ) (exit))
        (progn (_publish-dsd dsdPath outPdf) (_log "Publish complete.") (_log-close)
              (princ (strcat "\nPublished: " outPdf)) (princ)))
    ))
)
