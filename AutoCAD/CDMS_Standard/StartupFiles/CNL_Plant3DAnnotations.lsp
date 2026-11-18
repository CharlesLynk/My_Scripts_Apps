
;	(defun c:CNL_InstrumentDivision (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments with their Division
;		(if (and (setq ssgr3 (ssget))
;				(setq i -1)
;			)
;			(repeat (sslength ssgr3)
;				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
;				(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
;				(setq x (car pt))
;				(setq y (cadr pt))
;				(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
;				(command-s "._PIDANNOTATE" ss3 "Instrument Division" crds)
;			)
;		)
;		(princ)
;	)
   
;	(defun c:CNL_InstrumentSetpoint (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments with their Setpoints
;		(if (and (setq ssgr3 (ssget))
;				(setq i -1)
;			)
;			(repeat (sslength ssgr3)
;				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
;				(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
;				(setq x (car pt))
;				(setq y (cadr pt))
;				(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
;				(command-s "._PIDANNOTATE" ss3 "Setpoint" crds)
;			)
;		)
;		(princ)
;	)

;	(defun c:CNL_HandSwitchTopRight (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments Top Right
;		(if (and (setq ssgr3 (ssget))
;				(setq i -1)
;			)
;			(repeat (sslength ssgr3)
;				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
;				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
;				;(setq x (car pt))
;				;(setq y (cadr pt))
;				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
;				(command-s "._PIDANNOTATE" ss3 "Device Function (Bubble Top Right)")
;			)
;		)
;		(princ)
;	)
   
;	(defun c:CNL_HandSwitchTopCenter (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments Top Center
;		(if (and (setq ssgr3 (ssget))
;				(setq i -1)
;			)
;			(repeat (sslength ssgr3)
;				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
;				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
;				;(setq x (car pt))
;				;(setq y (cadr pt))
;				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
;				(command-s "._PIDANNOTATE" ss3 "Device Function (Bubble Top Center)")
;			)
;		)
;		(princ)
;	)
   
;	(defun c:CNL_ValveTag (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Valves
;		(if (and (setq ssgr3 (ssget))
;				(setq i -1)
;			)
;			(repeat (sslength ssgr3)
;				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
;				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
;				;(setq x (car pt))
;				;(setq y (cadr pt))
;				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
;				(command-s "._PIDANNOTATE" ss3 "CDM Valve Tag")
;			)
;		)
;		(princ)
;	)
   
	(defun c:CNL_ReducerAnnotate (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Reducers
		(if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
			(repeat (sslength ssgr3)
				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
				;(setq x (car pt))
				;(setq y (cadr pt))
				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
				(command-s "._PIDANNOTATE" ss3 "Adjacent Pipe Reducer Style")
			)
		)
		(princ)
	)
   
	(defun c:CNL_VerticalIOAnno (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Vertical IO
	(if (and (setq ssgr3 (ssget))
			(setq i -1)
		)
		(repeat (sslength ssgr3)
			(setq ss3 (ssname ssgr3 (setq i (1+ i))))
			;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
			;(setq x (car pt))
			;(setq y (cadr pt))
			;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
			(command-s "._PIDANNOTATE" ss3 "CDM IO Description Vertical")
		)
	)
	(princ)
   )
   
	(defun c:CNL_CDMPipelineTag (/ pt x y crds ss3 ss4 rotn) ; Prints CDM Pipeline tag with all info. Offset automatically dependent on orientation.
		(setq ss3 (ssget))
		(setq pt (getpoint "\n Select Base Point")) ; get the coordinates of the first selected object
		(setq x (car pt))
		(setq y (cadr pt))
		(setq crds (strcat (rtos x 2 2) "," (rtos y 2 2)))
		(command-s "._PIDANNOTATE" ss3 "CDM Tag [Size-Service-Spec-LineNumber]" crds)
		;(command-s "._SELECT" "LAST")
		(setq ss4 (ssname (ssget "L") 0))
		(setq rotn (cdr (assoc 50  (entget ss4)))) ; get the rotation of the created annotation
		; Move the object based on the rotation angle
		(cond
		  ((and (<= (- rotn 0.0174533) 0.0174533) (>= (- rotn 0.0174533) -0.0174533)) (command "._MOVE" ss4 "" "D" "0.0,0.12"))
		  ((and (<= (- rotn 1.5708 0.0174533) 0.0174533) (>= (- rotn 1.5708) -0.0174533)) (command "._MOVE" ss4 "" "D" "-0.12,0.0"))
		  ((and (<= (- rotn 4.71239 0.0174533) 0.0174533) (>= (- rotn 4.71239) -0.0174533)) (command "._MOVE" ss4 "" "D" "0.12,0.0"))
		)
		(princ)
	)
   
	(defun c:CNL_CDMPipelineSSL (/ pt x y crds ss3 ss4 rotn) ; Prints CDM Pipeline tag with Size, Service, Line Number info. Offset automatically dependent on orientation.
		(setq ss3 (ssget))
		(setq pt (getpoint "\n Select Base Point")) ; get the coordinates of the first selected object
		(setq x (car pt))
		(setq y (cadr pt))
		(setq crds (strcat (rtos x 2 2) "," (rtos y 2 2)))
		(command-s "._PIDANNOTATE" ss3 "CDM Tag [Size-Service-Line Number]" crds)
		;(command-s "._SELECT" "LAST")
		(setq ss4 (ssname (ssget "L") 0))
		(setq rotn (cdr (assoc 50  (entget ss4)))) ; get the rotation of the created annotation
		; Move the object based on the rotation angle
		(cond
		  ((and (<= (- rotn 0.0174533) 0.0174533) (>= (- rotn 0.0174533) -0.0174533)) (command "._MOVE" ss4 "" "D" "0.0,0.12"))
		  ((and (<= (- rotn 1.5708 0.0174533) 0.0174533) (>= (- rotn 1.5708) -0.0174533)) (command "._MOVE" ss4 "" "D" "-0.12,0.0"))
		  ((and (<= (- rotn 4.71239 0.0174533) 0.0174533) (>= (- rotn 4.71239) -0.0174533)) (command "._MOVE" ss4 "" "D" "0.12,0.0"))
		)
		(princ)
	)
   
	(defun c:CNL_CDMPipelineSS (/ pt x y crds ss3 ss4 rotn) ; Prints CDM Pipeline tag with Size and Service info. Offset automatically dependent on orientation.
		(setq ss3 (ssget))
		(setq pt (getpoint "\n Select Base Point")) ; get the coordinates of the first selected object
		(setq x (car pt))
		(setq y (cadr pt))
		(setq crds (strcat (rtos x 2 2) "," (rtos y 2 2)))
		(command-s "._PIDANNOTATE" ss3 "CDM Tag [Size-Service]" crds)
		;(command-s "._SELECT" "LAST")
		(setq ss4 (ssname (ssget "L") 0))
		(setq rotn (cdr (assoc 50  (entget ss4)))) ; get the rotation of the created annotation
		; Move the object based on the rotation angle
		(cond
		  ((and (<= (- rotn 0.0174533) 0.0174533) (>= (- rotn 0.0174533) -0.0174533)) (command "._MOVE" ss4 "" "D" "0.0,0.12"))
		  ((and (<= (- rotn 1.5708 0.0174533) 0.0174533) (>= (- rotn 1.5708) -0.0174533)) (command "._MOVE" ss4 "" "D" "-0.12,0.0"))
		  ((and (<= (- rotn 4.71239 0.0174533) 0.0174533) (>= (- rotn 4.71239) -0.0174533)) (command "._MOVE" ss4 "" "D" "0.12,0.0"))
		)
		(princ)
	)
   
	(defun c:CNL_QueryAssocVals (/ entl ent ct code item curval) ; Gets names of properties
		  (setq ent (ssname (ssget) 0))      ; Select entity.
		  (setq entl (entget ent))           ; Set entl to association list of last entity.
		  (setq ct 0)                        ; Set ct (a counter) to 0.
		  (textpage)                         ; Switch to the text screen.
		  (princ "\nentget of last entity:")
		  (repeat (length entl)              ; Repeat for number of members in list:
			(print (nth ct entl))            ; Output a newline, then each list member.
			(setq ct (1+ ct))                ; Increments the counter by one.
		  )
		 (princ)                             ; Exit quietly.
		
	)
	
	(defun c:CNL_zDumpObj (/ obj ssgr3 ss3 i temppos) ; Prints properties of object
	    (if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
		(repeat (sslength ssgr3)
			(setq ss3 (ssname ssgr3 (setq i (1+ i))))
			(setq obj (vlax-ename->vla-object (ssname ssgr3 i)))
			(vlax-dump-object obj)
			;(command-s "._PIDTAG" ss3 "")
		)
		)
	)
	
	(defun c:asdf (/ entl ent ct code item curval) ; Reloads this file in AutoCAD to avoid manually using appload
		 (load "CNL_Plant3DAnnotations.lsp")
		 (princ)                             ; Exit quietly.
		
	)
	
	
	
	(defun CNL_DeleteValveTag (pt / half ss i ent data ip del) ; Erases INSERT annotations whose base point lies in a square centered on a valve base point
		(setq half 0.125) ; half of the 0.25 x 0.25 inch square
		(setq ss (ssget "_C"
					(list (- (car pt) half) (- (cadr pt) half))
					(list (+ (car pt) half) (+ (cadr pt) half))
				)
		)
		(setq del (ssadd)
			  i -1
		)
		(if ss
			(repeat (sslength ss)
				(setq ent  (ssname ss (setq i (1+ i)))
					  data (entget ent)
				)
				(if (and (= (cdr (assoc 0 data)) "INSERT") ; INSERT excludes the valve itself, which is an ACPPDYNAMICASSET
						 (setq ip (cdr (assoc 10 data)))
						 (<= (abs (- (car ip) (car pt))) half)
						 (<= (abs (- (cadr ip) (cadr pt))) half)
					)
					(ssadd ent del)
				)
			)
		)
		(if (> (sslength del) 0)
			(command-s "._ERASE" del "") ; ERASE rather than entdel so Plant 3D removes the annotation cleanly
		)
		(princ)
	)

	(defun CNL_COH_PlaceValveAnno (ent pt style / crds) ; Places a PIDANNOTATE annotation of the given style at a valve base point
		(setq crds (strcat (rtos (car pt) 2 2) "," (rtos (cadr pt) 2 2)))
		(command-s "._PIDANNOTATE" ent style crds)
		(princ)
	)

	(defun c:CNL_COH_ValveTag (/ pt ss3 ssgr3 i) ; Replaces existing CDM Valve Tags on all selected Valves with CDM Valve Tag Only and CDM Valve Size Only
		(if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
			(repeat (sslength ssgr3)
				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
				(setq pt (cdr (assoc 10 (entget ss3)))) ; get the base coordinates of the valve
				(CNL_DeleteValveTag pt)
				(CNL_COH_PlaceValveAnno ss3 pt "CDM Valve Tag Only")
				(CNL_COH_PlaceValveAnno ss3 pt "CDM Valve Size Only")
			)
		)
		(princ)
	)
