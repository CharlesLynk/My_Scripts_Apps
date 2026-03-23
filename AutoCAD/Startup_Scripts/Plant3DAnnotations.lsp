
	(defun c:InstrumentDivision (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments with their Division
		(if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
			(repeat (sslength ssgr3)
				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
				(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
				(setq x (car pt))
				(setq y (cadr pt))
				(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
				(command-s "._PIDANNOTATE" ss3 "Instrument Division" crds)
			)
		)
		(princ)
	)
   
	(defun c:HandSwitchTopRight (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments Top Right
		(if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
			(repeat (sslength ssgr3)
				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
				;(setq x (car pt))
				;(setq y (cadr pt))
				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
				(command-s "._PIDANNOTATE" ss3 "Device Function (Bubble Top Right)")
			)
		)
		(princ)
	)
   
	(defun c:HandSwitchTopCenter (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Discrete Instruments Top Center
		(if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
			(repeat (sslength ssgr3)
				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
				;(setq x (car pt))
				;(setq y (cadr pt))
				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
				(command-s "._PIDANNOTATE" ss3 "Device Function (Bubble Top Center)")
			)
		)
		(princ)
	)
   
	(defun c:ValveTag (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Valves
		(if (and (setq ssgr3 (ssget))
				(setq i -1)
			)
			(repeat (sslength ssgr3)
				(setq ss3 (ssname ssgr3 (setq i (1+ i))))
				;(setq pt (cdr (assoc 10 (entget ss3)))) ; get the coordinates of the first selected object
				;(setq x (car pt))
				;(setq y (cadr pt))
				;(setq crds (strcat (rtos x 2 2) "," (rtos (+ y 0.25) 2 2)))
				(command-s "._PIDANNOTATE" ss3 "CDM Valve Tag")
			)
		)
		(princ)
	)
   
	(defun c:ReducerAnnotate (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Reducers
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
   
	(defun c:VerticalIOAnno (/ pt x y crds ss3 i ssgr3) ; Annotates all selected Vertical IO
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
   
	(defun c:CDMPipelineTag (/ pt x y crds ss3 ss4 rotn) ; Prints CDM Pipeline tag with all info. Offset automatically dependent on orientation.
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
   
	(defun c:CDMPipelineSSL (/ pt x y crds ss3 ss4 rotn) ; Prints CDM Pipeline tag with Size, Service, Line Number info. Offset automatically dependent on orientation.
		(setq ss3 (ssget))
		(setq pt (getpoint "\n Select Base Point")) ; get the coordinates of the first selected object
		(setq x (car pt))
		(setq y (cadr pt))
		(setq crds (strcat (rtos x 2 2) "," (rtos y 2 2)))
		(command-s "._PIDANNOTATE" ss3 "CDM Tag [Size-Service-LineNumber]" crds)
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
   
	(defun c:CDMPipelineSS (/ pt x y crds ss3 ss4 rotn) ; Prints CDM Pipeline tag with Size and Service info. Offset automatically dependent on orientation.
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
   
	(defun c:QueryAssocVals (/ entl ent ct code item curval) ; Gets names of properties
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
	
	(defun c:dumpobj (/ obj ssgr3 ss3 i temppos) ; Prints properties of object
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
		 (load "Plant3DAnnotations.lsp")
		 (princ)                             ; Exit quietly.
		
	)
	
	
	
	