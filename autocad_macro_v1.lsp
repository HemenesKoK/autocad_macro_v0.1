; Error Handler
(defun MyErrorHandler (msg)
  (princ (strcat "\nError: " msg))
  (CloseExcel)
  (princ)
)
(setq *error* MyErrorHandler)


; Excel Cell Reader Sub-Routine
; by Leonard Lorden
; July 26, 2017

; Open Excel file
(defun OpenExcel (Exfile)

  ; Check if the file exists
  (setq MyFile (findfile Exfile))

  ; if exists > open / > return nil
  (if (/= MyFile nil)
    ; open process
    (progn
      ; create Excel application object
      (setq MyXL (vlax-get-or-create-object "Excel.Application"))
      ; set properties
      (vla-put-visible MyXL :vlax-false)
      (vlax-put-property MyXL 'DisplayAlerts :vlax-false)
      ; open file
      (setq MyBook (vl-catch-all-apply 'vla-open (list (vlax-get-property MyXL "WorkBooks") MyFile)))
      
      ; file error
      (if (vl-catch-all-error-p MyBook)
        (*error* "\nError opening Excel file.")
      )
    )
    ; file not found
    (*error* "\nExcel file not found.") 
  )
)	

; Routine to CLOSE Excel file & Session
; Assumes previously opened with OpenExcel function
(defun CloseExcel ()

  ; Close the workbook
  (if MyBook
    (vl-catch-all-apply 'vlax-invoke-method (list MyBook "Close")))

  ; Quit Excel
  (if MyXL
    (vl-catch-all-apply 'vlax-invoke-method (list MyXL "Quit")))

  ; Release cell object
  (if MyCell
    (vl-catch-all-apply 'vlax-release-object MyCell))

  ; Release range object             
  (if MyRange
    (vl-catch-all-apply 'vlax-release-object MyRange))

  ; Release sheet object             
  (if MySheet
    (vl-catch-all-apply 'vlax-release-object MySheet))

  ; Release workbook object
  (if MyBook
    (vl-catch-all-apply 'vlax-release-object MyBook))

  ; Release Excel application object          
  (if MyXL
    (vl-catch-all-apply 'vlax-release-object MyXL))       

  ; Clear variables
  (setq MyFile nil MyXL nil MyBook nil MySheet nil MyRange nil
        MyTab nil MyCell nil ExCell nil)

  ; Garbage cleanup
  (gc) 
)

; Set Working Tab
(defun GetTab
    (progn
        ; Get the first sheet
        (setq MySheet (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property MyBook "Sheets") "Item" 1)))
        ; Activate the sheet
        (if (not (vl-catch-all-error-p MySheet)) (vlax-invoke-method MySheet "Activate")
            (*error* "\nError activating sheet.")
        )
    )
  MySheet)

; Get the value of a cell
(defun GetCell (ExCell)
    (progn
        ; Get the cell range
        (setq MyRange (vlax-get-property (vlax-get-property MySheet 'Cells) "Range" ExCell))
         ; Get the cell value
        (setq MyCell (vlax-variant-value (vlax-get-property MyRange 'Value2)))

        ; Convert numeric values to strings without decimals   WHY?
        (if (numberp MyCell)
            (setq MyCell (rtos MyCell 2 0))
        )

        ; debug
        (princ (strcat "\nValue of cell " ExCell ": " (vl-princ-to-string MyCell)))
  )
  MyCell ; Return the cell value
)

;--- Append for PK Ossendorf

;--------------
;--- LAYERS ---
;--------------

;--- layer CURRENT by name
(defun slaycurr (layername)
  (command "_-layer" "S" layername "")
)

;--- layer ON by name
(defun slayon (layername)
  (command "_-layer" "ON" layername "")
)

;--- layer NEW by name
(defun slaynew (layername)
  (command "_-layer" "N" layername "" "C")
)

;--- layer OFF by name
(defun slayoff (layername)
  (if (not (equal (getvar "CLAYER") layername)) ; Check if the layer is not the current layer
    (command "_-layer" "OFF" layername "") ; Turn off the layer
    (progn
      (command "_-layer" "S" "0" "") ; Switch to layer "0"
      (command "_-layer" "OFF" layername "") ; Turn off the layer
    )
  )
)

;-------------
;--- EXCEL ---
;-------------
;--- TODO: Error handling

(defun SetCellValue (cellAddress value)
  (if (/= MyXL nil) ; Check if Excel is open
    (progn
      ;; Convert the value to a string with no decimal places
      (setq formattedValue (if (numberp value) (rtos value 2 0) value)) ; Use rtos for numbers, keep strings as is
      (setq MyRange (vlax-get-property (vlax-get-property MySheet 'Cells) "Range" cellAddress)) ; Get the cell range
      (vlax-put-property MyRange 'Value2 formattedValue) ; Set the cell value
      (princ (strcat "\nSet value of cell " cellAddress " to: " (vl-princ-to-string formattedValue))) ; Print confirmation
      ;; Save the workbook after updating the cell
      (vlax-invoke-method MyBook "Save")
    )
    (*error* "\nExcel is not open.") ; Error message if Excel is not open
  )
)

;--------------
;--- SEARCH ---
;--------------
; TODO: set search filter to Current layer

; Parameters: 
; searchString: The string to search for

; Local Variables:
; ss:         The selection set of entities to search in
; count:      The number of entities in the selection set
; ent:        The current entity being processed
; entData:    The data of the current entity
; layerName:  The name of the layer of the current entity
; layerDef:   The definition of the layer of the current entity
; textValue:  The text value of the current entity

(defun txtSearch (searchString / ss count ent entData layerName layerDef textValue)
  (setq ss (ssget "_X" '((0 . "TEXT,MTEXT")))) ; Select all TEXT and MTEXT
  (if ss
    (progn
      (setq count (sslength ss))
      (while (> count 0)
        (setq ent (ssname ss (setq count (1- count))))
        (setq entData (entget ent)) ; Get entity data once
        (setq layerName (cdr (assoc 8 entData))) ; Get layer name

        ;; Check if the layer is ON
        (setq layerDef (tblsearch "LAYER" layerName))
        (if (and layerDef
                 (or (not (assoc 62 layerDef))         ; No color means it's ON
                     (> (cdr (assoc 62 layerDef)) 0))   ; Color is positive => layer ON
            )
          (progn
            ;; Now get text value
            (setq textValue
              (cond
                ((= (cdr (assoc 0 entData)) "TEXT")
                 (cdr (assoc 1 entData)))
                ((= (cdr (assoc 0 entData)) "MTEXT")
                 (vlax-get-property (vlax-ename->vla-object ent) 'TextString))
              )
            )
            ;; Exact match
            (if (and textValue
                     (equal textValue searchString))
              (progn
                (setq count 0) ; Exit loop
                ent ; Return entity
              )
            )
          )
        )
      )
    )
    nil
  )
)

;-------------
;--- COORD ---
;-------------

(defun GetTextCoordinates (textObj)
  (if textObj
    (progn
      (setq coords (cdr (assoc 10 (entget textObj)))) ; Get the insertion point (coordinates)
      (if coords
        (progn
          ;; Return the coordinates as a list
          (list (car coords) (cadr coords))
        )
        nil ; Handle missing coordinates
      )
    )
    nil ; Return nil if no text object is provided
  )
)

;------------
;--- MISC ---
;------------

(defun CreateCircle (x y radius)
  (princ (strcat "\nCreating circle at: " (rtos x 2 2) ", " (rtos y 2 2))) ; Print the coordinates
  (command "_CIRCLE" (list x y) radius) ; Create a circle at the specified coordinates with the given radius
)

(defun GetFileInput (prompt)
  (setq fileName (getfiled prompt "" "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx" 1))
  (if fileName
    (progn
      (setq MyFile fileName)
      (princ (strcat "\nSelected file: " MyFile))
    )
    (princ "\nNo file selected.")
  )
)

;------------
;--- MAIN ---
;------------

(defun c:MacroKolecka
  (GetFileInput "Select Excel file")
  (OpenExcel MyFile)
  (GetTab)

  

)