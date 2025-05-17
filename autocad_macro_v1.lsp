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
(defun GetTab (MyTab)
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
    nil ; Return nil if Excel is not open
  )
)

;--------------
;--- SEARCH ---
;--------------

