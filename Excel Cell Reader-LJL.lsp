;--- Excel Cell Reader Sub-Routine
;--- by Leonard Lorden
;--- July 26, 2017
;--- Routine to OPEN an existing Excel file (best to confirm filename & path prior to calling)
;--- Call routine with path & filename (Example:  OpenExcel "C:/Test.xlsx")
;--- File will be opened hidden
(defun OpenExcel (Exfile)
  (setq MyFile (findfile Exfile)) ; Double-check file exists at location
  (if (/= MyFile nil) ; nil = file not found
    (progn ; If file found, open it
      (setq MyXL (vlax-get-or-create-object "Excel.Application")) ; Find Excel application
      (vla-put-visible MyXL :vlax-false) ; Hide application from view
      (vlax-put-property MyXL 'DisplayAlerts :vlax-false) ; Hide Excel alerts
      (setq MyBook (vl-catch-all-apply 'vla-open (list (vlax-get-property MyXL "WorkBooks") MyFile)))
      (if (vl-catch-all-error-p MyBook)
        (progn
          (princ "\nError opening Excel file.")
          (setq MyBook nil)
        )
      )
    )
    (progn
      (princ "\nExcel file not found.")
      (setq MyBook nil)
    )
  )
)																		;return	- MyFile = nil if file not found	
		
;--- Routine to CLOSE Excel file & Session
;--- Assumes previously opened with OpenExcel function
(defun CloseExcel ()
  (if MyBook
    (vl-catch-all-apply 'vlax-invoke-method (list MyBook "Close"))) ; Close the workbook
  (if MyXL
    (vl-catch-all-apply 'vlax-invoke-method (list MyXL "Quit")))    ; Quit Excel
  (if MyCell
    (vl-catch-all-apply 'vlax-release-object MyCell))               ; Release cell object
  (if MyRange
    (vl-catch-all-apply 'vlax-release-object MyRange))              ; Release range object
  (if MySheet
    (vl-catch-all-apply 'vlax-release-object MySheet))              ; Release sheet object
  (if MyBook
    (vl-catch-all-apply 'vlax-release-object MyBook))               ; Release workbook object
  (if MyXL
    (vl-catch-all-apply 'vlax-release-object MyXL))                 ; Release Excel application object
  ;; Clear variables
  (setq MyFile nil MyXL nil MyBook nil MySheet nil MyRange nil
        MyTab nil MyCell nil ExCell nil)
  (gc) ; Garbage cleanup
)

;--- Routine to set Worksheet Tab
;--- Call using GetTab "Tabname"  (Example:  GetTab Sheet1)
;--- If MySheet = nil on return then requested TAB not found in Excel file or Excel file was not open
(defun GetTab (MyTab)
  (if (/= MyXL nil) ; Ensure Excel is open
    (progn
      (setq MySheet
            (vl-catch-all-apply
              'vlax-get-property
              (list (vlax-get-property MyBook "Sheets") "Item" MyTab))) ; Get the sheet by name
      (if (not (vl-catch-all-error-p MySheet)) ; Check if the sheet was found
        (vlax-invoke-method MySheet "Activate") ; Activate the sheet
        (progn
          (princ (strcat "\nSheet not found: " MyTab)) ; Debugging output
          (setq MySheet nil))) ; Set MySheet to nil if not found
    )
    (setq MySheet nil)) ; Set MySheet to nil if Excel is not open
  MySheet) ; Return the sheet object or nil

;--- Routine to READ an Excel Cell on the current active tab
;--- Call using GetCell "Cell Name" (Example:  GetCell A1)
;--- MyCell returns cell value (nil = empty)
(defun GetCell (ExCell)
  (if (/= MyXL nil) ; Ensure file is open
    (progn
      (setq MyRange (vlax-get-property (vlax-get-property MySheet 'Cells) "Range" ExCell)) ; Get the cell range
      (setq MyCell (vlax-variant-value (vlax-get-property MyRange 'Value2))) ; Get the cell value
      ;; Convert numeric values to strings without decimals
      (if (numberp MyCell)
        (setq MyCell (rtos MyCell 2 0)) ; Convert number to string with no decimals
      )
      (princ (strcat "\nValue of cell " ExCell ": " (vl-princ-to-string MyCell))) ; Print cell value
    )
    (setq MyCell nil) ; Set MyCell to nil if file is not open
  )
  MyCell ; Return the cell value
)

;--- Append for PK Ossendorf


;--- Routine to set current layer by param
(defun slaycurr (layername)
  (command "_-layer" "S" layername "")
)

;--- Routine to set layer on by param
(defun slayon (layername)
  (command "_-layer" "ON" layername "")
)

;--- Routine to create a new layer by param
(defun slaynew (layername)
  (command "_-layer" "N" layername "" "C")
)

;--- Routine to set layer off by param
(defun slayoff (layername)
  (if (not (equal (getvar "CLAYER") layername)) ; Check if the layer is not the current layer
    (command "_-layer" "OFF" layername "") ; Turn off the layer
    (progn
      (command "_-layer" "S" "0" "") ; Switch to layer "0" (or any other default layer)
      (command "_-layer" "OFF" layername "") ; Turn off the layer
    )
  )
)

;--- Routine to write a value to a cell in Excel
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


;--- Routine to Get the coordinates of a text object
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

;--- Routine to create a circle at the specified coordinates
(defun CreateCircle (x y radius)
  (princ (strcat "\nCreating circle at: " (rtos x 2 2) ", " (rtos y 2 2))) ; Print the coordinates
  (command "_CIRCLE" (list x y) radius) ; Create a circle at the specified coordinates with the given radius
)

;--- Run sequence
(defun c:RunExcel ()
  ;(CloseExcel) ; Close any previously opened Excel file
  
  (princ "\nEnter the full path to the Excel file (e.g., D:\\Macro1.xlsx): ")
  (setq excelFile (getstring)) ; Prompt user to input the file path
  (princ "\nExcel file entered: ")
  (princ excelFile) ; Print the entered file path
  (if (not (findfile excelFile))
    (progn
      (princ "\nExcel file not found.")
      (exit)
    )
  )
  (setq sheetName "List1") ; Replace with your sheet name
  
  (OpenExcel excelFile) ; Open the Excel file
  (GetTab sheetName) ; Get the specified sheet
  
  (setq i 1) ; Initialize row counter
  (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
  
  (while cellValueLayer
    (progn
      (princ (strcat "\nProcessing Layer: " cellValueLayer)) ; Print the layer name


      (setq cellValueNumber (GetCell (strcat "C" (itoa i)))) ; Get the number from column B
      (setq textObj (txtSearch cellValueNumber)) ; Get the text object
      (setq coords (GetTextCoordinates textObj)) ; Get the coordinates of the text object
      
      (if (not (and 
            textObj
            (tblsearch "LAYER" cellValueLayer)))
        (progn
          (princ (strcat "\nLayer does not exist. Creating layer: " cellValueLayer))
          (slaynew cellValueLayer) ; Create the layer with color 7 (default white)
        )
      )
      
    
      (slayon cellValueLayer) ; Turn on the layer
      (slaycurr cellValueLayer) ; Set the layer as current

      (SetCellValue (strcat "D" (itoa i)) (car coords)) ; Write X-coordinate to column D
      (SetCellValue (strcat "E" (itoa i)) (cadr coords)) ; Write Y-coordinate to column E

      (CreateCircle (car coords) (cadr coords) 7) ; Create a circle at the coordinates with radius 5
      (slayoff cellValueLayer) ; Turn off the layer

      (setq i (1+ i)) ; Increment the row counter
      (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
    )
  )
  (CloseExcel) ; Close the Excel file
)

;---(OpenExcel "D:\\Macro1.xlsx")							
;--- (GetTab "List1")

;--- LOOP until eof (var i)

;---   (GetCell "A" + i)												
;---   (slayon MyCell)
;---   (slaycurr MyCell)

;---   (GetCell "B" + i)
;---   TODO: find MyCell (duplicate handling)
;---   TODO: Create circle on position (MyCell) of radius 5

;---   (GetCell "A" + i)
;---   (slayoff MyCell)

;--- END LOOP

;--- (CloseExcel)