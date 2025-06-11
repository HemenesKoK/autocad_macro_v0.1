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
(defun GetTab ()
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
  (command "_-layer" "N" layername "")
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

(defun GetLayerOfObject (obj)
  (if (and obj (tblsearch "LAYER" (cdr (assoc 8 (entget obj)))))
    (cdr (assoc 8 (entget obj))) ; Return the layer name
    (*error* "\nObject not found or does not have a layer.")
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

(defun txtSearch (searchString inputLayerName / ss count ent entData textValue foundEnt)
  (setq filterList
    (list
      '(-4 . "<OR")
        '(0 . "TEXT")
        '(0 . "MTEXT")
      '(-4 . "OR>")
      (cons 8 inputLayerName) ; (8 . "LayerName")
    )
  )
  (setq ss (ssget "X" filterList))
  (if ss
    (progn
      (setq count (sslength ss))
      (while (> count 0)
        (setq ent (ssname ss (setq count (1- count))))
        (setq entData (entget ent)) ; Get entity data once

        ; get text value of ent
        (setq textValue
          (cond
            ((= (cdr (assoc 0 entData)) "TEXT")
             (cdr (assoc 1 entData)))
            ((= (cdr (assoc 0 entData)) "MTEXT")
             (vlax-get-property (vlax-ename->vla-object ent) 'TextString))
          )
        )
        ;; Exact match
        (if (and textValue (equal textValue searchString))
          (progn
            (setq count 0) ; Exit loop
            (setq foundEnt ent) ; Store the found entity
          )
        )
      )
      foundEnt
    )
    (*error* "\nNo text objects found in the drawing.")
  )
)

;(defun FindBlocks ( BlockNameReq LayerNameReq / ss count ent entData layerName layerDef textValue)
;  (setq ss (ssget "_X" '((0 . "INSERT"))))
;  (if ss
;    (progn
;      (setq count (sslength ss))
;      (while (> count 0)
;        (setq ent (ssname ss (setq count (1- count))))
;        (setq entData (entget ent))
;        (setq layerName (cdr (assoc 8 entData)))
;        (if (and layerName (tblsearch "LAYER" layerName) (not (equal layerName LayerNameReq)))
;          (setq blockName (cdr (assoc 2 entData)))
;          (if (and blockName (equal blockName BlockNameReq))
;            (progn
;              (setq entList (cons ent entList))
;            )
;          )
;        )
;                 
;      )
;      entList
;    )
;    (*error* "\nNo blocks found in the drawing.")
;  )
;)

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
    (*error* "\nText object not found.")
  )
)

;------------
;--- MISC ---
;------------

(defun CreateCircle (x y radius)
  (entmake
    (list
      (cons 0 "CIRCLE")                     ; Entity type
      (cons 10 (list x y 0.0))              ; Center point (must be 3D)
      (cons 40 radius)                      ; Radius
      (cons 62 256)                         ; Color (256 = BYLAYER, optional)
    )
  )
)

(defun InsertBlock (name x y)
  (entmake
    (list
      (cons 0 "INSERT")                     ; Entity type
      (cons 2 name)                         ; Block name
      (cons 10 (list x y 0.0))              ; Insertion point (3D)
      (cons 41 1)                           ; X scale
      (cons 42 1)                           ; Y scale
      (cons 43 1)                           ; Z scale
      (cons 50 0)                           ; Rotation (in radians)
      (cons 66 0)                           ; Has attributes? 0 = no
    )
  )
)

(defun GetFileInput (prompt)
  (setq fileName (getfiled prompt "" "xlsx" 0))
  (if fileName
    (progn
      (setq MyFile fileName)
      (princ (strcat "\nSelected file: " MyFile))
    )
    (*error* "\nNo file selected.")
  )
)

(defun GetUserInput (prompt)
  (setq userInput (getstring (strcat "\n" prompt ": ")))
  (if (not (equal userInput ""))
    userInput
    (*error* "\nNo input provided.")
  )
)

(defun FindBlockBBox (block / bbox)
  (setq entData (entget block))
  (setq entBBox (vlax-get-property (vlax-ename->vla-object block) 'BoundingBox))
  (if entBBox
    (progn
      (setq minPt (vlax-safearray->list (vlax-variant-value (car entBBox))))
      (setq maxPt (vlax-safearray->list (vlax-variant-value (cadr entBBox))))
      ; Create a bounding box list: (minX minY minZ maxX maxY maxZ)
      (setq bbox (list (car minPt) (cadr minPt) 0.0 ; Min point
                       (car maxPt) (cadr maxPt) 0.0)) ; Max point
    )
    (*error* "\nBounding box not found for the block.")
  )
  bbox
)

;------------
;--- PLOT ---
;------------

(defun PlotBlock (blockName layerName / blockObj bbox)
  (setq blockObj (tblsearch "BLOCK" blockName))
  (foreach block blockObj
    (if block
      (progn
        (setq bbox (FindBlockBBox block))
        (if bbox
          (progn
            ; Plot the block using the bounding box coordinates
            (command "_-plot" "Y" "N" "N" "N" "N" "N" "N" "N" "N" "N"
                     (car bbox) (cadr bbox) (caddr bbox) (cadddr bbox)
                     layerName)
            (princ (strcat "\nBlock plotted on layer: " layerName))
          )
          (*error* "\nBounding box not found for the block.")
        )
      )
      (*error* "\nBlock not found.")
    )
  )
)



;------------
;--- MAIN ---
;------------

; TODO:
; - Clear coords in excel file before running
; - Verify search function on non visible layers (in autocad)

(defun c:MacroKolecka ()
  (GetFileInput "Select Excel file")
  (OpenExcel MyFile)
  (GetTab)
  (GetUserInput "Enter layer name containing propery numbers")

  ;loop
  (setq i 1)
  (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
  (while cellValueLayer
    (progn
      (setq cellValueNumber (GetCell (strcat "C" (itoa i))))
      (setq textObj (txtSearch cellValueNumber userInput))
      (setq coords (GetTextCoordinates textObj))
      
      (if (not (and 
            textObj
            (tblsearch "LAYER" cellValueLayer)))
        (progn
          (princ (strcat "\nLayer does not exist. Creating layer: " cellValueLayer))
          (slaynew cellValueLayer)
        )
      )
      (slayon cellValueLayer)
      (slaycurr cellValueLayer)
      
      (SetCellValue (strcat "D" (itoa i)) (car coords))
      (SetCellValue (strcat "E" (itoa i)) (cadr coords))
      
      (CreateCircle (car coords) (cadr coords) 7)
      (princ (strcat "\nCircle created at: " (rtos (car coords) 2 2) ", " (rtos (cadr coords) 2 2)))
      
      (slayoff cellValueLayer)
      (setq i (1+ i))
      (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
    )
  )
  (CloseExcel)
)

(defun c:MacroBloky ()
  (GetFileInput "Select Excel file")
  (OpenExcel MyFile)
  (GetTab)
  (GetUserInput "Enter block name")
  
  ;loop
  (setq i 1)
  (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
  (while cellValueLayer
    (progn
      (setq cellValueCoorX (GetCell (strcat "D" (itoa i))))
      (setq cellValueCoorY (GetCell (strcat "E" (itoa i))))
      
      (if (not (tblsearch "LAYER" cellValueLayer))
        (progn
          (princ (strcat "\nLayer does not exist. Creating layer: " cellValueLayer))
          (slaynew cellValueLayer) ;
        )
      )
      (slayon cellValueLayer)
      (slaycurr cellValueLayer)
      
      (InsertBlock userInput (atof cellValueCoorX) (atof cellValueCoorY))
      
      (slayoff cellValueLayer)
      (setq i (1+ i))
      (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
      (setq cellValueCoorX nil)
      (setq cellValueCoorY nil)
    )
  )
  (CloseExcel)
  )

(defun c:MacroPlot (/ blockList bboxesList)
  (GetFileInput "Select Excel file")
  (OpenExcel MyFile)
  (GetTab)
  (GetUserInput "Enter block name")
  
  ;loop
  (setq i 1)
  (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
  (while cellValueLayer
    (progn
      (slayon cellValueLayer)
      (slaycurr cellValueLayer)
      (PlotBlock userInput cellValueLayer)
      (slayoff cellValueLayer)
      (setq i (1+ i))
      (setq cellValueLayer (GetCell (strcat "A" (itoa i))))
    )
  )
  
)
; loop through the layers
; find blocks
; plot blocks
