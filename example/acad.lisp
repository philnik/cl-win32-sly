
(in-package :cl-user)


;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))


(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)
(use-package :cffi)
(in-package :cl-win32ole)

(co-initialize-multithreaded)
(defparameter acad (create-object "BricscadApp.AcadApplication"))
(defparameter doc (ole acad :ActiveDocument))
(defparameter model (ole doc :ModelSpace))
(setf (ole acad :Visible) 1)
(setf n (ole acad :ActiveDocument :Name))


(format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))

(invoke (ole doc :Application) :ZoomAll)

(defparameter application  (ole doc :Application))

(defparameter selectionset  (ole doc :SelectionSets))

(ole selectionset :Add "myselectionset")

(ole (ole doc :ModelSpace) :Count)



(setf d2 (make-array 3 :element-type 'double-float :initial-contents #(3.0d0 4.0d0 0.0d0)))
(setf v0 (alloc-variant))
(setf (variant-type v0) (logior VT_ARRAY VT_SAFEARRAY))
;(setf (variant-type v0) 4096)
(setf p1 (cffi:foreign-alloc :double :initial-contents d2))
(setf (variant-value v0) p1)
(to-lisp v0)
(variant-array-to-lisp v0)


(setf v1 (alloc-variant))
(setf (variant-type v1) VT_R8)
(setf (variant-value v1) 10.0d0)
(to-lisp v1)

(setf aaa(make-variant p1 8204))

(cffi:mem-aref p1 :double 0)
(cffi:mem-aref p1 :double 1)
(cffi:mem-aref p1 :double 2)

(cffi:mem-aref (variant-value v0) :double 0)
(cffi:mem-aref (variant-value v0) :double 1)
(cffi:mem-aref (variant-value v0) :double 2)
(cffi:mem-ref (variant-value v0) :double 0)

(invoke (ole doc :ModelSpace) :AddArc aaa 25.0 0.0 1.5)


(defparameter l '(1 2))

(loop for i from 1 to 20
      do (progn 
	   (setf ent1 (ole (ole doc :ModelSpace) :Item 2))
	   (format t "~a~%" (property ent1 :Layer))
	   ))
(elt selectionset 0)

      

(ole application :RunCommand "circle")
(ole application :RunCommand "0,0,0")
(ole application :RunCommand "(setq c (+ 1 2)")

(loop for i from 11 to 21
      do (ole application :RunCommand
	      (format nil "circle~%0,0,0~%~s~%" i)
	      )
      )
