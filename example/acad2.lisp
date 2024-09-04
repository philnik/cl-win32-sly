
(in-package :cl-user)


;; (eval-when (:compile-toplevel :load-toplevel :execute)
;;   (asdf:oos 'asdf:load-op :cl-win32ole)
;;   (use-package :cl-win32ole))

(defparameter cl_directory "c:/Users/filip/AppData/Roaming/lisp/cl-win32-sly")
(push cl_directory ql:*local-project-directories*)

(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)
(use-package :cffi)
(in-package :cl-win32ole)

(cffi:define-foreign-library winapi
  (:windows (:or "ole32.dll")))

(cffi:use-foreign-library winapi)


(defun ex ()
  (progn
    (cl-win32ole-sys::coinitialize)
    (defparameter acad (create-object "BricscadApp.AcadApplication"))
    (defparameter doc (ole acad :ActiveDocument))
    (defparameter model (ole doc :ModelSpace))
    (setf (ole acad :Visible) 1)
    (format t "acad.ActiveDocument.Name:~a~%" (ole acad :ActiveDocument :Name))

;    (invoke model :AddLine "0,0" "100,100")
    
    (loop for i from 11 to 21
	  do (ole doc :SendCommand
		  (format nil "circle~%0,0,0~%~s~%" i)
		  )
	  )

    (invoke acad :ZoomAll)
    (cl-win32ole-sys::co-uninitialize)
    (list acad model doc)))

(setq mydata (ex))


