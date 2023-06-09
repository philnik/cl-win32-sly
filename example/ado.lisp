

;;; miss link here

(ql:quickload :cl-win32ole)
(use-package :cl-win32ole)

;(defparameter *load-truename* "c:/")


(push "c:/users/me/Application Data/lisp/cl-win32ole/" asdf:*central-registry*)
(eval-when (:compile-toplevel :load-toplevel :execute)
  (asdf:oos 'asdf:load-op :cl-win32ole)
  (use-package :cl-win32ole))

(defvar *src-dir* (format nil "~a:~a"
                        (pathname-device *load-truename*)
                        (directory-namestring *load-truename*)))
(defun csv-example ()
  (co-initialize-multithreaded)
  (let ((cn (create-object "ADODB.Connection"))
        (rs (create-object "ADODB.Recordset"))
        (cs (format nil "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=~a;Extended Properties=\"text;HDR=Yes;FMT=Delimited;\";" *src-dir*)))
    (property cn :ConnectionString cs)
    (invoke cn :open)
    (invoke rs :open "select * from aaa.csv" cn)
    (unwind-protect
         (loop until (property rs :eof)
            do (progn
                 (format t "key: ~a, value: ~a, etc: ~a~%"
                         (property (invoke (property rs :fields)
                                           :item "key") :value)
                         (property (invoke (property rs :fields)
                                           :item "value") :value)
                         (property (invoke (property rs :fields)
                                           :item "etc") :value))
                 (invoke rs :movenext))))
      (invoke rs :close)
      (invoke cn :close)))

(defvar *mdb-path* (merge-pathnames "test.mdb" *src-dir*))

(defvar *conn-str* (format nil "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=~a;Jet OLEDB:Engine Type=5;" *mdb-path*))

(defun create-test-mdb ()
  (let ((adox (create-object "ADOX.Catalog")))
    (invoke adox :create *conn-str*)))

(defun mdb-example ()
  (unless (probe-file *mdb-path*)
    (create-test-mdb))
  (let ((cn (create-object "ADODB.Connection")))
    (invoke cn :open *conn-str*)
    (unwind-protect
         (progn
           (ole cn :execute "drop table table1")
           (invoke cn :execute
                   "create table table1 (
                      str1 integer primary key,
                      str varchar(10))")
           )
      (invoke cn :close))))

;;(csv-example)
;;(mdb-example)
