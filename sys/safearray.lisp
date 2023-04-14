(in-package :cl-win32ole-sys)

(defun ->safe-array (variant-ptr-list)
  (let* ((element-sizes (compute-elements-list variant-ptr-list))
         (dim (length element-sizes)))
    (cffi:with-foreign-objects ((safe-array-bounds `(:struct ,'SAFEARRAYBOUND) dim)
                                (indices :long dim)
                                (l-bounds :long dim))
      ;; Set SAFEARRAYBOUND
      (dotimes (i dim)
        (setf (cffi:mem-aref indices :long i) 0)
        (let ((p (cffi:mem-aref safe-array-bounds `(:struct ,'SAFEARRAYBOUND) i)))
          (setf (cffi:foreign-slot-value p `(:struct ,'SAFEARRAYBOUND) 'elements)
                (nth i element-sizes)
                (cffi:foreign-slot-value p `(:struct ,'SAFEARRAYBOUND) 'l-bound)
                0)))
      ;; Creat Safe Array and set lisp value in lisp list
      (let ((psa (SafeArrayCreate VT_VARIANT dim safe-array-bounds)))
        (when (cffi-sys:null-pointer-p psa)
          (error "SafeArrayCreate return NULL."))
        (set-lisp-list-to-safe-array variant-ptr-list psa dim element-sizes indices)
        psa
        ))))

(defun compute-elements-list (list)
  (if (null list)
      (list 0)
      (loop for each on list by #'car
	 unless (atom each)
	 collect (length each))))

(defun set-lisp-list-to-safe-array (list psa dim element-sizes indices)
  (labels ((elm-of-list ()
             (%elm-of-list list 0 dim indices))
           (%elm-of-list (list i dim indices)
             (let ((index (cffi:mem-aref indices :long i)))
               (if (= i (1- dim ))
                   (nth index list)
                   (%elm-of-list (nth index list) (1+ i) dim indices))))
           (put-element ()
             (succeeded (SafeArrayPutElement psa indices (elm-of-list))))
           (run (n)
             (if (= n dim)
                 (put-element)
                 (dotimes (i (nth n element-sizes))
                   (setf (cffi:mem-aref indices :long n) i)
                   (run (1+ n))))))
    (run 0)))


(defun safe-array->variant-ptr-list (safe-array)
  (let* ((vt (logior (logand (variant-type* safe-array) (lognot VT_ARRAY))
                     VT_BYREF))
         (psa (variant-array-value safe-array))
         (dim (SafeArrayGetDim psa)))
    (cffi:with-foreign-objects ((variant `(:struct ,'VARIANT))
                                (l-bounds :long dim)
                                (u-bounds :long dim)
                                (indices :long dim))
      (VariantInit variant)
      (setf (variant-type variant) vt)
      (dotimes (i dim)
					;        do (let ((offset (* i (cffi:foreign-type-size :long))))
	do (let ((offset (* i (cffi:foreign-type-size  :long))))
             (succeeded (SafeArrayGetLBound
                         psa (1+ i) (cffi-sys:inc-pointer l-bounds offset)))
             (succeeded (SafeArrayGetUBound
                         psa (1+ i) (cffi-sys:inc-pointer u-bounds offset)))
             (succeeded (SafeArrayGetLBound
                         psa (1+ i) (cffi-sys:inc-pointer indices offset)))))
      (set-safe-array-to-lisp-list psa dim l-bounds u-bounds indices
                                   variant))))

(defun set-safe-array-to-lisp-list (psa dim l-bounds u-bounds indices
                                    variant)
  (labels ((run (n)
             (if (= n dim)
                 (get-val)
                 (loop for i from (l-b n) to (u-b n)
                    collect (progn (setf (cffi:mem-aref indices :long n) i)
                                   (run (1+ n))))))
           (get-val ()
             (succeeded (SafeArrayGetElement psa indices variant))
             (unwind-protect
                  (variant-copy (alloc-variant) variant)
               (succeeded (VariantClear variant))))
           (l-b (n)
             (cffi:mem-aref l-bounds :long n))
           (u-b (n)
             (cffi:mem-aref u-bounds :long n)))
    (run 0)))
