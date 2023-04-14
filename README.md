# cl-win32-sly
cl-win32-sly

DO:
Add (co-initialize-multithreaded) to work with sly/slime


convert
(cffi:foreign-slot-value variant 'VARIANT 'value)
--->
(cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'value)

to stop bare warning, may miss some spot



TODO:
Export executable
opens up a new application each time like calling DispatchEx
sometimes may need to just use an already opened application

