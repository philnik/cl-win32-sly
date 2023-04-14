# cl-win32-sly
cl-win32-sly

## DO:

### sly/slime repl operation
Add (co-initialize-multithreaded) before create object to work with sly/slime


### Remove bare warning
convert
(cffi:foreign-slot-value variant 'VARIANT 'value)
--->
(cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'value)

stops bare warning, may miss some structure spot



## TODO:

### Export executable
Gives error at dispatch

### connect to opened window
opens up a new application each time like calling DispatchEx
sometimes may need to just use an already opened application

