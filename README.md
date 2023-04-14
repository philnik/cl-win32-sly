# cl-win32-sly
cl-win32-sly

## DO:
Add (co-initialize-multithreaded) to work with sly/slime


### Remove warning
convert
(cffi:foreign-slot-value variant 'VARIANT 'value)
--->
(cffi:foreign-slot-value variant `(:struct ,'VARIANT) 'value)

stops bare warning, may miss some structure spot



## TODO:

### Export executable

### opens up a new application each time like calling DispatchEx
sometimes may need to just use an already opened application

