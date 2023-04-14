;;;; -*- Mode: LISP; Syntax: COMMON-LISP; Coding: shift_jis; -*-
(in-package :cl-win32ole-sys)

(defconstant FORMAT_MESSAGE_ALLOCATE_BUFFER #x00000100)
(defconstant FORMAT_MESSAGE_IGNORE_INSERTS  #x00000200)
(defconstant FORMAT_MESSAGE_FROM_STRING     #x00000400)
(defconstant FORMAT_MESSAGE_FROM_HMODULE    #x00000800)
(defconstant FORMAT_MESSAGE_FROM_SYSTEM     #x00001000)
(defconstant FORMAT_MESSAGE_ARGUMENT_ARRAY  #x00002000)
(defconstant FORMAT_MESSAGE_MAX_WIDTH_MASK  #x000000FF)


(defconstant COINIT_APARTMENTTHREADED #x2)
(defconstant COINIT_MULTITHREADED #x0)
(defconstant COINIT_DISABLE_OLE1DDE #x4)
(defconstant COINIT_SPEED_OVER_MEMORY  #x8)


(defconstant CLSCTX_INPROC_SERVER           #x1)
(defconstant CLSCTX_INPROC_HANDLER          #x2)
(defconstant CLSCTX_LOCAL_SERVER            #x4)
(defconstant CLSCTX_INPROC_SERVER16         #x8)
(defconstant CLSCTX_REMOTE_SERVER           #x10)
(defconstant CLSCTX_INPROC_HANDLER16        #x20)
(defconstant CLSCTX_RESERVED1               #x40)
(defconstant CLSCTX_RESERVED2               #x80)
(defconstant CLSCTX_RESERVED3               #x100)
(defconstant CLSCTX_RESERVED4               #x200)
(defconstant CLSCTX_NO_CODE_DOWNLOAD        #x400)
(defconstant CLSCTX_RESERVED5               #x800)
(defconstant CLSCTX_NO_CUSTOM_MARSHAL       #x1000)
(defconstant CLSCTX_ENABLE_CODE_DOWNLOAD    #x2000)
(defconstant CLSCTX_NO_FAILURE_LOG          #x4000)
(defconstant CLSCTX_DISABLE_AAA             #x8000)
(defconstant CLSCTX_ENABLE_AAA              #x10000)
(defconstant CLSCTX_FROM_DEFAULT_CONTEXT    #x20000)
(defconstant CLSCTX_ACTIVATE_32_BIT_SERVER  #x40000)
(defconstant CLSCTX_ACTIVATE_64_BIT_SERVER  #x80000)


(defconstant DISPATCH_METHOD         #x1)
(defconstant DISPATCH_PROPERTYGET    #x2)
(defconstant DISPATCH_PROPERTYPUT    #x4)
(defconstant DISPATCH_PROPERTYPUTREF #x8)

(defconstant DISPID_UNKNOWN -1)
(defconstant DISPID_VALUE 0)
(defconstant DISPID_PROPERTYPUT -3)
(defconstant DISPID_NEWENUM -4)
(defconstant DISPID_EVALUATE -5)
(defconstant DISPID_CONSTRUCTOR	-6)
(defconstant DISPID_DESTRUCTOR -7)
(defconstant DISPID_COLLECT -8)

;; FUNCKIND
(defconstant FUNC_VIRTUAL 0)
(defconstant FUNC_PUREVIRTUAL (1+ FUNC_VIRTUAL))
(defconstant FUNC_NONVIRTUAL (1+ FUNC_PUREVIRTUAL))
(defconstant FUNC_STATIC (1+ FUNC_NONVIRTUAL))
(defconstant FUNC_DISPATCH (1+ FUNC_STATIC))

;; INVOKEKIND
(defconstant INVOKE_FUNC 1)
(defconstant INVOKE_PROPERTYGET 2)
(defconstant INVOKE_PROPERTYPUT 4)
(defconstant INVOKE_PROPERTYPUTREF 8)

;; CALLCONV
(defconstant CC_FASTCALL 0)
(defconstant CC_CDECL 1)
(defconstant CC_MSCPASCAL (1+ CC_CDECL))
(defconstant CC_PASCAL CC_MSCPASCAL)
(defconstant CC_MACPASCAL (1+ CC_PASCAL))
(defconstant CC_STDCALL (1+ CC_MACPASCAL))
(defconstant CC_FPFASTCALL (1+ CC_STDCALL))
(defconstant CC_SYSCALL (1+ CC_FPFASTCALL))
(defconstant CC_MPWCDECL (1+ CC_SYSCALL))
(defconstant CC_MPWPASCAL (1+ CC_MPWCDECL))
(defconstant CC_MAX (1+ CC_MPWPASCAL))

(defconstant VAR_PERINSTANCE 0)
(defconstant VAR_STATIC (1+ VAR_PERINSTANCE))
(defconstant VAR_CONST (1+ VAR_STATIC))
(defconstant VAR_DISPATCH (1+ VAR_CONST))


(defconstant VARFLAG_FREADONLY #x1)
(defconstant VARFLAG_FSOURCE #x2)
(defconstant VARFLAG_FBINDABLE #x4)
(defconstant VARFLAG_FREQUESTEDIT #x8)
(defconstant VARFLAG_FDISPLAYBIND #x10)
(defconstant VARFLAG_FDEFAULTBIND #x20)
(defconstant VARFLAG_FHIDDEN #x40)
(defconstant VARFLAG_FRESTRICTED #x80)
(defconstant VARFLAG_FDEFAULTCOLLELEM #x100)
(defconstant VARFLAG_FUIDEFAULT #x200)
(defconstant VARFLAG_FNONBROWSABLE #x400)
(defconstant VARFLAG_FREPLACEABLE #x800)
(defconstant VARFLAG_FIMMEDIATEBIND #x1000)
