#include "RCOMObject.h"
#include <oleauto.h>
#include <oaidl.h>
#include <tchar.h>

// #undef ERROR
extern "C" {
#include "RUtils.h"
#include <Rdefines.h>

  SEXP R_getDynamicVariantValue(SEXP ref);
  SEXP R_setDynamicVariantValue(SEXP ref, SEXP value);
}

#include "converters.h"

static SEXP convertArrayToR(const VARIANT *var);
void GetScodeString(HRESULT hr, LPTSTR buf, int bufSize);
SEXP UnList(SEXP ans);

BSTR
AsBstr(const char *str)
{
  BSTR ans = NULL;
  if(!str)
    return(NULL);

  int size = strlen(str);
  int wideSize = 2 * size;
  LPOLESTR wstr = (LPWSTR) S_alloc(wideSize, sizeof(OLECHAR)); 
  if(MultiByteToWideChar(CP_ACP, 0, str, size, wstr, wideSize) == 0 && str[0]) {
    PROBLEM "Can't create BSTR for '%s'", str
    ERROR;
  }

  ans = SysAllocStringLen(wstr, size);

  return(ans);
}

char *
FromBstr(BSTR str)
{
  char *ptr = NULL;
  DWORD len;

  if(!str)
    return(NULL);

  len = wcslen(str);

  if(len < 1)
    len = 0;

  ptr = (char *) S_alloc(len+1, sizeof(char));
  ptr[len] = '\0';
  if(len > 0) {
    DWORD ok = WideCharToMultiByte(CP_ACP, 0, str, len, ptr, len, NULL, NULL);
    if(ok == 0) 
      ptr = NULL;
  }

  return(ptr);
}


/*
 Get the number of dimensions.
 For each of these dimensions, get the lower and upper bound and iterate
 over the elements.
*/
static SEXP
convertArrayToR(const VARIANT *var)
{
  SAFEARRAY *arr;
  SEXP ans;
  UINT dim;

  if(V_ISBYREF(var))
    arr = *V_ARRAYREF(var);
  else
    arr = V_ARRAY(var);

  dim = SafeArrayGetDim(arr);
  long *indices = (long*) S_alloc(dim, sizeof(long)); // new long[dim];
  ans = getArray(arr, dim, dim, indices);

  return(ans);
}

SEXP
getArray(SAFEARRAY *arr, int dimNo, int numDims, long *indices)
{
  long lb, ub, n,  i;
  HRESULT status;
  SEXP ans;
  int rtype = -1;

  status = SafeArrayGetLBound(arr, dimNo, &lb);
  if(FAILED(status)) {
    TCHAR buf[512];
    GetScodeString(status, buf, sizeof(buf)/sizeof(buf[0]));
    PROBLEM "Can't get lower bound of array: %s", buf
    ERROR;
  }
  status = SafeArrayGetUBound(arr, dimNo, &ub);
  if(FAILED(status)) {
    TCHAR buf[512];
    GetScodeString(status, buf, sizeof(buf)/sizeof(buf[0]));
    PROBLEM "Can't get upper bound of array: %s", buf
    ERROR;
  }

  n = ub-lb+1;
  PROTECT(ans = NEW_LIST(n));

  for(i = 0; i < n; i++) {
    SEXP el;
    indices[dimNo - 1] = lb + i;
    if(dimNo == 1) {
      VARIANT variant;
      VariantInit(&variant);
      status = SafeArrayGetElement(arr, indices, &variant);
      if(FAILED(status)) {
        TCHAR buf[512];
        GetScodeString(status, buf, sizeof(buf)/sizeof(buf[0]));
        PROBLEM "Can't get element %d of array %s", (int) indices[dimNo-1], buf
        ERROR;
      } 
      el = R_convertDCOMObjectToR(&variant);
    } else {
      el = getArray(arr, dimNo - 1, numDims, indices);
    }
    if(i == 0)
      rtype = TYPEOF(el);
    else if(rtype != -1 ){
      if(TYPEOF(el) != rtype)
	rtype = -1;
    }
    SET_VECTOR_ELT(ans, i, el);
  }

  UNPROTECT(1);

  return(ans);
}

SEXP
UnList(SEXP ans)
{
  SEXP e, val;
  int errorOccurred;

  PROTECT(e = allocVector(LANGSXP, 2));
  SETCAR(e, Rf_install("unlist"));
  SETCAR(CDR(e), ans);
  val = R_tryEval(e, R_GlobalEnv, &errorOccurred);
  UNPROTECT(1);

  return(errorOccurred ? ans : val);
}

void 
R_typelib_finalizer(SEXP obj)
{
    R_ClearExternalPtr(obj);
}


void
R_Variant_finalizer(SEXP s)
{
  VARIANT *var;
  var = (VARIANT *) R_ExternalPtrAddr(s);
  if(var) {
    VariantClear(var);
    free(var);
    R_ClearExternalPtr(s);
  }
}

SEXP
createRVariantObject(VARIANT *var,  VARTYPE kind)
{
  const char *className;
  SEXP klass, ans, tmp;
  VARIANT *dupvar;  
  switch(kind) {
    case VT_DATE:
      className = "DateVARIANT";
      break;
    case VT_CY:
      className = "CurrencyVARIANT";
      break;

    default:
      className = "VARIANT";
  }

  PROTECT(klass = MAKE_CLASS(className));
  if(klass == NULL || klass == R_NilValue) {
     PROBLEM  "Can't locate S4 class definition %s", className
     ERROR;
  }
  
  dupvar = (VARIANT *) malloc(sizeof(VARIANT));
  VariantCopyInd(dupvar, var);
  
  PROTECT(ans = NEW(klass));
  PROTECT(tmp = R_MakeExternalPtr(dupvar, Rf_install(className), R_NilValue));
  R_RegisterCFinalizer(tmp, R_Variant_finalizer);
  SET_SLOT(ans, Rf_install("ref"), tmp);
  UNPROTECT(1);

  PROTECT(tmp = NEW_INTEGER(1));
  INTEGER(tmp)[0] = kind;
  SET_SLOT(ans, Rf_install("kind"), tmp);
  
  UNPROTECT(3);
  return(ans);
}

/**
  Turn a variant into an S object with a special class
  such as COMDate or COMCurrency which is simply an extension
  of numeric.
*/
SEXP
numberFromVariant(VARIANT *var, VARTYPE type)
{
  SEXP ans;
  SEXP klass;
  const char *tmpName = NULL;

  switch(type) {
  case VT_CY:
    tmpName = "COMCurrency";
    break;
  case VT_DATE:
    tmpName = (char *) "COMDate";
    break;
  case VT_HRESULT:
    tmpName = (char *) "HResult";
    break;
  case VT_DECIMAL:
    tmpName = (char *) "COMDecimal";
    break;
  default:
    PROBLEM "numberFromVariant called with unsupported variant type."
     ERROR;
  }
  PROTECT(klass = MAKE_CLASS(tmpName));
  PROTECT(ans = NEW(klass));
  ans = R_do_slot_assign(ans, mkString(".Data"), R_scalarReal(V_R8(var)));
      // SET_SLOT(ans, Rf_install(".Data"), R_scalarReal(V_R8(var)));
  UNPROTECT(2);

  return(ans);
}


static SEXP
createVariantRef(VARIANT *var, VARTYPE baseType)
{
  SEXP e, ans = R_NilValue, ref;
  PROTECT(e = allocVector(LANGSXP, 3));
  SETCAR(e, Rf_install("createDynamicVariantReference"));
  ref = R_MakeExternalPtr((void *) var, Rf_install("VARIANTReference"), R_NilValue);
  SETCAR(CDR(e), ref);
  SETCAR(CDR(CDR(e)), ScalarInteger(baseType));

  ans = R_tryEval(e, R_GlobalEnv, NULL);
  UNPROTECT(1);

  return(ans);
}

static VARIANT *
R_getVariantRef(SEXP ref)
{
  VARIANT *p;

  if(TYPEOF(ref) != EXTPTRSXP) {
    PROBLEM "Argument to R_getVariantRef must be an external pointer"
      ERROR;
  }

  if(EXTPTR_TAG(ref) != Rf_install("VARIANTReference")) {
    PROBLEM "Argument to R_getVariantRef does not have the correct tag."
    ERROR;
  }

  p = (VARIANT *) R_ExternalPtrAddr(ref);
  return(p);
}

SEXP
R_getDynamicVariantValue(SEXP ref)
{
  VARIANT *var;
  VARTYPE rtype;

  var = R_getVariantRef(ref);
  rtype = V_VT(var) & (~ VT_BYREF);
  switch(rtype) {
  case VT_BOOL:
    return(ScalarLogical(*V_BOOLREF(var)));
    break;
  case VT_I4:
    return(ScalarInteger(*V_I4REF(var)));
    break;
  case VT_R8:
    return(ScalarReal(*V_R8REF(var)));
    break;
  default:
    return(R_NilValue);
  }

  return(R_NilValue);
}



SEXP
R_setDynamicVariantValue(SEXP ref, SEXP val)
{
  VARIANT *var;
  VARTYPE rtype;

  var = R_getVariantRef(ref);
  rtype = V_VT(var) & (~ VT_BYREF);
  switch(rtype) {
  case VT_BOOL:
    *V_BOOLREF(var) = LOGICAL(val)[0];
    break;
  case VT_I4:
    *V_I4REF(var) = INTEGER(val)[0];
    break;
  case VT_R8:
    *V_R8REF(var) = REAL(val)[0];
    break;
  default:
    return(R_NilValue);
  }

  return(R_NilValue);
}


/* Taken from connect.cpp in RDCOMClient. */

SEXP
R_convertDCOMObjectToR(VARIANT *var)
{
  SEXP ans = R_NilValue;
  HRESULT hr;

  VARTYPE type = V_VT(var);

#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
  errorLog("Converting VARIANT to R %d\n", V_VT(var));
#endif


  if(V_ISARRAY(var)) {
#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
  errorLog("Finishing convertDCOMObjectToR - convert array\n");
#endif
    return(convertArrayToR(var));
  } else if(V_VT(var) == VT_DISPATCH || (V_ISBYREF(var) && ((V_VT(var) & (~ VT_BYREF)) == VT_DISPATCH)) ) {
    IDispatch *ptr;
    if(V_ISBYREF(var)) {

#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
      errorLog("BYREF and DISPATCH in convertDCOMObjectToR\n");
#endif

      IDispatch **tmp = V_DISPATCHREF(var);
      if(!tmp)
	return(ans);
      ptr = *tmp;
    } else
      ptr = V_DISPATCH(var);
       //xxx
    if(ptr) 
      ptr->AddRef();
    ans = R_createRCOMUnknownObject((void*) ptr, "COMIDispatch");
#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
    errorLog("Finished convertDCOMObjectToR  COMIDispatch\n");
#endif
    return(ans);
  }



  if(V_ISBYREF(var)) {
    VARTYPE rtype = type & (~ VT_BYREF);

#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
    errorLog("ISBYREF() in convertDCOMObjectToR: ref type %d\n", rtype);
#endif

    if(rtype == VT_BSTR) {
        BSTR *tmp;
        const char *ptr = "";
#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
	errorLog("BYREF and BSTR convertDCOMObjectToR  (scalar string)\n");
#endif
        tmp = V_BSTRREF(var);
        if(tmp)
  	  ptr = FromBstr(*tmp);
        ans = R_scalarString(ptr);
	return(ans);
    } else if(rtype == VT_BOOL || rtype == VT_I4 || rtype == VT_R8){
      return(createVariantRef(var, rtype));
    } else {
        fprintf(stderr, "Unhandled by-reference conversion type %d\n", V_VT(var));fflush(stderr);
	return(R_NilValue);        
    }
  }

  switch(type) {

    case VT_BOOL:
      ans = R_scalarLogical( (Rboolean) (V_BOOL(var) ? TRUE : FALSE));
      break;

    case VT_UI1:
    case VT_UI2:
    case VT_UI4:
    case VT_UINT:
      hr = VariantChangeType(var, var, 0, VT_I4);
      ans = R_scalarReal((double) V_I4(var));
      break;

    case VT_I1:
    case VT_I2:
    case VT_I4:
    case VT_INT:
      hr = VariantChangeType(var, var, 0, VT_I4);
      ans = R_scalarInteger(V_I4(var));
      break;

    case VT_R4:
    case VT_R8:
    case VT_I8:
      hr = VariantChangeType(var, var, 0, VT_R8);
      ans = R_scalarReal(V_R8(var));
      break;

    case VT_CY:
    case VT_DATE:
    case VT_HRESULT:
    case VT_DECIMAL:
      hr = VariantChangeType(var, var, 0, VT_R8);
      ans = numberFromVariant(var, type);
      break;

    case VT_BSTR:
      {
	char *ptr = FromBstr(V_BSTR(var));
        ans = R_scalarString(ptr);
      }
      break;

    case VT_UNKNOWN:
      {
       IUnknown *ptr = V_UNKNOWN(var);
       //xxx
       if(ptr)
         ptr->AddRef();
       ans = R_createRCOMUnknownObject((void**) ptr, "COMUnknown");
      }
       break;
  case VT_EMPTY:
  case VT_NULL:
  case VT_VOID:
    return(R_NilValue);
    break;


/*XXX Need to fill these in */
  case VT_RECORD:
  case VT_FILETIME:
  case VT_BLOB:
  case VT_STREAM:
  case VT_STORAGE:
  case VT_STREAMED_OBJECT:
    /*  case LPSTR: */
  case VT_LPWSTR:
  case VT_PTR:
  case VT_ERROR:
  case VT_VARIANT:
  case VT_CARRAY:
  case VT_USERDEFINED:
  default:
    fprintf(stderr, "Unhandled conversion type %d\n", V_VT(var));fflush(stderr);
    //XXX this consumes the variant. So the variant clearance in Invoke() does it again!
    ans = createRVariantObject(var, V_VT(var));
  }

#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
  errorLog("Finished convertDCOMObjectToR\n");
#endif

  return(ans);
}

VARTYPE
getDCOMType(SEXP obj)
{
  VARTYPE val = VT_UNKNOWN;

  switch(TYPEOF(obj)) {
    case REALSXP:
      val = VT_R8;
      break;
    case LGLSXP:
      val = VT_BOOL;
      break;
    case INTSXP:
      val = VT_I4;
      break;
    case STRSXP:
      val = VT_BSTR;
      break;
    case VECSXP:
      val = VT_VARIANT;
      break;
    default:
      break;
  }

  return(val);
}

SAFEARRAY*
createRDCOMArray(SEXP obj, VARIANT *var)
{
  VARTYPE type;
  unsigned int cDims = 1, len;
  SAFEARRAYBOUND bounds[1];
  SAFEARRAY *arr;
  void *data;

  len = Rf_length(obj);
  bounds[0].lLbound = 0;
  bounds[0].cElements = len;

  type = getDCOMType(obj);
  arr = SafeArrayCreate(type, cDims, bounds);

  HRESULT hr = SafeArrayAccessData(arr, (void**) &data);
  if(hr != S_OK) {
    std::cerr <<"Problems accessing data" << std::endl;
    SafeArrayDestroy(arr);
    return(NULL);
  }

  switch(TYPEOF(obj)) {
    case REALSXP:
      memcpy(data, REAL(obj), sizeof(double) * len);
      break;
    case INTSXP:
      memcpy(data, INTEGER(obj), sizeof(LOGICAL(obj)[0]) * len);
      break;
    case LGLSXP:
      for(unsigned int i = 0 ; i < len ; i++)
	((bool *) data)[i] = LOGICAL(obj)[i];
      break;
    case STRSXP:
      for(unsigned int i = 0 ; i < len ; i++)
	((BSTR *) data)[i] = AsBstr(getRString(obj, i));
      break;
    case VECSXP:
      for(unsigned int i = 0 ; i < len ; i++) {
	VARIANT *v = &(((VARIANT *) data)[i]);
	VariantInit(v);
	R_convertRObjectToDCOM(VECTOR_ELT(obj, i), v);
      }
      break;

    default:
      std::cerr <<"Array case not handled yet for R type " << TYPEOF(obj) << std::endl;
    break;
  }

 SafeArrayUnaccessData(arr);

  if(var) {
    V_VT(var) = VT_ARRAY | type;
    V_ARRAY(var) = arr;
  }

  return(arr);
}

HRESULT
createGenericCOMObject(SEXP obj, VARIANT *var)
{
  SEXP e, val;
  int errorOccurred;

  /* Make certain RDCOMServer is loaded as this might be invoked
     as part of RDCOMClient. */
  PROTECT(e = allocVector(LANGSXP, 3));
  SETCAR(e, Rf_install("require"));
  SETCAR(CDR(e), Rf_install("RDCOMServer"));
  SETCAR(CDR(CDR(e)), val = allocVector(LGLSXP, 1));
  INTEGER(val)[0] = TRUE;
  SET_TAG(CDR(CDR(e)), Rf_install("quiet"));

  val = R_tryEval(e, R_GlobalEnv, &errorOccurred);
  UNPROTECT(1);
  if(!LOGICAL(val)[0]) {
    PROBLEM  "Can't attach the RDCOMServer package needed to create a generic COM object"
    ERROR;
    return(S_FALSE);
  }

  PROTECT(e = allocVector(LANGSXP, 2));
  SETCAR(e, Rf_install("createCOMObject"));
  SETCAR(CDR(e), obj);
  val = R_tryEval(e, R_GlobalEnv, &errorOccurred);
  if(errorOccurred) {
    UNPROTECT(1);
    PROBLEM "Can't create COM object"
    ERROR;
    return(S_FALSE);
  }

  RCOMObject *robj;
  if(TYPEOF(val) != EXTPTRSXP)
    return(S_FALSE);

  robj = (RCOMObject *) R_ExternalPtrAddr(val);
  V_VT(var) = VT_DISPATCH;
  V_DISPATCH(var) = robj;

  return(S_OK);
}

HRESULT
R_convertRObjectToDCOM(SEXP obj, VARIANT *var)
{
  HRESULT status;
  int type = R_typeof(obj);

  if(!var)
    return(S_FALSE);

#ifdef RDCOM_VERBOSE
  errorLog("Type of argument %d\n", type);
#endif

 if(type == EXTPTRSXP && EXTPTR_TAG(obj) == Rf_install("R_VARIANT")) {
   VARIANT *tmp;
   tmp = (VARIANT *) R_ExternalPtrAddr(obj);
   if(tmp) {
     //XXX
     VariantCopy(var, tmp);
     return(S_OK);
   }
 }

 if(ISCOMIDispatch(obj)) {
   IDispatch *ptr;
   ptr = (IDispatch *) derefRIDispatch(obj);
   V_VT(var) = VT_DISPATCH;
   V_DISPATCH(var) = ptr;
   //XX
   ptr->AddRef();
   return(S_OK);
 }

 if(ISSInstanceOf(obj, "COMDate")) {
    double val;
    val = NUMERIC_DATA(GET_SLOT(obj, Rf_install(".Data")))[0];
    V_VT(var) = VT_DATE;
    V_DATE(var) = val;
    return(S_OK);
 } else if(ISSInstanceOf(obj, "COMCurrency")) {
    double val;
    val = NUMERIC_DATA(GET_SLOT(obj, Rf_install(".Data")))[0];
    V_VT(var) = VT_R8;
    V_R8(var) = val;
    VariantChangeType(var, var, 0, VT_CY);
    return(S_OK);
 } else if(ISSInstanceOf(obj, "COMDecimal")) {
    double val;
    val = NUMERIC_DATA(GET_SLOT(obj, Rf_install(".Data")))[0];
    V_VT(var) = VT_R8;
    V_R8(var) = val;
    VariantChangeType(var, var, 0, VT_DECIMAL);
    return(S_OK);
 }


 /* We have a complex object and we are not going to try to convert it directly
    but instead create an COM server object to represent it to the outside world. */
  if((type == VECSXP && Rf_length(GET_NAMES(obj))) || Rf_length(GET_CLASS(obj)) > 0  || isMatrix(obj)) {
    status = createGenericCOMObject(obj, var);
    if(status == S_OK)
      return(S_OK);
  }

  if(Rf_length(obj) == 0) {
   V_VT(var) = VT_VOID;
   return(S_OK);
  }

  if(type == VECSXP || Rf_length(obj) > 1) {
      createRDCOMArray(obj, var);
      return(S_OK);
  }

  switch(type) {
    case STRSXP:
      V_VT(var) = VT_BSTR;
      V_BSTR(var) = AsBstr(getRString(obj, 0));
      break;

    case INTSXP:
      V_VT(var) = VT_I4;
      V_I4(var) = R_integerScalarValue(obj, 0);
      break;

    case REALSXP:
	V_VT(var) = VT_R8;
	V_R8(var) = R_realScalarValue(obj, 0);
      break;

    case LGLSXP:
      V_VT(var) = VT_BOOL;
      V_BOOL(var) = R_logicalScalarValue(obj, 0) ? VARIANT_TRUE : VARIANT_FALSE;
      break;

    case VECSXP:
      break;
  }
  
  return(S_OK);
}

extern "C" {
  void registerCOMObject(void *, int);
}

void
RDCOM_finalizer(SEXP s)
{
 IUnknown *ptr = (IUnknown*) derefRDCOMPointer(s);
 if(ptr) {
#ifdef ANNOUNCE_COM_CALLS
   fprintf(stderr, "Releasing COM object %p\n", ptr);fflush(stderr);
#endif

#ifdef REGISTER_COM_OBJECTS_WITH_S
   registerCOMObject(ptr, 0);
#endif

   //XXX 
  ptr->Release();
#ifdef ANNOUNCE_COM_CALLS
   fprintf(stderr, "Released COM object %p\n", ptr);fflush(stderr);
#endif
   R_ClearExternalPtr(s);
 }
}

void
RDCOM_SafeArray_finalizer(SEXP s)
{
  SAFEARRAY *arr;
  arr = (SAFEARRAY*) R_ExternalPtrAddr(s);
  if(arr) {
    SafeArrayDestroy(arr);
    R_ClearExternalPtr(s);
  }
}

SEXP
R_create2DArray(SEXP obj)
{
  SAFEARRAYBOUND bounds[2] =  {{0, 0}, {0, 0}};;
  SAFEARRAY *arr;
  void *data, *el;
  VARTYPE type = VT_R8;
  SEXP dim = GET_DIM(obj);
  int integer;
  double real;
  BSTR bstr;


  bounds[0].cElements = INTEGER(dim)[0];
  bounds[1].cElements = INTEGER(dim)[1];

  type = getDCOMType(obj);

  arr = SafeArrayCreate(type, 2, bounds);
  SafeArrayAccessData(arr, (void**) &data);

  long indices[2];
  UINT i, j, ctr = 0;
  for(j = 0 ; j < bounds[1].cElements; j++) {
    indices[1] = j;
    for(i = 0; i < bounds[0].cElements; i++, ctr++) {
      indices[0] = i;
      switch(TYPEOF(obj)) {
        case LGLSXP:
	  integer =  (LOGICAL(obj)[ctr] ? 1:0);
          el = &integer;
	  break;
        case REALSXP:
	  real = REAL(obj)[ctr];
          el = &real;
	  break;
        case INTSXP:
	  integer = INTEGER(obj)[ctr];
          el = &integer;
	  break;
        case STRSXP:
	  bstr = AsBstr(CHAR(STRING_ELT(obj, ctr)));
          el = (void*) bstr;
	  break;
        default:
	  continue;
	  break;
      }

      SafeArrayPutElement(arr, indices, el);
    }
  }
  SafeArrayUnaccessData(arr);

  VARIANT *var;
  var = (VARIANT*) malloc(sizeof(VARIANT));
  VariantInit(var);
  V_VT(var) = VT_ARRAY | type;
  V_ARRAY(var) = arr;

  SEXP ans;
  PROTECT(ans = R_MakeExternalPtr((void*) var, Rf_install("R_VARIANT"), R_NilValue));
  R_RegisterCFinalizer(ans, RDCOM_SafeArray_finalizer);  
  UNPROTECT(1);
  return(ans);
}

extern "C"
SEXP
R_createVariant(SEXP type)
{
  VARIANT var;
  VariantInit(&var);
  return(createRVariantObject(&var, INTEGER_DATA(type)[0]));
}


SEXP
R_setVariant(SEXP svar, SEXP value, SEXP type)
{
  VARIANT *var;
  var = (VARIANT *)R_ExternalPtrAddr(GET_SLOT(svar, Rf_install("ref")));
  if(!var) {
    PROBLEM "Null VARIANT value passed to R_setVariant. Was this saved in another session\n"
   ERROR;
  }

  HRESULT hr;
  hr = R_convertRObjectToDCOM(value, var);

  SEXP ans;
  ans = NEW_LOGICAL(1);
  LOGICAL_DATA(ans)[0] = hr == S_OK ? TRUE : FALSE;
  return(ans);
}


