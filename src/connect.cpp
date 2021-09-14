#ifndef _GNU_
#include "stdafx.h"
#include <objbase.h>
#else
#include <windows.h>
#include <oaidl.h>
#include <oleauto.h>
#ifndef V_UI4
# define V_UI4(X) V_UNION((X), uintVal)
#endif

#ifdef ERROR
#undef ERROR
#endif

#endif


extern "C" {
#include "RUtils.h"
#include <Rinternals.h>
#include <Rdefines.h>
}



#include "converters.h"


#ifdef _GNU_
#include "RUtils.h"
#define R_logicalScalarValue(x, i) LOGICAL((x))[(i)]
#define R_integerScalarValue(x, i) INTEGER((x))[(i)]
#define R_realScalarValue(x, i)    REAL((x))[(i)]
#endif

extern HRESULT checkErrorInfo(IUnknown *obj, HRESULT status, SEXP *serr);


extern "C" {

  __declspec(dllexport) SEXP R_initCOM(SEXP);

  __declspec(dllexport) SEXP R_connect(SEXP className, SEXP raiseError);

  __declspec(dllexport) SEXP R_create(SEXP className);

#ifdef UNUSED
  __declspec(dllexport) SEXP R_invoke(SEXP obj, SEXP methodName, SEXP args);
#endif

  __declspec(dllexport) SEXP R_Invoke(SEXP obj, SEXP methodName, SEXP args, SEXP type, SEXP sreturn, SEXP ids);

  __declspec(dllexport) SEXP R_getInvokeTypes();

  __declspec(dllexport) SEXP R_getProperty(SEXP obj, SEXP propertyName, SEXP args, SEXP ids);
  __declspec(dllexport) SEXP R_setProperty(SEXP obj, SEXP propertyName, SEXP value, SEXP ids);


  __declspec(dllexport) SEXP R_getCLSIDFromName(SEXP className);

  __declspec(dllexport) SEXP R_isValidCOMObject(SEXP obj);

  /* Doesn't work. need to figure out how to get R symbols to work here
     in this C++ compiled code.
     extern int * INTEGER(struct SEXPREC *);
  */
} /* end of extern "C" */

#ifndef _GNU_
#define GET_NAMES(x) getRNames((x))
#endif

SEXP R_COM_Invoke(SEXP obj, SEXP methodName, SEXP args, WORD callType, WORD doReturn, SEXP ids);
SEXP R_convertDCOMObjectToR(const VARIANT *var);
HRESULT R_getCOMArgs(SEXP args, DISPPARAMS *parms, DISPID *, int numNamedArgs, int *namedArgPos);
HRESULT R_convertRObjectToDCOM(SEXP obj, VARIANT *var);

void COMError(HRESULT hr);

void GetScodeString(HRESULT hr, LPTSTR buf, int bufSize);

__declspec(dllexport)
SEXP
R_initCOM(SEXP val)
{
  SEXP ans = R_NilValue;
  if(R_logicalScalarValue(val, 0)) {
    CoInitialize(NULL);
  } else 
    CoUninitialize();

  return(ans);
}

HRESULT
R_getCLSIDFromString(SEXP className, CLSID *classId)
{
  HRESULT hr;
  const char *ptr;
  int status = FALSE;
  BSTR str;

  ptr = CHAR(STRING_ELT(className, 0)); 
  str = AsBstr(ptr);
   
  hr = CLSIDFromString(str, classId);
  if(SUCCEEDED(hr)) {
    SysFreeString(str);
    return(S_OK);
  }

  status = CLSIDFromProgID(str, classId);
  SysFreeString(str);

  return status;
}

__declspec(dllexport)
SEXP 
R_getCLSIDFromName(SEXP className)
{
  CLSID classId;
  HRESULT hr;
  SEXP ans;

  hr = R_getCLSIDFromString(className, &classId);
  if(!SUCCEEDED(hr)) {
    COMError(hr);
  }

  LPOLESTR str;
  hr = StringFromCLSID(classId, &str);
  if(!SUCCEEDED(hr))
    COMError(hr);

  //???
  ans = mkString(FromBstr(str));
  CoTaskMemFree(str);

  return(ans);
}


     /*XXXX This needs some work! Does it still? */
__declspec(dllexport)
SEXP
R_connect(SEXP className, SEXP raiseError)
{
  IUnknown *unknown = NULL;
  HRESULT hr;
  SEXP ans = R_NilValue;
  CLSID classId;

  if(R_getCLSIDFromString(className, &classId) == S_OK) {
    hr = GetActiveObject(classId, NULL, &unknown);
    if(SUCCEEDED(hr)) {
      void *ptr;
      hr = unknown->QueryInterface(IID_IDispatch, &ptr);
      ans = R_createRCOMUnknownObject((void *) ptr, "COMIDispatch");
    } else {
      if(LOGICAL(raiseError)[0]) {
	/* From COMError.cpp  - COMError */
	TCHAR buf[512];
	GetScodeString(hr, buf, sizeof(buf)/sizeof(buf[0]));
	PROTECT(ans = mkString(buf));
	SET_CLASS(ans, mkString("COMErrorString"));
	UNPROTECT(1);
        return(ans);
      } else
	return(R_NilValue);
    }
  } else {
      PROBLEM "Couldn't get clsid from the string"
	WARN;
  }
  return(ans);
}

/*
 Routine that is used to create a COM object from its name or CLSID.
*/
__declspec(dllexport)
SEXP
R_create(SEXP className)
{
  DWORD context = CLSCTX_SERVER;
  SEXP ans;
  CLSID classId;
  IID refId = IID_IDispatch;
  IUnknown *unknown, *punknown = NULL;

  HRESULT hr = R_getCLSIDFromString(className, &classId);
  if(FAILED(hr))  
    COMError(hr);

  SCODE sc = CoCreateInstance(classId,  punknown,  context, refId, (void **) &unknown);

  if(FAILED(sc)) {
   TCHAR buf[512];
   GetScodeString(sc, buf, sizeof(buf)/sizeof(buf[0]));
   PROBLEM "Failed to create COM object: %s", buf
   ERROR;
  }

  //Already AddRef in the CoCreateInstance
  // so no need to do it now ( unknown->AddRef())

  ans = R_createRCOMUnknownObject((void *) unknown, "COMIDispatch");

  return(ans);
}

/*
 General interface from R to invoke a COM method or property accessor.
  @type integer giving the invocation kind/style (e.g. property get, property put, invoke)
  @sreturn logical indicating whether the return value should be converted or not.
 */
__declspec(dllexport) SEXP 
R_Invoke(SEXP obj, SEXP methodName, SEXP args, SEXP stype, SEXP sreturn, SEXP ids)
{
  WORD type = R_integerScalarValue(stype, 0);
  WORD doReturn = R_logicalScalarValue(sreturn, 0);
  return(R_COM_Invoke(obj, methodName, args, type, doReturn, ids));
}

/* Utility routine to clear the variants in the argument structure of the COM call. */
static int
clearVariants(DISPPARAMS *params)
{
 if(params->cArgs) {
   for(unsigned int i = 0; i < params->cArgs; i++) {
     VariantClear(&params->rgvarg[i]);
   }
 } 
 return(params->cArgs);
}




void
freeSysStrings(BSTR *els, int num)
{
  if(els) {
    for(int i = 0; i < num ; i++)
      SysFreeString(els[i]);
  }
}


/* 
 The real invoke mechanism that handles all the details.
*/
SEXP
R_COM_Invoke(SEXP obj, SEXP methodName, SEXP args, WORD callType, WORD doReturn,
             SEXP ids)
{
 IDispatch* disp;
 SEXP ans = R_NilValue;
 int numNamedArgs = 0, *namedArgPositions = NULL, i;
 HRESULT hr;

 // callGC();
 disp = (IDispatch *) getRDCOMReference(obj);

#ifdef ANNOUNCE_COM_CALLS
 fprintf(stderr, "<COM> %s %d %p\n", CHAR(STRING_ELT(methodName, 0)), (int) callType, 
                                     disp);fflush(stderr);
#endif

 DISPID *methodIds;
 const char *pmname = CHAR(STRING_ELT(methodName, 0));
 BSTR *comNames = NULL;

 SEXP names = GET_NAMES(args);
 int numNames = Rf_length(names) + 1;

 SetErrorInfo(0L, NULL);

 methodIds = (DISPID *) S_alloc(numNames, sizeof(DISPID));
 namedArgPositions = (int*) S_alloc(numNames, sizeof(int)); // we may not use all of these, but we allocate them

 if(Rf_length(ids) == 0) {
     comNames = (BSTR*) S_alloc(numNames, sizeof(BSTR));

     comNames[0] = AsBstr(pmname);
     for(i = 0; i < Rf_length(names); i++) {
       const char *str = CHAR(STRING_ELT(names, i));
       if(str && str[0]) {
         comNames[numNamedArgs+1] = AsBstr(str);
         namedArgPositions[numNamedArgs] = i;
         numNamedArgs++;
       }
     }
     numNames = numNamedArgs + 1;

     hr = disp->GetIDsOfNames(IID_NULL, comNames, numNames, LOCALE_USER_DEFAULT, methodIds);

     if(FAILED(hr) || hr == DISP_E_UNKNOWNNAME /* || DISPID mid == DISPID_UNKNOWN */) {
       PROBLEM "Cannot locate %d name(s) %s in COM object (status = %d)", numNamedArgs, pmname, (int) hr
	 ERROR;
     }
 } else {
   for(i = 0; i < Rf_length(ids); i++) {
     methodIds[i] = (MEMBERID) NUMERIC_DATA(ids)[i];
     //XXX What about namedArgPositions here.
   }
 }


 DISPPARAMS params = {NULL, NULL, 0, 0};
 
 if(args != NULL && Rf_length(args) > 0) {

   hr = R_getCOMArgs(args, &params, methodIds, numNamedArgs, namedArgPositions);

   if(FAILED(hr)) {
     clearVariants(&params);
     freeSysStrings(comNames, numNames);
     PROBLEM "Failed in converting arguments to DCOM call"
     ERROR;
   }
   if(callType & DISPATCH_PROPERTYPUT) {
     params.rgdispidNamedArgs = (DISPID*) S_alloc(1, sizeof(DISPID));
     params.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;
     params.cNamedArgs = 1;
   }
 }

 VARIANT varResult, *res = NULL;

 if(doReturn && callType != DISPATCH_PROPERTYPUT)
   VariantInit(res = &varResult);

 EXCEPINFO exceptionInfo;
 memset(&exceptionInfo, 0, sizeof(exceptionInfo));
 unsigned int nargErr = 100;

#ifdef RDCOM_VERBOSE
 if(params.cNamedArgs) {
   errorLog("# of named arguments to %d: %d\n", (int) methodIds[0], 
                                                (int) params.cNamedArgs);
   for(int p = params.cNamedArgs; p > 0; p--)
     errorLog("%d) id %d, type %d\n", p, 
	                             (int) params.rgdispidNamedArgs[p-1],
                                     (int) V_VT(&(params.rgvarg[p-1])));
 }
#endif

 hr = disp->Invoke(methodIds[0], IID_NULL, LOCALE_USER_DEFAULT, callType, &params, res, &exceptionInfo, &nargErr);
 if(FAILED(hr)) {
   if(hr == DISP_E_MEMBERNOTFOUND) {
     errorLog("Error because member not found %d\n", nargErr);
   }

#ifdef RDCOM_VERBOSE
   errorLog("Error (%d): <in argument %d>, call type = %d, call = \n",  
	   (int) hr, (int)nargErr, (int) callType, pmname);
#endif

    clearVariants(&params);
    freeSysStrings(comNames, numNames);

    if(checkErrorInfo(disp, hr, NULL) != S_OK) {
 fprintf(stderr, "checkErrorInfo %d\n", (int) hr);fflush(stderr);
      COMError(hr);
    }
 }

 if(res) {
   ans = R_convertDCOMObjectToR(&varResult);
   VariantClear(&varResult);
 }
 clearVariants(&params);
 freeSysStrings(comNames, numNames);

#ifdef ANNOUNCE_COM_CALLS
 fprintf(stderr, "</COM>\n", (int) callType);fflush(stderr);
#endif

 return(ans);
}

__declspec(dllexport)
SEXP
R_getProperty(SEXP obj, SEXP propertyName, SEXP args, SEXP ids)
{
 return(R_COM_Invoke(obj, propertyName, args, DISPATCH_PROPERTYGET, 1, ids));
}

__declspec(dllexport)
SEXP
R_setProperty(SEXP obj, SEXP propertyName, SEXP value, SEXP ids)
{
 return(R_COM_Invoke(obj, propertyName, value, DISPATCH_PROPERTYPUT, 0, ids));
}

SEXP getArray(SAFEARRAY *arr, int dimNo, int numDims, long *indices);


HRESULT
R_getCOMArgs(SEXP args, DISPPARAMS *parms, DISPID *ids, int numNamedArgs, int *namedArgPositions)
{
 int numArgs = Rf_length(args), i, ctr;
 if(numArgs == 0)
   return(S_OK);

#ifdef RDCOM_VERBOSE
 errorLog("Converting arguments (# %d, # %d named)\n", numArgs, numNamedArgs);
#endif


 parms->rgvarg = (VARIANT *) S_alloc(numArgs, sizeof(VARIANT));
 parms->cArgs = numArgs;

 /* If there are named arguments, then put these at the beginning of the
    rgvarg*/
 if(numNamedArgs > 0) {
   int namedArgCtr = 0;
   VARIANT *var;
   SEXP el;
   SEXP names = GET_NAMES(args);

   parms->rgdispidNamedArgs = (DISPID *) S_alloc(numNamedArgs, sizeof(DISPID));
   parms->cNamedArgs = numNamedArgs;

   for(i = 0, ctr = numArgs-1; i < numArgs ; i++) {
     if(strcmp(CHAR(STRING_ELT(names, i)), "") != 0) {
       var = &(parms->rgvarg[namedArgCtr]);
       parms->rgdispidNamedArgs[namedArgCtr] = ids[namedArgCtr + 1];
#ifdef RDCOM_VERBOSE
       errorLog("Putting named argument %d into %d\n", i+1, namedArgCtr);
       Rf_PrintValue(VECTOR_ELT(args, i));
#endif
       namedArgCtr++;
     } else {
       var = &(parms->rgvarg[ctr]);
       ctr--;       
     }
     el = VECTOR_ELT(args, i);
     VariantInit(var);
     R_convertRObjectToDCOM(el, var);
   }
 } else {

   parms->cNamedArgs = 0;
   parms->rgdispidNamedArgs = NULL;

   for(i = 0, ctr = numArgs-1; i < numArgs; i++, ctr--) {
     SEXP el = VECTOR_ELT(args, i);
     VariantInit(&parms->rgvarg[ctr]);
     R_convertRObjectToDCOM(el, &(parms->rgvarg[ctr]));
   }
 }

 return(S_OK);
}

SEXP
R_isValidCOMObject(SEXP obj)
{
  SEXP el = GET_SLOT(obj, Rf_install("ref"));
  void *ptr = NULL;

  if(TYPEOF(el) != EXTPTRSXP || el == R_NilValue) 
      return(ScalarLogical(FALSE));

   ptr = R_ExternalPtrAddr(el);
   return(ScalarLogical(ptr != NULL));
}

