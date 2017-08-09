#include "RCOMObject.h"
#include <iostream>

#undef ERROR

extern "C" {
#include <Rdefines.h>
}

bool isCOMError(SEXP obj);
bool isClass(SEXP obj, char *name);
HRESULT processCOMError(SEXP obj, EXCEPINFO *excep, UINT *argNum);
static SEXP callQueryInterfaceMethod(SEXP obj, char *guid);


HRESULT __stdcall 
RCOMObject::QueryInterface(const IID& iid, void** ppv)
{
  HRESULT hr;

  LPOLESTR ostr; 
  char *gname;
  StringFromCLSID(iid, &ostr);
  gname = FromBstr(ostr);

#ifdef RDCOM_VERBOSE
  errorLog("[RCOMObject::QueryInterface] %s\n", gname);
#endif

#if 0
      hr = this->unknown->QueryInterface(iid, ppv);
      if(hr == S_OK) {
	SysFreeString(ostr);
	return(hr);
      }
#endif


      if(iid == IID_IUnknown)  {
	*ppv = static_cast<IUnknown*>(this);
	hr = S_OK;
      } else if(iid == IID_IDispatch) {
	*ppv = static_cast<IDispatch*>(this);
	hr = S_OK;
      } else {
#if 1
        /* Call the generic R function to see if it wants to respond.
           It can query TRUE or FALSE or give an external pointer.
           If it returns TRUE, we return this object. If it returns FALSE,
           we return an error. And if it returns an external pointer, we
           return that. It is up to the function to get that right!!!!
         */
	SEXP ans = callQueryInterfaceMethod(this->obj, gname);
	if(TYPEOF(ans) == LGLSXP) {
	  hr = LOGICAL(ans)[0] ? S_OK : E_NOINTERFACE;
	  if(hr == S_OK)
	    *ppv = static_cast<IDispatch*>(this);
          else
	    *ppv = NULL;
	} else if(TYPEOF(ans) == EXTPTRSXP) {
	  /* If it is an external ptr. */
          *ppv = R_ExternalPtrAddr(ans);
	  hr = S_OK;
        } else {
	  *ppv = NULL;
 	  hr = E_NOINTERFACE;
#ifdef RDCOM_VERBOSE
	  errorLog("iid not handled in RCOMObject::QueryInterface %s\n", gname);
#endif
	}

#else
	*ppv = NULL;
	hr = E_NOINTERFACE;
	/*	*ppv = static_cast<IDispatch*>(this); */
#endif
      }

      if(*ppv)
	reinterpret_cast<IUnknown*>(*ppv)->AddRef();

#ifdef RDCOM_VERBOSE
      errorLog("[end RCOMObject::QueryInterface]\n");
#endif
	SysFreeString(ostr);

	return hr;
}


/* 
  Call the QueryInterface generic function.
*/
static SEXP
callQueryInterfaceMethod(SEXP obj, char *guid)
{
  SEXP e, ans = R_NilValue;
  int errorOccurred = 0;

  PROTECT(e = allocVector(LANGSXP, 3));
  SETCAR(e, Rf_install("QueryInterface"));
  SETCAR(CDR(e), obj);
  SETCAR(CDR(CDR(e)), mkString(guid));

  ans = R_tryEval(e, R_GlobalEnv, &errorOccurred);

  UNPROTECT(1);
  if(errorOccurred) {
    return(R_NilValue);
  }

  return(ans);
}


// Need to handle case where cNames > 1

HRESULT __stdcall
RCOMEnvironmentObject::GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT cNames, LCID local, DISPID *id)
{
  SEXP sid;
  int i, n;
  SEXP names;
  char str[90];
  HRESULT hr;

#ifdef RDCOM_VERBOSE
  errorLog("[RCOMEnvironment::GetIDsOfNames] %p\n", this->obj);
#endif

    /* Loop over the names of the elements in the function and find the index of the 
       one that matches the method. */
  names = GET_NAMES(this->obj);

  hr = lookupRName(names, name, id);

  return(hr);
  /*XXX  Now look for properties. */
}

HRESULT
RCOMObject::lookupRName(SEXP names, LPOLESTR *name, DISPID *id)
{
  char str[1000];
   memset(str, '\0', sizeof(str)/sizeof(str[0]));
   WideCharToMultiByte(CP_ACP, 0, *name, -1, str, sizeof(str)/sizeof(str[0]), NULL, NULL);
   return(lookupRName(names, str, id));
}

HRESULT
RCOMObject::lookupRName(SEXP names, const char * const str, DISPID *id)
{
  int i, n;

  n = Rf_length(names);

#ifdef RDCOM_VERBOSE
  errorLog("Looking for method '%s'\n", str);
#endif

  for(i = 0; i < n; i++) {
    if(strcmp(str, CHAR(STRING_ELT(names, i))) == 0) {
      *id = i;
      break;
    }
  }

#ifdef RDCOM_VERBOSE
  if(i == n) {
    errorLog("Couldn't find method %s\n", str);
  } else {
    errorLog("Method id for %s = %d\n", str,  *id);
  }
#endif

  return(i < n ? S_OK : S_FALSE);
}

HRESULT  __stdcall
RCOMEnvironmentObject::Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				      VARIANT *var, EXCEPINFO *excep, UINT *argNumErr)
{
 SEXP func;

#ifdef RDCOM_VERBOSE
 errorLog("Method id %d, method = %s",  id, method);
#endif
 func = VECTOR_ELT(this->obj, id);

 return(callRFunc(func, id, refId, locale, method, parms, var, excep, argNumErr));
}

int
RCOMObject::getCallLength(DISPPARAMS *parms)
{
 return(parms->cArgs + 1);
}

SEXP
RCOMObject::getCallExpression(SEXP func, DISPPARAMS *parms, DISPID id, WORD method, SEXP *ptr)
{
 SEXP e;
 PROTECT(e = allocVector(LANGSXP, getCallLength(parms)));
 SETCAR(e, func);
 *ptr = CDR(e);
 UNPROTECT(1);
 return(e);
}

HRESULT __stdcall 
RCOMObject::callRFunc(SEXP func, DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
		      VARIANT *var, EXCEPINFO *excep, UINT *argNumErr)
{
 SEXP e, ptr, val;
 int errorOccurred, i, nargs;
 HRESULT hr = S_OK;

 PROTECT(e = getCallExpression(func, parms, id, method, &ptr));
 nargs = parms->cArgs;

 for(i = 0; i < nargs; i++) {
   SETCAR(ptr, convertToR(parms->rgvarg[i]));
   ptr = CDR(ptr);
 }

 val = R_tryEval(e, R_GlobalEnv, &errorOccurred);
 if(errorOccurred) {
   // Fill in excep 
   UNPROTECT(1);
   return(S_FALSE);
 }
 PROTECT(val);

 if(var)
   hr = convertToCOM(val, var);

 if(FAILED(hr)) {
   //XXX Fill in excep.
   return(S_FALSE);
 }

 UNPROTECT(2);

 return(S_OK);
}


HRESULT __stdcall
RCOMFunctionsObject::GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT huh, LCID locale, DISPID *id)
{
  int which;
  which = lookupRName(this->obj, name, id);
  if(which < 0) {
    return(S_FALSE);
  }

  *id = which;
  
  return(S_OK);
}


HRESULT __stdcall 
RCOMFunctionsObject::Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr)
{
 const char *funName;
 funName = CHAR(STRING_ELT(this->obj, id));

 return(callRFunc(Rf_install(funName), id, refId, locale, method, parms, var, excep, argNumErr));
}


/* Add an extra one. */
int
RCOMS4Object::getCallLength(DISPPARAMS *params)
{
  return(RCOMObject::getCallLength(params) + 1);
}

SEXP
RCOMS4Object::getCallExpression(SEXP func, DISPPARAMS *parms, DISPID id, WORD method, SEXP *ptr)
{
  SEXP e;
  PROTECT(e = RCOMFunctionsObject::getCallExpression(func, parms, id, method, ptr));
  SETCAR(*ptr, Sthis);
  *ptr = CDR(*ptr);
  UNPROTECT(1);
  return(e);
}


SEXP
asRStringVector(LPOLESTR *name, UINT cNames)
{
  SEXP tmp;
  int i;
  char str[1000];
  PROTECT(tmp = allocVector(STRSXP, cNames));
  for(i = 0; i < cNames; i++) {
    memset(str, '\0', sizeof(str)/sizeof(str[0]));
    WideCharToMultiByte(CP_ACP, 0, name[i], -1, str, sizeof(str)/sizeof(str[0]), NULL, NULL);
    SET_STRING_ELT(tmp, i, COPY_TO_USER_STRING(str));
  }
  UNPROTECT(1);
  return(tmp);
}

HRESULT __stdcall
RCOMSObject::GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT cNames, LCID locale, DISPID *id)
{
  SEXP e, val, tmp;
  int errorOccurred, i;

  PROTECT(e = allocVector(LANGSXP, 2));
  SETCAR(e, VECTOR_ELT(this->obj, IDS_OF_NAMES));
  SETCAR(CDR(e), asRStringVector(name, cNames));

  PROTECT(val = R_tryEval(e, R_GlobalEnv, &errorOccurred));
  if(!errorOccurred && val != R_NilValue) {
    //XXX Must be an integer. Need to coerce
    for(i = 0; i < cNames; i++) {
      id[i] = INTEGER(val)[i];
    }
  }
  UNPROTECT(2);

  return(errorOccurred ? S_FALSE : S_OK);
}

#if 0
/*
 Call the invoke function for this instance and pass it 
 the method id and the arguments from COM.
*/
SEXP 
RCOMSObject::getCallExpression(SEXP func, DISPPARAMS *parms, DISPID id, WORD method, SEXP *ptr)
{
 SEXP e, tmp;
 PROTECT(e = allocVector(LANGSXP, getCallLength(parms)));
 SETCAR(e, func);
 SETCAR(CDR(e), tmp = allocVector(INTSXP, 1));
 INTEGER(tmp)[0] = id;
 SETCAR(CDR(CDR(e)), tmp = allocVector(INTSXP, 1));
 INTEGER(tmp)[0] = method;
 *ptr = CDR(CDR(CDR(e)));
 UNPROTECT(1);
 return(e);
}
#endif

HRESULT __stdcall 
RCOMSObject::Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr)
{
  int errorOccurred, i;
  SEXP e, ptr, val, tmp;
  HRESULT hr;

#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
  errorLog("About to call RCOMSObject::Invoke\n");
#endif

  PROTECT(e = ptr = allocVector(LANGSXP, 5));
  SETCAR(e, VECTOR_ELT(obj, INVOKE));

  ptr = CDR(e);
  SETCAR(ptr, tmp = allocVector(INTSXP, 1));
  INTEGER(tmp)[0] = id;

  ptr = CDR(ptr);
  SETCAR(ptr, tmp = allocVector(LGLSXP, 4));
    LOGICAL(tmp)[0] = (method & INVOKE_FUNC) ? TRUE : FALSE;
    LOGICAL(tmp)[1] = (method & INVOKE_PROPERTYGET) ? TRUE : FALSE;
    LOGICAL(tmp)[2] = (method & INVOKE_PROPERTYPUT) ? TRUE : FALSE;
    LOGICAL(tmp)[3] = (method & INVOKE_PROPERTYPUTREF) ? TRUE : FALSE;

  ptr = CDR(ptr);
  if(parms->cArgs > 0) {
    PROTECT(tmp = allocVector(VECSXP, parms->cArgs));
    for(i = 0 ; i < parms->cArgs; i++) {
      SET_VECTOR_ELT(tmp, i, convertToR(parms->rgvarg[i]));
    }
    SETCAR(ptr, tmp);
    UNPROTECT(1);
  } else
    SETCAR(ptr, R_NilValue);


  ptr = CDR(ptr);
  if(parms->cNamedArgs) {
    PROTECT(tmp = allocVector(INTSXP, parms->cNamedArgs));
    for(i = 0; i < parms->cNamedArgs   ; i++) {
      INTEGER(tmp)[i] = parms->rgdispidNamedArgs[i];
    }
    SETCAR(ptr, tmp);
    UNPROTECT(1);
  } else
    SETCAR(ptr, R_NilValue);
  

  val = R_tryEval(e, R_GlobalEnv, &errorOccurred);

  if(errorOccurred) {
    UNPROTECT(1);
#if defined(RDCOM_VERBOSE) && RDCOM_VERBOSE
    errorLog("<RCOMSObject::Invoke> failed\n");
#endif
    return(S_FALSE);
  }
  PROTECT(val);

  if(isCOMError(val)) {
    HRESULT status = processCOMError(val, excep, argNumErr);
    UNPROTECT(2);
    return(status);
  }
  
  hr = convertToCOM(val, var);
  UNPROTECT(2);

  return(S_OK);
}

HRESULT 
processCOMError(SEXP obj, EXCEPINFO *excep, UINT *argNum)
{
  HRESULT status;
 
  if(isClass(obj, "COMReturnValue")) {
    //XXX    status = (HRESULT) REAL(VECTOR_ELT(obj, 0))[0];
    status = DISP_E_MEMBERNOTFOUND;
  }

  return(status);
}

bool
isCOMError(SEXP obj)
{
  return(isClass(obj, "COMError"));
}

bool
isClass(SEXP obj, char *name)
{
 SEXP klass;
 klass = GET_CLASS(obj);
 for(int i = 0; i < Rf_length(klass); i++) {
   if(strcmp(name, CHAR(STRING_ELT(klass, i))) == 0)
     return(TRUE);
 }

 return(FALSE);
}


void
RCOMSObject::destroy()
{
  if(Rf_length(obj) < 3 || VECTOR_ELT(obj, DESTRUCTOR) == R_NilValue) 
    return;

  SEXP e;
  PROTECT(e =  allocVector(LANGSXP, 1));
  SETCAR(e, VECTOR_ELT(obj, DESTRUCTOR));
  R_tryEval(e, R_GlobalEnv, NULL);
  UNPROTECT(1);
}
