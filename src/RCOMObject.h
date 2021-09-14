/*

 There are several classes defined here to represent a COM server object 
 within R.  The class hierarchy is

                                RCOMObject
 
  RCOMEnvironmentObject     RCOMFunctionsObject      RCOMSObject

                        RCOMOOPObject  RCOMS4Object

  
  RCOMObject

  RCOMEnvironmentObject
  RCOMFunctionsObject
  RCOMSObject
   
  RCOMOOPObject
  RCOMS4Object

*/
#include <windows.h>
#include <objbase.h>
#include <oaidl.h>

/* 
*/
#define RDCOM_VERBOSE 1 



#undef ERROR

extern "C" {
#include <Rinternals.h>
  //#include <Defn.h>
#include <Rdefines.h>
  extern void R_PreserveObject(SEXP);
  extern void R_ReleaseObject(SEXP);
}

#ifdef length
#undef length
#endif

#include <iostream>

#include "converters.h"

#include "RUtils.h" /* For errorLog() */


class RCOMObject : public IDispatch
{
  public: 
    // IUnknown
    virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
    virtual ULONG __stdcall AddRef() {return InterlockedIncrement(&m_cRef);}
    virtual ULONG __stdcall Release() { if(InterlockedDecrement(&m_cRef) == 0) {
                                          delete this;
					  return 0;
                                        }
                                        return m_cRef;
                                      }


    //IDispatch
  virtual HRESULT __stdcall GetTypeInfoCount(UINT *n) { *n = 0; return S_OK; };
  virtual HRESULT __stdcall GetTypeInfo(UINT i, LCID locale, LPTYPEINFO *typeInfo) {
                          if(typeInfo)
                            *typeInfo = NULL;
			  return E_NOTIMPL;
                        }

   /* One needs to implement these. */

  virtual HRESULT __stdcall GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT huh, LCID locale, DISPID *id) = 0;
  virtual HRESULT __stdcall Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr) = 0;

  RCOMObject() : m_cRef(1) { 
                             setObject(R_NilValue); 
                           }

  RCOMObject(SEXP def) : m_cRef(1) {
                          if(def != R_NilValue)
                             setObject(def);
                       };
  virtual ~RCOMObject() {
                  m_cRef--;
                  if(m_cRef == 0) {
		    destroy();
		    R_ReleaseObject(obj);
                  }
                };

 SEXP R_getObject() { return(obj);}; 
 SEXP R_setObject(SEXP def) { SEXP old = obj; 
                              obj = def; 
                              R_PreserveObject(obj); 
                              /* R_ReleaseObject(old); */
                              return(old); 
                            }; 


  virtual const char *getClassName() {
    return("RCOMObject");
  };

 protected:
   // 
  SEXP convertToR(VARIANT &var) { return(R_convertDCOMObjectToR(&var));}
  HRESULT convertToCOM(SEXP obj, VARIANT *var) { return(R_convertRObjectToDCOM(obj, var));}

  void setObject(SEXP def) {
    obj = def;
    if(obj && obj != R_NilValue) {
      R_PreserveObject(obj);
    }
  }

  /* Return the length of the call expression */
  virtual int  getCallLength(DISPPARAMS *parms);

  /* Create the call expression and populate it with the initial element(s),
     i.e. function or function name, or that and the S-level 'this' object. */
  virtual SEXP getCallExpression(SEXP func, DISPPARAMS *parms, DISPID id, WORD method, SEXP *ptr);

  HRESULT lookupRName(SEXP names, const char * const str, DISPID *id);
  HRESULT lookupRName(SEXP names, LPOLESTR *name, DISPID *id);

  HRESULT __stdcall callRFunc(SEXP func, DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
			    VARIANT *var, EXCEPINFO *excep, UINT *argNumErr);

  virtual void destroy() {};


 protected:
  SEXP obj;

 private:
   long m_cRef;
};


class RCOMEnvironmentObject : public RCOMObject
{
 public:
  HRESULT __stdcall GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT huh, LCID locale, DISPID *id);
  HRESULT __stdcall Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr);


  virtual const char *getClassName() {
    return("RCOMEnvironmentObject");
  };

  RCOMEnvironmentObject(SEXP def) {
    SEXP f, e, val;
    int errorOccurred;
#ifdef RDCOM_VERBOSE
    errorLog("[RCOMEnvironment]\n");
#endif
    f = GET_SLOT(def, Rf_install("generator"));
    PROTECT(e = allocVector(LANGSXP, 1));
    SETCAR(e, f);

    val = R_tryEval(e, R_GlobalEnv, &errorOccurred);

    setObject(val);
    UNPROTECT(1);
#ifdef RDCOM_VERBOSE
    errorLog("[end RCOMEnvironment]\n");
#endif
  };
};



/**
  Collection of names identifying top-level functions.
*/
class RCOMFunctionsObject : public RCOMObject
{
 public:

  virtual const char *getClassName() {
    return("RCOMFunctionsObject");
  };

  HRESULT __stdcall GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT huh, LCID locale, DISPID *id);
  HRESULT __stdcall Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr);

  RCOMFunctionsObject(SEXP def) : RCOMObject(def) { };
};


class RCOMOOPObject : public RCOMFunctionsObject
{

 public:

  virtual const char *getClassName() {
    return("RCOMOOPObject");
  };


   RCOMOOPObject(SEXP def) : RCOMFunctionsObject(def) {
     /* Now call the new method in the OOP object. */
     //XXX
   }
};


/**
  A named collection of top-level methods/functions
  but with an additional object that is passed as the 
  first argument for each function call, acting
  like an immutable this object.

 */
class RCOMS4Object : public RCOMFunctionsObject
{
 
 public:

  virtual const char *getClassName() {
    return("RCOMS4Object");
  };

  RCOMS4Object(SEXP def, SEXP sobj) : RCOMFunctionsObject(def) {
    R_PreserveObject(sobj);
    Sthis = sobj;
  }

  /* Adds the Sthis as the first argument to the function. */
  HRESULT __stdcall Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr);


 protected:
  SEXP Sthis; // S level object passed as first argument.  

  /* Add an extra argument to accomodate the Sthis object as the first argument. */
  virtual int  getCallLength(DISPPARAMS *parms);
  /* Insert the Sthis as the second argument and advance the ptr to the 3rd slot
     for the additional arguments. */
  virtual SEXP getCallExpression(SEXP func, DISPPARAMS *parms, DISPID id, WORD method, SEXP *ptr);  
};



/**
 An implementation of the Invoke and GetIdsOfNames 
 via S-level functions, not internal C++-level methods.
 This allows one to control the exact way the invocation
 of methods is performed.
 These two methods are provided as an S list of functions 
 in the order 
     invoke, getIdsOfNames
*/
class RCOMSObject : public RCOMObject
{

  enum {INVOKE = 0, IDS_OF_NAMES, DESTRUCTOR};

 public:


  virtual const char *getClassName() {
    return("RCOMSObject");
  };

  RCOMSObject(SEXP def) : RCOMObject(def) {}

  virtual HRESULT __stdcall GetIDsOfNames(REFIID refId, LPOLESTR *name, UINT cNames, LCID locale, DISPID *id);

  HRESULT __stdcall Invoke(DISPID id, REFIID refId, LCID locale, WORD method, DISPPARAMS *parms, 
				     VARIANT *var, EXCEPINFO *excep, UINT *argNumErr);

 protected:
  void destroy();

#if 0
  int getCallLength(DISPPARAMS *parms) {
                                         return RCOMObject::getCallLength(parms) + 2;
                                       };

  SEXP getCallExpression(SEXP func, DISPPARAMS *parms, DISPID id, WORD method, SEXP *ptr);  
#endif
};
