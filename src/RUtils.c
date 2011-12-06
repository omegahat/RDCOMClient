#include <RUtils.h>
#include <R_ext/Rdynload.h>

static SEXP R_IDispatchSym, R_IUnknownSym, R_ITypeLibSym;

void
R_RDCOMClient_init(DllInfo *info)
{
  R_IUnknownSym = Rf_install("COMIUnknown");
  R_IDispatchSym = Rf_install("COMIDispatch");
  R_ITypeLibSym = Rf_install("ITypeLib");
}

void *
derefRDCOMPointer(SEXP el)
{
 void *ptr = NULL;

 if(TYPEOF(el) != EXTPTRSXP || el == R_NilValue) {
   PROBLEM "Looking at a COM object that does not have an external pointer in the ref slot"
   ERROR;
 }

#if USE_COM_SYMBOLS
 if(R_ExternalPtrTag(el) != R_IDispatchSym || R_ExternalPtrTag(el) != R_IUnknownSym) {
   PROBLEM "Unusual RCOM object since the internal tag is not one we have seen."
   WARN;
 }
#endif

 ptr = R_ExternalPtrAddr(el);

 if(!ptr) {
   PROBLEM "RDCOM Reference object is not valid (NULL). This may be due to restoring it from a previous session."
   ERROR;
 } 

 return(ptr);
}

void*
getRDCOMReference(SEXP obj)
{
 SEXP el = GET_SLOT(obj, Rf_install("ref"));
 return(derefRDCOMPointer(el));
}

void *
derefRIDispatch(SEXP obj)
{
  return(derefRDCOMPointer(GET_SLOT(obj, Rf_install("ref"))));
}


#ifdef LOCAL_NEW

SEXP RR_do_new_object(SEXP class_def)
{
    static SEXP s_virtual = NULL, s_prototype, s_className;
    SEXP e, value;
    if(!s_virtual) {
	s_virtual = Rf_install("virtual");
	s_prototype = Rf_install("prototype");
	s_className = Rf_install("className");
    }
    if(!class_def)
	error("C level NEW macro called with null class definition pointer");
    e = R_do_slot(class_def, s_virtual);
    if(asLogical(e) != 0)  { /* includes NA, TRUE, or anything other than FALSE */
	e = R_do_slot(class_def, s_className);
	error("Trying to generate an object in C from a virtual class (\"%s\")",
	      CHAR(asChar(e)));
    }
    e = R_do_slot(class_def, s_className);
    value = duplicate(R_do_slot(class_def, s_prototype));
    setAttrib(value, R_ClassSymbol, e);
    return value;
}
#define NEW(klass)  RR_do_new_object(klass)
#endif


void
callGC()
{
  SEXP e;
  PROTECT(e = allocVector(LANGSXP, 1));
  SETCAR(e, Rf_install(".gcAll"));
  eval(e, R_GlobalEnv);
  UNPROTECT(1);
}

/*
 Debugging mechanism.
*/
void
registerCOMObject(void *ref, int reg)
{
  char buf[100];
  SEXP e, tmp;
  sprintf(buf, "%p", ref);
  PROTECT(e = allocVector(LANGSXP, 3));
  SETCAR(e, Rf_install(".comRegistry"));
  SETCAR(CDR(e), mkString(buf));
  SETCAR(CDR(CDR(e)), tmp = allocVector(LGLSXP, 1));
    LOGICAL(tmp)[0] = (reg ? TRUE : FALSE);
  eval(e, R_GlobalEnv);
  UNPROTECT(1); 
}


SEXP
createCOMReferenceObject(SEXP ptr, const char *tag)
{
  SEXP e, val;
  PROTECT(e = allocVector(LANGSXP, 3));
  SETCAR(e, Rf_install("createCOMReference")); /* in RDCOMClient code. */
  SETCAR(CDR(e), ptr);
  SETCAR(CDR(CDR(e)), mkString(tag));
  val = eval(e, R_GlobalEnv);
  UNPROTECT(1);
  return(val);
}

SEXP
R_createRCOMUnknownObject(void *ref, const char *tag)
{
  SEXP obj, ans; /*, classNames, klass, sym */;

  if(!ref)
    return(R_NilValue);


#ifdef ANNOUNCE_COM_CALLS
  fprintf(stderr, "Creating %s %p\n", tag, ref);fflush(stderr);
#endif

#ifdef REGISTER_COM_OBJECTS_WITH_S
  registerCOMObject(ref, 1);
#endif

  //XXX do we need this? Probably to ensure that if we get back the same
  // value that has already been used, that we don't use it before calling the
  // finalizer.  R_gc() doesn't seem to do it, so we may need more.
  // 
  // This is not in fact necessary. Left here as a reminder.
  // callGC();

  PROTECT(ans = R_MakeExternalPtr(ref, Rf_install(tag), R_NilValue));
  //XXX  R_RegisterCFinalizer(ans, RDCOM_finalizer);
  R_RegisterCFinalizerEx(ans, RDCOM_finalizer, TRUE);
  
#if 1
  obj = createCOMReferenceObject(ans, tag);
  UNPROTECT(1);
#else
  klass = MAKE_CLASS("COMIDispatch");
  if(klass == NULL || klass == R_NilValue) {
   PROBLEM  "Can't locate S4 class definition COMIDispatch"
     ERROR;
  }

  /*XX  the call to duplicate is needed until 1.6.2 is released
        because of a bug in the NEW() mechanism in < 1.6.2! Removed now. */
  PROTECT(obj = NEW(klass));
  SET_SLOT(obj, Rf_install("ref"), ans);

  UNPROTECT(2);
#endif

  return(obj);
}

/*XXX Is this used or is it in SWinTypeLibs. */
SEXP
R_createRTypeLib(void *ref)
{
  SEXP ans, obj, klass;

  PROTECT(ans = R_MakeExternalPtr((void*) ref, R_ITypeLibSym, R_ITypeLibSym));
  R_RegisterCFinalizer(ans, R_typelib_finalizer);

  klass = MAKE_CLASS("ITypeLib");
  PROTECT(obj = duplicate(NEW(klass)));
  SET_SLOT(obj, Rf_install("ref"), ans);

  UNPROTECT(2);

  return(obj);
}



SEXP
R_scalarLogical(Rboolean v)
{
  SEXP  ans = allocVector(LGLSXP, 1);
  LOGICAL(ans)[0] = v;
  return(ans);
}


SEXP
R_scalarReal(double v)
{
  SEXP  ans = allocVector(REALSXP, 1);
  REAL(ans)[0] = v;
  return(ans);
}

SEXP
R_scalarInteger(int v)
{
  SEXP  ans = allocVector(INTSXP, 1);
  INTEGER(ans)[0] = v;
  return(ans);
}

SEXP 
R_scalarString(const char * const v)
{
  SEXP ans = allocVector(STRSXP, 1);
  PROTECT(ans);
  if(v)
    SET_STRING_ELT(ans, 0, mkChar(v));
  UNPROTECT(1);
  return(ans);
}


Rboolean
ISSInstanceOf(SEXP obj, const char *name)
{
  SEXP e, val, tmp;
  Rboolean status;

  PROTECT(e = allocVector(LANGSXP, 3));
  SETCAR(e, Rf_install("is"));
  SETCAR(CDR(e), obj);
  PROTECT(tmp = allocVector(STRSXP, 1));
  SET_STRING_ELT(tmp, 0, COPY_TO_USER_STRING(name));
  SETCAR(CDR(CDR(e)), tmp);

  val = eval(e, R_GlobalEnv);
  status = LOGICAL(val)[0];
  UNPROTECT(2);

  return(status);
}

Rboolean
ISCOMIDispatch(SEXP obj)
{
  return(ISSInstanceOf(obj, "COMIDispatch"));
}

SEXP
R_createList(int n)
{
  SEXP ans;
  PROTECT(ans = allocVector(VECSXP, n));
  return(ans);
}

void
R_letgo(SEXP s)
{
  UNPROTECT(1);
}



#if 1

SEXP
getRNilValue(void)
{
  return(R_NilValue);
}

const char *
getRString(SEXP str, int which)
{
 return(CHAR(STRING_ELT(str, which)));
}

SEXP getRNames(SEXP obj)
{
  return(GET_NAMES(obj));
}

void 
clearRDCOMObject(SEXP s)
{
 fprintf(stderr,"Clearing the RDCOM object\n");
 R_ClearExternalPtr(s);
}

int
getRLength(SEXP obj)
{
  return(Rf_length(obj));
}

SEXP 
getRListElement(SEXP o, int index)
{
  return(VECTOR_ELT(o, index));
}

void
setRListElement(SEXP o, int index, SEXP el)
{
  SET_VECTOR_ELT(o, index, el);
}


int
R_typeof(SEXP obj)
{
  return(TYPEOF(obj));
}

long
R_integerScalarValue(SEXP obj, int index)
{
  return(INTEGER(obj)[index]);
}

double
R_realScalarValue(SEXP obj, int index)
{
  return(REAL(obj)[index]);
}

int
R_logicalScalarValue(SEXP obj, int index)
{
  return((int) LOGICAL(obj)[index]);
}


#endif
