#include <Rdefines.h>
#include <stdio.h>

#ifdef __cplusplus
extern "C" {
#endif

 extern FILE *getErrorFILE();

  SEXP R_createRCOMUnknownObject(void *ref, const char *tag);
  void* getRDCOMReference(SEXP);
  SEXP R_scalarInteger(int v);
  SEXP R_scalarReal(double v);
  SEXP R_scalarLogical(Rboolean v);
  SEXP R_scalarString(const char * const v);

#if 0
 __declspec(dllexport) SEXP getRNilValue(void);
  int getRLength(SEXP obj);
#endif

 __declspec(dllexport) const char * getRString(SEXP str, int which);
  int R_typeof(SEXP obj);

  int R_logicalScalarValue(SEXP obj, int index);
  double R_realScalarValue(SEXP obj, int index);
  long R_integerScalarValue(SEXP obj, int index);



  // __declspec(dllexport) SEXP R_createRCOMUnknownObject(IUnknown *ref, IID refId);



  SEXP getRListElement(SEXP, int);


  void RDCOM_finalizer(SEXP s);
  void *derefRDCOMPointer(SEXP el);

  void clearRDCOMObject(SEXP);

  Rboolean ISSInstanceOf(SEXP obj, const char *name);

  Rboolean ISCOMIDispatch(SEXP obj);
  void * derefRIDispatch(SEXP obj);

  SEXP R_createList(int n);
  void R_letgo(SEXP s);
  void setRListElement(SEXP o, int index, SEXP el);

  SEXP getRNames(SEXP);

  void R_typelib_finalizer(SEXP obj);

#ifdef __cplusplus
}
#endif


#define errorLog(a,...) fprintf(getErrorFILE(), a, ##__VA_ARGS__); fflush(getErrorFILE());



