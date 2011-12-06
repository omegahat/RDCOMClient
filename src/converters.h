
HRESULT R_convertRObjectToDCOM(SEXP obj, VARIANT *var);
SEXP R_convertDCOMObjectToR(VARIANT *var);
char *FromBstr(BSTR str);
BSTR AsBstr(const char *str);
SEXP getArray(SAFEARRAY *arr, int dimNo, int numDims, long *indices);

extern "C" {
  void RDCOM_finalizer(SEXP s);
  SEXP R_create2DArray(SEXP obj);
  SEXP R_createVariant(SEXP type);
  SEXP R_setVariant(SEXP svar, SEXP value, SEXP type);
}

