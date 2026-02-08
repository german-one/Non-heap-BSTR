#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include "non_heap_bstr.h"

#define MAKE_WIDE_CAT(L_, str_) L_##str_
#define MAKE_WIDE(str_) MAKE_WIDE_CAT(L, str_)

#define STR "1234567890"
#define WSTR MAKE_WIDE(STR)
#define UUID_PATTERN L"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}" // format of the string created by StringFromGUID2()

// *** For the sake of clarity, the test code below does not include error checking. ***

int main(void)
{
  CoInitialize(NULL);

  // *** use the MAKE_INITIALIZED_BSTR macro ***

  // SysStringLen() is an example for a function that has a BSTR parameter (in
  // contrast to LPBSTR or BSTR*).
  MAKE_INITIALIZED_BSTR(bstrNum, ARRAYSIZE(WSTR), WSTR);
  printf_s("%-6s %p: %2u, L\"%S\"\n\n", "init", (void *)bstrNum, SysStringLen(bstrNum), bstrNum);

  // *** use the MAKE_BSTR and SET_BSTR_BUF_LEN macros ***

  GUID uuid;
  CoCreateGuid(&uuid);
  MAKE_BSTR(bstrUuid, ARRAYSIZE(UUID_PATTERN) /* 38 + 1 (terminating NUL) */); // uninizialized
#ifdef __cplusplus // just in case you want to test how it compiles as C++ code
  StringFromGUID2(uuid, bstrUuid, ARRAYSIZE(UUID_PATTERN)); // fill string buffer
#else
  StringFromGUID2(&uuid, bstrUuid, ARRAYSIZE(UUID_PATTERN)); // fill string buffer
#endif
  SET_BSTR_LEN(bstrUuid, ARRAYSIZE(UUID_PATTERN) - 1); // define string length
  printf_s("%-6s %p: %2u, L\"%S\"\n\n", "raw", (void *)bstrUuid, SysStringLen(bstrUuid), bstrUuid);

  // *** use the BSTR buffers to create a system-allocated BSTR ***
  // *** use macro GET_BSTR_LEN to show that it works like SysStringLen() ***

  // VarBstrCat() demonstrates nicely that the two BSTR parameters are not
  // changed, while the pointer referenced by the BSTR* parameter is updated
  // (newly allocated in this case).
  BSTR concat;
  VarBstrCat(bstrNum, bstrUuid, &concat);
  printf_s("%-6s %p: %2u, L\"%S\"\n\n", "concat", (void *)concat, GET_BSTR_LEN(concat), concat);
  SysFreeString(concat);

  // *** use the MAKE_INITIALIZED_BSTR_BYTE and GET_BSTR_BYTE_LEN macros ***

  MAKE_INITIALIZED_BSTR_BYTE(bstrByte, ARRAYSIZE(STR), STR);
  printf_s("%-6s %p: %2u, \"%s\"\n\n", "bytes", (void *)bstrByte, GET_BSTR_BYTE_LEN(bstrByte), (char *)bstrByte);

  // *** use the SET_BSTR_BYTE_LEN macro ***

  ((char *)bstrByte)[5] = 0;
  SET_BSTR_BYTE_LEN(bstrByte, 5);
  printf_s("%-6s %p: %2u, \"%s\"\n\n", "update", (void *)bstrByte, SysStringByteLen(bstrByte), (char *)bstrByte);

  CoUninitialize();
  return 0;
}
