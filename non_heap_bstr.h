// =============================================================================
/// @important
///   This file has no copyright assigned and is placed in the Public Domain.
///   This file is provided "as is" without any warranty or liability,
///   including for merchantability or fitness for a particular purpose.
///   Users assume all risks, as the author disclaims all damages.
/// @file    non_heap_bstr.h
/// @brief   Windows `BSTR` with automatic or static storage duration.
/// @author  Steffen Illhardt
/// @date    February 2026
/// @version 1.0
/// @pre     Requires compiler support for at least C11.
/// @details
///   The macros in this header-only library may help to declare, initialize and
///   update `BSTR` containers on the stack frame or in static storage. They can
///   be used in both C and C++ code. <br>
///   The official documentation describes only heap-allocated BSTRs. However,
///   statements such as "BSTRs are allocated using COM memory allocation
///   functions ..." are a little too bold after all, although I'm not in the
///   position to change this definition. <br>
///   It turns out that a `BSTR` with automatic or static storage duration is
///   typically safe to use if passed to a `BSTR` parameter of functions (as
///   opposed to a `BSTR*` or `LPBSTR` parameter, where it would be completely
///   unsuitable). `SysFreeString()`, as an exception of this rule, is not
///   applicable in this context anyway.
///   <hr>
///   This depiction shapes a simplified template of a `BSTR` container as
///   created by the macros. The type pattern has been derived from Microsoft's
///   documentation: <br>
///   https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr
///   @code
///     struct {
///         // Four-byte prefix containing the string length in bytes,
///         // null-terminating character not counted.
///         UINT length;
///         // The pointer to the first character in this buffer is the BSTR.
///         // `bufcount` is at least the length of the character string,
///         // plus null-terminating character.
///         OLECHAR bstr[bufcount];
///     } varname;
///   @endcode
///   <hr>
///   There cannot be just a single type declaration because the length of the
///   buffer array varies depending on the string to represent. In the
///   consequence, the macros specify new types, tailored to a specific buffer
///   size each. <br>
///   The actual implementation of the container also allows for binary content
///   and supports `SysAllocStringByteLen()`-like functionality. <br>
///   Furthermore, native memory alignment is taken into account, just like with
///   a heap-allocated `BSTR`. <br>
///   To extend the flexibility of this library, length-related operations are
///   wrapped into macros, too.
// =============================================================================
#ifndef HEADER_NON_HEAP_BSTR_63E45A1A_6124_4281_9104_C3B113C2A312_1_0
#define HEADER_NON_HEAP_BSTR_63E45A1A_6124_4281_9104_C3B113C2A312_1_0
#include <windows.h>
// =============================================================================
/// @defgroup detail    Implementation Detail
///                     Memory alignment guard and generic template. Do not use.
/// @{
// -----------------------------------------------------------------------------
/// @brief Implementation detail - DO NOT USE.
/// @details A heap-allocated `BSTR` does always point to a buffer with native
///          alignment (i.e. 4 or 8 bytes in a 32-bit or 64-bit process,
///          respectively). This is emulated by the conditional definition of
///          the INTERNAL_BSTR_CONTAINER_LENGTH_PREFIX__ macro. It specifies an
///          object with native size and alignment, that also maintains the
///          position of the four-byte `length` field, which has to appear
///          adjacent to the following character array.
/// @note As the name indicates, this macro is only **internally** used as
///       length prefix of a `BSTR` container.
#if defined(_WIN64)
#  define INTERNAL_BSTR_CONTAINER_LENGTH_PREFIX__ /* 64-bit */           \
    union {                                                              \
      struct {                                                           \
        /* unused, its size defines the offset of the `length` member */ \
        __int32 margin_dummy;                                            \
        /* length of the string in bytes, null-terminator not counted */ \
        UINT length;                                                     \
      } prefix;                                                          \
      /* unused, its size defines the memory alignment */                \
      __int64 alignment_dummy;                                           \
    }
#else
#  define INTERNAL_BSTR_CONTAINER_LENGTH_PREFIX__ /* 32-bit */ \
    struct {                                                   \
      UINT length;                                             \
    } prefix
#endif
// -----------------------------------------------------------------------------
/// @brief Implementation detail - DO NOT USE.
/// @details The INTERNAL_BSTR_CONTAINER__ macro creates a container on the
///          stack frame or in static storage.
/// @remark This macro is directly or subsequently used in the creation-related
///         macros, as it contains the generic structure and components of the
///         implementation.
/// @note As the name indicates, this macro is only **internally** used.
/// @param varname_   Name of the container to be instantiated.
/// @param bytecount_ Size of the buffer, in bytes.
#define INTERNAL_BSTR_CONTAINER__(varname_, bytecount_)                                                                          \
  struct tag_##varname_ {                                                                                                        \
    /* contains the `length` member */                                                                                           \
    INTERNAL_BSTR_CONTAINER_LENGTH_PREFIX__;                                                                                     \
    union {                                                                                                                      \
      /* wide string buffer, natively aligned */                                                                                 \
      WCHAR bstr[((bytecount_) / sizeof(WCHAR) + sizeof(__int3264) / sizeof(WCHAR)) & ~(sizeof(__int3264) / sizeof(WCHAR) - 1)]; \
      /* byte-string buffer that shares its memory with `bstr`; used for the initialization with arbitrary data */               \
      char bytestr[((bytecount_) + sizeof(__int3264)) & ~(sizeof(__int3264) - 1)];                                               \
    };                                                                                                                           \
  } varname_
// -----------------------------------------------------------------------------
/// @}
// =============================================================================
/// @defgroup wcreate    BSTR Wide String Creation
///                      Create a BSTR of wide characters with automatic or
///                      static storage duration.
/// @{
// -----------------------------------------------------------------------------
/// @brief Create a `BSTR` container.
/// @details The BSTR_CONTAINER macro creates a `BSTR` container on the stack
///          frame (where it is uninitialized) or in static storage (where
///          it is zero-initialized by default).
/// @param varname_  Name of the container to be instantiated.
/// @param bufcount_ Size of the buffer, in wide characters, that must be large
///                  enough for the string to represent, including the
///                  null-terminating character.
#define BSTR_CONTAINER(varname_, bufcount_) \
  INTERNAL_BSTR_CONTAINER__(varname_, (bufcount_) * sizeof(WCHAR))
// -----------------------------------------------------------------------------
/// @brief Create an initialized `BSTR` container.
/// @details Aim of the INITIALIZED_BSTR_CONTAINER macro is both the creation
///          and the initialization of a `BSTR` container on the stack frame or
///          in static storage.
/// @param varname_  Name of the container to be instantiated.
/// @param bufcount_ Size of the represented string, in wide characters,
///                  including the null-terminating character. This might not be
///                  the length of the initializer, but must meet the total
///                  length of the string before used.
/// @param ...       Variadic expression to initialize the string buffer. <br>
///                  This can be a string literal or brace-enclosed list of
///                  characters that fill the whole buffer. <br>
///                  This can be L"" or { 0 } to zero-initialize the buffer in
///                  order to copy characters into it later. <br>
///                  This can be a substring like L"ab" or { L'a', L'b' } to
///                  which remaining characters are appended later.
#define INITIALIZED_BSTR_CONTAINER(varname_, bufcount_, /*initializer*/...) \
  BSTR_CONTAINER(varname_, bufcount_) = { .prefix = { .length = ((bufcount_) - 1) * sizeof(WCHAR) }, .bstr = __VA_ARGS__ }
// -----------------------------------------------------------------------------
/// @brief Declare a `BSTR` variable.
/// @details The MAKE_BSTR macro declares a `BSTR` variable in the current scope
///          but restricts the visibility of the container implementation to the
///          body block of a wrapping while-loop. The container object has
///          static storage duration and is therefore zero-initialized. <br>
///          For the description of the parameters, see @ref BSTR_CONTAINER().
#define MAKE_BSTR(varname_, bufcount_)                           \
  BSTR varname_;                                                 \
  do {                                                           \
    static BSTR_CONTAINER(bstr_container_##varname_, bufcount_); \
    varname_ = bstr_container_##varname_.bstr;                   \
  } while (0)
// -----------------------------------------------------------------------------
/// @brief Declare and initialize a `BSTR` variable.
/// @details The MAKE_INITIALIZED_BSTR macro declares a `BSTR` variable in the
///          current scope but restricts the visibility of the container
///          implementation to the body block of a wrapping while-loop. The
///          container object has static storage duration and the variadic
///          arguments are used to initialize it. <br>
///          For the description of the parameters, see
///          @ref INITIALIZED_BSTR_CONTAINER().
#define MAKE_INITIALIZED_BSTR(varname_, bufcount_, /*initializer*/...)                    \
  BSTR varname_;                                                                          \
  do {                                                                                    \
    static INITIALIZED_BSTR_CONTAINER(bstr_container_##varname_, bufcount_, __VA_ARGS__); \
    varname_ = bstr_container_##varname_.bstr;                                            \
  } while (0)
// -----------------------------------------------------------------------------
/// @}
// =============================================================================
/// @defgroup bcreate    BSTR Byte String Creation
///                      Create a BSTR for binary data with automatic or static
///                      storage duration.
/// @{
// -----------------------------------------------------------------------------
/// @brief Create a `BSTR` container for binary data.
/// @details The BSTR_BYTE_CONTAINER macro creates a `BSTR` container on the
///          stack frame (where it is uninitialized) or in static storage (where
///          it is zero-initialized by default).
/// @param varname_ Name of the container to be instantiated.
/// @param bufsize_ Size of the buffer, in bytes, that must be large enough for
///                 the data to represent, including the null-terminating
///                 character.
#define BSTR_BYTE_CONTAINER(varname_, bufsize_) \
  INTERNAL_BSTR_CONTAINER__(varname_, bufsize_)
// -----------------------------------------------------------------------------
/// @brief Create an initialized `BSTR` container for binary data.
/// @details Aim of the INITIALIZED_BSTR_BYTE_CONTAINER macro is both the
///          creation and the initialization of a `BSTR` container on the stack
///          frame or in static storage.
/// @param varname_ Name of the container to be instantiated.
/// @param bufsize_ Size of the represented data, in bytes, including the
///                 null-terminating character. This might not be the length of
///                 the initializer, but must meet the total length of the data
///                 before used.
/// @param ...      Variadic expression to initialize the string buffer. <br>
///                 This can be a string literal or brace-enclosed list of
///                 characters that fill the whole buffer. <br>
///                 This can be "" or { 0 } to zero-initialize the buffer in
///                 order to copy characters into it later. <br>
///                 This can be a substring like "ab" or { 'a', 'b' } to which
///                 remaining bytes are appended later.
#define INITIALIZED_BSTR_BYTE_CONTAINER(varname_, bufsize_, /*initializer*/...) \
  BSTR_BYTE_CONTAINER(varname_, bufsize_) = { .prefix = { .length = (bufsize_) - 1 }, .bytestr = __VA_ARGS__ }
// -----------------------------------------------------------------------------
/// @brief Declare a `BSTR` variable containing binary data.
/// @details The MAKE_BSTR_BYTE macro declares a `BSTR` variable in the current
///          scope but restricts the visibility of the container implementation
///          to the body block of a wrapping while-loop. The container object
///          has static storage duration and is therefore zero-initialized. <br>
///          For the description of the parameters, see
///          @ref BSTR_BYTE_CONTAINER().
#define MAKE_BSTR_BYTE(varname_, bufsize_)                           \
  BSTR varname_;                                                     \
  do {                                                               \
    static BSTR_BYTE_CONTAINER(bstr_container_##varname_, bufsize_); \
    varname_ = bstr_container_##varname_.bstr;                       \
  } while (0)
// -----------------------------------------------------------------------------
/// @brief Declare and initialize a `BSTR` variable containing binary data.
/// @details The MAKE_INITIALIZED_BSTR_BYTE macro declares a `BSTR` variable in
///          the current scope but restricts the visibility of the container
///          implementation to the body block of a wrapping while-loop. The
///          container object has static storage duration and the variadic
///          arguments are used to initialize it. <br>
///          For the description of the parameters, see
///          @ref INITIALIZED_BSTR_BYTE_CONTAINER().
#define MAKE_INITIALIZED_BSTR_BYTE(varname_, bufsize_, /*initializer*/...)                    \
  BSTR varname_;                                                                              \
  do {                                                                                        \
    static INITIALIZED_BSTR_BYTE_CONTAINER(bstr_container_##varname_, bufsize_, __VA_ARGS__); \
    varname_ = bstr_container_##varname_.bstr;                                                \
  } while (0)
// -----------------------------------------------------------------------------
/// @}
// =============================================================================
/// @defgroup wlength    BSTR Wide String Length
///                      Get or set the length of a BSTR.
/// @{
// -----------------------------------------------------------------------------
/// @brief Retrieve the length of a `BSTR` containing wide characters.
/// @details This is just a simple macro alternative for SysStringLen() to get
///          the length of a `BSTR` as number of wide characters. The
///          null-terminating character is not counted.
/// @param bstr_ Non-NULL `BSTR`.
#define GET_BSTR_LEN(bstr_) \
  ((UINT)(((UINT *)(void *)(bstr_))[-1] / sizeof(WCHAR)))
// -----------------------------------------------------------------------------
/// @brief Update the length of a `BSTR` containing wide characters.
/// @details This is necessary for uninitialized or default-initialized
///          containers as soon as the content of the string buffer was updated,
///          also if a `BSTR` is reused with new content of a different length.
/// @note No matter if the buffer of the updated string was heap-allocated or
///       not, ensure that the memory boundaries were not violated, the
///       null-terminating character was appended properly and the length to set
///       meets the actual length of the represented data. An update of the
///       length prefix using this macro does not change the size of the
///       allocated memory space.
/// @param bstr_   Non-NULL `BSTR`.
/// @param length_ Length of the represented string, in wide characters. The
///                null-terminating character is not counted.
#define SET_BSTR_LEN(bstr_, length_) \
  ((UINT *)(void *)(bstr_))[-1] = (UINT)((length_) * sizeof(WCHAR))
// -----------------------------------------------------------------------------
/// @}
// =============================================================================
/// @defgroup blength    BSTR Byte String Length
///                      Get or set the byte length of a BSTR.
/// @{
// -----------------------------------------------------------------------------
/// @brief Retrieve the length of a `BSTR` containing binary data.
/// @details This is just a simple macro alternative for SysStringByteLen() to
///          get the length of a `BSTR` as number of bytes. The null-terminating
///          character is not counted.
/// @param bstr_ Non-NULL `BSTR`.
#define GET_BSTR_BYTE_LEN(bstr_) \
  (((UINT *)(void *)(bstr_))[-1])
// -----------------------------------------------------------------------------
/// @brief Update the length of a `BSTR` containing binary data.
/// @details This is necessary for uninitialized or default-initialized
///          containers as soon as the content of the string buffer was updated,
///          also if a `BSTR` is reused with new content of a different length.
/// @note No matter if the buffer of the updated string was heap-allocated or
///       not, ensure that the memory boundaries were not violated, the
///       null-terminating character was appended properly and the length to set
///       meets the actual length of the represented data. An update of the
///       length prefix using this macro does not change the size of the
///       allocated memory space.
/// @param bstr_   Non-NULL `BSTR`.
/// @param length_ Length of the represented data, in bytes. The
///                null-terminating character is not counted.
#define SET_BSTR_BYTE_LEN(bstr_, length_) \
  ((UINT *)(void *)(bstr_))[-1] = (UINT)(length_)
// -----------------------------------------------------------------------------
/// @}
// =============================================================================
#endif /* header guard */
