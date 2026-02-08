The macros in this header-only library may help to declare, initialize and
update `BSTR` containers on the stack frame or in static storage. They can
be used in both C and C++ code. The pointer to the string buffer in such
a container is `BSTR`-_like_ (because _by definition_, a `BSTR` is
allocated using COM memory allocation functions).  

This library has no copyright assigned and is placed in the Public Domain.
It is provided "as is" without any warranty or liability, including for
merchantability or fitness for a particular purpose.  
Users assume all risks, as the author disclaims all damages.  

#### Non-heap BSTR  

The official documentation describes only heap-allocated BSTRs. However,
statements such as "BSTRs are allocated using COM memory allocation
functions ..." are a little too bold after all, although I'm not in the
position to change this definition.  
It turns out that a `BSTR` with automatic or static storage duration is
typically safe to use if passed to a `BSTR` parameter of functions (as
opposed to a `BSTR*` or `LPBSTR` parameter, where it would be completely
unsuitable). `SysFreeString()`, as an exception of this rule, is not
applicable in this context anyway.  

----
This depiction shapes a simplified template of a `BSTR` container as
created by the macros. The type pattern has been derived from [Microsoft's
documentation](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr).

```c
  struct {
      // Four-byte prefix containing the string length in bytes,
      // null-terminating character not counted.
      UINT length;
      // The pointer to the first character in this buffer is the BSTR.
      // `bufcount` is at least the length of the character string,
      // plus null-terminating character.
      OLECHAR bstr[bufcount];
  } varname;
```

----
There cannot be just a single type declaration because the length of the
buffer array varies depending on the string to represent. In the
consequence, the macros specify new types, tailored to a specific buffer
size each.  
The actual implementation of the container also allows for binary content
and supports `SysAllocStringByteLen()`-like functionality.  
Furthermore, native memory alignment is taken into account, just like with
a heap-allocated `BSTR`.  
To extend the flexibility of this library, length-related operations are
wrapped into macros, too.  

#### How to use  

Include __non_heap_bstr.h__ in your code. No other file of this
repository is necessary.  
The code in __main.c__ contains a few examples that demonstrate the use of
this little API.  
What you need to keep in mind is that you must not use a non-heap `BSTR`
where the Windows COM API may change it (e.g. allocate or reallocate).
You can easily recognize this though:  
- You may use a non-heap `BSTR` where a function receives a copy of the
pointer, i.e., the parameter is declared as `BSTR`.  
- Do NOT use a non-heap `BSTR` where a function receives a reference,
i.e., the parameter is declared as `LPBSTR`, `BSTR*` or `BSTR&`.  

Do not try to deallocate a non-heap `BSTR` using `SysFreeString()`. The
whole point of the containers that are generated using this macro lib is
that they live on the stack frame or in static storage.  

PDF prints of Doxygen-generated descriptions of the relevant macros are
placed in the __doc__ folder. More detailed information, including
information about implementation details, can be found in the comments of
the __non_heap_bstr.h__ header file.  
The comments are Doxygen-compatible, so you can create your own individual
documentation at any time.  

#### Why to use  

As an example, there is no alternative way other than using COM to access
Scheduled Tasks in C or C++. Just like a file in the file system, a
Scheduled Task is specified by its path and name. These are required as
`BSTR` though (refer to `ITaskService::GetFolder()` or
`ITaskFolder::GetTask()`).  
While it's quite likely that you would have used string literals in case
of file system elements, the COM API wants you to allocate the names of a
task folder and a task using `SysAllocString()` and family, only to make a
copy of the passed strings with a length prefix. You'll instantly
deallocate them using `SysFreeString()` after they were used in the above
mentioned interface methods.  
This adds an unnecessary amount of complexity, not only by the heap
allocation/deallocation, but also by the error checking of the returned
values in your code. So, that's a typical case where non-heap `BSTR`s
would come in handy.  
Other interfaces, like the COM API for WMI, make heavily use of `BSTR`s,
too. `IWbemServices::ExecQuery()` is a nice example. The language
specifier is always the "WQL" `BSTR` anyway. But also the WQL statement
doesn't need to be heap-allocated if you want to query the same
information every time again ...  
