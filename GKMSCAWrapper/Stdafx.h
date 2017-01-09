// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#define CHECK_NULL_AND_FREE(funcFree, variable) \
	if (NULL != variable) \
	{ \
		funcFree( variable ); \
		variable = NULL; \
	}

#define CHECK_COM_NULL_AND_RELEASE(variable) \
	if (NULL != variable) \
	{ \
		variable->Release(); \
		variable = NULL; \
	}

#define ASSERT_COM_SUCCESS(hr, message) \
	if (FAILED(hr)) \
		throw gcnew System::Runtime::InteropServices::COMException(	\
			message, hr);

	// Marshals a .NET Byte[] into a binary string, represented as a char *
#define NETBYTEARRAY_2_CHARSTRING(netByteArray, charString) \
	charString = new char[netByteArray->Length]; \
	System::Runtime::InteropServices::Marshal::Copy(netByteArray,0, \
		static_cast<System::IntPtr>( \
		const_cast<char *>( \
		static_cast<const char *>(charString))), \
		netByteArray->Length);

#define CHARARRAY_2_NETBYTEARRAY(charArray, length, netByteArray) { \
	IntPtr ^binaryPointer = gcnew IntPtr(const_cast<void *>(reinterpret_cast<const void *>(charArray))); \
	netByteArray = gcnew array<Byte>(length); \
	System::Runtime::InteropServices::Marshal::Copy(*binaryPointer,netByteArray,0,length); \
	}